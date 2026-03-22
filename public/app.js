const state = {
  brandFile: null,
  brandB64: null,
  contentFile: null,
  contentB64: null,
  slideCount: 12
}

function handleUpload(type, input) {
  const file = input.files[0]
  if (!file) return
  const reader = new FileReader()
  reader.onload = (e) => {
    const b64 = e.target.result.split(',')[1]
    if (type === 'brand') {
      state.brandFile = file
      state.brandB64 = b64
      document.getElementById('label-brand').innerHTML = '✓ <span>' + file.name + '</span>'
      document.getElementById('zone-brand').classList.add('ready')
    } else {
      state.contentFile = file
      state.contentB64 = b64
      document.getElementById('label-content').innerHTML = '✓ <span>' + file.name + '</span>'
      document.getElementById('zone-content').classList.add('ready')
    }
    checkReady()
  }
  reader.readAsDataURL(file)
}

function updateSlides(val) {
  state.slideCount = parseInt(val)
  document.getElementById('slide-count').textContent = val
}

function checkReady() {
  const btn = document.getElementById('run-btn')
  btn.disabled = !(state.brandB64 && state.contentB64)
}

function setStep(num, status) {
  const el = document.getElementById('s' + num)
  const labels = [
    '', 
    'Agent 1 — Collecting inputs',
    'Agent 2 — Parsing brand guidelines',
    'Agent 3 — Analysing content structure',
    'Agent 4 — Writing slide content',
    'Agent 5 — Applying brand to slides'
  ]
  if (status === 'done') el.textContent = '✅ ' + labels[num]
  if (status === 'active') el.textContent = '🔵 ' + labels[num]
  if (status === 'wait') el.textContent = '⬜ ' + labels[num]
  el.className = 'step ' + status
}

function setProgress(pct) {
  document.getElementById('prog').style.width = pct + '%'
}

async function callClaude(system, messages, max_tokens = 1000) {
  const res = await fetch('/api/claude', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ system, messages, max_tokens })
  })
  if (!res.ok) {
    const err = await res.json()
    throw new Error(err.error || 'API call failed')
  }
  const data = await res.json()
  return data.content.map(b => b.text || '').join('')
}

async function runPipeline() {
  const btn = document.getElementById('run-btn')
  btn.disabled = true
  document.getElementById('status-box').classList.add('show')

  try {
    // AGENT 1
    setStep(1, 'active')
    setProgress(5)
    const ext = state.brandFile.name.split('.').pop().toLowerCase()
    let brandContent = ''
    if (ext === 'pptx' || ext === 'ppt') {
      brandContent = 'Brand file: ' + state.brandFile.name + ' (PPTX - extract rules from filename context)'
    } else {
      brandContent = state.brandB64
    }
    setStep(1, 'done')
    setProgress(15)

    // AGENT 2
    setStep(2, 'active')
    const agent2System = `You are an expert brand designer. Extract brand design rules and return as JSON only with these fields: primary_colors (array of hex codes), secondary_colors (array), background_colors (array), title_font (object: family/size/weight/color), body_font (object: family/size/weight), slide_width, slide_height, logo_position, layout_patterns (array), visual_style, spacing_notes, chart_colors (array). Return ONLY valid JSON. No explanation. No markdown.`
    
    const agent2Response = await callClaude(agent2System, [{
      role: 'user',
      content: 'Extract brand rules from this brand deck content and return as JSON:\n\n' + brandContent.slice(0, 3000)
    }])
    
    let brandRulebook
    try {
      brandRulebook = JSON.parse(agent2Response.replace(/```json|```/g, '').trim())
    } catch(e) {
      brandRulebook = { visual_style: 'corporate', primary_colors: ['#1A3C6E'], title_font: { family: 'Calibri', size: '32pt', weight: 'bold' }, body_font: { family: 'Calibri', size: '18pt' } }
    }
    setStep(2, 'done')
    setProgress(35)

    // AGENT 3
    setStep(3, 'active')
    const agent3System = `You are a management consultant structuring a board-level presentation. Create a logical slide outline for senior management. Return a numbered list of slides with title and one-line description. Total slides must equal the number requested.`
    
    const agent3Response = await callClaude(agent3System, [{
      role: 'user',
      content: `Create a ${state.slideCount} slide presentation outline from this content. Include title slide, section dividers, and content slides:\n\n` + atob(state.contentB64).slice(0, 3000)
    }], 1000)
    
    setStep(3, 'done')
    setProgress(55)

    // AGENT 4
    setStep(4, 'active')
    const agent4System = `You are a business writer. Expand each slide in the outline into full content. For each slide return a JSON array where each object has: slide_number, type (title/divider/content), title, bullets (array), visual_type (text/chart/icons/table/image). Return ONLY a valid JSON array.`
    
    const agent4Response = await callClaude(agent4System, [{
      role: 'user',
      content: 'Expand this outline into full slide content as a JSON array:\n\n' + agent3Response
    }], 1000)
    
    let slideManifest
    try {
      slideManifest = JSON.parse(agent4Response.replace(/```json|```/g, '').trim())
    } catch(e) {
      slideManifest = [{ slide_number: 1, type: 'title', title: 'Presentation', bullets: [] }]
    }
    setStep(4, 'done')
    setProgress(75)

    // AGENT 5
    setStep(5, 'active')
    const agent5System = `You are a senior presentation designer. Combine the slide content and brand rules to produce a final slide specification. For each slide return a JSON array where each object has: slide_number, title, bullets, background_color, title_color, title_font, body_font, layout, visual_type. Follow the brand rules strictly. Return ONLY a valid JSON array.`
    
    const agent5Response = await callClaude(agent5System, [{
      role: 'user',
      content: `Brand rules:\n${JSON.stringify(brandRulebook)}\n\nSlide content:\n${JSON.stringify(slideManifest)}\n\nProduce final slide spec as JSON array.`
    }], 1000)
    
    let finalSpec
    try {
      finalSpec = JSON.parse(agent5Response.replace(/```json|```/g, '').trim())
    } catch(e) {
      finalSpec = slideManifest
    }
    setStep(5, 'done')
    setProgress(100)

    // Store for Agent 6
    localStorage.setItem('finalSpec', JSON.stringify(finalSpec))
    localStorage.setItem('brandRulebook', JSON.stringify(brandRulebook))

    setTimeout(() => {
      alert('✅ All 5 agents complete! Final spec ready for Agent 6 (PPTX generation).\n\nOpen browser console to see the output.')
      console.log('FINAL SPEC:', finalSpec)
      console.log('BRAND RULEBOOK:', brandRulebook)
    }, 500)

  } catch(err) {
    alert('Error: ' + err.message)
    btn.disabled = false
  }
}