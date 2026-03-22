// ─── STATE ───────────────────────────────────────────────────────────────────
const state = {
  brandFile: null,
  brandB64:  null,
  brandExt:  null,
  contentB64: null,
  slideCount: 12,
  brandRulebook: null,
  slideManifest: null,
  finalSpec: null
}

// ─── UI HELPERS ──────────────────────────────────────────────────────────────
function $(id) { return document.getElementById(id) }

function handleUpload(type, input) {
  const file = input.files[0]
  if (!file) return
  const ext = file.name.split('.').pop().toLowerCase()
  const reader = new FileReader()
  reader.onload = (e) => {
    const b64 = e.target.result.split(',')[1]
    if (type === 'brand') {
      state.brandFile = file
      state.brandB64  = b64
      state.brandExt  = ext
      $('label-brand').innerHTML = '✓ <span>' + file.name + '</span>'
      $('zone-brand').classList.add('ready')
    } else {
      state.contentB64 = b64
      $('label-content').innerHTML = '✓ <span>' + file.name + '</span>'
      $('zone-content').classList.add('ready')
    }
    checkReady()
  }
  reader.readAsDataURL(file)
}

function updateSlides(val) {
  state.slideCount = parseInt(val)
  $('slide-count').textContent = val
}

function checkReady() {
  $('run-btn').disabled = !(state.brandB64 && state.contentB64)
}

function setStep(num, status) {
  const labels = [
    '',
    'Agent 1 — Collecting & packaging inputs',
    'Agent 2 — Parsing brand guidelines',
    'Agent 3 — Analysing content & building structure',
    'Agent 4 — Writing detailed slide content',
    'Agent 5 — Merging brand + content into final spec'
  ]
  const el = $('s' + num)
  const icons = { done: '✅', active: '🔵', wait: '⬜', error: '❌' }
  el.textContent = (icons[status] || '⬜') + ' ' + labels[num]
  el.className = 'step ' + status
}

function setProgress(pct) {
  $('prog').style.width = pct + '%'
}

function showStatus() {
  $('status-box').classList.add('show')
}

function showError(msg) {
  $('error-msg').textContent = msg
  $('error-box').classList.add('show')
}

function hideError() {
  $('error-box').classList.remove('show')
}

// ─── PPTX TEXT EXTRACTOR (runs in browser via JSZip) ─────────────────────────
async function extractPptxText(b64) {
  // JSZip is loaded via CDN in index.html
  const binary = atob(b64)
  const bytes  = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)

  const zip = await JSZip.loadAsync(bytes.buffer)

  const slideFiles = Object.keys(zip.files)
    .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => {
      const na = parseInt(a.match(/\d+/)[0])
      const nb = parseInt(b.match(/\d+/)[0])
      return na - nb
    })

  let extracted = ''

  for (const sf of slideFiles.slice(0, 12)) {
    const xml   = await zip.files[sf].async('string')
    const texts = [...xml.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)]
      .map(m => m[1].trim()).filter(Boolean)
    if (texts.length) extracted += '\n[' + sf + ']\n' + texts.join('\n')
  }

  // Theme colors
  const themeFile = zip.files['ppt/theme/theme1.xml']
  if (themeFile) {
    const themeXml = await themeFile.async('string')
    const hexColors = [...new Set(
      [...themeXml.matchAll(/val="([0-9A-Fa-f]{6})"/g)].map(m => '#' + m[1])
    )].slice(0, 20)
    if (hexColors.length) extracted += '\n\n[Theme colors]\n' + hexColors.join(', ')
  }

  // Font names from layouts
  const layoutFiles = Object.keys(zip.files)
    .filter(n => /ppt\/slideLayouts\/slideLayout\d+\.xml/.test(n))
    .slice(0, 3)
  for (const lf of layoutFiles) {
    const xml   = await zip.files[lf].async('string')
    const fonts = [...new Set(
      [...xml.matchAll(/typeface="([^"+][^"]*)"/g)].map(m => m[1])
    )]
    if (fonts.length) extracted += '\n[Fonts] ' + fonts.join(', ')
  }

  return extracted.trim().slice(0, 9000)
}

// ─── CLAUDE API CALLER ───────────────────────────────────────────────────────
async function callClaude(system, messages, max_tokens = 1500) {
  const res = await fetch('/api/claude', {
    method:  'POST',
    headers: { 'Content-Type': 'application/json' },
    body:    JSON.stringify({ system, messages, max_tokens })
  })

  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.error || 'API call failed with status ' + res.status)
  }

  const data = await res.json()
  return data.content.map(b => b.text || '').join('')
}

// ─── JSON SAFE PARSER ────────────────────────────────────────────────────────
function safeParseJSON(raw, fallback) {
  try {
    const cleaned = raw
      .replace(/```json\s*/g, '')
      .replace(/```\s*/g, '')
      .trim()
    return JSON.parse(cleaned)
  } catch (e) {
    console.warn('JSON parse failed:', e.message, '\nRaw:', raw.slice(0, 300))
    return fallback
  }
}

// ─── MAIN PIPELINE ───────────────────────────────────────────────────────────
async function runPipeline() {
  const btn = $('run-btn')
  btn.disabled = true
  hideError()
  showStatus()
  setProgress(0)

  // Reset all steps
  for (let i = 1; i <= 5; i++) setStep(i, 'wait')

  try {

    // ── AGENT 1 — Package inputs ──────────────────────────────────────────
    setStep(1, 'active')
    setProgress(5)

    let brandContent = ''

    if (state.brandExt === 'pptx' || state.brandExt === 'ppt') {
      try {
        brandContent = await extractPptxText(state.brandB64)
        if (!brandContent) throw new Error('empty')
      } catch(e) {
        // fallback if extraction fails
        brandContent = 'Brand PPTX file uploaded: ' + state.brandFile.name +
          '. Extract brand design rules based on common corporate brand guidelines.'
      }
    } else if (state.brandExt === 'pdf') {
      // PDF sent directly to Claude as document
      brandContent = '__PDF__'
    } else {
      // Image (png/jpg) sent directly
      brandContent = '__IMAGE__'
    }

    setStep(1, 'done')
    setProgress(15)

    // ── AGENT 2 — Parse brand guidelines ─────────────────────────────────
    setStep(2, 'active')

    const agent2System = `You are an expert brand designer and design systems consultant.
Extract ALL brand design rules from the provided content and return as a single valid JSON object with these exact fields:
- primary_colors: array of hex codes
- secondary_colors: array of hex codes
- background_colors: array of hex codes
- title_font: object with family, size, weight, color
- body_font: object with family, size, weight, color
- caption_font: object with family, size, weight, color
- slide_width: number
- slide_height: number
- logo_position: string
- layout_patterns: array of strings describing recurring slide layouts
- visual_style: string (e.g. flat, minimal, bold, corporate)
- spacing_notes: string
- chart_colors: array of hex codes
If exact hex values are not visible infer reasonable hex codes from color names or context.
Return ONLY valid JSON. No explanation. No markdown fences.`

    let agent2Messages

    if (brandContent === '__PDF__') {
      agent2Messages = [{
        role: 'user',
        content: [
          { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: state.brandB64 } },
          { type: 'text', text: 'Extract all brand design rules from this document. Return JSON only.' }
        ]
      }]
    } else if (brandContent === '__IMAGE__') {
      const mime = state.brandExt === 'png' ? 'image/png' : 'image/jpeg'
      agent2Messages = [{
        role: 'user',
        content: [
          { type: 'image', source: { type: 'base64', media_type: mime, data: state.brandB64 } },
          { type: 'text', text: 'Extract all brand design rules from this slide image. Return JSON only.' }
        ]
      }]
    } else {
      agent2Messages = [{
        role: 'user',
        content: 'Brand guideline content extracted from PPTX:\n\n' + brandContent + '\n\nReturn brand rules as JSON only.'
      }]
    }

    const agent2Raw = await callClaude(agent2System, agent2Messages, 1000)

    state.brandRulebook = safeParseJSON(agent2Raw, {
      visual_style: 'corporate',
      primary_colors: ['#1A3C6E'],
      secondary_colors: ['#F4A300'],
      background_colors: ['#FFFFFF'],
      title_font: { family: 'Calibri', size: '32pt', weight: 'bold', color: '#1A3C6E' },
      body_font:  { family: 'Calibri', size: '18pt', weight: 'regular', color: '#333333' },
      caption_font: { family: 'Calibri', size: '12pt', weight: 'regular', color: '#666666' },
      slide_width: 1280,
      slide_height: 720,
      logo_position: 'top-right',
      layout_patterns: ['title slide', 'two-column', 'full-bleed image', 'bullet list'],
      spacing_notes: '0.5 inch margins, 0.3 inch between blocks',
      chart_colors: ['#1A3C6E', '#F4A300', '#22C55E']
    })

    console.log('Agent 2 — Brand Rulebook:', state.brandRulebook)
    setStep(2, 'done')
    setProgress(30)

    // ── AGENT 3 — Analyse content & build structure ───────────────────────
    setStep(3, 'active')

    const agent3System = `You are a management consultant structuring a board-level presentation for senior leadership.
Read the content carefully and create a logical presentation outline.
Rules:
- Total slides must be exactly ${state.slideCount}
- First slide is always a Title slide
- Include section divider slides between major topics
- Last slide is always Next Steps or Conclusion
- Each slide gets a clear title and one-line description of what it covers
- Think about narrative flow: context → problem → insight → recommendation → next steps
Return a numbered list. Format each line exactly as:
Slide N: [Title] — [one line description]`

    const agent3Messages = [{
      role: 'user',
      content: [
        { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: state.contentB64 } },
        { type: 'text', text: `Create a ${state.slideCount} slide presentation outline for senior management from this content. Follow the format exactly.` }
      ]
    }]

    const agent3Raw = await callClaude(agent3System, agent3Messages, 1500)

    console.log('Agent 3 — Outline:', agent3Raw)
    setStep(3, 'done')
    setProgress(50)

    // ── AGENT 4 — Write detailed slide content ────────────────────────────
    setStep(4, 'active')

    const agent4System = `You are a senior business writer creating content for a board-level presentation.
Expand each slide in the outline into full content.
Return a JSON array where each object has:
- slide_number: number
- type: "title" or "divider" or "content"
- title: string
- subtitle: string (for title slides only, empty string otherwise)
- bullets: array of strings (3-5 bullets for content slides, empty for dividers)
- visual_type: one of "text", "three-column", "two-column", "chart", "table", "icons", "quote", "full-image"
- speaker_note: one sentence summary for the presenter
Return ONLY a valid JSON array. No explanation. No markdown fences.`

    const agent4Messages = [{
      role: 'user',
      content: 'Expand this presentation outline into full slide content as a JSON array:\n\n' + agent3Raw
    }]

    const agent4Raw = await callClaude(agent4System, agent4Messages, 1500)

    state.slideManifest = safeParseJSON(agent4Raw, [{
      slide_number: 1,
      type: 'title',
      title: 'Presentation',
      subtitle: '',
      bullets: [],
      visual_type: 'text',
      speaker_note: 'Opening slide'
    }])

    console.log('Agent 4 — Slide Manifest:', state.slideManifest)
    setStep(4, 'done')
    setProgress(70)

    // ── AGENT 5 — Merge brand + content into final spec ───────────────────
    setStep(5, 'active')

    const agent5System = `You are a senior presentation designer. 
Combine the slide content and brand rules to produce a final slide-by-slide design specification.
For each slide return a JSON array where each object has:
- slide_number: number
- type: string
- title: string
- subtitle: string
- bullets: array of strings
- visual_type: string
- background_color: hex code from brand rules
- title_color: hex code from brand rules
- title_font: string (font family from brand rules)
- title_size: string (font size)
- body_font: string
- body_size: string
- body_color: hex code
- accent_color: hex code (for dividers, icons, highlights)
- layout: string describing placement
- speaker_note: string
Follow the brand rules strictly for all colors and fonts.
Return ONLY a valid JSON array. No explanation. No markdown fences.`

    const agent5Messages = [{
      role: 'user',
      content: 'Brand rules:\n' + JSON.stringify(state.brandRulebook) +
        '\n\nSlide content:\n' + JSON.stringify(state.slideManifest) +
        '\n\nProduce the final branded slide specification as a JSON array.'
    }]

    const agent5Raw = await callClaude(agent5System, agent5Messages, 1500)

    state.finalSpec = safeParseJSON(agent5Raw, state.slideManifest)

    console.log('Agent 5 — Final Spec:', state.finalSpec)
    setStep(5, 'done')
    setProgress(100)

    // ── PIPELINE COMPLETE ─────────────────────────────────────────────────
    showResults()

  } catch (err) {
    console.error('Pipeline error:', err)
    showError('Something went wrong: ' + err.message)
    btn.disabled = false
  }
}

// ─── SHOW RESULTS ─────────────────────────────────────────────────────────────
function showResults() {
  const box = $('results-box')
  box.classList.add('show')

  // Summary cards
  const totalSlides = state.finalSpec ? state.finalSpec.length : 0
  const colors      = state.brandRulebook ? (state.brandRulebook.primary_colors || []) : []

  $('res-slides').textContent  = totalSlides + ' slides'
  $('res-style').textContent   = state.brandRulebook ? (state.brandRulebook.visual_style || '—') : '—'
  $('res-font').textContent    = state.brandRulebook && state.brandRulebook.title_font ? state.brandRulebook.title_font.family : '—'
  $('res-colors').innerHTML    = colors.map(c =>
    '<span style="display:inline-block;width:16px;height:16px;background:' + c + ';border-radius:3px;margin-right:4px;vertical-align:middle;border:1px solid #ccc"></span>'
  ).join('')

  // Slide list preview
  if (state.finalSpec && state.finalSpec.length) {
    $('slide-preview').innerHTML = state.finalSpec.map(s =>
      '<div class="slide-row">' +
        '<span class="slide-num">S' + s.slide_number + '</span>' +
        '<span class="slide-type ' + s.type + '">' + s.type + '</span>' +
        '<span class="slide-title">' + (s.title || '—') + '</span>' +
        '<span class="slide-layout">' + (s.visual_type || '—') + '</span>' +
      '</div>'
    ).join('')
  }
}

// ─── DOWNLOAD SPEC (for Agent 6) ──────────────────────────────────────────────
function downloadSpec() {
  const output = {
    brandRulebook: state.brandRulebook,
    finalSpec:     state.finalSpec,
    slideCount:    state.slideCount,
    generatedAt:   new Date().toISOString()
  }
  const blob = new Blob([JSON.stringify(output, null, 2)], { type: 'application/json' })
  const url  = URL.createObjectURL(blob)
  const a    = document.createElement('a')
  a.href     = url
  a.download = 'presentation-spec.json'
  a.click()
  URL.revokeObjectURL(url)
}
async function generatePPTX() {
  const btn        = document.getElementById('pptx-btn')
  const statusEl   = document.getElementById('agent6-status')
  const progressEl = document.getElementById('agent6-progress')
  const cardEl     = document.getElementById('pptx-download-card')

  if (!state.finalSpec || !state.brandRulebook) {
    alert('Please run the full pipeline first.')
    return
  }

  btn.disabled = true
  cardEl.style.display   = 'none'
  statusEl.textContent   = '⏳ Sending spec to Agent 6...'
  progressEl.style.width = '20%'

  try {
    progressEl.style.width = '50%'
    statusEl.textContent   = '⏳ python-pptx building slides on server...'

    const res  = await fetch('/api/generate-pptx', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({
        finalSpec:     state.finalSpec,
        brandRulebook: state.brandRulebook
      })
    })

    const json = await res.json()
    if (!res.ok || !json.success) throw new Error(json.error || 'Agent 6 failed')

    progressEl.style.width = '90%'
    statusEl.textContent   = '⏳ Preparing download...'

    // Decode base64 → Blob → Object URL
    const binary = atob(json.data)
    const bytes  = new Uint8Array(binary.length)
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
    const blob = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    })
    const url      = URL.createObjectURL(blob)
    const filename = json.filename || 'presentation.pptx'

    // Show green download card on screen
    const link     = document.getElementById('pptx-link')
    link.href      = url
    link.download  = filename
    link.textContent = '⬇  Download ' + filename + ' — ' + json.slides + ' slides'
    cardEl.style.display = 'block'

    // Also auto-trigger download
    const a    = document.createElement('a')
    a.href     = url
    a.download = filename
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)

    progressEl.style.width = '100%'
    statusEl.textContent   = '✅ ' + json.slides + ' slides generated successfully.'
    btn.disabled           = false
    btn.textContent        = '↺ Regenerate PPTX'

  } catch (err) {
    statusEl.textContent   = '❌ Error: ' + err.message
    progressEl.style.width = '0%'
    btn.disabled           = false
    console.error('Agent 6 error:', err)
  }
}
