// ─── STATE ───────────────────────────────────────────────────────────────────
const state = {
  brandFile:     null,
  brandB64:      null,
  brandExt:      null,
  contentB64:    null,
  slideCount:    12,
  brandRulebook: null,
  outline:       null,
  slideManifest: null,
  finalSpec:     null
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
  const el    = $('s' + num)
  const icons = { done: '✅', active: '🔵', wait: '⬜', error: '❌' }
  el.textContent = (icons[status] || '⬜') + ' ' + labels[num]
  el.className   = 'step ' + status
}

function setProgress(pct) {
  $('prog').style.width = pct + '%'
  const lbl = $('prog-label')
  if (lbl) lbl.textContent = pct + '%'
}

function showStatus()      { $('status-box').classList.add('show') }
function showError(msg)    { $('error-msg').textContent = msg; $('error-box').classList.add('show') }
function hideError()       { $('error-box').classList.remove('show') }

// ─── SHARED: CLAUDE API CALLER ───────────────────────────────────────────────
async function callClaude(system, messages, max_tokens = 2000) {
  const res = await fetch('/api/claude', {
    method:  'POST',
    headers: { 'Content-Type': 'application/json' },
    body:    JSON.stringify({ system, messages, max_tokens })
  })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(err.error || 'API call failed — status ' + res.status)
  }
  const data = await res.json()
  return data.content.map(b => b.text || '').join('')
}

// ─── SHARED: SAFE JSON PARSER ────────────────────────────────────────────────
function safeParseJSON(raw, fallback) {
  try {
    const cleaned = raw.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim()
    return JSON.parse(cleaned)
  } catch (e) {
    console.warn('JSON parse failed:', e.message)
    console.warn('Raw response was:', raw.slice(0, 500))
    return fallback
  }
}

// ─── MAIN PIPELINE ORCHESTRATOR ──────────────────────────────────────────────
async function runPipeline() {
  const btn = $('run-btn')
  btn.disabled = true
  hideError()
  showStatus()
  setProgress(0)
  for (let i = 1; i <= 5; i++) setStep(i, 'wait')

  try {
    // Each agent is in its own file
    setStep(1, 'active')
    const brandContent = await runAgent1(state)
    setStep(1, 'done')
    setProgress(15)

    setStep(2, 'active')
    state.brandRulebook = await runAgent2(state, brandContent)
    console.log('Agent 2 output — Brand Rulebook:', JSON.stringify(state.brandRulebook, null, 2))
    setStep(2, 'done')
    setProgress(30)

    setStep(3, 'active')
    state.outline = await runAgent3(state)
    console.log('Agent 3 output — Outline:\n', state.outline)
    setStep(3, 'done')
    setProgress(50)

    setStep(4, 'active')
    state.slideManifest = await runAgent4(state)
    console.log('Agent 4 output — Slide Manifest:', JSON.stringify(state.slideManifest, null, 2))
    setStep(4, 'done')
    setProgress(70)

    setStep(5, 'active')
    state.finalSpec = await runAgent5(state)
    console.log('Agent 5 output — Final Spec:', JSON.stringify(state.finalSpec, null, 2))
    setStep(5, 'done')
    setProgress(100)

    showResults()

  } catch (err) {
    console.error('Pipeline error:', err)
    showError('Something went wrong: ' + err.message)
    btn.disabled = false
  }
}

// ─── SHOW RESULTS ────────────────────────────────────────────────────────────
function showResults() {
  $('results-box').classList.add('show')

  const totalSlides = state.finalSpec ? state.finalSpec.length : 0
  const colors      = state.brandRulebook ? (state.brandRulebook.primary_colors || []) : []

  $('res-slides').textContent = totalSlides + ' slides'
  $('res-style').textContent  = state.brandRulebook ? (state.brandRulebook.visual_style || '—') : '—'
  $('res-font').textContent   = state.brandRulebook && state.brandRulebook.title_font
    ? (state.brandRulebook.title_font.family || '—') : '—'
  $('res-colors').innerHTML   = colors.map(c =>
    '<span style="display:inline-block;width:16px;height:16px;background:' + c +
    ';border-radius:3px;margin-right:4px;vertical-align:middle;border:1px solid #ccc"></span>'
  ).join('')

  if (state.finalSpec && state.finalSpec.length) {
    $('slide-preview').innerHTML = state.finalSpec.map(s =>
      '<div class="slide-row">' +
        '<span class="slide-num">S' + s.slide_number + '</span>' +
        '<span class="slide-type ' + (s.type || 'content') + '">' + (s.type || 'content') + '</span>' +
        '<span class="slide-title">' + (s.title || '—') + '</span>' +
        '<span class="slide-layout">' + (s.visual_type || '—') + '</span>' +
      '</div>'
    ).join('')
  }
}

// ─── DOWNLOAD SPEC JSON ──────────────────────────────────────────────────────
function downloadSpec() {
  const output = {
    brandRulebook: state.brandRulebook,
    outline:       state.outline,
    slideManifest: state.slideManifest,
    finalSpec:     state.finalSpec,
    slideCount:    state.slideCount,
    generatedAt:   new Date().toISOString()
  }
  const blob = new Blob([JSON.stringify(output, null, 2)], { type: 'application/json' })
  const url  = URL.createObjectURL(blob)
  const a    = document.createElement('a')
  a.href = url; a.download = 'presentation-spec.json'; a.click()
  URL.revokeObjectURL(url)
}
