// ─── STATE ───────────────────────────────────────────────────────────────────
const state = {
  brandFile:     null,
  brandB64:      null,
  brandExt:      null,
  contentB64:    null,
  slideCount:    12,
  // Agent outputs — in pipeline order
  brandRulebook: null,   // Agent 2
  outline:       null,   // Agent 3
  slideManifest: null,   // Agent 4 — zones, artifacts, archetypes
  designedSpec:  null,   // Agent 5 — final blocks-first slide specs
  reviewedSpec:  null,   // Agent 5.1 — partner-reviewed final spec
  reviewReport:  null,   // Agent 5.1 — structured critique
  finalSpec:     null,   // alias → reviewedSpec (consumed by Agent 6)
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

// Step numbers: 1=Input, 2=Brand, 3=Structure, 4=Content, 5=Design, 6=Review
function setStep(num, status) {
  const labels = [
    '',
    'Agent 1 — Collecting & packaging inputs',
    'Agent 2 — Parsing brand guidelines',
    'Agent 3 — Analysing content & building structure',
    'Agent 4 — Writing detailed slide content',
    'Agent 5 — Building design specification',
    'Agent 5.1 — Senior partner review'
  ]
  const el = $('s' + num)
  if (!el) return
  const icons = { done: '✅', active: '🔵', wait: '⬜', error: '❌' }
  el.textContent = (icons[status] || '⬜') + ' ' + labels[num]
  el.className   = 'step ' + status
}

function setProgress(pct) {
  $('prog').style.width = pct + '%'
  const lbl = $('prog-label')
  if (lbl) lbl.textContent = pct + '%'
}

function showStatus()   { $('status-box').classList.add('show') }
function showError(msg) { $('error-msg').textContent = msg; $('error-box').classList.add('show') }
function hideError()    { $('error-box').classList.remove('show') }

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

// ─── SHARED: SAFE JSON PARSER ─────────────────────────────────────────────────
// Attempts three strategies in order:
//   1. Strip markdown fences and parse directly
//   2. Extract the outermost { } or [ ] block and parse that
//      (handles preamble text like "Here is the analysis: {...}")
//   3. Return fallback
function safeParseJSON(raw, fallback) {
  if (!raw) { console.warn('safeParseJSON: empty response'); return fallback }

  // Strategy 1 — strip fences and parse
  try {
    const cleaned = raw.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim()
    return JSON.parse(cleaned)
  } catch (e1) {
    // Strategy 2 — find first { or [ and last matching closer
    try {
      const isArr   = raw.indexOf('[') !== -1 && (raw.indexOf('[') < (raw.indexOf('{') === -1 ? Infinity : raw.indexOf('{')))
      const opener  = isArr ? '[' : '{'
      const closer  = isArr ? ']' : '}'
      const start   = raw.indexOf(opener)
      const end     = raw.lastIndexOf(closer)
      if (start !== -1 && end > start) {
        return JSON.parse(raw.slice(start, end + 1))
      }
    } catch (e2) {}

    console.warn('safeParseJSON failed:', e1.message)
    console.warn('Raw response (first 600):', raw.slice(0, 600))
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
  for (let i = 1; i <= 6; i++) setStep(i, 'wait')

  try {

    // ── AGENT 1 — Package inputs ─────────────────────────────────────────────
    setStep(1, 'active')
    setProgress(5)
    const brandContent = await runAgent1(state)
    setStep(1, 'done')
    setProgress(10)

    // ── AGENT 2 — Parse brand guidelines ─────────────────────────────────────
    setStep(2, 'active')
    state.brandRulebook = await runAgent2(state, brandContent)
    console.log('Agent 2 — Brand Rulebook:', JSON.stringify(state.brandRulebook, null, 2))
    setStep(2, 'done')
    setProgress(22)

    // ── AGENT 3 — Analyse content & build structure ───────────────────────────
    setStep(3, 'active')
    state.outline = await runAgent3(state)
    console.log('Agent 3 — Outline:\n', state.outline)
    setStep(3, 'done')
    setProgress(35)

    // ── AGENT 4 — Write detailed slide content ────────────────────────────────
    setStep(4, 'active')
    state.slideManifest = await runAgent4(state)
    console.log('Agent 4 — Slide Manifest:', JSON.stringify(state.slideManifest, null, 2))
    setStep(4, 'done')
    setProgress(55)

    // ── AGENT 5 — Layout & Design Engine ────────────────────────────────────────
    // Calls Claude in batches and emits final blocks-first slide specs
    // Content from Agent 4 manifest is merged in automatically
    setStep(5, 'active')
    state.designedSpec = await runAgent5(state)
    const totalBlocks = state.designedSpec.reduce((s, sl) => s + ((sl.blocks || []).length), 0)
    const totalArts  = state.designedSpec.reduce((s, sl) =>
      s + ((sl.zones_summary || []).reduce((m, z) => m + ((z.artifact_types || []).length), 0)), 0)
    console.log('Agent 5 — Designed Spec:', state.designedSpec.length, 'slides |',
      totalBlocks, 'blocks |', totalArts, 'artifacts')
    setStep(5, 'done')
    setProgress(72)

    // ── AGENT 5.1 — Senior Partner Review ────────────────────────────────────
    // Makes ONE Claude API call — reviews deck quality, applies targeted fixes
    setStep(6, 'active')
    const { reviewedSpec, reviewReport } = await runAgent51(state)
    state.reviewedSpec = reviewedSpec
    state.reviewReport = reviewReport
    state.finalSpec    = reviewedSpec   // Agent 6 reads state.finalSpec

    const rating   = reviewReport ? reviewReport.overall_rating : 'not_reviewed'
    const issueCount = reviewReport ? (reviewReport.issues || []).length : 0
    console.log('Agent 5.1 — Review complete:', rating, '|', issueCount, 'issues')
    setStep(6, 'done')
    setProgress(100)

    // ── PIPELINE COMPLETE ─────────────────────────────────────────────────────
    showResults()

  } catch (err) {
    console.error('Pipeline error:', err)
    showError('Something went wrong: ' + err.message)
    btn.disabled = false
  }
}

// ─── SHOW RESULTS ─────────────────────────────────────────────────────────────
function showResults() {
  $('results-box').classList.add('show')

  const totalSlides = state.reviewedSpec ? state.reviewedSpec.length : 0
  const colors      = state.brandRulebook ? (state.brandRulebook.primary_colors || []) : []
  const rr          = state.reviewReport

  $('res-slides').textContent = totalSlides + ' slides'
  $('res-style').textContent  = state.brandRulebook ? (state.brandRulebook.visual_style || '—') : '—'
  $('res-font').textContent   = state.brandRulebook && state.brandRulebook.title_font
    ? (state.brandRulebook.title_font.family || '—') : '—'
  $('res-colors').innerHTML   = colors.map(c =>
    '<span style="display:inline-block;width:16px;height:16px;background:' + c +
    ';border-radius:3px;margin-right:4px;vertical-align:middle;border:1px solid #ccc"></span>'
  ).join('')

  // Review badge
  if (rr) {
    const ratingLabel = {
      approved:        '✅ Approved',
      minor_revisions: '⚠ Minor Revisions',
      major_revisions: '❌ Major Revisions',
      not_reviewed:    '— Not Reviewed'
    }
    const reviewEl = $('res-review')
    if (reviewEl) reviewEl.textContent = ratingLabel[rr.overall_rating] || rr.overall_rating || '—'
    const issueEl = $('res-review-issues')
    if (issueEl) {
      const issues    = rr.issues || []
      const criticals = issues.filter(i => i.severity === 'critical').length
      issueEl.textContent = issues.length + ' issues' + (criticals ? ' (' + criticals + ' critical)' : '')
    }
  }

  // Slide preview list
  if (state.reviewedSpec && state.reviewedSpec.length) {
    $('slide-preview').innerHTML = state.reviewedSpec.map(s => {
      const flagged = (s.partner_review || {}).flagged
      return '<div class="slide-row">' +
        '<span class="slide-num">S' + s.slide_number + '</span>' +
        '<span class="slide-type ' + (s.slide_type || 'content') + '">' + (s.slide_type || 'content') + '</span>' +
        '<span class="slide-title">' + (s.title || '—') + '</span>' +
        '<span class="slide-layout">' + (s.slide_archetype || '—') + (flagged ? ' ⚠' : '') + '</span>' +
      '</div>'
    }).join('')
  }
}

// ─── DOWNLOAD SPEC JSON ───────────────────────────────────────────────────────
function downloadSpec() {
  const output = {
    brandRulebook:  state.brandRulebook,
    outline:        state.outline,
    slideManifest:  state.slideManifest,
    designedSpec:   state.designedSpec,
    reviewedSpec:   state.reviewedSpec,
    reviewReport:   state.reviewReport,
    slideCount:     state.slideCount,
    generatedAt:    new Date().toISOString()
  }
  const blob = new Blob([JSON.stringify(output, null, 2)], { type: 'application/json' })
  const url  = URL.createObjectURL(blob)
  const a    = document.createElement('a')
  a.href = url; a.download = 'presentation-spec.json'; a.click()
  URL.revokeObjectURL(url)
}
