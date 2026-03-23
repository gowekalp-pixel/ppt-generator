// ─── AGENT 6 — PPTX GENERATOR ────────────────────────────────────────────────
// Input:  state.finalSpec     — reviewed spec from Agent 5 / 5.1
//         state.brandRulebook — from Agent 2 (passed to backend for fallbacks)
// Output: downloadable .pptx via /api/generate-pptx (Python serverless)
//
// No Claude API call. Pure POST to backend → decode base64 → download.
// Schema expected: canvas, brand_tokens, title_block, subtitle_block,
//                  zones[].artifacts[], global_elements

async function generatePPTX() {
  const btn        = $('pptx-btn')
  const statusEl   = $('agent6-status')
  const progressEl = $('agent6-progress')
  const cardEl     = $('pptx-download-card')

  // ── Guard ─────────────────────────────────────────────────────────────────
  if (!state.finalSpec || !state.brandRulebook) {
    alert('Please run the full pipeline first (click Run AI Pipeline).')
    return
  }
  if (!Array.isArray(state.finalSpec) || state.finalSpec.length === 0) {
    alert('Final spec is empty. Please run the pipeline again.')
    return
  }

  // ── Schema validation ─────────────────────────────────────────────────────
  const firstSlide   = state.finalSpec[0]
  const hasNewSchema = firstSlide && firstSlide.canvas && firstSlide.zones !== undefined
  if (!hasNewSchema) {
    alert('Spec schema mismatch. Expected canvas/zones schema from the new Agent 5. Please re-run the pipeline.')
    return
  }

  const slideCount = state.finalSpec.length
  console.log('Agent 6 starting —', slideCount, 'slides')
  console.log('  Brand primary:', (state.brandRulebook.primary_colors || [])[0] || 'none')
  console.log('  Slide size:', (firstSlide.canvas.width_in || '?') + '" x ' + (firstSlide.canvas.height_in || '?') + '"')

  // ── UI: reset ─────────────────────────────────────────────────────────────
  btn.disabled           = true
  btn.textContent        = '⏳ Generating...'
  cardEl.style.display   = 'none'
  statusEl.textContent   = '⏳ Sending ' + slideCount + ' slides to python-pptx...'
  progressEl.style.width = '10%'

  try {
    progressEl.style.width = '30%'
    statusEl.textContent   = '⏳ python-pptx building ' + slideCount + ' slides on server...'

    const payload = {
      finalSpec:     state.finalSpec,
      brandRulebook: state.brandRulebook
    }

    const res = await fetch('/api/generate-pptx', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify(payload)
    })

    progressEl.style.width = '70%'
    statusEl.textContent   = '⏳ Decoding response...'

    // Safe JSON parse — server may return HTML on 502
    let json
    try {
      json = await res.json()
    } catch(e) {
      throw new Error('Server returned non-JSON (status ' + res.status + '). Check server logs.')
    }

    if (!res.ok || !json.success) {
      throw new Error(json.error || json.message || 'Server error — status ' + res.status)
    }
    if (!json.data || json.data.length < 100) {
      throw new Error('Server returned empty PPTX data.')
    }

    console.log('Agent 6 — success:', json.slides, 'slides')
    progressEl.style.width = '88%'
    statusEl.textContent   = '⏳ Preparing download...'

    // ── base64 → Blob → download ──────────────────────────────────────────
    const binary  = atob(json.data)
    const bytes   = new Uint8Array(binary.length)
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)

    const blob     = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    })
    const objectUrl = URL.createObjectURL(blob)
    const filename  = json.filename || ('presentation_' + todayStr6() + '.pptx')

    // Persistent download card
    const link       = $('pptx-link')
    link.href        = objectUrl
    link.download    = filename
    link.textContent = '⬇  Download ' + filename + ' (' + json.slides + ' slides)'
    cardEl.style.display = 'block'

    // Auto-trigger browser download
    const a    = document.createElement('a')
    a.href     = objectUrl
    a.download = filename
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)

    progressEl.style.width = '100%'
    statusEl.textContent   = '✅ ' + json.slides + ' slides generated — ' + filename
    btn.disabled    = false
    btn.textContent = '↺ Regenerate PPTX'

  } catch (err) {
    console.error('Agent 6 error:', err)
    let msg = err.message
    if (msg.includes('Failed to fetch') || msg.includes('NetworkError')) {
      msg = 'Network error — cannot reach server.'
    } else if (msg.includes('500') || msg.includes('502')) {
      msg = 'Server error — python-pptx may be starting up. Wait 10 seconds and retry.'
    } else if (msg.includes('413')) {
      msg = 'Payload too large — try fewer slides.'
    }
    statusEl.textContent   = '❌ ' + msg
    progressEl.style.width = '0%'
    btn.disabled    = false
    btn.textContent = '▶ Generate PPTX — Agent 6'
  }
}

function todayStr6() {
  const d = new Date()
  const p = n => n < 10 ? '0' + n : String(n)
  return d.getFullYear() + p(d.getMonth() + 1) + p(d.getDate())
}
