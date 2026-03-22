// ─── AGENT 6 — PPTX GENERATOR ────────────────────────────────────────────────
// Input:  state.finalSpec from Agent 5, state.brandRulebook from Agent 2
// Output: downloadable .pptx file (via python-pptx on Vercel backend)
// Makes ONE call to /api/generate-pptx (Python serverless function)

async function generatePPTX() {
  const btn        = $('pptx-btn')
  const statusEl   = $('agent6-status')
  const progressEl = $('agent6-progress')
  const cardEl     = $('pptx-download-card')

  // Guard — must have run pipeline first
  if (!state.finalSpec || !state.brandRulebook) {
    alert('Please run the full pipeline first (click Run AI Pipeline).')
    return
  }

  if (!Array.isArray(state.finalSpec) || state.finalSpec.length === 0) {
    alert('Final spec is empty. Please run the pipeline again.')
    return
  }

  console.log('Agent 6 starting — sending', state.finalSpec.length, 'slides to python-pptx')

  btn.disabled           = true
  btn.textContent        = '⏳ Generating...'
  cardEl.style.display   = 'none'
  statusEl.textContent   = '⏳ Agent 6 starting — sending ' + state.finalSpec.length + ' slides to server...'
  progressEl.style.width = '10%'

  try {
    progressEl.style.width = '30%'
    statusEl.textContent   = '⏳ python-pptx building ' + state.finalSpec.length + ' slides on Vercel...'

    const payload = {
      finalSpec:     state.finalSpec,
      brandRulebook: state.brandRulebook
    }

    console.log('Agent 6 — payload slides:', payload.finalSpec.length)
    console.log('Agent 6 — brand primary color:', payload.brandRulebook.primary_colors)

    const res = await fetch('/api/generate-pptx', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify(payload)
    })

    progressEl.style.width = '70%'
    statusEl.textContent   = '⏳ Decoding response...'

    const json = await res.json()

    if (!res.ok || !json.success) {
      console.error('Agent 6 — server error:', json)
      throw new Error(json.error || 'Agent 6 python-pptx failed with status ' + res.status)
    }

    console.log('Agent 6 — success:', json.slides, 'slides generated')

    progressEl.style.width = '90%'
    statusEl.textContent   = '⏳ Preparing download...'

    // Decode base64 → Uint8Array → Blob → Object URL
    const binary  = atob(json.data)
    const bytes   = new Uint8Array(binary.length)
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)

    const blob     = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    })
    const url      = URL.createObjectURL(blob)
    const filename = json.filename || 'presentation.pptx'

    // Show persistent green download card on screen
    const link       = $('pptx-link')
    link.href        = url
    link.download    = filename
    link.textContent = '⬇  Download ' + filename + ' — ' + json.slides + ' slides'
    cardEl.style.display = 'block'

    // Also auto-trigger browser download
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
    console.error('Agent 6 error:', err)
    statusEl.textContent   = '❌ Error: ' + err.message
    progressEl.style.width = '0%'
    btn.disabled           = false
    btn.textContent        = '▶ Generate PPTX — Agent 6 (python-pptx)'
  }
}
