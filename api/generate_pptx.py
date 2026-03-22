// ─── AGENT 6 — ADD THIS TO THE BOTTOM OF public/app.js ──────────────────────

async function generatePPTX() {
  const btn       = document.getElementById('pptx-btn')
  const statusEl  = document.getElementById('agent6-status')
  const progressEl= document.getElementById('agent6-progress')
  const cardEl    = document.getElementById('pptx-download-card')

  if (!state.finalSpec || !state.brandRulebook) {
    alert('Please run the full pipeline first.')
    return
  }

  btn.disabled = true
  statusEl.textContent = '⏳ Sending spec to Agent 6 (python-pptx)...'
  progressEl.style.width = '20%'
  if (cardEl) cardEl.style.display = 'none'

  try {
    progressEl.style.width = '50%'
    statusEl.textContent = '⏳ Building slides on server...'

    const res = await fetch('/api/generate-pptx', {
      method:  'POST',
      headers: { 'Content-Type': 'application/json' },
      body:    JSON.stringify({
        finalSpec:     state.finalSpec,
        brandRulebook: state.brandRulebook
      })
    })

    const json = await res.json()

    if (!res.ok || !json.success) {
      throw new Error(json.error || 'Agent 6 failed')
    }

    progressEl.style.width = '90%'
    statusEl.textContent = '⏳ Preparing download...'

    // Decode base64 → blob → object URL
    const binary = atob(json.data)
    const bytes  = new Uint8Array(binary.length)
    for (let i = 0; i < binary.length; i++) {
      bytes[i] = binary.charCodeAt(i)
    }
    const blob     = new Blob([bytes], {
      type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    })
    const url      = URL.createObjectURL(blob)
    const filename = json.filename || 'presentation.pptx'

    // Show download card on screen
    if (cardEl) {
      cardEl.style.display = 'block'
      const link = document.getElementById('pptx-link')
      if (link) {
        link.href     = url
        link.download = filename
        link.textContent = '⬇  Download ' + filename + '  (' + json.slides + ' slides)'
      }
    }

    // Also auto-trigger download
    const a    = document.createElement('a')
    a.href     = url
    a.download = filename
    document.body.appendChild(a)
    a.click()
    document.body.removeChild(a)

    progressEl.style.width = '100%'
    statusEl.textContent   = '✅ Done! ' + json.slides + ' slides generated.'
    btn.disabled           = false
    btn.textContent        = '↺ Regenerate PPTX'

  } catch (err) {
    statusEl.textContent   = '❌ Error: ' + err.message
    progressEl.style.width = '0%'
    btn.disabled           = false
    console.error('Agent 6 error:', err)
  }
}


// ─── ADD THIS BLOCK TO public/index.html  ────────────────────────────────────
// Place it inside the results-box div, after the existing download-btn button
//
// <div id="pptx-download-card" style="display:none; margin-top:16px;
//      background:#f0fdf4; border:1.5px solid #22c55e; border-radius:10px;
//      padding:16px 20px;">
//   <div style="font-size:12px;font-weight:700;text-transform:uppercase;
//        letter-spacing:0.06em;color:#16a34a;margin-bottom:8px;">
//     ✅ Agent 6 Complete — Your PPTX is Ready
//   </div>
//   <a id="pptx-link" href="#" download
//      style="display:block;padding:12px 16px;background:#16a34a;color:white;
//             border-radius:8px;font-size:15px;font-weight:700;
//             text-decoration:none;text-align:center;">
//     ⬇  Download presentation.pptx
//   </a>
//   <div style="font-size:12px;color:#555;margin-top:8px;text-align:center;">
//     File stays available until you refresh the page
//   </div>
// </div>
//
// <button class="download-btn" id="pptx-btn" onclick="generatePPTX()"
//   style="margin-top:10px;background:#7c3aed;">
//   ▶ Generate PPTX (Agent 6)
// </button>
// <div id="agent6-status"
//   style="font-size:13px;color:#555;margin-top:8px;min-height:18px;"></div>
// <div style="height:4px;background:#e5e7eb;border-radius:99px;margin-top:6px;">
//   <div id="agent6-progress"
//     style="height:100%;background:#7c3aed;border-radius:99px;
//            width:0%;transition:width 0.4s ease;"></div>
// </div>