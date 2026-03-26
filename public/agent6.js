// ─── AGENT 6 — PPTX GENERATOR ────────────────────────────────────────────────
// Input:  state.finalSpec     — reviewed spec from Agent 5 / 5.1
//         state.brandRulebook — from Agent 2 (passed to backend for fallbacks)
// Output: downloadable .pptx via /api/generate-pptx (Python serverless)
//
// No Claude API call. Pure POST to backend → decode base64 → download.
// Schema expected: canvas, brand_tokens, title_block, subtitle_block,
//                  blocks[] as the primary render contract.
//                  zones[].artifacts[] remains as backward-compatible metadata.
//
// ─── RENDERING CONTRACT: insight_text ────────────────────────────────────────
//
// artifact.insight_mode determines the render path:
//
// ── "standard" mode ──────────────────────────────────────────────────────────
// Render flat bullet list using artifact.body_style:
//   - list_style "bullet"/"tick_cross"/"numbered"
//   - vertical_distribution "spread": evenly space points across full artifact height
//   - header_block rendered above the list (if not null)
//
// ── "grouped" mode ───────────────────────────────────────────────────────────
// All dimensions come from the spec as calculated by Agent 5 — do NOT hardcode.
// Use the formulas below to reconstruct positions from spec values.
//
// artifact.group_layout === "columns":
//   n        = groups.length
//   col_w    = (artifact.w - (n-1) × group_gap_in) / n          [equal width]
//   header_h = group_header_style.h                              [from spec]
//   box_h    = artifact.h - header_block_h - header_h - header_to_box_gap_in  [fills rest]
//
//   Per group column i (x_offset = artifact.x + i × (col_w + group_gap_in)):
//     1. GROUP HEADER at (x_offset, artifact.y + header_block_h):
//        - "rounded_rect": rect w=col_w, h=header_h, filled, corner_radius from spec
//        - "circle_badge": circle diameter=header_h, centered horizontally in col_w;
//          show 1-based priority number in bold text inside
//        Text: group_header_style font, text_color, centered
//     2. GAP of header_to_box_gap_in
//     3. BULLET BOX at (x_offset, above_y + header_h + header_to_box_gap_in):
//        w=col_w, h=box_h; rounded rect, fill + border from group_bullet_box_style
//        Bullet list from groups[i].bullets using artifact.bullet_style
//        Inner padding from group_bullet_box_style.padding
//
// artifact.group_layout === "rows":
//   n             = groups.length
//   total_bullets = sum of groups[i].bullets.length
//   total_row_h   = artifact.h - header_block_h - (n-1) × group_gap_in
//   row_h[i]      = total_row_h × (groups[i].bullets.length / total_bullets)
//                   [PROPORTIONAL to bullet count — rows with more bullets get more height]
//                   enforce minimum: group_header_style.h + header_to_box_gap_in + one text line
//   header_w      = group_header_style.w                         [from spec]
//   box_w         = artifact.w - header_w - header_to_box_gap_in [fills rest]
//
//   Per group row i (y_offset = artifact.y + header_block_h + sum of prior row_h + i × group_gap_in):
//     1. GROUP HEADER at (artifact.x, y_offset):
//        - "rounded_rect": rect w=header_w, h=row_h[i], filled, corner_radius from spec
//        - "circle_badge": circle diameter=group_header_style.h, centered vertically in row_h[i];
//          show 1-based priority number inside
//        Text: group_header_style font, text_color, centered
//     2. GAP of header_to_box_gap_in to the right of header
//     3. BULLET BOX at (artifact.x + header_w + header_to_box_gap_in, y_offset):
//        w=box_w, h=row_h[i]; rounded rect, fill + border from group_bullet_box_style
//        Bullet list from groups[i].bullets using artifact.bullet_style
//        Inner padding from group_bullet_box_style.padding
//
// SPACING CONSISTENCY:
//   - group_gap_in and header_to_box_gap_in are uniform across all groups
//   - All padding values are uniform across all bullet boxes
//
// BRAND COLORS:
//   group_header_style.fill_color — use exactly as specified; do NOT substitute
//   group_bullet_box_style.border_color — use exactly as specified; do NOT darken

// ─── TITLE BLOCK SANITISER ────────────────────────────────────────────────────
// Agent 5 occasionally outputs title_block / subtitle_block coordinates that
// push text off-canvas or produce a very narrow w — causing python-pptx to
// render the title vertically down the left edge (each character on its own line).
//
// Rules enforced per slide:
//   title_block.x  ≤ 15 % of canvas width           (not pushed to far left)
//   title_block.w  ≥ 70 % of canvas width            (wide enough for horizontal text)
//   title_block.y  in [0.05, 1.5]                    (within top region)
//   title_block.h  ≥ 0.35                            (enough height for at least one line)
//   subtitle_block follows same x/w rules, y ≤ 2.5
//
// Coordinates that already satisfy the rules are left untouched.
function sanitiseTitleBlocks(spec) {
  if (!Array.isArray(spec)) return spec
  return spec.map(slide => {
    const cvs      = slide.canvas || {}
    const slideW   = cvs.width_in  || 10

    const minW     = slideW * 0.70
    const maxX     = slideW * 0.15
    const defaultX = 0.4
    const defaultW = slideW - 0.8

    function fixBlock(blk, maxY, defaultY, defaultH) {
      if (!blk || !blk.text) return blk
      const out = Object.assign({}, blk)

      // x: must not be so large that w becomes too narrow
      if ((out.x == null) || out.x > maxX) {
        console.warn('Agent 6 sanitise: title_block.x', out.x, '→', defaultX)
        out.x = defaultX
      }

      // w: must cover at least 70% of slide width
      if ((out.w == null) || out.w < minW) {
        console.warn('Agent 6 sanitise: title_block.w', out.w, '→', defaultW)
        out.w = defaultW
      }

      // y: must be in the title region
      if ((out.y == null) || out.y < 0.05 || out.y > maxY) {
        console.warn('Agent 6 sanitise: title_block.y', out.y, '→', defaultY)
        out.y = defaultY
      }

      // h: minimum single-line height
      if ((out.h == null) || out.h < 0.35) {
        console.warn('Agent 6 sanitise: title_block.h', out.h, '→', defaultH)
        out.h = defaultH
      }

      return out
    }

    return Object.assign({}, slide, {
      title_block:    fixBlock(slide.title_block,    1.5, 0.15, 0.75),
      subtitle_block: fixBlock(slide.subtitle_block, 2.5, 1.0,  0.45)
    })
  })
}

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
  const hasNewSchema = !!(
    firstSlide &&
    firstSlide.canvas &&
    (
      Array.isArray(firstSlide.blocks) ||
      firstSlide.zones !== undefined
    )
  )
  if (!hasNewSchema) {
    alert('Spec schema mismatch. Expected canvas plus blocks[] from Agent 5 (zones[] accepted only as legacy metadata). Please re-run the pipeline.')
    return
  }

  const slideCount    = state.finalSpec.length
  const useTemplate   = (state.brandExt === 'pptx' || state.brandExt === 'ppt') && !!state.brandB64
  console.log('Agent 6 starting —', slideCount, 'slides')
  console.log('  Brand primary:', (state.brandRulebook.primary_colors || [])[0] || 'none')
  console.log('  Slide size:', (firstSlide.canvas.width_in || '?') + '" x ' + (firstSlide.canvas.height_in || '?') + '"')
  console.log('  Template mode:', useTemplate ? 'YES — brand PPTX master will be used' : 'NO — building from blank')

  // ── UI: reset ─────────────────────────────────────────────────────────────
  btn.disabled           = true
  btn.textContent        = '⏳ Generating...'
  cardEl.style.display   = 'none'
  statusEl.textContent   = useTemplate
    ? '⏳ Sending ' + slideCount + ' slides to python-pptx (template mode — brand master preserved)...'
    : '⏳ Sending ' + slideCount + ' slides to python-pptx...'
  progressEl.style.width = '10%'

  try {
    progressEl.style.width = '30%'
    statusEl.textContent   = '⏳ python-pptx building ' + slideCount + ' slides on server...'

    const payload = {
      finalSpec:     sanitiseTitleBlocks(state.finalSpec),
      brandRulebook: state.brandRulebook,
      templateB64:   useTemplate ? state.brandB64 : null
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
