// ─── AGENT 6 — PPTX GENERATOR ────────────────────────────────────────────────
// Input:  state.finalSpec     — reviewed spec from Agent 5 / 5.1
//         state.brandRulebook — from Agent 2 (passed to backend for fallbacks)
// Output: downloadable .pptx via /api/generate-pptx (Python serverless)
//
// No Claude API call. Pure POST to backend → decode base64 → download.
// Schema expected: canvas, brand_tokens, and blocks[] as the render contract.
//                  blocks[] is the PRIMARY rendering path — Agent 5 flattens all
//                  non-chart/table artifacts (insight_text, cards, workflow, matrix,
//                  driver_tree, prioritization) into primitive render blocks.
//                  (text_box, rect, line, rule, image) with native chart/table
//                  blocks retained only where needed. zones[] is diagnostic-only.
//
// ─── RENDERING CONTRACT: insight_text ────────────────────────────────────────
//
// ─── RENDERING CONTRACT: workflow ────────────────────────────────────────────
//
// Workflow blocks are emitted as primitive rect / text_box / line blocks by Agent 5:
//
// ── left_to_right / timeline ─────────────────────────────────────────────────
//   Node box    → rect block
//   Label       → text_box INSIDE the node, valign: "middle"
//   Value       → text_box ABOVE the node (y = node.y - valueH - externalGap), valign: "bottom"
//   Description → text_box BELOW the node (y = node.y + node.h + externalGap), valign: "top"
//   Connectors  → line blocks (segment per path waypoint pair)
//
// ── top_to_bottom / bottom_up ────────────────────────────────────────────────
//   Node box    → rect block (occupies ~40% of container width)
//   Label       → text_box INSIDE the node, valign: "middle"
//   Description → text_box to the RIGHT of the node (x = node.x + node.w + externalGap)
//                 width = remaining container width; never place text above/below in vertical flows
//   Connectors  → line blocks (bottom-center → top-center of adjacent nodes)
//
// ── top_down_branching / hierarchy ───────────────────────────────────────────
//   Same as left_to_right for label/description band — description placed BELOW the node
//
// All coordinates (x, y, w, h) are pre-computed by Agent 5. The renderer writes
// them as-is; no layout recalculation is needed in Agent 6 or generate_pptx.py.
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

function slimBlockForRender(block) {
  if (!block || typeof block !== 'object') return block
  const out = {}
  const keep = new Set([
    'block_type',
    'x', 'y', 'w', 'h', 'x1', 'y1', 'x2', 'y2',
    'text', 'points', 'rows', 'headers', 'cells',
    'font_family', 'font_size', 'bold', 'italic', 'underline', 'color', 'font_color',
    'align', 'valign', 'fill_color', 'border_color', 'border_width', 'corner_radius',
    'line_width', 'line_style', 'arrowhead', 'rotation', 'padding', 'body_style', 'heading_style', 'style',
    'chart_type', 'chart_header', 'chart_title', 'categories', 'series', 'dual_axis', 'secondary_series',
    'show_data_labels', 'show_legend', 'x_label', 'y_label', 'secondary_y_label', 'chart_style', 'series_style',
    'legend_position', 'data_label_size', 'category_label_rotation',
    'table_header', 'table_style', 'column_widths', 'column_x_positions', 'row_heights', 'row_y_positions',
    'header_row_height', 'header_cell_frames', 'body_cell_frames', 'column_types', 'column_alignments',
    'workflow_style', 'flow_direction', 'workflow_type', 'nodes', 'connections',
    'artifact_id', 'artifact_type', 'artifact_subtype', 'artifact_header_text', 'block_role',
    'fallback_policy', 'sentiment', 'image_b64', 'image_path', 'src'
  ])
  for (const [k, v] of Object.entries(block)) {
    if (keep.has(k) && v !== undefined) out[k] = v
  }
  return out
}

function slimSlideForRender(slide, useTemplate) {
  if (!slide || typeof slide !== 'object') return slide
  const out = {}
  const keep = new Set([
    'slide_number', 'slide_type', 'slide_archetype',
    'canvas', 'global_elements',
    'layout_mode', 'selected_layout_name',
    'title', 'subtitle', 'speaker_note',
    'artifact_groups', 'blocks'
  ])
  for (const [k, v] of Object.entries(slide)) {
    if (!keep.has(k) || v === undefined) continue
    if (k === 'artifact_groups' && Array.isArray(v)) {
      // Slim the nested blocks within each artifact group
      out[k] = v.map(ag => ({
        ...ag,
        blocks: (ag.blocks || []).map(slimBlockForRender)
      }))
      continue
    }
    if (k === 'blocks' && Array.isArray(v)) {
      out[k] = v.map(slimBlockForRender)
      continue
    }
    if (k === 'global_elements' && v && typeof v === 'object') {
      const ge = { ...v }
      if (useTemplate && ge.logo) delete ge.logo
      out[k] = ge
      continue
    }
    out[k] = v
  }
  return out
}

function slimBrandRulebookForRender(rulebook, finalSpec) {
  const rb = rulebook || {}
  return {
    title_layout_name: rb.title_layout_name || '',
    divider_layout_name: rb.divider_layout_name || '',
    primary_colors: rb.primary_colors || [],
    slide_width_inches: rb.slide_width_inches || finalSpec?.[0]?.canvas?.width_in || 10,
    slide_height_inches: rb.slide_height_inches || finalSpec?.[0]?.canvas?.height_in || 7.5
  }
}

function estimateJsonBytes(obj) {
  try {
    return new TextEncoder().encode(JSON.stringify(obj)).length
  } catch (_) {
    return 0
  }
}

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
  if (!state.finalSpec || !state.brandRulebook || !state.brandTokens) {
    alert('Please run the full pipeline first (click Run AI Pipeline).')
    return
  }
  if (!Array.isArray(state.finalSpec) || state.finalSpec.length === 0) {
    alert('Final spec is empty. Please run the pipeline again.')
    return
  }

  // ── Schema validation ─────────────────────────────────────────────────────
  const hasNewSchema = state.finalSpec.every(slide =>
    slide && slide.canvas && (
      (Array.isArray(slide.artifact_groups) && slide.artifact_groups.length > 0) ||
      (Array.isArray(slide.blocks) && slide.blocks.length > 0)
    )
  )
  if (!hasNewSchema) {
    alert('Spec schema mismatch. Agent 6 only accepts Agent 5 output with canvas plus non-empty artifact_groups[] (or blocks[]) on every slide.')
    return
  }
  const firstSlide   = state.finalSpec[0]

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

    // Sort by slide_number before rendering to guarantee title → content → closing order,
    // regardless of which path (Agent 5 / 5.1 / fallback) produced the spec.
    const sortedFinalSpec = Array.isArray(state.finalSpec)
      ? [...state.finalSpec].sort((a, b) => (a?.slide_number || 0) - (b?.slide_number || 0))
      : []
    let renderSpec = sortedFinalSpec.map(s => slimSlideForRender(s, useTemplate))

    // ── Layout-mode sanitisation ─────────────────────────────────────────────
    // layout_mode and selected_layout_name are ONLY meaningful for content slides.
    // Title / divider / thank-you slides use their dedicated master layouts
    // (title_layout_name / divider_layout_name from the brand rulebook), so
    // we strip any content-layout fields from them here to prevent the backend
    // from accidentally applying a content layout to a title or divider slide.
    //
    // Additionally, if no PPTX template is loaded there are no named layouts at
    // all — strip layout_mode from content slides too so the backend renders
    // them in scratch/coordinate mode using the blocks[] coordinates from Agent 5.
    const NON_CONTENT_TYPES = new Set(['title', 'divider', 'thank_you', 'thankyou', 'end'])
    const strippedNoTemplate = []
    const strippedNonContent = []
    renderSpec = renderSpec.map(slide => {
      const isNonContent = NON_CONTENT_TYPES.has((slide.slide_type || '').toLowerCase())
      const hasLayoutFields = slide.layout_mode || slide.selected_layout_name

      if (isNonContent && hasLayoutFields) {
        // Title/divider/thank-you must never carry content-layout fields
        strippedNonContent.push('S' + slide.slide_number + ' (' + slide.slide_type + ')')
        const out = Object.assign({}, slide)
        delete out.layout_mode
        delete out.selected_layout_name
        return out
      }
      if (!useTemplate && slide.slide_type === 'content' && slide.layout_mode) {
        // No template loaded — named layout cannot be resolved
        strippedNoTemplate.push('S' + slide.slide_number + ' (' + (slide.selected_layout_name || '?') + ')')
        const out = Object.assign({}, slide)
        delete out.layout_mode
        delete out.selected_layout_name
        return out
      }
      return slide
    })
    if (strippedNonContent.length > 0) {
      console.log('Agent 6 — stripped layout_mode from non-content slides:', strippedNonContent.join(', '))
    }
    if (strippedNoTemplate.length > 0) {
      console.log('Agent 6 — no template: cleared layout_mode on', strippedNoTemplate.length, 'content slide(s):', strippedNoTemplate.join(', '))
      statusEl.textContent = '⚠ No brand template — layout_mode cleared on ' + strippedNoTemplate.length + ' content slide(s); rendering in scratch mode...'
    }

    const basePayload = {
      finalSpec:     renderSpec,
      brand_tokens:  state.brandTokens,
      brandRulebook: slimBrandRulebookForRender(state.brandRulebook, renderSpec),
      templateB64:   null
    }
    let payload = {
      ...basePayload,
      templateB64: useTemplate ? state.brandB64 : null
    }

    const MAX_SAFE_BYTES = 4_000_000
    const payloadBytes = estimateJsonBytes(payload)
    if (useTemplate && payloadBytes > MAX_SAFE_BYTES) {
      console.warn('Agent 6 -- payload too large with template, retrying without template', payloadBytes)
      payload = basePayload
      statusEl.textContent = '⚠ Template omitted because request exceeded endpoint size limit. Rendering from blank theme instead...'
    }

    const slimBytes = estimateJsonBytes(payload)
    if (slimBytes > MAX_SAFE_BYTES) {
      throw new Error('Payload too large even after slimming. Try fewer slides or remove verbose content.')
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
      msg = 'Payload too large — most likely the uploaded PPTX template or too many slides. Retry without template or with fewer slides.'
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
