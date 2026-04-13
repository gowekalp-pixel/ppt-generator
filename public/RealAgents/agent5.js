// ********* AGENT 5 **” SLIDE LAYOUT & VISUAL DESIGN ENGINE *********************************************************************************
// Input:  state.slideManifest   **” output from Agent 4
//         state.brandRulebook   **” brand guideline JSON from Agent 2
//
// Output: designedSpec **” flat JSON array, one render-ready object per slide
//
// Architecture: Claude API call per batch of 3 slides.
// Claude receives the Agent 4 manifest + brand guideline and returns
// a precise layout spec: canvas, title_block, subtitle_block,
// zones (with fully positioned artifacts), and global_elements.
// brand_tokens are derived from the brand rulebook and hoisted to the
// top-level return value **” Claude does NOT output them per slide.
//
// Agent 5.1 then reviews this spec and applies targeted fixes.
// Agent 6 (python-pptx) consumes the final reviewed spec.

// Defensive global fallback: some older browser-served bundles or indirect
// preview/eval paths may still reference `bt` without a local binding.
// Keeping a harmless global object prevents a hard ReferenceError while the
// local per-function brand token bindings continue to be the primary source.
if (typeof globalThis.bt === 'undefined') globalThis.bt = {}
var bt = globalThis.bt

// ─── PROMPT TEXT ──────────────────────────────────────────────────────────────────────────────
// Phase fragments loaded via <script> tags (see index.html):
//   prompts/agent5/P0-SystemHeader.js            → _A5_HEADER
//   prompts/agent5/A2-SlideRules.js              → _A5_SLIDE_RULES
//   prompts/agent5/A5-SlideSchema.js             → _A5_OUTPUT_SCHEMA
//   prompts/agent5/A1-FlattenedArtifactSchema.js → _A5_ARTIFACTS  ← edit when changing any artifact schema
//   prompts/agent5/A7-QualityCheck.js            → _A5_QUALITY_RULES
//   prompts/agent5/A6-fallback-system.js         → _A5_FALLBACK_SYSTEM
//   prompts/agent5/A3-BatchProcess.js            → _A5_BATCH_INSTRUCTIONS
//   prompts/agent5/A4-FallbackRules.js           → _A5_FALLBACK_INSTRUCTIONS
//
// Repair / fallback logic → repair/agent5-fallback.js
//   buildFallbackDesign, manifestZoneArtifactSignature, validateFallbackStructure

var AGENT5_SYSTEM = [
  _A5_HEADER,
  _A5_SLIDE_RULES,
  _A5_OUTPUT_SCHEMA,
  _A5_ARTIFACTS,
  _A5_QUALITY_RULES,
].join('\n\n')

var AGENT5_FALLBACK_SYSTEM = _A5_FALLBACK_SYSTEM





// *********************************************************************************************************************************************************************************************************************************************
// UTILITIES
// *********************************************************************************************************************************************************************************************************************************************

// Round to 2 decimal places **” kills JS float drift (10.219999... -> 10.22)
function r2(n) { return Math.round(n * 100) / 100 }


// *********************************************************************************************************************************************************************************************************************************************
// BRAND TOKEN EXTRACTOR
// Strips slide_layouts and other bulky fields before sending to Claude.
// slide_layouts contains full placeholder XML **” can be 10K+ tokens alone.
// Agent 5 only needs design tokens: colors, fonts, slide size.
// *********************************************************************************************************************************************************************************************************************************************

function extractBrandTokens(brand) {
  return {
    slide_width_inches:   r2(brand.slide_width_inches  || 13.33),
    slide_height_inches:  r2(brand.slide_height_inches || 7.50),
    primary_colors:       brand.primary_colors       || [],
    secondary_colors:     brand.secondary_colors     || [],
    background_colors:    brand.background_colors    || ['#FFFFFF'],
    text_colors:          brand.text_colors          || ['#111111'],
    accent_colors:        brand.accent_colors        || [],
    chart_colors:         brand.chart_colors         || [],
    chart_color_sequence: brand.chart_color_sequence || brand.chart_colors || [],
    all_colors:           brand.all_colors           || {},
    title_font:           brand.title_font           || {},
    body_font:            brand.body_font            || {},
    caption_font:         brand.caption_font         || {},
    typography_hierarchy: brand.typography_hierarchy || {},
    bullet_style:         brand.bullet_style         || { char: '**¢', indent_inches: 0.12, space_before_pt: 4 },
    insight_box_style:    brand.insight_box_style    || { fill_color: null, border_color: null, corner_radius: 2 },
    visual_style:         brand.visual_style         || 'corporate',
    color_scheme_name:    brand.color_scheme_name    || '',
    spacing_notes:        brand.spacing_notes        || '',
    uses_template:        brand.uses_template        || false,
    // Compact layout map: name â†’ { title_placeholder, body_placeholder, ph_count, content_areas, usage_guidance }
    // content_areas: large body placeholders (h > 0.5") ordered leftâ†’right, topâ†’bottom.
    // Used in LAYOUT MODE **” the pipeline maps zone[i] â†’ content_areas[i] for frame + placeholder_idx.
    layout_map:           (brand.slide_layouts || []).reduce((acc, l) => {
      if (l.name) {
        const contentAreas = (l.placeholders || [])
          .filter(p => p.type === 'body' && (p.h_in || 0) > 0.5)
          .sort((a, b) => {
            const rowA = Math.round((a.y_in || 0) * 2)  // 0.5" row buckets
            const rowB = Math.round((b.y_in || 0) * 2)
            if (rowA !== rowB) return rowA - rowB
            return (a.x_in || 0) - (b.x_in || 0)       // left before right
          })
          .map(p => ({ idx: p.idx, x_in: p.x_in, y_in: p.y_in, w_in: p.w_in, h_in: p.h_in }))
        acc[l.name] = {
          ph_count:          l.ph_count          || 0,
          usage_guidance:    l.usage_guidance    || '',
          title_placeholder: l.title_placeholder || (l.master_summary || {}).title_placeholder || null,
          body_placeholder:  l.body_placeholder  || null,
          content_areas:     contentAreas
        }
      }
      return acc
    }, {})
    // slide_masters, layout_blueprints, master_blueprints intentionally excluded **” too large for API
  }
}

function buildBrandBrief(brand) {
  const tokens = extractBrandTokens(brand)
  return 'BRAND DESIGN TOKENS:\n' +
    JSON.stringify(tokens, null, 2) +
    '\n\nBRAND COMPLIANCE RULES (MUST follow exactly):' +
    '\n- Fonts: use title_font.family for all titles/headings, body_font.family for all body/bullet text' +
    '\n- Title font size: typography_hierarchy.title_size_pt (content slides), larger for title/divider slides' +
    '\n- Body font size: typography_hierarchy.body_size_pt **” do NOT guess; use the extracted value' +
    '\n- Chart colors: use chart_color_sequence in order **” each series/segment gets a DIFFERENT color' +
    '\n- Pie charts: series_style must have one entry PER DATA POINT (category), each with a unique fill_color' +
    '\n- Bullet char: bullet_style.char **” use exactly this character, not substitutes' +
    '\n- Bullet spacing: bullet_style.space_before_pt **” pass directly into body_style.space_before_pt' +
    '\n- Insight boxes: insight_box_style.fill_color and border_color **” a left accent bar is always rendered; do NOT add a full perimeter border' +
    (tokens.uses_template
      ? '\n- TEMPLATE MODE ACTIVE: master provides background/logo/footer **” set global_elements:{}, canvas.background:null' +
        '\n- Title/divider slides: text only on title_block/subtitle_block **” omit x/y/w/h, set layout_mode:true' +
        '\n- Content slides with selected_layout_name: LAYOUT MODE **” set layout_mode:true, zone.frame:null; do NOT set placeholder_idx (pipeline assigns from layout content_areas)' +
        '\n- Content slides without selected_layout_name: SCRATCH MODE **” compute all coordinates from layout_hint'
      : '\n- SCRATCH MODE: compute all coordinates; specify background and footer in global_elements')
}


// *********************************************************************************************************************************************************************************************************************************************
// BATCH WRITER
// Sends one batch of slides to Claude and returns the array of layout specs
// *********************************************************************************************************************************************************************************************************************************************

async function designSlideBatch(batchManifest, brand, batchNum) {
  const slideNums = batchManifest.map(s => s.slide_number)
  console.log('Agent 5 batch', batchNum, ':', slideNums.join(', '))

  // Annotate each slide with its mode so Claude doesn't have to infer it.
  // Strip internal pipeline flags (_was_repaired) **” Claude doesn't need them.
  const annotatedManifest = batchManifest.map(({ _was_repaired: _r, ...s }) => ({
    ...s,
    _mode: (brand.uses_template && s.selected_layout_name)
      ? 'layout_mode'
      : (brand.uses_template && (s.slide_type === 'title' || s.slide_type === 'divider' || s.slide_type === 'thank_you'))
        ? 'template_title_divider'
        : 'scratch_mode'
  }))
  const compactManifest = JSON.stringify(annotatedManifest)

  const prompt =
    buildBrandBrief(brand) +
    '\n\nSLIDE BATCH ' + batchNum + ' (' + annotatedManifest.length + ' slides):\n' +
    compactManifest +
    '\n\n' + _A5_BATCH_INSTRUCTIONS.replace(/__SLIDE_COUNT__/g, batchManifest.length)

  const raw    = await callClaude(AGENT5_SYSTEM, [{ role: 'user', content: prompt }], 8000)
  const parsed = safeParseJSON(raw, null)

  if (!Array.isArray(parsed)) {
    console.warn('Agent 5 batch', batchNum, '-- parse failed. Raw length:', raw.length)
    console.warn('  First 400 chars:', raw.slice(0, 400))
    return null
  }

  console.log('Agent 5 batch', batchNum, '-- got', parsed.length, 'slides')
  return parsed
}


// ═══════════════════════════════════════════════════════════════════════════════
// VALIDATION → repair/agent5-validate.js
// validateDesignedSlide, validateRenderCompleteness,
// rectsOverlap, rectArea
// Note: clamp, rectWithin, getZoneInnerBounds remain here — used by normaliseDesignedSlide
// ═══════════════════════════════════════════════════════════════════════════════

function clamp(n, min, max) {
  return Math.min(Math.max(n, min), max)
}

function rectWithin(rect, bounds) {
  const x = clamp(rect.x || 0, bounds.x, bounds.x + Math.max(0, bounds.w - 0.05))
  const y = clamp(rect.y || 0, bounds.y, bounds.y + Math.max(0, bounds.h - 0.05))
  const maxW = Math.max(0.05, bounds.x + bounds.w - x)
  const maxH = Math.max(0.05, bounds.y + bounds.h - y)
  return {
    ...rect,
    x: r2(x),
    y: r2(y),
    w: r2(clamp(rect.w || maxW, 0.05, maxW)),
    h: r2(clamp(rect.h || maxH, 0.05, maxH))
  }
}

function getZoneInnerBounds(zone) {
  const frame = zone.frame || {}
  const p = frame.padding || {}
  return {
    x: r2((frame.x || 0) + (p.left || 0)),
    y: r2((frame.y || 0) + (p.top || 0)),
    w: r2(Math.max(0.1, (frame.w || 0) - (p.left || 0) - (p.right || 0))),
    h: r2(Math.max(0.1, (frame.h || 0) - (p.top || 0) - (p.bottom || 0)))
  }
}

function normalizeTableSizing(artifact) {
  const cols = artifact.column_widths || []
  const rows = artifact.row_heights || []
  const totalW = cols.reduce((s, n) => s + (+n || 0), 0)
  const totalH = rows.reduce((s, n) => s + (+n || 0), 0)
  const colCount = Math.max(1, (artifact.headers || []).length || cols.length)
  const rowCount = Math.max(1, ((artifact.rows || []).length + 1) || rows.length)

  artifact.column_widths = totalW > 0
    ? cols.map(v => r2((+v || 0) * artifact.w / totalW))
    : Array.from({ length: colCount }, () => r2(artifact.w / colCount))

  artifact.row_heights = totalH > 0
    ? rows.map(v => r2((+v || 0) * artifact.h / totalH))
    : Array.from({ length: rowCount }, () => r2(artifact.h / rowCount))

  const ts = artifact.table_style || {}
  const density = rowCount * colCount
  const bodySize = density > 30 ? 8.5 : density > 20 ? 9 : (ts.body_font_size || 10)
  artifact.table_style = {
    ...ts,
    body_font_size: Math.max(8, bodySize),
    header_font_size: Math.max(9, Math.min(ts.header_font_size || 11, bodySize + 1))
  }
}

function enforceArtifactBounds(zone) {
  const inner = getZoneInnerBounds(zone)
  zone.artifacts = (zone.artifacts || []).map(artifact => {
    const a = { ...artifact }

    if (['insight_text', 'chart', 'table'].includes(a.type)) {
      Object.assign(a, rectWithin(a, inner))
    }

    if (a.type === 'insight_text') {
      if (a.insight_mode === 'grouped') {
        // For grouped mode: scale bullet font based on total bullet count across all groups
        const groups = a.groups || []
        const totalBullets = groups.reduce((s, g) => s + (g.bullets || []).length, 0)
        const maxBullets = groups.reduce((m, g) => Math.max(m, (g.bullets || []).length), 0)
        const baseSize = (a.bullet_style || {}).font_size || 10
        const fitted = totalBullets > 20 || maxBullets > 6 ? Math.max(8, baseSize - 2)
          : totalBullets > 12 || maxBullets > 4 ? Math.max(8.5, baseSize - 1)
          : baseSize
        a.bullet_style = { ...(a.bullet_style || {}), font_size: fitted }
        // Clamp spacing proportional to artifact dimensions **” no hardcoded defaults
        const dimForGap = a.group_layout === 'rows' ? (a.h || 5) : (a.w || 10)
        const minGap = r2(Math.max(dimForGap * 0.01, 0.04))
        const maxGap = r2(Math.min(dimForGap * 0.03, 0.18))
        a.group_gap_in = r2(Math.min(maxGap, Math.max(minGap, a.group_gap_in || dimForGap * 0.015)))
        const f = (a.bullet_style || {}).font_size || 10
        a.header_to_box_gap_in = r2(Math.min(0.10, Math.max(f * 0.5 / 72, 0.02)))
      } else {
        const pointCount = (a.points || []).length
        const avgChars = pointCount ? Math.round((a.points || []).join(' ').length / pointCount) : 0
        const baseSize = (a.body_style || {}).font_size || 10
        const fitted = pointCount > 6 || avgChars > 120 ? Math.max(8, baseSize - 2)
          : pointCount > 4 || avgChars > 80 ? Math.max(8.5, baseSize - 1)
          : baseSize
        a.body_style = { ...(a.body_style || {}), font_size: fitted }
      }
    }

    if (a.type === 'cards') {
      // Use the artifact's own x/y/w/h as the authoritative container (respects allocated position/size
      // within the zone), clamped within zone inner bounds. Ignore a.container **” LLM sets it inconsistently.
      const artBounds = (a.x != null && a.y != null && a.w != null && a.h != null)
        ? { x: a.x, y: a.y, w: a.w, h: a.h }
        : inner
      const container = rectWithin(artBounds, inner)
      a.container = container
      // Always recompute card_frames to fill the full container **” never trust LLM-output frames
      const _cards    = a.cards || []
      const _count    = _cards.length
      const _cs       = a.card_style || {}
      const _gap      = _cs.gap || 0.12
      const _ax       = container.x
      const _ay       = container.y
      const _aw       = container.w
      const _ah       = container.h
      const _layout   = String(a.cards_layout || '').toLowerCase()
      const _aspect   = _ah > 0 ? _aw / _ah : 1
      const _rowCW    = _count > 0 ? (_aw - _gap * (_count - 1)) / Math.max(_count, 1) : _aw
      const _colCH    = _count > 0 ? (_ah - _gap * (_count - 1)) / Math.max(_count, 1) : _ah
      const _gridCols = _count > 1 ? 2 : 1
      const _gridRows = Math.ceil(_count / Math.max(_gridCols, 1))
      const _gridCW   = (_aw - _gap * (_gridCols - 1)) / Math.max(_gridCols, 1)
      const _gridCH   = (_ah - _gap * (_gridRows - 1)) / Math.max(_gridRows, 1)
      const minCW     = 1.45
      const minCH     = 1.10
      let _l = _layout
      if (!['row', 'column', 'grid'].includes(_l)) {
        if (_count <= 1)      _l = 'row'
        else if (_count <= 3) _l = _aspect >= 1 ? 'row' : 'column'
        else if (_count === 4) _l = 'grid'
        else                  _l = _aspect >= 1.15 ? 'row' : 'grid'
      }
      if (_l === 'row'    && _rowCW  < minCW)  _l = _count <= 3 ? 'column' : 'grid'
      if (_l === 'grid'   && (_gridCW < minCW || _gridCH < minCH)) _l = _count <= 3 ? 'column' : (_rowCW >= minCW && _aspect >= 1.15 ? 'row' : 'grid')
      if (_l === 'column' && _colCH  < minCH && _rowCW >= minCW)   _l = 'row'
      const _frames = []
      if (_l === 'row') {
        const cw = r2((_aw - _gap * (_count - 1)) / Math.max(_count, 1))
        for (let i = 0; i < _count; i++) _frames.push({ x: r2(_ax + i * (cw + _gap)), y: r2(_ay), w: cw, h: r2(_ah) })
      } else if (_l === 'column') {
        const ch = r2((_ah - _gap * (_count - 1)) / Math.max(_count, 1))
        for (let i = 0; i < _count; i++) _frames.push({ x: r2(_ax), y: r2(_ay + i * (ch + _gap)), w: r2(_aw), h: ch })
      } else {
        const cols = _count > 1 ? 2 : 1
        const rows = Math.ceil(_count / cols)
        const cw   = r2((_aw - _gap * (cols - 1)) / Math.max(cols, 1))
        const ch   = r2((_ah - _gap * (rows - 1)) / Math.max(rows, 1))
        for (let i = 0; i < _count; i++) {
          _frames.push({ x: r2(_ax + (i % cols) * (cw + _gap)), y: r2(_ay + Math.floor(i / cols) * (ch + _gap)), w: cw, h: ch })
        }
      }
      a.cards_layout = _l
      a.card_frames  = _frames
      const longestBody = Math.max(0, ...(a.cards || []).map(c => String(c.body || '').length))
      const bodyBase = (a.body_style || {}).font_size || 10
      a.body_style = { ...(a.body_style || {}), font_size: longestBody > 180 ? Math.max(8, bodyBase - 1.5) : bodyBase }
    }

    if (a.type === 'workflow') {
      const container = rectWithin(a.container || inner, inner)
      a.container = container
      a.nodes = (a.nodes || []).map(node => rectWithin(node, container))
      a.connections = (a.connections || []).map(conn => ({
        ...conn,
        path: (conn.path || []).map(pt => ({
          x: r2(clamp(pt.x || 0, container.x, container.x + container.w)),
          y: r2(clamp(pt.y || 0, container.y, container.y + container.h))
        }))
      }))
      const nodeCount = (a.nodes || []).length
      if (nodeCount >= 5) {
        a.workflow_style = {
          ...(a.workflow_style || {}),
          node_title_font_size: Math.max(8, ((a.workflow_style || {}).node_title_font_size || 11) - 1),
          node_value_font_size: Math.max(7, ((a.workflow_style || {}).node_value_font_size || 10) - 1)
        }
      }
    }

    // NOTE: table sizing is handled exclusively by computeArtifactInternals() **” do not call normalizeTableSizing() here

    return a
  })
  return zone
}


function applyBrandGuidelineOverrides(slide, manifestSlide, brand) {
  if (!slide || !brand) return slide

  const normalized = JSON.parse(JSON.stringify(slide))
  normalized.global_elements = normalized.global_elements || {}
  if (brand.uses_template) {
    normalized.global_elements = {}
  }

  // In layout mode or template title/divider, coordinates are driven by the template.
  // Skip bounds enforcement **” enforcing would corrupt placeholder-derived positions.
  const isLayoutMode = normalized.layout_mode === true ||
    (brand.uses_template && (normalized.slide_type === 'title' || normalized.slide_type === 'divider'))

  if (isLayoutMode) {
    // Just ensure text content is preserved from manifest
    if (manifestSlide?.title && normalized.title_block) {
      normalized.title_block.text = normalized.title_block.text || manifestSlide.title
    }
    if (manifestSlide?.subtitle && normalized.subtitle_block) {
      normalized.subtitle_block.text = normalized.subtitle_block.text || manifestSlide.subtitle
    }
    return normalized
  }

  const slideW = normalized.canvas?.width_in || brand.slide_width_inches || 13.33
  const slideH = normalized.canvas?.height_in || brand.slide_height_inches || 7.5
  const margin = normalized.canvas?.margin || { left: 0.4, right: 0.4, top: 0.15, bottom: 0.3 }
  const contentBounds = {
    x: margin.left || 0.4,
    y: margin.top || 0.15,
    w: slideW - (margin.left || 0.4) - (margin.right || 0.4),
    h: slideH - (margin.top || 0.15) - (margin.bottom || 0.3)
  }

  normalized.title_block = normalized.title_block ? rectWithin(normalized.title_block, contentBounds) : normalized.title_block
  normalized.subtitle_block = normalized.subtitle_block ? rectWithin(normalized.subtitle_block, contentBounds) : normalized.subtitle_block

  normalized.zones = (normalized.zones || []).map(zone => {
    const z = { ...zone, frame: rectWithin(zone.frame || contentBounds, contentBounds) }
    return enforceArtifactBounds(z)
  })

  if (manifestSlide?.title && normalized.title_block) normalized.title_block.text = normalized.title_block.text || manifestSlide.title
  if (manifestSlide?.subtitle && normalized.subtitle_block) normalized.subtitle_block.text = normalized.subtitle_block.text || manifestSlide.subtitle

  return normalized
}

function applyLayoutTitleFrames(slide, layoutName, brand) {
  if (!slide || !layoutName || !brand) return slide
  const layouts = brand.slide_layouts || []
  const layout = layouts.find(l => (l.name || '').toLowerCase() === layoutName.toLowerCase())
    || layouts.find(l => (l.name || '').toLowerCase().includes(layoutName.toLowerCase()))
  if (!layout) return slide

  const placeholders = layout.placeholders || []
  const isTitleType = p => {
    const t = String(p?.type || '').toLowerCase()
    return t === 'title' || t === 'center_title' || t === 'centertitle' || t === 'ctrtitle'
  }
  const titlePh = layout.title_placeholder
    || (layout.master_summary || {}).title_placeholder
    || placeholders.find(isTitleType)
    || placeholders
      .filter(p => String(p?.type || '').toLowerCase() !== 'body')
      .sort((a, b) => {
        const ay = a?.y_in ?? 99
        const by = b?.y_in ?? 99
        if (ay !== by) return ay - by
        return (b?.w_in ?? 0) - (a?.w_in ?? 0)
      })[0]
  if (!titlePh) return slide

  const out = JSON.parse(JSON.stringify(slide))
  const r2 = x => Math.round(x * 100) / 100
  const contentAreas = (layout.placeholders || [])
    .filter(p => p.type === 'body' && (p.h_in || 0) > 0.5)
    .sort((a, b) => {
      const rowA = Math.round((a.y_in || 0) * 2)
      const rowB = Math.round((b.y_in || 0) * 2)
      if (rowA !== rowB) return rowA - rowB
      return (a.x_in || 0) - (b.x_in || 0)
    })
  const topContentY = contentAreas.length ? Math.min(...contentAreas.map(p => p.y_in || 99)) : null
  if (out.title_block) {
    const titleX = r2(titlePh.x_in != null ? titlePh.x_in : out.title_block.x || 0.4)
    const titleY = r2(titlePh.y_in != null ? titlePh.y_in : out.title_block.y || 0.15)
    const titleW = r2(titlePh.w_in != null ? titlePh.w_in : out.title_block.w || 9.2)
    let titleH = r2(titlePh.h_in != null ? titlePh.h_in : out.title_block.h || 0.7)
    if (topContentY != null) {
      titleH = Math.max(0.18, Math.min(titleH, topContentY - titleY - 0.08))
    }
    out.title_block = {
      ...out.title_block,
      x: titleX,
      y: titleY,
      w: titleW,
      h: titleH,
      align: out.title_block.align || 'left',
      valign: out.title_block.valign || 'top',
      wrap: out.title_block.wrap !== false
    }
  }

  if (out.subtitle_block) {
    const subtitleX = titlePh.x_in != null ? titlePh.x_in : (out.subtitle_block.x != null ? out.subtitle_block.x : 0.4)
    const subtitleW = titlePh.w_in != null ? titlePh.w_in : (out.subtitle_block.w != null ? out.subtitle_block.w : 9.2)
    const defaultSubtitleY = (titlePh.y_in || 0.15) + (titlePh.h_in || 0.7) + 0.08
    let subtitleY = out.subtitle_block.y != null ? out.subtitle_block.y : defaultSubtitleY
    let subtitleH = out.subtitle_block.h != null ? out.subtitle_block.h : 0.35
    if (topContentY != null) {
      const maxBottom = topContentY - 0.08
      subtitleY = Math.min(subtitleY, Math.max(defaultSubtitleY, maxBottom - subtitleH))
      subtitleH = Math.max(0.18, Math.min(subtitleH, maxBottom - subtitleY))
    }
    out.subtitle_block = {
      ...out.subtitle_block,
      x: r2(subtitleX),
      y: r2(subtitleY),
      w: r2(subtitleW),
      h: r2(subtitleH),
      align: out.subtitle_block.align || 'left',
      valign: out.subtitle_block.valign || 'top',
      wrap: out.subtitle_block.wrap !== false
    }
  }

  return out
}


// *********************************************************************************************************************************************************************************************************************************************
// CLAUDE-BASED FALLBACK
// When a slide fails (bad parse, missing artifacts, truncation) **” ask Claude to
// build the best possible layout for just that one slide.
// Uses a tight focused prompt: no batch overhead, brand tokens only, one slide.
// *********************************************************************************************************************************************************************************************************************************************


// buildFallbackDesign, manifestZoneArtifactSignature, validateFallbackStructure
// → moved to repair/agent5-fallback.js

function normalizeArtifactType(type, chartType) {
  const t = String(type || '').toLowerCase()
  const ct = String(chartType || '').toLowerCase()
  if (t === 'stat_bar' || t === 'star_bar') return 'stat_bar'
  if (t === 'chart' && (ct === 'stat_bar' || ct === 'star_bar')) return 'stat_bar'
  return type || 'unknown'
}

function normalizeChartSubtype(type, chartType) {
  const t = String(type || '').toLowerCase()
  if (t === 'stat_bar' || t === 'star_bar') return 'stat_bar'
  return chartType || ''
}

function artifactSignatureType(artifact) {
  const normalizedType = normalizeArtifactType(artifact?.type, artifact?.chart_type)
  return normalizedType || 'unknown'
}

function normalizeArtifactDefinition(artifact) {
  if (!artifact || typeof artifact !== 'object') return artifact
  const normalizedType = normalizeArtifactType(artifact.type, artifact.chart_type)
  const normalizedChartType = normalizeChartSubtype(artifact.type, artifact.chart_type)
  if (normalizedType === artifact.type && normalizedChartType === (artifact.chart_type || '')) return artifact
  return {
    ...artifact,
    type: normalizedType,
    ...(normalizedChartType ? { chart_type: normalizedChartType } : {})
  }
}

function getArtifactHeader(artifact) {
  return (
    artifact?.artifact_header ||
    artifact?.insight_header ||
    artifact?.stat_header ||
    artifact?.chart_header ||
    artifact?.table_header ||
    artifact?.comparison_header ||
    artifact?.initiative_header ||
    artifact?.profile_header ||
    artifact?.risk_header ||
    artifact?.workflow_header ||
    artifact?.matrix_header ||
    artifact?.tree_header ||
    artifact?.priority_header ||
    artifact?.heading ||
    ''
  )
}

function syncArtifactHeaderBlock(artifact, headerText) {
  if (!artifact || !headerText) return artifact
  return {
    ...artifact,
    artifact_header: artifact.artifact_header || headerText,
    // Only update text if header_block already exists **” never create one here.
    // Charts render chart_header internally; creating a block would duplicate it.
    // Types that need a guaranteed header_block (profile_card_set, comparison_table,
    // initiative_map, risk_register) call ensureArtifactHeaderBlock explicitly.
    header_block: artifact.header_block
      ? { ...artifact.header_block, text: headerText }
      : artifact.header_block
  }
}

// Creates a header_block if one is absent. Only called for artifact types that
// render their own header separately (not charts/tables that embed it internally).
function ensureArtifactHeaderBlock(artifact, headerText, bt) {
  if (!artifact || !headerText) return artifact
  if (artifact.header_block) return syncArtifactHeaderBlock(artifact, headerText)
  return {
    ...artifact,
    artifact_header: artifact.artifact_header || headerText,
    header_block: {
      text: headerText,
      x: null, y: null, w: null, h: 0.30,
      font_family: (bt && bt.title_font_family) || 'Arial',
      font_size: 11, font_weight: 'semibold',
      color: (bt && bt.primary_color) || '#0078AE',
      style: 'underline',
      accent_color: (bt && bt.primary_color) || '#0078AE'
    }
  }
}

function normalizeWorkflowNodes(nodes) {
  return (Array.isArray(nodes) ? nodes : []).map((node, i) => ({
    id: node?.id || `n${i + 1}`,
    node_label: node?.node_label || node?.label || '',
    primary_message: node?.primary_message || node?.value || '',
    secondary_message: node?.secondary_message || node?.description || '',
    level: node?.level != null ? node.level : 1
  }))
}

function normalizeComparisonTableManifest(artifact) {
  // ****** New flat schema: columns[] (strings) + rows[].cells[{value, subtext, tone}] ******
  if (Array.isArray(artifact?.columns) && artifact.columns.length) {
    const columns  = artifact.columns.map(c => String(c || ''))
    const criteria = columns.slice(1)   // first column is the label column
    const rows     = Array.isArray(artifact?.rows) ? artifact.rows : []
    const toneToRating = t => t === 'positive' ? 'yes' : t === 'negative' ? 'no' : t === 'neutral' ? 'partial' : 'text'
    const options = rows.map((row, ri) => {
      const cells     = Array.isArray(row?.cells) ? row.cells : []
      const labelCell = cells[0] || {}
      const dataCells = cells.slice(1)
      return {
        id:         `row_${ri + 1}`,
        name:       String(labelCell.value || ''),
        badge_text: row?.is_recommended ? String(row?.badge || 'Recommended') : '',
        row_tone:   row?.is_recommended ? 'recommended' : 'neutral',
        ratings:    dataCells.map((cell, ci) => ({
          criterion:           criteria[ci] || '',
          column_id:           criteria[ci] || '',
          rating:              toneToRating(cell?.tone),
          display_value:       cell?.value     != null ? String(cell.value)     : '',
          icon_type:           cell?.icon_type != null ? String(cell.icon_type) : '',
          note:                cell?.subtext   != null ? String(cell.subtext)   : '',
          representation_type: cell?.icon_type ? 'icon' : 'text',
          tonality:            cell?.tone === 'positive' ? 'positive' : cell?.tone === 'negative' ? 'negative' : 'neutral'
        }))
      }
    })
    const recommendedRow = options.find(o => o.row_tone === 'recommended')
    return { criteria, options, recommended_option: recommendedRow?.name || '', recommended_row_id: '' }
  }

}

function normalizeInitiativeMapManifest(artifact) {
  const dimension_labels = (Array.isArray(artifact?.column_headers) ? artifact.column_headers : [])
    .map(c => ({ id: String(c?.id || c?.label || c?.name || ''), label: _displayLabel(c) }))
  // Normalize a tags entry to {label, tone} regardless of agent4 format (string or object)
  const normInitTag = t => typeof t === 'string'
    ? { label: t, tone: 'neutral' }
    : { label: String(t?.label || t?.text || t?.name || ''), tone: String(t?.tone || 'neutral') }
  const initiatives = Array.isArray(artifact?.rows) && artifact.rows.length
    ? artifact.rows.map((row, ri) => ({
        id: row?.id || `initiative_${ri + 1}`,
        name: row?.initiative_name || row?.name || '',
        subtitle: row?.initiative_subtitle || row?.subtitle || '',
        placements: (Array.isArray(row?.cells) ? row.cells : []).map((cell, ci) => ({
          lane_id: cell?.column_id || dimension_labels[ci]?.id || '',
          title: cell?.primary_message || '',
          subtitle: cell?.secondary_message || '',
          tags: Array.isArray(cell?.tags) ? cell.tags.map(normInitTag) : [],
          cell_tone: cell?.cell_tone || ''
        })),
        dimensions: (Array.isArray(row?.cells) ? row.cells : []).map((cell, ci) => ({
          label: dimension_labels.find(h => String(h.id) === String(cell?.column_id || ''))?.label || dimension_labels[ci]?.label || '',
          lane_id: cell?.column_id || dimension_labels[ci]?.id || '',
          value: cell?.primary_message || '',
          subtitle: cell?.secondary_message || '',
          tags: Array.isArray(cell?.tags) ? cell.tags.map(normInitTag) : [],
          cell_tone: cell?.cell_tone || ''
        }))
      }))
    : []
  return { dimension_labels, initiatives }
}

function normalizeRiskRegisterManifest(artifact) {
  // New schema: severity_levels[]
  if (Array.isArray(artifact?.severity_levels) && artifact.severity_levels.length) {
    const severityLevels = artifact.severity_levels.map((lvl, li) => ({
      id: lvl.id || `level_${li + 1}`,
      label: lvl.label || '',
      tone: String(lvl.tone || lvl.severity || 'medium').toLowerCase(),
      pip_levels: typeof lvl.pip_levels === 'number' ? Math.max(1, Math.round(lvl.pip_levels)) : 5,
      item_details: (Array.isArray(lvl.item_details) ? lvl.item_details : []).map(item => ({
        primary_message: item.primary_message || item.risk_title || item.title || '',
        secondary_message: item.secondary_message || item.risk_detail || item.detail || '',
        tags: (Array.isArray(item.tags) ? item.tags : []).map(t => ({
          value: String(t.value || t.label || ''),
          tone: String(t.tone || 'neutral').toLowerCase()
        })),
        pips: (Array.isArray(item.pips) ? item.pips : []).map(p => ({
          label: String(p.label || p.value || ''),
          intensity: typeof p.intensity === 'number' ? p.intensity : p.intensity
        }))
      }))
    }))
    return { severity_levels: severityLevels }
  }
  return { severity_levels: [] }
}

function normalizeMatrixManifest(artifact) {
  const quadrants = (Array.isArray(artifact?.quadrants) ? artifact.quadrants : []).map((q, i) => ({
    id: q?.id || `q${i + 1}`,
    name: q?.title || q?.name || '',
    primary_message: q?.primary_message || '',
    tone: String(q?.tone || 'neutral').toLowerCase()
  }))
  const points = (Array.isArray(artifact?.points) ? artifact.points : []).map((pt) => ({
    label: pt?.label || '',
    short_label: pt?.short_label || '',
    quadrant_id: pt?.quadrant_id || '',
    x: typeof pt?.x === 'number' ? pt.x : 50,
    y: typeof pt?.y === 'number' ? pt.y : 50,
    emphasis: pt?.emphasis || 'medium'
  }))
  return { quadrants, points }
}

function normalizeDriverTreeNode(node) {
  if (!node || typeof node !== 'object') return { label: '', value: '', description: '' }
  return {
    ...node,
    label: node?.label || '',
    value: node?.value || '',
    description: node?.description || ''
  }
}

function normalizeDriverTreeManifest(artifact) {
  const root = normalizeDriverTreeNode(artifact?.root)
  const branches = (Array.isArray(artifact?.branches) ? artifact.branches : []).map((branch) => ({
    ...normalizeDriverTreeNode(branch),
    children: (Array.isArray(branch?.children) ? branch.children : []).map(child => normalizeDriverTreeNode(child))
  }))
  return { root, branches }
}

function makeHeaderBlockFromManifestArtifact(artifact, bt) {
  const text = getArtifactHeader(artifact)
  if (!text) return null
  return {
    text,
    x: null, y: null, w: null, h: 0.3,
    font_family: bt.title_font_family || 'Arial',
    font_size: 11,
    font_weight: 'semibold',
    color: bt.primary_color || '#0078AE',
    style: 'underline',
    accent_color: bt.primary_color || '#0078AE'
  }
}

function buildSafeArtifactShell(manifestArt, bt) {
  const t = normalizeArtifactType(manifestArt?.type, manifestArt?.chart_type) || 'insight_text'
  const header_block = makeHeaderBlockFromManifestArtifact(manifestArt, bt)
  const artifact_coverage_hint = manifestArt?.artifact_coverage_hint
  const artifact_header = getArtifactHeader(manifestArt)
  if (t === 'stat_bar') {
    return {
      type: 'stat_bar',
      artifact_coverage_hint,
      x: null, y: null, w: null, h: null,
      artifact_header,
      stat_header: manifestArt?.stat_header || artifact_header || manifestArt?.chart_header || '',
      stat_decision: manifestArt?.stat_decision || manifestArt?.chart_insight || '',
      column_headers: Array.isArray(manifestArt?.column_headers) ? manifestArt.column_headers : [],
      scale_UL: manifestArt?.scale_UL ?? null,
      rows: Array.isArray(manifestArt?.rows) ? manifestArt.rows : [],
      annotation_style: manifestArt?.annotation_style || 'trailing',
      header_block
    }
  }
  if (t === 'chart') {
    const palette = bt.chart_palette || bt.accent_colors || ['#1A3C8F', '#E8A020', '#2E9E5B', '#C82333']
    const chartType = normalizeChartSubtype(manifestArt?.type, manifestArt?.chart_type) || 'bar'
    const isPie = chartType === 'pie' || chartType === 'donut'
    const isGroupPie = chartType === 'group_pie'
    const seriesArr = Array.isArray(manifestArt?.series) ? manifestArt.series : []
    const categories = Array.isArray(manifestArt?.categories) ? manifestArt.categories : []
    const defaultLabelColor = bt.body_color || bt.primary_color || '#111111'
    // group_pie and pie both use one series_style entry PER SLICE (category), not per entity/series
    const autoSeriesStyle = (isPie || isGroupPie)
      ? categories.map((cat, i) => ({
          series_name: String(cat || ''), fill_color: palette[i % palette.length],
          border_color: null, border_width: 0, data_label_color: defaultLabelColor, data_label_size: 9
        }))
      : (seriesArr.length > 0 ? seriesArr : [{ name: '' }]).map((s, i) => ({
          series_name: s.name || '', fill_color: palette[i % palette.length],
          border_color: null, border_width: 0, data_label_color: defaultLabelColor, data_label_size: 9
        }))
    return {
      type: 'chart',
      artifact_coverage_hint,
      artifact_header,
      x: null, y: null, w: null, h: null,
      chart_type:       chartType,
      rows:             Array.isArray(manifestArt?.rows) ? manifestArt.rows : [],
      annotation_style: manifestArt?.annotation_style || 'trailing',
      categories:       categories,
      series:           seriesArr,
      chart_title:      manifestArt?.chart_title  || '',
      chart_header:     manifestArt?.chart_header || artifact_header || manifestArt?.stat_header || '',
      chart_insight:    manifestArt?.chart_insight || manifestArt?.stat_decision || '',
      show_data_labels: manifestArt?.show_data_labels !== false,
      // combo charts always need a legend to distinguish bar vs line series
      show_legend:      chartType === 'combo' ? true : !!(manifestArt?.show_legend),
      x_label:          manifestArt?.x_label || '',
      y_label:          manifestArt?.y_label || '',
      secondary_y_label: manifestArt?.secondary_y_label || '',
      chart_style: {
        title_font_family: bt.title_font_family || 'Arial',
        title_font_size: 12,
        axis_font_family: bt.body_font_family || 'Arial',
        axis_font_size: 9,
        label_font_family: bt.body_font_family || 'Arial',
        label_font_size: 9,
        title_color: bt.primary_color || '#0078AE',
        axis_color: bt.body_color || '#111111',
        gridline_color: '#DDDDDD',
        legend_font_family: bt.body_font_family || 'Arial',
        legend_font_size: 9,
        legend_color: bt.body_color || '#111111',
        show_gridlines: false,
        show_border: false,
        border_color: null,
        background_color: null,
        legend_position: 'top',
        data_label_size: 9,
        category_label_rotation: 0
      },
      series_style: autoSeriesStyle,
      header_block
    }
  }
  if (t === 'table') {
    return {
      type: 'table',
      artifact_coverage_hint,
      artifact_header,
      x: null, y: null, w: null, h: null,
      table_header:    manifestArt?.table_header  || artifact_header || '',
      headers:         Array.isArray(manifestArt?.headers)        ? manifestArt.headers        : [],
      rows:            Array.isArray(manifestArt?.rows)           ? manifestArt.rows           : [],
      highlight_rows:  Array.isArray(manifestArt?.highlight_rows) ? manifestArt.highlight_rows : [],
      note:            manifestArt?.note || '',
      table_style: {
        header_fill_color: bt.primary_color || '#0078AE',
        header_text_color: '#FFFFFF',
        header_font_family: bt.title_font_family || 'Arial',
        header_font_size: 10,
        body_fill_color: '#FFFFFF',
        body_alt_fill_color: '#F7F8FA',
        body_text_color: bt.body_color || '#111111',
        body_font_family: bt.body_font_family || 'Arial',
        body_font_size: 9,
        grid_color: '#D7DEE8',
        grid_width: 0.5,
        highlight_fill_color: '#FFF4BF',
        cell_padding: 0.06
      },
      column_widths: [],
      row_heights: [],
      header_row_height: null,
      column_types: [],
      column_alignments: [],
      header_block
    }
  }
  if (t === 'comparison_table') {
    const normalized = normalizeComparisonTableManifest(manifestArt)
    // Preserve the label column header from the flat schema (columns[0]) for the renderer
    const labelColumnHeader = Array.isArray(manifestArt?.columns) && manifestArt.columns.length
      ? String(manifestArt.columns[0] || 'Option')
      : String((manifestArt?.column_headers || [])[0]?.label || 'Option')
    return {
      type: 'comparison_table',
      artifact_coverage_hint,
      artifact_header,
      x: null, y: null, w: null, h: null,
      comparison_header: manifestArt?.comparison_header || artifact_header || manifestArt?.table_header || '',
      _label_column_header: labelColumnHeader,
      criteria: normalized.criteria,
      options: normalized.options,
      recommended_option: normalized.recommended_option,
      comparison_style: {
        container_fill_color: '#FFFFFF',
        container_border_color: '#D7DEE8',
        container_border_width: 0.6,
        container_corner_radius: 8,
        header_fill_color: bt.primary_color || '#0078AE',
        header_text_color: '#FFFFFF',
        row_fill_color: '#FFFFFF',
        row_alt_fill_color: '#F7F8FA',
        recommended_fill_color: '#EEF4E2',
        grid_color: '#D7DEE8',
        yes_fill_color: '#E4F2DE',
        partial_fill_color: '#FFF4BF',
        no_fill_color: '#FDE8E8',
        neutral_fill_color: '#F4F5F7',
        label_font_family: bt.title_font_family || 'Arial',
        label_font_size: 10,
        body_font_family: bt.body_font_family || 'Arial',
        body_font_size: 9
      },
      header_block
    }
  }
  if (t === 'initiative_map') {
    const normalized = normalizeInitiativeMapManifest(manifestArt)
    return {
      type: 'initiative_map',
      artifact_coverage_hint,
      artifact_header,
      x: null, y: null, w: null, h: null,
      initiative_header: manifestArt?.initiative_header || artifact_header || manifestArt?.table_header || '',
      dimension_labels: normalized.dimension_labels,
      initiatives: normalized.initiatives,
      initiative_style: {
        row_fill_color: '#FFFFFF',
        row_border_color: '#D7DEE8',
        row_border_width: 0.6,
        row_corner_radius: 8,
        header_fill_color: bt.primary_color || '#0078AE',
        header_text_color: '#FFFFFF',
        label_font_family: bt.title_font_family || 'Arial',
        label_font_size: 10,
        body_font_family: bt.body_font_family || 'Arial',
        body_font_size: 9,
        accent_color: bt.secondary_color || bt.primary_color || '#0078AE'
      },
      header_block
    }
  }
  if (t === 'profile_card_set') {
    return {
      type: 'profile_card_set',
      artifact_coverage_hint,
      artifact_header,
      x: null, y: null, w: null, h: null,
      profile_header: manifestArt?.profile_header || artifact_header || manifestArt?.heading || '',
      profiles: Array.isArray(manifestArt?.profiles) ? manifestArt.profiles : [],
      layout_direction: manifestArt?.layout_direction || 'horizontal',
      profile_style: {
        card_fill_color: '#FFFFFF',
        card_border_color: '#D7DEE8',
        card_border_width: 0.6,
        card_corner_radius: 2,
        header_fill_color: '#EDF4FF',
        header_text_color: bt.primary_color || '#0078AE',
        key_fill_color: '#F4F5F7',
        key_text_color: '#4B5563',
        label_font_family: bt.title_font_family || 'Arial',
        label_font_size: 11,
        body_font_family: bt.body_font_family || 'Arial',
        body_font_size: 9,
        positive_color: '#2D7F5E',
        negative_color: '#C2410C',
        warning_color: '#B45309',
        neutral_color: bt.body_color || '#111111'
      },
      header_block
    }
  }
  if (t === 'risk_register') {
    const normalized = normalizeRiskRegisterManifest(manifestArt)
    return {
      type: 'risk_register',
      artifact_coverage_hint,
      artifact_header,
      x: null, y: null, w: null, h: null,
      risk_header: manifestArt?.risk_header || manifestArt?.table_header || '',
      severity_levels: normalized.severity_levels,
      risk_style: {
        // Structural defaults **” LLM sets all color/font values via prompt
        row_border_color: null,
        row_border_width: 0.5,
        row_corner_radius: 8,
        // All color fields null **” LLM fills these from brand tokens per prompt instructions
        critical_fill_color: null, high_fill_color: null, medium_fill_color: null, low_fill_color: null,
        critical_badge_color: null, high_badge_color: null, medium_badge_color: null, low_badge_color: null,
        badge_text_color: null,
        critical_text_color: null, high_text_color: null, medium_text_color: null, low_text_color: null,
        critical_pip_fill: null, high_pip_fill: null, medium_pip_fill: null, low_pip_fill: null,
        tag_positive_fill: null, tag_positive_border: null, tag_positive_text: null,
        tag_negative_fill: null, tag_negative_border: null, tag_negative_text: null,
        tag_warning_fill:  null, tag_warning_border:  null, tag_warning_text:  null,
        tag_neutral_fill:  null, tag_neutral_border:  null, tag_neutral_text:  null,
        // Font families from brand tokens (structural, not style)
        label_font_family: bt.title_font_family || 'Arial',
        label_font_size: null,           // LLM decides
        body_font_family: bt.body_font_family || 'Arial',
        body_font_size: null,            // LLM decides
        primary_message_font_size: null, // LLM decides **” consistent across all rows
        secondary_message_font_size: null// LLM decides **” consistent across all rows
      },
      header_block
    }
  }
  if (t === 'workflow') {
    const manifestNodes = normalizeWorkflowNodes(manifestArt?.nodes)
    const manifestConns = Array.isArray(manifestArt?.connections) ? manifestArt.connections : []
    return {
      type: 'workflow',
      artifact_coverage_hint,
      artifact_header,
      x: null, y: null, w: null, h: null,
      flow_direction:  manifestArt?.flow_direction  || 'left_to_right',
      workflow_type:   manifestArt?.workflow_type   || 'process_flow',
      workflow_style: {
        node_fill_color: bt.primary_color || '#0078AE',
        node_border_color: '#FFFFFF',
        node_border_width: 1,
        node_title_font_family: bt.title_font_family || 'Arial',
        node_title_font_size: 10,
        node_title_color: '#FFFFFF',
        node_value_font_family: bt.body_font_family || 'Arial',
        node_value_font_size: 9,
        node_value_color: bt.body_color || '#111111',
        node_inner_padding: 0.08,
        external_label_gap: 0.08,
        connector_color: bt.primary_color || '#0078AE',
        connector_width: 0.5,
        node_corner_radius: 4
      },
      nodes: manifestNodes.map((node, i) => ({
        id: node?.id || `n${i + 1}`,
        label: node?.node_label || node?.label || '',
        value: node?.primary_message || node?.value || '',
        description: node?.secondary_message || node?.description || '',
        level: node?.level != null ? node.level : 1
      })),
      connections: manifestConns.map(conn => ({
        from: conn?.from || '',
        to: conn?.to || '',
        type: conn?.type || 'arrow'
      })),
      container: null,
      header_block
    }
  }
  if (t === 'cards') {
    return {
      type: 'cards',
      artifact_coverage_hint,
      artifact_header,
      x: null, y: null, w: null, h: null,
      cards_layout: manifestArt.cards_layout || 'column',
      container: null,
      card_frames: [],
      card_style: {
        fill_color: '#F5F5F5',
        border_color: '#DDDDDD',
        border_width: 0.75,
        corner_radius: 0,
        shadow: false,
        internal_padding: 0.12
      },
      title_style: {
        font_family: bt.title_font_family || 'Arial',
        font_size: 12,
        font_weight: 'bold',
        color: bt.primary_color || '#0078AE'
      },
      subtitle_style: {
        font_family: bt.body_font_family || 'Arial',
        font_size: 22,
        font_weight: 'bold',
        color: bt.body_color || '#111111'
      },
      body_style: {
        font_family: bt.body_font_family || 'Arial',
        font_size: 9,
        font_weight: 'regular',
        color: bt.body_color || '#111111',
        line_spacing: 1.2
      }
    }
  }
  if (t === 'matrix') {
    const normalized = normalizeMatrixManifest(manifestArt)
    return {
      type: 'matrix',
      artifact_coverage_hint,
      artifact_header,
      x: null, y: null, w: null, h: null,
      matrix_style: {
        // Outer grid
        border_color:    '#D7DEE8',
        border_width:    0.8,
        // Center dividers (thin dashed)
        divider_color:   '#AAAAAA',
        divider_width:   0.5,
        // Quadrant fills **” brand-derived tints at ~10% intensity (computed in _matrixToBlocks via _hexTint)
        positive_quadrant_fill: null,
        negative_quadrant_fill: null,
        neutral_quadrant_fill:  null,
        // Quadrant text colors
        positive_title_color:   bt.primary_color || '#1B5E20',
        negative_title_color:   '#B91C1C',
        neutral_title_color:    bt.body_color || '#374151',
        positive_body_color:    bt.primary_color || '#2D7F5E',
        negative_body_color:    '#B91C1C',
        neutral_body_color:     bt.body_color || '#374151',
        // Point dot fills **” tone-driven
        positive_point_fill:    bt.primary_color || '#2D7F5E',
        negative_point_fill:    '#C53030',
        neutral_point_fill:     bt.secondary_color || '#6B7280',
        // Axis labels
        axis_label_font_family: bt.body_font_family || 'Arial',
        axis_label_font_size:   9,
        axis_label_color:       bt.caption_color || bt.body_color || '#6B7280',
        // Quadrant labels
        quadrant_title_font_family: bt.title_font_family || 'Arial',
        quadrant_title_font_size:   11,
        quadrant_body_font_family:  bt.body_font_family  || 'Arial',
        quadrant_body_font_size:    9,
        // Point label bubble
        point_label_font_family: bt.body_font_family || 'Arial',
        point_label_font_size:   9
      },
      matrix_type: manifestArt?.matrix_type || '2x2',
      matrix_header: manifestArt?.matrix_header || artifact_header || '',
      x_axis: manifestArt?.x_axis || { label: '', low_label: '', high_label: '' },
      y_axis: manifestArt?.y_axis || { label: '', low_label: '', high_label: '' },
      quadrants: normalized.quadrants,
      points: normalized.points,
      header_block
    }
  }
  if (t === 'driver_tree') {
    const normalized = normalizeDriverTreeManifest(manifestArt)
    return {
      type: 'driver_tree',
      artifact_coverage_hint,
      artifact_header,
      x: null, y: null, w: null, h: null,
      tree_style: {},
      tree_header: manifestArt?.tree_header || artifact_header || '',
      root: normalized.root,
      branches: normalized.branches,
      header_block
    }
  }
  if (t === 'prioritization') {
    return {
      type: 'prioritization',
      artifact_coverage_hint,
      artifact_header,
      x: null, y: null, w: null, h: null,
      priority_style: {},
      priority_header: manifestArt?.priority_header || artifact_header || '',
      items: Array.isArray(manifestArt?.items) ? manifestArt.items : [],
      header_block
    }
  }

  const grouped = !!(manifestArt?.groups && manifestArt.groups.length)
  return grouped ? {
    type: 'insight_text',
    artifact_coverage_hint,
    artifact_header,
    insight_mode: 'grouped',
    x: null, y: null, w: null, h: null,
    style: { fill_color: null, border_color: null, border_width: 0, corner_radius: 0 },
    heading_style: { font_family: bt.title_font_family || 'Arial', font_size: 12, font_weight: 'bold', color: bt.primary_color || '#0078AE' },
    group_layout: 'rows',
    group_header_style: { shape: 'rounded_rect', fill_color: bt.primary_color || '#0078AE', text_color: '#FFFFFF', font_family: bt.title_font_family || 'Arial', font_size: 10, font_weight: 'bold', corner_radius: 0.04, w: 1.4, h: 0.28 },
    group_bullet_box_style: { fill_color: null, border_color: '#CCCCCC', border_width: 0.75, corner_radius: 0.04, padding: { top: 0.08, right: 0.1, bottom: 0.08, left: 0.1 } },
    bullet_style: { font_family: bt.body_font_family || 'Arial', font_size: 10, font_weight: 'regular', color: bt.body_color || '#111111', line_spacing: 1.35, indent_inches: 0.1, space_before_pt: 3, char: 'â–¶' },
    group_gap_in: 0.08,
    header_to_box_gap_in: 0.04,
    heading: manifestArt?.heading || artifact_header || manifestArt?.insight_header || '',
    groups: Array.isArray(manifestArt?.groups) ? manifestArt.groups : [],
    sentiment: manifestArt?.sentiment || 'neutral',
    header_block
  } : {
    type: 'insight_text',
    artifact_coverage_hint,
    artifact_header,
    insight_mode: 'standard',
    x: null, y: null, w: null, h: null,
    style: { fill_color: null, border_color: null, border_width: 0, corner_radius: 0 },
    heading_style: { font_family: bt.title_font_family || 'Arial', font_size: 12, font_weight: 'bold', color: bt.primary_color || '#0078AE' },
    body_style: { font_family: bt.body_font_family || 'Arial', font_size: 11, font_weight: 'regular', color: bt.body_color || '#111111', line_spacing: 1.4, indent_inches: 0.15, list_style: 'bullet', space_before_pt: 5, vertical_distribution: 'spread' },
    heading: manifestArt?.heading || artifact_header || manifestArt?.insight_header || 'Key Insight',
    points: Array.isArray(manifestArt?.points) ? manifestArt.points : [],
    sentiment: manifestArt?.sentiment || 'neutral',
    header_block
  }
}


// ********* Minimal safe slide **” only used if the fallback Claude call itself fails ******
// This is the true last resort. It still uses real brand values and real content.
function buildMinimalSafeSlide(manifestSlide, tokens) {
  const w         = r2(tokens.slide_width_inches  || 13.33)
  const h         = r2(tokens.slide_height_inches || 7.50)
  const primary   = (tokens.primary_colors    || ['#1A3C8F'])[0]
  const secondary = (tokens.secondary_colors  || ['#E8A020'])[0]
  const bg        = (tokens.background_colors || ['#FFFFFF'])[0]
  const titleFont = (tokens.title_font   || {}).family || 'Calibri'
  const bodyFont  = (tokens.body_font    || {}).family || 'Calibri'
  const isDark    = ['title', 'divider'].includes(manifestSlide.slide_type)
  const cw        = r2(w - 0.80)

  const fallbackBrandTokens = {
    title_font_family: titleFont,
    body_font_family: bodyFont,
    caption_font_family: bodyFont,
    title_color: isDark ? '#FFFFFF' : primary,
    body_color: isDark ? '#CCDDFF' : '#111111',
    caption_color: '#888888',
    primary_color: primary,
    secondary_color: secondary,
    accent_colors: tokens.accent_colors || [],
    chart_palette: tokens.chart_colors || [primary, secondary, '#2E9E5B', '#C82333']
  }

  const fallbackZones = manifestSlide.slide_type === 'content'
    ? (manifestSlide.zones || []).map((zone, zoneIdx) => ({
        zone_id: zone.zone_id || `z${zoneIdx + 1}` ,
        zone_role: zone.zone_role || (zoneIdx === 0 ? 'primary_proof' : 'supporting_evidence'),
        message_objective: zone.message_objective || manifestSlide.key_message || '',
        narrative_weight: zone.narrative_weight || (zoneIdx === 0 ? 'primary' : 'secondary'),
        frame: null,
        padding: zone.padding || null,
        layout_hint: zone.layout_hint || null,
        artifacts: (zone.artifacts || []).map(art => buildSafeArtifactShell(art, fallbackBrandTokens))
      }))
    : []

  const fallbackSlide = {
    slide_number: manifestSlide.slide_number,
    slide_type: manifestSlide.slide_type || 'content',
    canvas: {
      width_in: w, height_in: h,
      margin: { left: 0.40, right: 0.40, top: 0.15, bottom: 0.30 },
      background: { color: isDark ? primary : bg }
    },
    brand_tokens: {
      title_font_family: fallbackBrandTokens.title_font_family,
      body_font_family: fallbackBrandTokens.body_font_family,
      caption_font_family: fallbackBrandTokens.caption_font_family,
      title_color: fallbackBrandTokens.title_color,
      body_color: fallbackBrandTokens.body_color,
      caption_color: fallbackBrandTokens.caption_color,
      primary_color: fallbackBrandTokens.primary_color,
      secondary_color: fallbackBrandTokens.secondary_color,
      accent_colors: fallbackBrandTokens.accent_colors,
      chart_palette: fallbackBrandTokens.chart_palette
    },
    title_block: {
      text: manifestSlide.title || '',
      x: 0.40, y: 0.15, w: cw,
      h: isDark ? 2.00 : 0.75,
      font_family: titleFont,
      font_size: isDark ? 30 : 18,
      font_weight: 'bold',
      color: isDark ? '#FFFFFF' : primary,
      align: 'left', valign: 'middle', wrap: true
    },
    subtitle_block: manifestSlide.subtitle ? {
      text: manifestSlide.subtitle,
      x: 0.40, y: isDark ? 2.60 : 0.95,
      w: cw, h: 0.45,
      font_family: bodyFont, font_size: 14, font_weight: 'regular',
      color: isDark ? '#BBCCFF' : '#555555',
      align: 'left', valign: 'top', wrap: true
    } : null,
    zones: fallbackZones,
    global_elements: {
      footer: {
        show: true, x: 0.40, y: r2(h - 0.26), w: 3.00, h: 0.20,
        font_family: bodyFont, font_size: 8, color: '#AAAAAA', align: 'left'
      },
      page_number: {
        show: true, x: r2(w - 0.88), y: r2(h - 0.26), w: 0.65, h: 0.20,
        font_family: bodyFont, font_size: 8, color: '#AAAAAA', align: 'right'
      }
    },
    layout_mode: manifestSlide.layout_mode !== false,
    selected_layout_name: manifestSlide.selected_layout_name || manifestSlide.layout_name || '',
    title: manifestSlide.title || '',
    subtitle: manifestSlide.subtitle || '',
    key_message: manifestSlide.key_message || '',
    speaker_note: manifestSlide.speaker_note || '',
    _fallback: true
  }

  // Guarantee Agent 6's contract even on the deepest fallback path:
  // every slide gets a canvas plus non-empty blocks[].
  const framedZones = fallbackSlide.slide_type === 'content'
    ? buildScratchZoneFrames(fallbackSlide.zones || [], fallbackSlide)
    : (fallbackSlide.zones || [])

  if (framedZones.length > 0) {
    computeArtifactInternals(framedZones, fallbackSlide.canvas || {}, fallbackBrandTokens)
    normalizeArtifactHeaderBands(framedZones)
    framedZones.forEach((zone, zi) => {
      ;(zone.artifacts || []).forEach((art, ai) => {
        if (!art._artifact_id) art._artifact_id = 's' + (fallbackSlide.slide_number || '?') + '_z' + zi + '_a' + ai
      })
    })
  }

  fallbackSlide.zones = framedZones
  const fallbackRawBlocks = sanitizeBlocks(flattenToBlocks(
    fallbackSlide,
    fallbackBrandTokens
  ), fallbackSlide)
  fallbackSlide.artifact_groups = groupBlocksByArtifact(fallbackRawBlocks)
  fallbackSlide.zones_summary = framedZones.map(z => ({
    zone_id: z.zone_id,
    zone_role: z.zone_role,
    narrative_weight: z.narrative_weight,
    artifact_types: (z.artifacts || []).map(a => artifactSignatureType(a))
  }))

  return fallbackSlide
}

// CONTENT MERGER
// Injects Agent 4 manifest artifact content into Agent 5 designed artifacts.
// Agent 5 produces layout + style. Agent 4 holds the actual data.
// This merge produces a single self-contained object Agent 6 can render directly.
//
// Matching strategy: zone index first, then zone_id string match as fallback.
// Within a zone: artifact matched by position index (Agent 4 and Agent 5 should
// produce the same number of artifacts per zone in the same order).
//
// Content fields injected per artifact type:
//   insight_text : heading, artifact_header, insight_mode, points[] (standard), groups[] (grouped), sentiment
//   chart        : chart_type, chart_title, chart_insight, x_label, y_label,
//                  categories[], series[], show_data_labels, show_legend
//   cards        : cards[] (title, subtitle, body, sentiment per card)
//   workflow     : workflow_type, flow_direction, workflow_title, workflow_insight,
//                  node_label/primary_message/secondary_message mapped into node labels/values/descriptions/levels, connection from/to/type
//   table        : title, headers[], rows[][], highlight_rows[], note
//   matrix       : matrix_type, artifact_header, x_axis, y_axis, quadrants[id/title/primary_message/tone], points[label/short_label/quadrant_id/x(0-100)/y(0-100)/emphasis]
//   driver_tree  : artifact_header, root, branches[]
//   prioritization: artifact_header, items[]
// *********************************************************************************************************************************************************************************************************************************************

// ********* computeArtifactInternals ************************************************************************************************************************************************
// Post-processes merged zones and fills computed layout/sizing fields on each
// artifact IN PLACE, so generate_pptx.py can act as a pure renderer.
// Called after mergeContentIntoZones (and applyLayoutZoneFrames if used).
// ***************************************************************************************************************************************************************************************************************************************
function computeArtifactInternals(zones, canvas, brandTokens) {
  const round2 = x => Math.round(x * 100) / 100
  const bt = brandTokens || {}

  for (const zone of (zones || [])) {
    const artifacts = zone.artifacts || []
    const frame = zone.frame || {}
    if (artifacts.length === 1 && isValidFrame(frame)) {
      const art = artifacts[0]
      const pad = frame.padding || { top: 0.08, right: 0.08, bottom: 0.08, left: 0.08 }
      const inner = {
        x: round2((frame.x || 0) + (pad.left || 0)),
        y: round2((frame.y || 0) + (pad.top || 0)),
        w: round2(Math.max(0.1, (frame.w || 0) - (pad.left || 0) - (pad.right || 0))),
        h: round2(Math.max(0.1, (frame.h || 0) - (pad.top || 0) - (pad.bottom || 0)))
      }
      if (art.x == null || art.y == null || art.w == null || art.h == null || art.w <= 0 || art.h <= 0) {
        art.x = inner.x
        art.y = inner.y
        art.w = inner.w
        art.h = inner.h
      }
      if (art.type === 'workflow' || art.type === 'cards') {
        art.container = { x: inner.x, y: inner.y, w: inner.w, h: inner.h }
      }
    }

    // ****** 1. Multi-artifact zone stacking ******************************************************************************************************************
    if (artifacts.length >= 2) {
      const needsCompute = artifacts.some(a => a.h == null || a.w == null || a.x == null || a.y == null)
      if (needsCompute) {
        const zx = frame.x || 0
        const zy = frame.y || 0
        const zw = frame.w || 0
        const zh = frame.h || 0
        const gap = 0.12
        const splitHint =
          (Array.isArray(zone.artifact_split_hint) ? zone.artifact_split_hint : null) ||
          (Array.isArray(zone.split_hint) ? zone.split_hint : null) ||
          (Array.isArray((zone.layout_hint || {}).split_hint) ? (zone.layout_hint || {}).split_hint : null)
        const arrangement =
          zone.artifact_arrangement ||
          (zone.layout_hint || {}).artifact_arrangement ||
          'vertical'
        let coverage = artifacts.map(a => {
          const n = parseFloat(a?.artifact_coverage_hint)
          return Number.isFinite(n) && n > 0 ? n : null
        })
        if (coverage.some(v => v == null)) {
          if (splitHint && splitHint.length === artifacts.length) {
            coverage = splitHint.map(v => {
              const n = parseFloat(v)
              return Number.isFinite(n) && n > 0 ? n : 0
            })
          } else if (splitHint && splitHint.length >= 2 && artifacts.length === 2) {
            coverage = splitHint.slice(0, 2).map(v => {
              const n = parseFloat(v)
              return Number.isFinite(n) && n > 0 ? n : 0
            })
          } else {
            coverage = artifacts.map((_, idx) => idx === 0 ? 60 : (40 / Math.max(artifacts.length - 1, 1)))
          }
        }
        let totalCoverage = coverage.reduce((s, v) => s + (v || 0), 0)
        if (totalCoverage <= 0) {
          coverage = artifacts.map((_, idx) => idx === 0 ? 60 : (40 / Math.max(artifacts.length - 1, 1)))
          totalCoverage = coverage.reduce((s, v) => s + (v || 0), 0)
        }
        const fracs = coverage.map(v => (v || 0) / totalCoverage)
        const usableGap = gap * Math.max(artifacts.length - 1, 0)

        if (arrangement === 'horizontal') {
          const availW = Math.max(0.1, zw - usableGap)
          let cursorX = zx
          artifacts.forEach((art, i) => {
            const isLast = i === artifacts.length - 1
            const artW = isLast ? round2(zx + zw - cursorX) : round2(availW * fracs[i])
            art.x = round2(cursorX)
            art.y = round2(zy)
            art.w = round2(Math.max(0.1, artW))
            art.h = round2(zh)
            cursorX = round2(cursorX + art.w + gap)
          })
        } else {
          const availH = Math.max(0.1, zh - usableGap)
          let cursorY = zy
          artifacts.forEach((art, i) => {
            const isLast = i === artifacts.length - 1
            const artH = isLast ? round2(zy + zh - cursorY) : round2(availH * fracs[i])
            art.x = round2(zx)
            art.y = round2(cursorY)
            art.w = round2(zw)
            art.h = round2(Math.max(0.1, artH))
            cursorY = round2(cursorY + art.h + gap)
          })
        }
      }
    }

    // ****** Per-artifact computed fields ******************************************************************************************************************************
    for (const art of artifacts) {
      const artType = art.type

      // ****** 2. Chart: _computed sub-object ***************************************************************************************************************
      if (artType === 'chart') {
        if (!art._computed) art._computed = {}
        const computed = art._computed
        const canvasW = (canvas && canvas.width_in) ? canvas.width_in : (bt.slide_width_inches  || 13.33)
        const canvasH = (canvas && canvas.height_in) ? canvas.height_in : (bt.slide_height_inches || 7.50)
        const cs = art.chart_style || {}

        // legend_position
        // combo charts always need a legend **” force show_legend true here as a safety net
        if (art.chart_type === 'combo') art.show_legend = true
        if (art.show_legend) {
          if (art.chart_type === 'group_pie') {
            // group_pie always uses a single shared legend at the top
            computed.legend_position = 'top'
          } else if (art.chart_type === 'combo') {
            // combo: legend always at top to label bar vs line series
            computed.legend_position = 'top'
          } else {
            const widthRatio = (art.w || 0) / Math.max(canvasW, 0.1)
            const heightRatio = (art.h || 0) / Math.max(canvasH, 0.1)
            if (heightRatio > 0.60) computed.legend_position = 'top'
            else if (widthRatio > 0.60) computed.legend_position = 'right'
            else computed.legend_position = (art.chart_type === 'pie') ? 'right' : 'top'
          }
        } else {
          computed.legend_position = 'none'
        }

        const headerFs = ((art.header_block || {}).font_size) || cs.title_font_size || 11
        const maxLegendFs = Math.max(8, Math.min(headerFs - 1, 9))

        // data_label_size
        if (computed.data_label_size == null) {
          const base_size = cs.data_label_size || 9
          const n_cats = (art.categories || []).length || 1
          const density = Math.min(art.w || 0, art.h || 0) / n_cats
          const scale = Math.max(0.55, Math.min(1.0, density / 0.6))
          const computedSize = Math.round(base_size * scale)
          computed.data_label_size = Math.max(computedSize, Math.round(base_size * 0.55))
        }

        // category_label_rotation
        if (computed.category_label_rotation == null) {
          computed.category_label_rotation = (art.categories || []).length > 6 ? -45 : 0
        }

        // Persist computed readability choices on the artifact itself so
        // validation and downstream consumers see a complete chart spec.
        art.legend_position = computed.legend_position
        art.data_label_size = computed.data_label_size
        art.category_label_rotation = computed.category_label_rotation
        art.chart_style = {
          ...cs,
          legend_font_size: Math.min(cs.legend_font_size || maxLegendFs, maxLegendFs),
          legend_position: computed.legend_position,
          data_label_size: computed.data_label_size,
          category_label_rotation: computed.category_label_rotation
        }

        // Auto-repair series_style if missing or empty **” prevents criticalRenderIssues
        if (!art.series_style || art.series_style.length === 0) {
          const palette = bt.chart_palette || bt.accent_colors || ['#1A3C8F', '#E8A020', '#2E9E5B', '#C82333']
          const isPie = art.chart_type === 'pie' || art.chart_type === 'donut'
          const isGroupPie = art.chart_type === 'group_pie'
          if (isPie || isGroupPie) {
            // pie and group_pie both color per-slice (category) not per-series
            art.series_style = (art.categories || []).map((cat, i) => ({
              series_name: String(cat || ''),
              fill_color: palette[i % palette.length],
              border_color: null, border_width: 0,
              data_label_color: bt.body_color || bt.primary_color || '#111111',
              data_label_size: art.chart_style.data_label_size || 9
            }))
          } else {
            const seriesArr = art.series && art.series.length > 0
              ? art.series
              : [{ name: '' }]
            art.series_style = seriesArr.map((s, i) => ({
              series_name: s.name || '',
              fill_color: palette[i % palette.length],
              border_color: null, border_width: 0,
              data_label_color: bt.body_color || bt.primary_color || '#111111',
              data_label_size: art.chart_style.data_label_size || 9
            }))
          }
        }
      }

      // ****** 3. Table: column and row specs ***************************************************************************************************************
      if (artType === 'table') {
        const headers = art.headers || []
        const rows    = art.rows    || []
        const nCols   = headers.length

        if (nCols > 0) {
          const artW = art.w || 6
          const artH = art.h || 2
          const ts = art.table_style || {}
          const cellPadding = ts.cell_padding != null ? ts.cell_padding : 0.06
          const isNumericLike = (value) => /^[\sâ‚¹$â‚¬Â£Â¥\-+]?[\d,\.]+[%KMBcr\s]*$/i.test(String(value == null ? '' : value).trim())
          const countWrappedLines = (text, widthIn, fontSizePt) => {
            const raw = String(text == null ? '' : text).trim()
            if (!raw) return 1
            const usableWidth = Math.max(0.18, widthIn - cellPadding * 2)
            const charsPerLine = Math.max(4, Math.floor((usableWidth * 72) / (Math.max(7, fontSizePt) * 0.52)))
            let lines = 0
            for (const chunk of raw.split('\n')) {
              const words = String(chunk || '').trim().split(/\s+/).filter(Boolean)
              if (!words.length) {
                lines += 1
                continue
              }
              let lineLen = 0
              let chunkLines = 1
              for (const word of words) {
                const nextLen = lineLen === 0 ? word.length : lineLen + 1 + word.length
                if (nextLen <= charsPerLine) lineLen = nextLen
                else {
                  chunkLines += 1
                  lineLen = word.length
                }
              }
              lines += chunkLines
            }
            return Math.max(1, lines)
          }
          const lineHeightIn = (fontSizePt, factor) => (Math.max(7, fontSizePt) * factor) / 72
          const widthWeightForColumn = (ci) => {
            const headerLen = String(headers[ci] || '').length
            let maxLen = headerLen
            let avgLen = headerLen
            let samples = 1
            for (const row of rows) {
              if (ci < row.length) {
                const len = String(row[ci] || '').length
                maxLen = Math.max(maxLen, len)
                avgLen += len
                samples += 1
              }
            }
            avgLen = avgLen / Math.max(samples, 1)
            const colType = (art.column_types || [])[ci]
            const numericHits = rows.filter(row => ci < row.length && isNumericLike(row[ci])).length
            const isNumeric = colType === 'numeric' || (rows.length > 0 && numericHits / Math.max(rows.length, 1) > 0.6 && ci > 0)
            if (isNumeric) return Math.max(8, Math.min(20, avgLen * 0.7 + maxLen * 0.2))
            const firstColBoost = ci === 0 ? 1.18 : 1.0
            return Math.max(10, Math.min(42, (avgLen * 0.8 + maxLen * 0.45) * firstColBoost))
          }

          // column_widths
          if (!art.column_widths || art.column_widths.length === 0) {
            const weights = []
            for (let c = 0; c < nCols; c++) {
              weights.push(widthWeightForColumn(c))
            }
            const totalWeight = weights.reduce((s, w) => s + w, 0) || 1
            const colWidths = weights.map(w => round2(artW * w / totalWeight))
            // Fix rounding remainder on last column
            const widthSum = colWidths.reduce((s, w) => s + w, 0)
            colWidths[nCols - 1] = round2(colWidths[nCols - 1] + (artW - widthSum))
            art.column_widths = colWidths
          }

          // column_types
          if (!art.column_types) {
            const numPat = /^[\d,\.\%â‚¹\$\-\+]+$/
            const types = []
            for (let c = 0; c < nCols; c++) {
              if (c === 0) {
                types.push('text')
              } else {
                const hits = rows.filter(row => c < row.length && numPat.test(String(row[c] || '').trim())).length
                types.push((hits / Math.max(rows.length, 1)) > 0.5 ? 'numeric' : 'text')
              }
            }
            art.column_types = types
          }

          // column_alignments
          if (!art.column_alignments) {
            art.column_alignments = (art.column_types || []).map(t => t === 'numeric' ? 'right' : 'left')
          }

          // Content-aware font sizing + row heights
          const colWidths = art.column_widths || []
          let headerFs = ts.header_font_size || 10
          let bodyFs = ts.body_font_size || 9
          let headerRowHeight = 0.35
          let rowHeights = []
          let fitFound = false

          for (let attempt = 0; attempt < 6; attempt++) {
            const headerLines = headers.map((hdr, ci) => countWrappedLines(hdr, colWidths[ci] || (artW / Math.max(nCols, 1)), headerFs))
            const maxHeaderLines = headerLines.reduce((m, v) => Math.max(m, v), 1)
            headerRowHeight = round2(Math.max(0.35, maxHeaderLines * lineHeightIn(headerFs, 1.18) + cellPadding * 2 + 0.04))

            rowHeights = rows.map(row => {
              let maxLines = 1
              for (let ci = 0; ci < nCols; ci++) {
                const cellText = ci < row.length ? row[ci] : ''
                const cellLines = countWrappedLines(cellText, colWidths[ci] || (artW / Math.max(nCols, 1)), bodyFs)
                maxLines = Math.max(maxLines, cellLines)
              }
              return round2(Math.max(0.26, maxLines * lineHeightIn(bodyFs, 1.22) + cellPadding * 2 + 0.03))
            })

            const totalH = headerRowHeight + rowHeights.reduce((s, rh) => s + rh, 0)
            if (totalH <= artH + 0.01) {
              fitFound = true
              break
            }

            if (bodyFs > 8) bodyFs -= 0.5
            else if (headerFs > 8) headerFs -= 0.5
            else break
          }

          let finalEstimatedTotalH = headerRowHeight + rowHeights.reduce((s, rh) => s + rh, 0)
          if (!fitFound) {
            const totalH = finalEstimatedTotalH
            if (totalH > artH && totalH > 0) {
              const scale = artH / totalH
              headerRowHeight = round2(Math.max(0.30, headerRowHeight * scale))
              rowHeights = rowHeights.map(rh => round2(Math.max(0.24, rh * scale)))
            }
          }
          finalEstimatedTotalH = headerRowHeight + rowHeights.reduce((s, rh) => s + rh, 0)
          art._table_fit_failed = !fitFound && finalEstimatedTotalH > artH + 0.02
          art.table_style = {
            ...ts,
            header_font_size: headerFs,
            body_font_size: bodyFs,
            cell_padding: cellPadding
          }
          art.row_heights = rowHeights
          art.header_row_height = headerRowHeight

          // Explicit table grid geometry for Agent 6.
          const tableX = art.x || 0
          const tableY = art.y || 0
          const colWs = art.column_widths || []
          const dataRowHs = art.row_heights || []
          const headerH = art.header_row_height != null ? art.header_row_height : 0.35

          let curX = tableX
          art.column_x_positions = colWs.map(cw => {
            const x = round2(curX)
            curX += (+cw || 0)
            return x
          })

          let curY = tableY
          art.row_y_positions = [round2(curY)]
          curY += (+headerH || 0)
          for (const rh of dataRowHs) {
            art.row_y_positions.push(round2(curY))
            curY += (+rh || 0)
          }

          art.header_cell_frames = colWs.map((cw, ci) => ({
            col_index: ci,
            x: round2(art.column_x_positions[ci] || tableX),
            y: round2(tableY),
            w: round2(+cw || 0),
            h: round2(+headerH || 0)
          }))

          art.body_cell_frames = dataRowHs.map((rh, ri) =>
            colWs.map((cw, ci) => ({
              row_index: ri,
              col_index: ci,
              x: round2(art.column_x_positions[ci] || tableX),
              y: round2(art.row_y_positions[ri + 1] || tableY),
              w: round2(+cw || 0),
              h: round2(+rh || 0)
            }))
          )
        }
      }

      // ****** 4. Cards: pre-compute card_frames ******************************************************************************************************
      if (artType === 'cards') {
        if (!art.card_frames || art.card_frames.length === 0) {
          const cards  = art.cards  || []
          const requestedLayout = String(art.cards_layout || art.layout || '').toLowerCase()
          const cs     = art.card_style || {}
          const gap    = cs.gap              || 0.12
          const count  = cards.length
          const ax     = art.x || 0
          const ay     = art.y || 0
          const aw     = art.w || 0
          let ah       = art.h || 0
          // Note: do NOT shrink art.h here **” enforceArtifactBounds() recomputes frames
          // from the authoritative zone container and always fills the full allocated area.
          const aspect = ah > 0 ? aw / ah : 1
          const minReadableCardWidth = 1.45
          const minReadableCardHeight = 1.10
          const rowCardWidth = count > 0 ? (aw - gap * (count - 1)) / Math.max(count, 1) : aw
          const columnCardHeight = count > 0 ? (ah - gap * (count - 1)) / Math.max(count, 1) : ah
          const gridCols = count > 1 ? 2 : 1
          const gridRows = Math.ceil(count / Math.max(gridCols, 1))
          const gridCardWidth = (aw - gap * (gridCols - 1)) / Math.max(gridCols, 1)
          const gridCardHeight = (ah - gap * (gridRows - 1)) / Math.max(gridRows, 1)

          let layout = requestedLayout
          if (!['row', 'column', 'grid'].includes(layout)) {
            if (count <= 1) layout = 'row'
            else if (count === 2) layout = aspect >= 1 ? 'row' : 'column'
            else if (count === 3) layout = aspect >= 1 ? 'row' : 'column'
            else if (count === 4) layout = 'grid'
            else layout = aspect >= 1.15 ? 'row' : 'grid'
          }

          // Readability override: never keep a horizontal row when it makes KPI cards too narrow.
          if (layout === 'row' && rowCardWidth < minReadableCardWidth) {
            layout = count <= 3 ? 'column' : 'grid'
          }
          if (layout === 'grid' && (gridCardWidth < minReadableCardWidth || gridCardHeight < minReadableCardHeight)) {
            if (count <= 3) layout = 'column'
            else if (rowCardWidth >= minReadableCardWidth && aspect >= 1.15) layout = 'row'
          }
          if (layout === 'column' && columnCardHeight < minReadableCardHeight && rowCardWidth >= minReadableCardWidth) {
            layout = 'row'
          }

          const frames = []
          if (layout === 'row') {
            const card_w = round2((aw - gap * (count - 1)) / Math.max(count, 1))
            for (let i = 0; i < count; i++) {
              frames.push({ x: round2(ax + i * (card_w + gap)), y: round2(ay), w: card_w, h: round2(ah) })
            }
          } else if (layout === 'column') {
            const card_h = round2((ah - gap * (count - 1)) / Math.max(count, 1))
            for (let i = 0; i < count; i++) {
              frames.push({ x: round2(ax), y: round2(ay + i * (card_h + gap)), w: round2(aw), h: card_h })
            }
          } else {
            // grid or default
            const cols        = count > 1 ? 2 : 1
            const rows_count  = Math.ceil(count / cols)
            const card_w      = round2((aw - gap * (cols - 1)) / Math.max(cols, 1))
            const card_h      = round2((ah - gap * (rows_count - 1)) / Math.max(rows_count, 1))
            for (let i = 0; i < count; i++) {
              frames.push({
                x: round2(ax + (i % cols) * (card_w + gap)),
                y: round2(ay + Math.floor(i / cols) * (card_h + gap)),
                w: card_w,
                h: card_h
              })
            }
          }
          art.cards_layout = layout
          art.card_frames = frames
        }
      }

      if (artType === 'workflow') {
        const nodes = Array.isArray(art.nodes) ? art.nodes : []
        const ws = art.workflow_style || {}
        art.workflow_style = {
          // Level-1 (root / all process_flow nodes) **” brand primary
          node_fill_color:           ws.node_fill_color           || bt.primary_color   || '#0078AE',
          // Level-2 nodes (hierarchy/decomposition children) **” secondary brand or tinted
          node_fill_color_secondary: ws.node_fill_color_secondary || bt.secondary_color || '#3A6EA5',
          // Level-3+ nodes (leaves) **” very light neutral
          node_fill_color_leaf:      ws.node_fill_color_leaf      || '#EAF2FB',
          node_border_color:         ws.node_border_color         || '#FFFFFF',
          node_border_width:         ws.node_border_width    != null ? ws.node_border_width    : 1,
          node_corner_radius:        ws.node_corner_radius   != null ? ws.node_corner_radius   : 4,
          node_title_font_family:    ws.node_title_font_family    || bt.title_font_family || 'Arial',
          node_title_font_size:      ws.node_title_font_size      || 10,
          node_title_color:          ws.node_title_color          || '#FFFFFF',
          // Leaf nodes have dark text on light fill
          node_title_color_leaf:     ws.node_title_color_leaf     || bt.body_color || '#111111',
          node_value_font_family:    ws.node_value_font_family    || bt.body_font_family || 'Arial',
          node_value_font_size:      ws.node_value_font_size      || 9,
          node_value_color:          ws.node_value_color          || bt.body_color || '#111111',
          node_inner_padding:        ws.node_inner_padding   != null ? ws.node_inner_padding   : 0.08,
          external_label_gap:        ws.external_label_gap   != null ? ws.external_label_gap   : 0.08,
          connector_color:           ws.connector_color           || bt.primary_color   || '#0078AE',
          connector_width:           ws.connector_width      != null ? ws.connector_width      : 0.5,
          arrowhead_style:           ws.arrowhead_style           || 'triangle',
          // Timeline baseline bar color
          timeline_line_color:       ws.timeline_line_color       || bt.primary_color   || '#0078AE'
        }
        art.container = { x: art.x || 0, y: art.y || 0, w: art.w || 0, h: art.h || 0 }

        const flow = String(art.flow_direction || '').toLowerCase()
        const wtype = String(art.workflow_type || '').toLowerCase()
        const isHorizontal = flow === 'left_to_right' || flow === 'horizontal' || wtype === 'timeline' || wtype === 'roadmap' || wtype === 'process_flow'
        const isVerticalLinear = flow === 'top_to_bottom' || flow === 'bottom_up'
        const isTopDownBranching = !isVerticalLinear && (flow === 'top_down_branching' || flow === 'top_down' || flow === 'vertical' || wtype === 'decomposition' || wtype === 'hierarchy')

        if (nodes.length > 0 && isHorizontal) {
          const hasValues = nodes.some(n => String(n?.value || '').trim())
          const hasDescriptions = nodes.some(n => String(n?.description || '').trim())
          const ax = art.x || 0
          const ay = art.y || 0
          const aw = art.w || 0
          const ah = art.h || 0
          const hb = art.header_block || {}
          const headerLeft = hb.x != null ? hb.x : ax
          const headerRight = hb.x != null && hb.w != null ? (hb.x + hb.w) : (ax + aw)
          const railLeft = Math.max(ax, headerLeft)
          const railRight = Math.min(ax + aw, headerRight)
          const padX = Math.min(0.18, Math.max(0.10, aw * 0.02))

          // Compute where the body area starts **” after the header_block (if any).
          // Value labels in horizontal flows are rendered ABOVE nodeY, so nodeY must be
          // at least topBand inches below the body start, not below art.y.
          let bodyStartY = ay
          if (hb && hb.text) {
            const hbH = Math.max(hb.h != null ? +hb.h : 0, estimateHeaderBlockHeight(hb.text, aw, hb.font_size || 11))
            const hbRule = (hb.style === 'brand_fill') ? 0 : 0.005  // hairline rule
            const hbGap  = 0.06
            bodyStartY = round2(ay + hbH + hbRule + hbGap)
          }
          const effectiveBodyH = round2(ay + ah - bodyStartY)

          // topBand = space between bodyStartY and the top of the node box.
          // When value labels exist they are rendered in this band (above the box),
          // so size it from the actual value font rather than a hardcoded constant.
          //   valueLabelH  = one line of value text at node_value_font_size
          //   GAP_ABOVE    = breathing room between header-rule bottom and value text top
          //   GAP_BELOW    = breathing room between value text bottom and node box top
          const valueFs     = art.workflow_style.node_value_font_size || 9
          const valueLabelH = round2(valueFs * 1.35 / 72)          // one text line in inches
          const GAP_ABOVE   = 0.06                                  // header â†’ value label
          const GAP_BELOW   = 0.06                                  // value label â†’ node box
          const topBand = hasValues
            ? round2(Math.max(0.22, GAP_ABOVE + valueLabelH + GAP_BELOW))
            : 0.12
          const bottomBand = hasDescriptions ? Math.min(0.95, Math.max(0.60, effectiveBodyH * 0.26)) : 0.12
          const titleFs = art.workflow_style.node_title_font_size || 10
          const innerPad = art.workflow_style.node_inner_padding != null ? art.workflow_style.node_inner_padding : 0.08
          const avgCharW = titleFs * 0.58 / 72
          const longestLabel = Math.max(...nodes.map(n => String(n?.label || '').length), 8)
          const targetLines = longestLabel > 16 ? 3 : 2
          const minFitW = Math.max(0.92, Math.ceil(longestLabel / targetLines) * avgCharW + innerPad * 2 + 0.12)
          const gapMin = nodes.length >= 4 ? 0.16 : 0.20
          const alignmentSpan = Math.max(railRight - railLeft, aw - padX * 2)
          let nodeW = round2(Math.min(2.10, Math.max(minFitW, (alignmentSpan - gapMin * Math.max(nodes.length - 1, 0)) / Math.max(nodes.length, 1))))
          let gap = nodes.length > 1
            ? round2((alignmentSpan - nodeW * nodes.length) / Math.max(nodes.length - 1, 1))
            : 0
          if (gap < gapMin) {
            nodeW = round2((alignmentSpan - gapMin * Math.max(nodes.length - 1, 0)) / Math.max(nodes.length, 1))
            gap = gapMin
          }
          nodeW = round2(Math.max(0.88, nodeW))
          const nodeH = round2(Math.max(0.55, Math.min(1.00, effectiveBodyH - topBand - bottomBand - 0.08)))
          const nodeY = round2(bodyStartY + topBand + 0.04)
          const startX = nodes.length > 1 ? railLeft : round2(Math.max(ax + padX, railLeft + (alignmentSpan - nodeW) / 2))

          art.nodes = nodes.map((node, i) => ({
            ...node,
            x: round2(startX + i * (nodeW + gap)),
            y: nodeY,
            w: nodeW,
            h: nodeH
          }))

          art.connections = art.nodes.slice(0, -1).map((node, i) => {
            const next = art.nodes[i + 1]
            return {
              from: node.id,
              to: next.id,
              type: ((art.connections || [])[i] || {}).type || 'arrow',
              path: [
                { x: round2(node.x + node.w), y: round2(node.y + node.h / 2) },
                { x: round2(next.x), y: round2(next.y + next.h / 2) }
              ]
            }
          })

          const widthRatio = nodeW / Math.max(minFitW, 1.2)
          if (widthRatio < 1) {
            art.workflow_style = {
              ...art.workflow_style,
              node_title_font_size: Math.max(8, Math.floor(titleFs * Math.max(widthRatio, 0.88))),
              node_value_font_size: Math.max(7, Math.floor(valueFs * Math.max(widthRatio, 0.85)))
            }
          }
        } else if (nodes.length > 0 && isTopDownBranching) {
          const ax = art.x || 0
          const ay = art.y || 0
          const aw = art.w || 0
          const ah = art.h || 0
          const levels = [...new Set(nodes.map(n => Number.isFinite(+n?.level) ? +n.level : 1))].sort((a, b) => a - b)
          const levelNodes = levels.map(level => nodes.filter(n => (Number.isFinite(+n?.level) ? +n.level : 1) === level))
          const maxPerLevel = Math.max(...levelNodes.map(row => row.length), 1)
          const topPad = 0.10
          const bottomPad = nodes.some(n => String(n?.description || '').trim()) ? Math.min(0.95, Math.max(0.60, ah * 0.24)) : 0.16
          const sidePad = Math.min(0.18, Math.max(0.10, aw * 0.03))
          const rowGap = levels.length > 1 ? Math.max(0.24, Math.min(0.50, ah * 0.10)) : 0
          const usableH = Math.max(0.8, ah - topPad - bottomPad - rowGap * Math.max(levels.length - 1, 0))
          const nodeH = round2(Math.max(0.72, Math.min(1.00, usableH / Math.max(levels.length, 1))))
          const rowYByLevel = new Map()
          let curY = ay + topPad
          levels.forEach(level => {
            rowYByLevel.set(level, round2(curY))
            curY += nodeH + rowGap
          })

          const titleFs = art.workflow_style.node_title_font_size || 10
          const innerPad = art.workflow_style.node_inner_padding != null ? art.workflow_style.node_inner_padding : 0.08
          const avgCharW = titleFs * 0.58 / 72
          const longestLabel = Math.max(...nodes.map(n => String(n?.label || '').length), 8)
          const targetLines = longestLabel > 16 ? 3 : 2
          const minFitW = Math.max(0.92, Math.ceil(longestLabel / targetLines) * avgCharW + innerPad * 2 + 0.12)
          const usableW = Math.max(1.0, aw - sidePad * 2)
          const gapMin = maxPerLevel >= 4 ? 0.16 : 0.20
          let nodeW = round2(Math.max(0.88, Math.min(2.40, (usableW - gapMin * Math.max(maxPerLevel - 1, 0)) / Math.max(maxPerLevel, 1))))
          nodeW = round2(Math.max(nodeW, minFitW))
          if (maxPerLevel > 1 && nodeW * maxPerLevel + gapMin * (maxPerLevel - 1) > usableW) {
            nodeW = round2(Math.max(0.82, (usableW - gapMin * (maxPerLevel - 1)) / maxPerLevel))
          }

          const placedNodes = []
          levels.forEach(level => {
            const row = levelNodes.find(group => group[0] && (Number.isFinite(+group[0]?.level) ? +group[0].level : 1) === level) || []
            const rowCount = Math.max(row.length, 1)
            const totalRowW = rowCount * nodeW + Math.max(0, rowCount - 1) * gapMin
            const startX = round2(ax + sidePad + Math.max(0, (usableW - totalRowW) / 2))
            row.forEach((node, idx) => {
              placedNodes.push({
                ...node,
                x: round2(startX + idx * (nodeW + gapMin)),
                y: rowYByLevel.get(level),
                w: nodeW,
                h: nodeH
              })
            })
          })
          art.nodes = placedNodes

          const placedById = new Map((art.nodes || []).map(n => [n.id, n]))
          const originalConns = Array.isArray(art.connections) ? art.connections : []
          art.connections = originalConns.map((conn) => {
            const fromNode = placedById.get(conn.from)
            const toNode = placedById.get(conn.to)
            if (!fromNode || !toNode) return {
              ...conn,
              type: conn.type || 'arrow',
              path: Array.isArray(conn.path) ? conn.path : []
            }
            const startX = round2(fromNode.x + fromNode.w / 2)
            const startY = round2(fromNode.y + fromNode.h)
            const endX = round2(toNode.x + toNode.w / 2)
            const endY = round2(toNode.y)
            const midY = round2(startY + Math.max(0.10, (endY - startY) * 0.45))
            return {
              from: conn.from || fromNode.id || '',
              to: conn.to || toNode.id || '',
              type: conn.type || 'arrow',
              path: [
                { x: startX, y: startY },
                { x: startX, y: midY },
                { x: endX, y: midY },
                { x: endX, y: endY }
              ]
            }
          })
        } else if (nodes.length > 0 && isVerticalLinear) {
          // ****** top_to_bottom / bottom_up: linear vertical stack ***************************************************
          // Node box occupies ~40% of container width; right side reserved for description band.
          const ax = art.x || 0
          const ay = art.y || 0
          const aw = art.w || 0
          const ah = art.h || 0
          const hb = art.header_block || {}
          const headerH = (hb && hb.h) ? (hb.h + 0.06) : 0
          const topPad = 0.10
          const gapBetween = nodes.length > 1 ? Math.max(0.18, Math.min(0.35, ah * 0.06)) : 0
          const usableH = Math.max(0.8, ah - headerH - topPad - gapBetween * Math.max(nodes.length - 1, 0))
          const nodeH = round2(Math.max(0.60, Math.min(1.10, usableH / Math.max(nodes.length, 1))))
          // Node box takes left 40% of width; right 60% reserved for description text
          const hasDescs = nodes.some(n => String(n?.description || '').trim())
          const nodeWFraction = hasDescs ? 0.40 : 0.90
          const nodeW = round2(Math.max(0.80, Math.min(3.0, aw * nodeWFraction)))

          const titleFs = art.workflow_style.node_title_font_size || 10
          const innerPad = art.workflow_style.node_inner_padding != null ? art.workflow_style.node_inner_padding : 0.08
          const avgCharW = titleFs * 0.58 / 72
          const longestLabel = Math.max(...nodes.map(n => String(n?.label || '').length), 8)
          const targetLines = longestLabel > 16 ? 3 : 2
          const minFitW = Math.max(0.80, Math.ceil(longestLabel / targetLines) * avgCharW + innerPad * 2 + 0.12)
          const finalNodeW = round2(Math.max(nodeW, minFitW))

          const startY = round2(ay + headerH + topPad)
          art.nodes = nodes.map((node, i) => ({
            ...node,
            x: round2(ax),
            y: round2(startY + i * (nodeH + gapBetween)),
            w: finalNodeW,
            h: nodeH
          }))

          // Connections: bottom-center â†’ top-center for sequential pairs
          art.connections = art.nodes.slice(0, -1).map((node, i) => {
            const next = art.nodes[i + 1]
            return {
              from: node.id,
              to: next.id,
              type: ((art.connections || [])[i] || {}).type || 'arrow',
              path: [
                { x: round2(node.x + node.w / 2), y: round2(node.y + node.h) },
                { x: round2(next.x + next.w / 2), y: round2(next.y) }
              ]
            }
          })
        }
      }

      // ****** 5. insight_text (standard): font scaling *********************************************************************************
      if (artType === 'matrix') {
        const ms = art.matrix_style || {}
        art.matrix_style = {
          border_color:    ms.border_color    || '#D7DEE8',
          border_width:    ms.border_width    != null ? ms.border_width    : 0.8,
          divider_color:   ms.divider_color   || '#AAAAAA',
          divider_width:   ms.divider_width   != null ? ms.divider_width   : 0.5,
          axis_label_font_family: ms.axis_label_font_family || bt.body_font_family   || 'Arial',
          axis_label_font_size:   ms.axis_label_font_size   || 9,
          axis_label_color:       ms.axis_label_color || bt.caption_color || bt.body_color || '#6B7280',
          quadrant_title_font_family: ms.quadrant_title_font_family || bt.title_font_family || 'Arial',
          quadrant_title_font_size:   ms.quadrant_title_font_size   || 11,
          quadrant_body_font_family:  ms.quadrant_body_font_family  || bt.body_font_family  || 'Arial',
          quadrant_body_font_size:    ms.quadrant_body_font_size    || 9,
          // Tone-driven quadrant fills
          positive_quadrant_fill: ms.positive_quadrant_fill || null,   // resolved to brand tint in _matrixToBlocks
          negative_quadrant_fill: ms.negative_quadrant_fill || null,
          neutral_quadrant_fill:  ms.neutral_quadrant_fill  || null,
          // Tone-driven quadrant text colors
          positive_title_color:  ms.positive_title_color || bt.primary_color || '#1B5E20',
          negative_title_color:  ms.negative_title_color || '#B91C1C',
          neutral_title_color:   ms.neutral_title_color  || bt.body_color    || '#374151',
          positive_body_color:   ms.positive_body_color  || bt.primary_color || '#2D7F5E',
          negative_body_color:   ms.negative_body_color  || '#B91C1C',
          neutral_body_color:    ms.neutral_body_color   || bt.body_color    || '#374151',
          // Tone-driven point dot colors
          positive_point_fill:  ms.positive_point_fill || bt.primary_color  || '#2D7F5E',
          negative_point_fill:  ms.negative_point_fill || '#C53030',
          neutral_point_fill:   ms.neutral_point_fill  || bt.secondary_color || '#6B7280',
          // Point label bubble
          point_label_font_family: ms.point_label_font_family || bt.body_font_family || 'Arial',
          point_label_font_size:   ms.point_label_font_size   || 9
        }
      }

      if (artType === 'driver_tree') {
        const ts = art.tree_style || {}
        art.tree_style = {
          node_fill_color: ts.node_fill_color || '#EAF2FB',
          node_fill_color_secondary: ts.node_fill_color_secondary || '#EDF7F3',
          node_fill_color_leaf: ts.node_fill_color_leaf || '#F4F7FA',
          node_border_color: ts.node_border_color || '#D7DEE8',
          node_border_width: ts.node_border_width != null ? ts.node_border_width : 0.6,
          connector_color: ts.connector_color || '#7A8FA8',
          connector_width: ts.connector_width != null ? ts.connector_width : 0.5,
          label_font_family: ts.label_font_family || bt.title_font_family || 'Arial',
          label_font_size: ts.label_font_size || 11,
          label_color: ts.label_color || '#111111',
          value_font_family: ts.value_font_family || bt.body_font_family || 'Arial',
          value_font_size: ts.value_font_size || 10,
          value_color: ts.value_color || bt.primary_color || '#0078AE',
          corner_radius: ts.corner_radius != null ? ts.corner_radius : 6
        }
      }

      if (artType === 'prioritization') {
        const ps = art.priority_style || {}
        const rankPalette = [
          bt.secondary_color,
          bt.primary_color,
          ...(bt.accent_colors || []),
          ...(bt.chart_palette || [])
        ].filter(Boolean)
        const qualifierPalette = [
          bt.primary_color,
          bt.secondary_color,
          ...(bt.accent_colors || []),
          ...(bt.chart_palette || [])
        ].filter(Boolean)
        art.priority_style = {
          row_fill_color: ps.row_fill_color || '#FFFFFF',
          row_border_color: ps.row_border_color || '#D7DEE8',
          row_border_width: ps.row_border_width != null ? ps.row_border_width : 0.6,
          row_corner_radius: ps.row_corner_radius != null ? ps.row_corner_radius : 6,
          row_gap_in: ps.row_gap_in != null ? ps.row_gap_in : 0.16,
          rank_palette: (ps.rank_palette && ps.rank_palette.length ? ps.rank_palette : rankPalette),
          rank_font_family: ps.rank_font_family || bt.title_font_family || 'Arial',
          rank_font_size: ps.rank_font_size || 17,
          rank_text_color: ps.rank_text_color || '#FFFFFF',
          title_font_family: ps.title_font_family || bt.title_font_family || 'Arial',
          title_font_size: ps.title_font_size || 14,
          title_color: ps.title_color || '#1F2937',
          description_font_family: ps.description_font_family || bt.body_font_family || 'Arial',
          description_font_size: ps.description_font_size || 11,
          description_color: ps.description_color || '#374151',
          qualifier_fill_color: ps.qualifier_fill_color || '#EEF4E2',
          qualifier_text_color: ps.qualifier_text_color || '#1F2937',
          qualifier_value_palette: (ps.qualifier_value_palette && ps.qualifier_value_palette.length ? ps.qualifier_value_palette : qualifierPalette),
          qualifier_label_font_family: ps.qualifier_label_font_family || bt.body_font_family || 'Arial',
          qualifier_label_font_size: ps.qualifier_label_font_size || 10,
          qualifier_value_font_family: ps.qualifier_value_font_family || bt.title_font_family || 'Arial',
          qualifier_value_font_size: ps.qualifier_value_font_size || 10
        }
      }

      if (artType === 'insight_text' && art.insight_mode !== 'grouped') {
        const bs = art.body_style
        if (bs && bs.font_size != null) {
          const points      = art.points || []
          const n_points    = points.length
          if (n_points > 0) {
            const artH        = art.h || 0
            const body_h      = artH - 0.40 - 0.10 - 0.08
            const spec_fs     = bs.font_size
            const line_spacing  = bs.line_spacing    || 1.3
            const space_before  = bs.space_before_pt || 6
            const line_h_in   = spec_fs * line_spacing / 72
            const space_in    = space_before / 72
            const total_h     = n_points * line_h_in + (n_points - 1) * space_in

            if (total_h > body_h * 1.05) {
              const scaled_fs = Math.max(7, Math.floor(spec_fs * body_h / total_h))
              bs.font_size = scaled_fs
              if (bs.space_before_pt != null) {
                bs.space_before_pt = Math.round(bs.space_before_pt * scaled_fs / spec_fs)
              }
            }
          }
        }
      }
    } // end per-artifact loop
  } // end zones loop

  return zones
}


// *********************************************************************************************************************************************************************************************************************************************
// BLOCK FLATTENER
// Converts the final processed slide spec into a flat, ordered blocks[] array.
// Each block is self-contained: block_type + x/y/w/h + type-specific fields.
// Called after computeArtifactInternals in normaliseDesignedSlide.
// generate_pptx.py reads blocks[] and dispatches each to a typed renderer.
// *********************************************************************************************************************************************************************************************************************************************

function resolveArtifactSubtype(art) {
  if (!art || typeof art !== 'object') return 'generic'
  switch (art.type) {
    case 'chart':        return art.chart_type || 'generic'
    case 'stat_bar':     return 'stat_bar'
    case 'insight_text': return art.insight_mode || 'standard'
    case 'workflow':     return art.workflow_type || art.flow_direction || 'workflow'
    case 'cards':        return art.cards_layout || 'cards'
    case 'table':        return art.table_subtype || 'standard'
    case 'comparison_table': return 'comparison_table'
    case 'initiative_map': return 'initiative_map'
    case 'profile_card_set': return art.layout_direction || 'profile_card_set'
    case 'risk_register': return 'risk_register'
    case 'matrix':       return art.matrix_type || '2x2'
    case 'driver_tree':  return 'driver_tree'
    case 'prioritization': return 'ranked_list'
    default:             return art.type || 'generic'
  }
}

function resolveArtifactHeaderText(art) {
  if (!art || typeof art !== 'object') return ''
  return (
    art.artifact_header ||
    ((art.header_block || {}).text) ||
    art.insight_header ||
    art.chart_header ||
    art.table_header ||
    art.comparison_header ||
    art.initiative_header ||
    art.profile_header ||
    art.risk_header ||
    art.matrix_header ||
    art.tree_header ||
    art.priority_header ||
    art.workflow_header ||
    art.heading ||
    '')
}

function buildBlockFallbackPolicy(art, blockRole) {
  const artifactType = art?.type || 'generic'
  const artifactSubtype = resolveArtifactSubtype(art)
  return {
    allow_renderer_fallback: true,
    fallback_mode: 'subtype_default',
    trigger: 'missing_or_invalid_spec',
    artifact_type: artifactType,
    artifact_subtype: artifactSubtype,
    block_role: blockRole || 'artifact_body',
    fallback_key: artifactType + ':' + artifactSubtype
  }
}

function estimateHeaderBlockHeight(text, widthIn, fontSizePt) {
  const textStr = String(text || '').trim()
  if (!textStr) return 0.3
  const usableWidth = Math.max(0.6, Number(widthIn) || 0.6)
  const fontSize = Math.max(8, Number(fontSizePt) || 11)
  const charsPerLine = Math.max(10, Math.floor((usableWidth * 72) / (fontSize * 0.52)))
  const words = textStr.split(/\s+/).filter(Boolean)
  let lines = 1
  let lineLen = 0
  for (const word of words) {
    const nextLen = lineLen === 0 ? word.length : lineLen + 1 + word.length
    if (nextLen <= charsPerLine) lineLen = nextLen
    else {
      lines += 1
      lineLen = word.length
    }
  }
  const textHeight = lines * (fontSize * 1.28 / 72)
  return Math.max(0.3, Math.round((textHeight + 0.06) * 100) / 100)
}

function normalizeArtifactHeaderBands(zones) {
  // Align header_block bottom edges across artifacts whose headers start at the same y.
  // Applied to ALL header styles (underline and brand_fill) **” the bottom-edge alignment
  // is style-agnostic and prevents ragged-looking multi-zone slides.
  const items = []
  for (const zone of (zones || [])) {
    for (const art of (zone.artifacts || [])) {
      const hb = art && art.header_block
      if (!hb || !hb.text) continue
      const hy = Number(hb.y != null ? hb.y : art.y)
      const hw = Number(hb.w != null ? hb.w : art.w)
      const hfs = Number(hb.font_size || 11)
      if (!isFinite(hy) || !isFinite(hw)) continue
      const estimatedH = estimateHeaderBlockHeight(hb.text, hw, hfs)
      items.push({
        art,
        hb,
        y: hy,
        bottom: hy + Math.max(Number(hb.h || 0), estimatedH)
      })
    }
  }
  const groups = new Map()
  for (const item of items) {
    const key = String(Math.round(item.y * 8) / 8)
    if (!groups.has(key)) groups.set(key, [])
    groups.get(key).push(item)
  }
  for (const group of groups.values()) {
    if (group.length < 2) continue
    const alignedBottom = Math.max(...group.map(g => g.bottom))
    for (const g of group) {
      g.hb.h = Math.round(Math.max(Number(g.hb.h || 0), alignedBottom - g.y) * 100) / 100
    }
  }
}

function decorateArtifactBlocks(blocks, startIdx, endIdx, art, blockRole) {
  if (!art || startIdx >= endIdx) return
  const artifactType = art.type || 'generic'
  const artifactSubtype = resolveArtifactSubtype(art)
  const artifactHeaderText = resolveArtifactHeaderText(art)
  const artifactId = art._artifact_id || (artifactType + ':' + artifactSubtype + ':' + artifactHeaderText)
  const fallbackPolicy = buildBlockFallbackPolicy(art, blockRole)
  for (let i = startIdx; i < endIdx; i++) {
    blocks[i] = {
      ...blocks[i],
      artifact_id: blocks[i].artifact_id || artifactId,
      artifact_type: blocks[i].artifact_type || artifactType,
      artifact_subtype: blocks[i].artifact_subtype || artifactSubtype,
      artifact_header_text: blocks[i].artifact_header_text != null ? blocks[i].artifact_header_text : artifactHeaderText,
      block_role: blocks[i].block_role || blockRole,
      fallback_policy: blocks[i].fallback_policy || fallbackPolicy
    }
  }
}

function _sentimentColor(sentiment, style, bt) {
  const token = String(sentiment || '').toLowerCase()
  if (token === 'positive') return style?.positive_color || '#2D7F5E'
  if (token === 'negative') return style?.negative_color || '#C2410C'
  if (token === 'warning') return style?.warning_color || '#B45309'
  return style?.neutral_color || bt.body_color || '#111111'
}

function _ratingVisual(rating, note, cs) {
  const token = String(rating || '').toLowerCase()
  // Big-4 style: symbol + semantic color pill; text label only for free-text cells
  if (token === 'yes')     return { fill: cs.yes_fill_color     || '#D1FAE5', text: 'âœ“', textColor: cs.yes_text_color     || '#065F46', bold: true }
  if (token === 'partial') return { fill: cs.partial_fill_color || '#FEF3C7', text: 'â—‘', textColor: cs.partial_text_color || '#92400E', bold: true }
  if (token === 'no')      return { fill: cs.no_fill_color      || '#FEE2E2', text: 'âœ—', textColor: cs.no_text_color      || '#991B1B', bold: true }
  return { fill: cs.neutral_fill_color || '#F4F5F7', text: String(note || rating || ''), textColor: null, bold: false }
}

function _riskSeverityFill(severity, rs) {
  const token = String(severity || '').toLowerCase()
  if (token === 'critical') return rs.critical_fill_color || '#FEE2E2'
  if (token === 'high')     return rs.high_fill_color     || '#FFF1E5'
  if (token === 'medium')   return rs.medium_fill_color   || '#FFFBEB'
  return rs.low_fill_color || '#ECFDF5'
}

function _riskSeverityBadgeColor(severity, rs) {
  // Solid semantic badge colors **” severity IS the signal, not brand color
  const token = String(severity || '').toLowerCase()
  if (token === 'critical') return rs.critical_badge_color || '#DC2626'
  if (token === 'high')     return rs.high_badge_color     || '#EA580C'
  if (token === 'medium')   return rs.medium_badge_color   || '#D97706'
  return rs.low_badge_color || '#16A34A'
}

function _truncateText(text, maxChars) {
  const str = String(text || '')
  if (!maxChars || str.length <= maxChars) return str
  return str.slice(0, Math.max(0, maxChars - 1)).trimEnd() + '**¦'
}

function _displayLabel(value) {
  if (value == null) return ''
  if (typeof value === 'string') return value
  if (typeof value === 'object' && value !== null) {
    return String(value.label || value.name || value.title || value.id || '')
  }
  return String(value)
}


// risk_register renders as a severity-banded stack (new severity_levels schema):
// - Colored band per severity group: dot + label + item count
// - Item rows: bold primary_message + muted secondary_message (left); tags[] pills + pips[] squares (right)
// - Thin divider between items; section gap between groups
function _riskRegisterToBlocks(art, content_y, blocks, bt, r2) {
  const rs = art.risk_style || {}
  const severityLevels = Array.isArray(art.severity_levels) ? art.severity_levels : []
  const ax = art.x || 0
  const ay = content_y
  const aw = art.w || 0
  const ah = r2((art.y || 0) + (art.h || 0) - content_y)
  if (!severityLevels.length || aw <= 0 || ah <= 0) return

  const titleFont    = rs.label_font_family || bt.title_font_family || 'Arial'
  const bodyFont     = rs.body_font_family  || bt.body_font_family  || 'Arial'
  const bodyColor    = bt.body_color    || '#111111'
  const captionColor = bt.caption_color || bodyColor

  // Font sizes **” LLM-decided, consistent across all rows; fallback to sensible defaults
  const primaryFs   = rs.primary_message_font_size   || rs.label_font_size || 11
  const secondaryFs = rs.secondary_message_font_size || rs.body_font_size  || 9
  const bandLabelFs = rs.label_font_size || 10
  const pipLabelFs  = rs.body_font_size  || 8.5
  const countFs     = rs.body_font_size  || 9

  const pipSize    = 0.10
  const pipGap     = 0.03
  const bandH      = rs.band_height_in      != null ? +rs.band_height_in      : 0.34
  const rowH       = rs.row_height_in       != null ? +rs.row_height_in       : 0.90
  const dividerH   = 0.005
  const sectionGap = 0.18

  // Severity color fallbacks (LLM fills these; JS provides semantic fallbacks only)
  const badgeFb   = { critical: '#DC2626', high: '#EA580C', medium: '#D97706', low: '#6B7280' }
  const fillFb    = { critical: '#FDE8E8', high: '#FFF1F2', medium: '#FFF7E5', low: '#F3F4F6' }
  const txtFb     = { critical: '#8B2C23', high: '#8B2C23', medium: '#6E5712', low: '#6B7280' }
  const pipFb     = { critical: '#8B2C23', high: '#8B2C23', medium: '#6E5712', low: '#7A7A72' }

  const bandFill  = tone => rs[`${tone}_fill_color`]  || fillFb[tone]  || '#F5F5F5'
  const dotColor  = tone => rs[`${tone}_badge_color`] || badgeFb[tone] || '#6B7280'
  const bandTxt   = tone => rs[`${tone}_text_color`]  || txtFb[tone]   || '#374151'
  const pipFill   = tone => rs[`${tone}_pip_fill`]    || pipFb[tone]   || dotColor(tone)
  const pipEmpty  = '#FFFFFF'
  const pipBorder = '#C8C8C8'

  // intensity â†’ filled pip count; pipLevels = total squares in scale
  const intensityToCount = (v, pipLevels) => {
    const scale = pipLevels || 5
    if (typeof v === 'number') return Math.max(0, Math.min(scale, Math.round(v)))
    const n = Number(v)
    if (!isNaN(n)) return Math.max(0, Math.min(scale, Math.round(n)))
    // Legacy semantic fallback (maps to a 5-level scale, then scales to pipLevels)
    const t = String(v || '').toLowerCase().replace(/[\s_]/g, '')
    const sem5 = t === 'extreme' || t === 'vhigh' ? 5
               : t === 'high'                     ? 4
               : t === 'medium' || t === 'med'    ? 3
               : t === 'low'                       ? 2
               : t === 'vlow'                      ? 1
               :                                    0
    return Math.round(sem5 * scale / 5)
  }

  // Tag chip colors
  const tagColors = tone => {
    if (tone === 'positive') return { fill: rs.tag_positive_fill || '#F1F8E8', border: rs.tag_positive_border || '#7AA243', text: rs.tag_positive_text || '#386B2A' }
    if (tone === 'negative') return { fill: rs.tag_negative_fill || '#FFF7F6', border: rs.tag_negative_border || '#A33B32', text: rs.tag_negative_text || '#A33B32' }
    if (tone === 'warning')  return { fill: rs.tag_warning_fill  || '#FFF8E8', border: rs.tag_warning_border  || '#9A6B10', text: rs.tag_warning_text  || '#6E5712' }
    return                          { fill: rs.tag_neutral_fill  || '#F5F5F5', border: rs.tag_neutral_border  || '#D1D5DB', text: rs.tag_neutral_text  || '#374151' }
  }

  // Right column: enough for 2 tags + pip grid
  const rightColW = rs.right_col_width_in != null ? +rs.right_col_width_in : 1.90
  const leftX     = r2(ax + 0.16)
  const leftW     = r2(aw - rightColW - 0.28)

  let cursorY = ay

  severityLevels.forEach((level, li) => {
    const tone    = String(level.tone || level.severity || 'medium').toLowerCase()
    const bFill   = bandFill(tone)
    const dColor  = dotColor(tone)
    const txtColor= bandTxt(tone)
    const items   = Array.isArray(level.item_details) ? level.item_details : []

    // Severity band
    blocks.push({ block_type: 'rect', x: ax, y: cursorY, w: aw, h: bandH, fill_color: bFill, border_color: null, border_width: 0, corner_radius: 8 })
    blocks.push({ block_type: 'circle', x: r2(ax + 0.16), y: r2(cursorY + 0.12), w: 0.10, h: 0.10, fill_color: dColor, text: '' })
    blocks.push({ block_type: 'text_box', x: r2(ax + 0.36), y: r2(cursorY + 0.06), w: r2(aw - 1.8), h: 0.22,
      text: level.label || tone, font_family: titleFont, font_size: bandLabelFs, bold: true, color: txtColor, align: 'left', valign: 'middle' })
    blocks.push({ block_type: 'text_box', x: r2(ax + aw - 0.88), y: r2(cursorY + 0.06), w: 0.72, h: 0.22,
      text: `${items.length} item${items.length !== 1 ? 's' : ''}`, font_family: bodyFont, font_size: countFs, bold: false, color: txtColor, align: 'right', valign: 'middle' })
    cursorY = r2(cursorY + bandH + 0.08)

    items.forEach((item, ii) => {
      const rowY = cursorY

      // Pip layout constants **” pipLevels drives total squares per row
      const pips     = Array.isArray(item.pips) ? item.pips.slice(0, 4) : []
      const pipLevels= level.pip_levels || 5
      const pipRowH  = 0.22   // height per pip row
      const tagAreaH = 0.42   // height reserved for tags row (0.30 = tag bottom; extra 0.12 = gap before pips)
      // Right column fixed width; pip grid uses left part, pip squares right part
      const pipLblW  = 0.82
      const pipGridW = pipLevels * pipSize + (pipLevels - 1) * pipGap
      const pipStart = r2(ax + aw - pipGridW - 0.06)  // right-aligned within rightColW
      const pipLblX  = r2(pipStart - pipLblW - 0.04)

      // Left: primary_message (bold) + secondary_message (muted)
      const primY = r2(rowY + 0.06)
      const secY  = r2(rowY + 0.06 + primaryFs / 72 * 1.3 * 1.2 + 0.06)   // below primary with gap
      blocks.push({ block_type: 'text_box', x: leftX, y: primY, w: leftW, h: r2(primaryFs / 72 * 1.5 + 0.06),
        text: _truncateText(item.primary_message || '', 64), font_family: titleFont, font_size: primaryFs, bold: true, color: bodyColor, align: 'left', valign: 'top' })
      if (item.secondary_message) {
        blocks.push({ block_type: 'text_box', x: leftX, y: secY, w: leftW, h: r2(secondaryFs / 72 * 1.5 * 2 + 0.04),
          text: _truncateText(item.secondary_message, 120), font_family: bodyFont, font_size: secondaryFs, bold: false, color: captionColor, align: 'left', valign: 'top' })
      }

      // Right top: tag pills **” right-aligned so last tag flush with artifact right edge
      const tags     = Array.isArray(item.tags) ? item.tags : []
      const pillMaxX = r2(ax + aw - 0.04)          // hard right edge (same as pip grid)
      const pillMinX = r2(ax + aw - rightColW)      // don't encroach on left text column
      // Compute positions right-to-left, then render in visual (left-to-right) order
      const visibleTags = []
      let pillRightEdge = pillMaxX
      ;[...tags].reverse().forEach(tag => {
        const val = String(tag.value || '')
        const tW  = r2(Math.min(1.10, Math.max(0.58, val.length * 0.068 + 0.20)))
        const pX  = r2(pillRightEdge - tW)
        if (pX < pillMinX) return   // skip **” no room in right column
        visibleTags.unshift({ tag, pillX: pX, tW })
        pillRightEdge = r2(pX - 0.07)
      })
      visibleTags.forEach(({ tag, pillX, tW }) => {
        const tc  = tagColors(String(tag.tone || 'neutral').toLowerCase())
        const val = String(tag.value || '')
        blocks.push({ block_type: 'rect', x: pillX, y: r2(rowY + 0.04), w: tW, h: 0.26,
          fill_color: tc.fill, border_color: tc.border, border_width: 0.5, corner_radius: 10 })
        blocks.push({ block_type: 'text_box', x: r2(pillX + 0.07), y: r2(rowY + 0.08), w: r2(tW - 0.10), h: 0.16,
          text: _truncateText(val, 16), font_family: bodyFont, font_size: pipLabelFs, bold: false, color: tc.text, align: 'center', valign: 'middle' })
      })

      // Right bottom: pip rows **” each pip on its own row, right-aligned within rightColW
      pips.forEach((pip, pi) => {
        const pipY = r2(rowY + tagAreaH + pi * pipRowH)
        blocks.push({ block_type: 'text_box', x: pipLblX, y: pipY, w: pipLblW, h: 0.18,
          text: pip.label || '', font_family: bodyFont, font_size: pipLabelFs, bold: false, color: captionColor, align: 'right', valign: 'middle' })
        for (let i = 0; i < pipLevels; i++) {
          blocks.push({ block_type: 'rect', x: r2(pipStart + i * (pipSize + pipGap)), y: r2(pipY + 0.04), w: pipSize, h: pipSize,
            fill_color: i < intensityToCount(pip.intensity, pipLevels) ? pipFill(tone) : pipEmpty, border_color: pipBorder, border_width: 0.5, corner_radius: 2 })
        }
      })

      cursorY = r2(cursorY + rowH)
      if (ii < items.length - 1) {
        blocks.push({ block_type: 'rule', x: r2(ax + 0.16), y: r2(cursorY - 0.04), w: r2(aw - 0.32), h: dividerH,
          color: rs.row_border_color || '#D9D9D9', line_width: 0.5 })
      }
    })

    if (li < severityLevels.length - 1) cursorY = r2(cursorY + sectionGap)
  })
}

function _statBarToBlocks(art, content_y, blocks, bt, r2) {
  const cs = art.annotation_style || art.chart_style || {}
  const colHeaders = art.column_headers
  const isNewSchema = Array.isArray(colHeaders) && colHeaders.length > 0

  const ax = art.x || 0
  const ay = content_y
  const aw = art.w || 0
  const ah = r2((art.y || 0) + (art.h || 0) - content_y)

  // ****** Shared color tokens ************************************************************************************************************************************************************
  const bodyFont           = cs.label_font_family || bt.body_font_family || 'Arial'
  const bodyTextColor      = bt.body_color || '#111111'
  const captionColor       = bt.caption_color || bodyTextColor
  const headerColor        = cs.axis_color || captionColor
  const annotationColor    = cs.annotation_color || cs.axis_color || captionColor
  const dividerColor       = cs.gridline_color || cs.border_color || '#CFCFCF'
  const rowBorder          = cs.border_color || cs.gridline_color || '#E1E5EF'
  const trackFill          = cs.background_color || '#EEF1F5'
  const highlightBarFill   = (bt.chart_palette && bt.chart_palette[0])
    || (bt.accent_colors && bt.accent_colors[0])
    || bt.primary_color || '#CFE0A9'
  const highlightFill      = cs.background_color || '#EBF4E2'
  const neutralBarColor    = cs.axis_color || '#7B7B7B'
  const highlightTextColor = bodyTextColor

  if (isNewSchema) {
    // ****** New flexible-column schema ***************************************************************************************************************************************
    const rows  = Array.isArray(art.rows) ? art.rows : []
    const items = rows.slice(0, 8)
    if (!items.length || aw <= 0 || ah <= 0) return

    const headerH   = r2(Math.min(0.48, Math.max(0.30, ah * 0.11)))
    const headerGap = r2(Math.max(0.08, Math.min(0.16, ah * 0.03)))
    const bodyH     = Math.max(0.6, ah - headerH - headerGap)
    // Prefer LLM-computed values (set by agent 5 prompt formula, based on artifact_split_hint).
    // Fall back to JS computation only when absent.
    const rowGap = items.length > 1
      ? (art.row_gap_in != null
          ? r2(art.row_gap_in)
          : r2(Math.max(0.04, Math.min(0.18, bodyH * 0.05 / items.length))))
      : 0
    const rowH = art.row_height_in != null
      ? r2(art.row_height_in)
      : r2((bodyH - rowGap * Math.max(0, items.length - 1)) / Math.max(items.length, 1))
    const rowPadX   = r2(Math.max(0.08, Math.min(0.14, aw * 0.018)))
    const colGap    = 0.12
    const headerFontSize     = Math.max(10, Math.min(13,   rowH * 20))
    const labelFontSize      = Math.max(11, Math.min(15,   rowH * 21))
    const valueFontSize      = Math.max(11, Math.min(14.5, rowH * 20))
    const annotationFontSize = Math.max(10, Math.min(14,   rowH * 19))

    // ****** Column width allocation ************************************************************************************************************************************************
    const barCols_r  = colHeaders.filter(c => c.display_type === 'bar')   // 1**“3 bar columns
    const textCols   = colHeaders.filter(c => c.display_type === 'text')
    const normalCols = colHeaders.filter(c => c.display_type === 'normal')
    const numBars    = barCols_r.length || 1
    const totalGaps  = colGap * (colHeaders.length - 1)
    const normalColW = r2(Math.min(Math.max(0.92, aw * 0.12), 1.2))
    const totalNormalW = normalColW * normalCols.length
    // Bar cols collectively take ~40% of zone width, split equally (shrinks per-bar as numBars grows)
    const totalBarW  = r2(Math.max(numBars * 1.0, (aw - totalGaps - totalNormalW) * 0.40))
    const barColW    = r2(totalBarW / numBars)
    const totalTextW = aw - totalGaps - totalNormalW - totalBarW
    const perTextColW = textCols.length > 0 ? r2(totalTextW / textCols.length) : 0

    // Map col id â†’ { x, w, display_type, barCompanionOf?, barScale }
    // Per-column bar scale: column-level scale_UL â†’ artifact-level scale_UL â†’ auto max
    const artScaleUL = art.scale_UL > 0 ? +art.scale_UL : null
    const barScales  = {}
    for (const bc of barCols_r) {
      const colId  = String(bc.id)
      const colUL  = bc.scale_UL > 0 ? +bc.scale_UL : null
      const useUL  = colUL || artScaleUL
      const maxVal = useUL || Math.max(1, ...items.map(row => {
        const cell = (row?.cells || []).find(c => String(c.col_id) === colId)
        return Math.abs(+cell?.value || 0)
      }))
      barScales[colId] = maxVal
    }
    // Per-column lower limit for bar normalisation: scale_LL if set, else 50% of column minimum
    const barScaleLLs = {}
    for (const bc of barCols_r) {
      const colId = String(bc.id)
      if (bc.scale_LL != null) {
        barScaleLLs[colId] = Math.max(0, +bc.scale_LL)
      } else {
        const minVal = Math.min(...items.map(row => {
          const cell = (row?.cells || []).find(c => String(c.col_id) === colId)
          return Math.abs(+cell?.value || 0)
        }))
        barScaleLLs[colId] = Math.max(0, minVal * 0.5)
      }
    }

    const colLayout = {}
    let curX = ax
    for (let ci = 0; ci < colHeaders.length; ci++) {
      const col   = colHeaders[ci]
      const colId = String(col.id)
      const w = col.display_type === 'bar' ? barColW
              : col.display_type === 'normal' ? normalColW
              : perTextColW
      const prevCol = colHeaders[ci - 1]
      const isBarCompanion = col.display_type === 'normal' && prevCol?.display_type === 'bar' && !col.value
      colLayout[colId] = {
        x: r2(curX), w,
        display_type: col.display_type,
        barCompanionOf: isBarCompanion ? String(prevCol.id) : null
      }
      // Use a tight gap between a bar column and its value companion; standard gap everywhere else
      const nextCol = colHeaders[ci + 1]
      const nextIsCompanion = nextCol?.display_type === 'normal' && !nextCol?.value && col.display_type === 'bar'
      curX += w + (nextIsCompanion ? 0.03 : colGap)
    }

    const bodyTop  = r2(ay + headerH + headerGap)

    // ****** Column header row ******************************************************************************************************************************************************************
    for (const col of colHeaders) {
      const colId     = String(col.id)
      const lay       = colLayout[colId]
      const isNormal  = col.display_type === 'normal'
      blocks.push({
        block_type: 'text_box',
        x: lay.x + (isNormal ? 0 : rowPadX), y: ay,
        w: lay.w - (isNormal ? 0 : rowPadX), h: headerH,
        text: String(col.value || ''),
        font_family: bodyFont, font_size: headerFontSize, bold: true,
        color: headerColor, align: isNormal ? 'right' : 'left', valign: 'middle'
      })
    }
    blocks.push({ block_type: 'rule', x: ax, y: r2(ay + headerH + 0.01), w: aw, h: 0.005, color: dividerColor, line_width: 0.6 })

    // ****** Data rows ******************************************************************************************************************************************************************************************
    items.forEach((row, ri) => {
      const y            = r2(bodyTop + ri * (rowH + rowGap))
      const isHighlighted = row?.row_focus === 'Y'
      const trackY       = r2(y + rowH * 0.40)
      const trackH       = r2(Math.max(0.14, Math.min(0.20, rowH * 0.18)))

      blocks.push({
        block_type: 'rect',
        x: ax, y: r2(y + 0.01), w: aw, h: r2(rowH - 0.02),
        fill_color: isHighlighted ? highlightFill : '#FFFFFF',
        border_color: isHighlighted ? null : rowBorder,
        border_width: isHighlighted ? 0 : 0.5, corner_radius: 10
      })

      for (let ci = 0; ci < colHeaders.length; ci++) {
        const col       = colHeaders[ci]
        const colId     = String(col.id)
        const lay       = colLayout[colId]
        const rowCells  = row?.cells || []
        // Primary: match by col_id; fallback: match by column position if LLM used inconsistent ids
        const cell      = rowCells.find(c => String(c.col_id) === colId) || rowCells[ci]
        const cellValue = cell?.value ?? ''

        if (col.display_type === 'bar') {
          const rawVal = Math.abs(+cellValue || 0)
          const ll = barScaleLLs[colId] ?? 0
          const ul = barScales[colId] || 1
          const frac = Math.max(0.05, Math.min(1, (rawVal - ll) / Math.max(ul - ll, 1)))
          const barLen = r2(Math.max(0.06, lay.w * frac))
          blocks.push({ block_type: 'rect', x: lay.x, y: trackY, w: lay.w, h: trackH, fill_color: trackFill, border_color: null, border_width: 0, corner_radius: 8 })
          blocks.push({ block_type: 'rect', x: lay.x, y: trackY, w: Math.max(0.04, barLen), h: trackH, fill_color: isHighlighted ? highlightBarFill : neutralBarColor, border_color: null, border_width: 0, corner_radius: 8 })
        } else if (col.display_type === 'normal') {
          // If this is a bar's companion column and its cell is empty, fall back to the bar cell value
          let displayVal = cellValue
          if (!displayVal && lay.barCompanionOf) {
            const barCell = (row?.cells || []).find(c => String(c.col_id) === lay.barCompanionOf)
            displayVal = barCell?.value ?? ''
          }
          const isLastCol = ci === colHeaders.length - 1
          // Bar companion columns (value label immediately after a bar) are left-aligned so the
          // number sits tight against the bar end. Standalone normal columns (e.g. trailing %) stay
          // right-aligned.
          const isBarCompanion = !!lay.barCompanionOf
          const alignVal = isBarCompanion ? 'left' : 'right'
          blocks.push({
            block_type: 'text_box',
            x: lay.x, y, w: isLastCol ? lay.w - rowPadX : lay.w, h: rowH,
            text: _truncateText(String(displayVal), 20),
            font_family: bodyFont, font_size: valueFontSize, bold: true,
            color: isHighlighted ? highlightTextColor : bodyTextColor, align: alignVal, valign: 'middle'
          })
        } else {
          // text column **” first text col is the entity label (bold), others are annotations
          const isLabel = col === textCols[0]
          const isLastCol = ci === colHeaders.length - 1
          blocks.push({
            block_type: 'text_box',
            x: lay.x + rowPadX, y, w: lay.w - rowPadX - (isLastCol ? rowPadX : 0), h: rowH,
            text: _truncateText(String(cellValue), isLabel ? 34 : 38),
            font_family: bodyFont,
            font_size: isLabel ? labelFontSize : annotationFontSize,
            bold: isLabel,
            color: isHighlighted
              ? (isLabel ? highlightTextColor : highlightBarFill)
              : (isLabel ? bodyTextColor : annotationColor),
            align: 'left', valign: 'middle'
          })
        }
      }
    })
    return
  }

}

// Blend a hex colour with white at `ratio` intensity (0=white, 1=full colour)
function _hexTint(hex, ratio) {
  const h = String(hex || '').replace(/^#/, '')
  if (h.length !== 6) return hex
  const r = parseInt(h.slice(0,2), 16), g = parseInt(h.slice(2,4), 16), b = parseInt(h.slice(4,6), 16)
  const tr = Math.round(255 + (r - 255) * ratio)
  const tg = Math.round(255 + (g - 255) * ratio)
  const tb = Math.round(255 + (b - 255) * ratio)
  return '#' + [tr,tg,tb].map(x => x.toString(16).padStart(2,'0')).join('')
}

// matrix renders as a 2Ã—2 grid:
// - Quadrant fill color per tone (positive/negative/neutral) **” brand-derived tints
// - Dashed center dividers
// - Quadrant title (bold) + primary_message axis descriptor (per-tone text color); no secondary_message
// - Each point: circle badge (â‰¤2 chars) or pill badge (>2 chars); size varies with emphasis
// - Outer border rendered last (no fill) so rounded corners cleanly frame all fills
function _matrixToBlocks(art, content_y, blocks, bt, r2) {
  const ms = art.matrix_style || {}
  const xAxis = art.x_axis || {}
  const yAxis = art.y_axis || {}
  const quadrants = art.quadrants || []
  const points = art.points || []

  const ax = art.x || 0
  const ay = content_y
  const aw = art.w || 0
  const ah = r2((art.y || 0) + (art.h || 0) - content_y)
  if (aw <= 0 || ah <= 0) return

  // ****** Layout bands ******************************************************************************************************************************************************************************
  // Outer axis labels are NOT rendered **” low_label/high_label are self-descriptive.
  // leftBand and bottomBand are kept as minimal margins only (no rotated label / no bottom label).
  const leftBand  = 0.14
  const bottomBand= 0.16
  const topPad    = 0.02
  const rightPad  = 0.04

  const gridX = r2(ax + leftBand)
  const gridY = r2(ay + topPad)
  const gridW = r2(Math.max(1.6, aw - leftBand - rightPad))
  const gridH = r2(Math.max(1.4, ah - bottomBand - topPad))
  const midX  = r2(gridX + gridW / 2)
  const midY  = r2(gridY + gridH / 2)
  const quadW = r2(gridW / 2)
  const quadH = r2(gridH / 2)

  // ****** Brand tokens ******************************************************************************************************************************************************************************
  const axisFont      = ms.axis_label_font_family || bt.body_font_family   || 'Arial'
  const titleFont     = ms.quadrant_title_font_family || bt.title_font_family || 'Arial'
  const bodyFont      = ms.quadrant_body_font_family  || bt.body_font_family  || 'Arial'
  const axisFs        = ms.axis_label_font_size || 9
  const axisTextColor = ms.axis_label_color || bt.caption_color || bt.body_color || '#6B7280'

  // Per-tone helpers **” fills derived from brand tokens at low intensity (~10% tint)
  const brandPos  = bt.primary_color   || '#1E40AF'
  const brandNeg  = '#DC2626'
  const brandNeut = bt.secondary_color || '#6B7280'
  const toneQuadFill   = t => t === 'positive' ? (ms.positive_quadrant_fill || _hexTint(brandPos,  0.10))
                            : t === 'negative' ? (ms.negative_quadrant_fill || _hexTint(brandNeg,  0.10))
                            :                    (ms.neutral_quadrant_fill  || _hexTint(brandNeut, 0.08))
  const toneTitleColor = t => t === 'positive' ? (ms.positive_title_color || brandPos)
                            : t === 'negative' ? (ms.negative_title_color || brandNeg)
                            :                    (ms.neutral_title_color  || bt.body_color || '#374151')
  const toneBodyColor  = t => t === 'positive' ? (ms.positive_body_color || brandPos)
                            : t === 'negative' ? (ms.negative_body_color || brandNeg)
                            :                    (ms.neutral_body_color  || bt.body_color || '#374151')
  const tonePointFill  = t => t === 'positive' ? (ms.positive_point_fill || brandPos)
                            : t === 'negative' ? (ms.negative_point_fill || brandNeg)
                            :                    (ms.neutral_point_fill  || brandNeut)

  // ****** Quadrant data lookup *********************************************************************************************************************************************************
  // q1=top-left, q2=top-right, q3=bottom-left, q4=bottom-right
  const quadMap = Object.fromEntries(quadrants.map(q => [String(q.id || '').toLowerCase(), q]))
  const quadDefs = [
    { id: 'q1', x: gridX, y: gridY },
    { id: 'q2', x: midX,  y: gridY },
    { id: 'q3', x: gridX, y: midY  },
    { id: 'q4', x: midX,  y: midY  }
  ]

  // ****** 1. White grid background (no border **” border drawn last) ***************************************
  blocks.push({
    block_type: 'rect',
    x: gridX, y: gridY, w: gridW, h: gridH,
    fill_color: '#FFFFFF',
    border_color: null, border_width: 0, corner_radius: 8
  })

  // ****** 2. Quadrant fills **” drawn over white background ************************************************************************
  const quadFillDefs = [
    { id: 'q1', x: gridX, y: gridY,  w: quadW, h: quadH },  // top-left outer corner
    { id: 'q2', x: midX,  y: gridY,  w: r2(gridW - quadW), h: quadH },  // top-right
    { id: 'q3', x: gridX, y: midY,   w: quadW, h: r2(gridH - quadH) },  // bottom-left
    { id: 'q4', x: midX,  y: midY,   w: r2(gridW - quadW), h: r2(gridH - quadH) }  // bottom-right
  ]
  quadFillDefs.forEach((def, idx) => {
    const q    = quadMap[def.id] || quadrants[idx] || {}
    const tone = String(q.tone || 'neutral').toLowerCase()
    blocks.push({
      block_type: 'rect',
      x: def.x, y: def.y, w: def.w, h: def.h,
      fill_color: toneQuadFill(tone),
      border_color: null, border_width: 0, corner_radius: 0
    })
  })

  // ****** 3. Center dividers **” thin dashed lines *********************************************************************************************
  const divColor = ms.divider_color || '#AAAAAA'
  const divW     = ms.divider_width != null ? ms.divider_width : 0.5
  // Vertical center divider
  blocks.push({
    block_type: 'line',
    x1: midX, y1: gridY, x2: midX, y2: r2(gridY + gridH),
    color: divColor, line_width: divW, line_style: 'dashed'
  })
  // Horizontal center divider
  blocks.push({
    block_type: 'line',
    x1: gridX, y1: midY, x2: r2(gridX + gridW), y2: midY,
    color: divColor, line_width: divW, line_style: 'dashed'
  })

  // ****** 4a. Axis mid-labels (at the divider crosshair) ******************************************************************
  // Y-axis high/low labels at the vertical center divider
  if (yAxis.high_label) {
    blocks.push({
      block_type: 'text_box',
      x: r2(midX + 0.06), y: r2(gridY + 0.04), w: r2(quadW - 0.12), h: 0.18,
      text: yAxis.high_label,
      font_family: axisFont, font_size: axisFs, bold: false,
      color: axisTextColor, align: 'left', valign: 'middle'
    })
  }
  if (yAxis.low_label) {
    blocks.push({
      block_type: 'text_box',
      x: r2(midX + 0.06), y: r2(gridY + gridH - 0.22), w: r2(quadW - 0.12), h: 0.18,
      text: yAxis.low_label,
      font_family: axisFont, font_size: axisFs, bold: false,
      color: axisTextColor, align: 'left', valign: 'middle'
    })
  }
  // X-axis low/high labels at the horizontal center divider
  // Centered on midY (y = midY - h/2) so the label straddles the divider and exactly touches Q3/Q4
  if (xAxis.low_label) {
    blocks.push({
      block_type: 'text_box',
      x: r2(gridX + 0.06), y: r2(midY - 0.09), w: r2(quadW - 0.12), h: 0.18,
      text: xAxis.low_label,
      font_family: axisFont, font_size: axisFs, bold: false,
      color: axisTextColor, align: 'left', valign: 'middle'
    })
  }
  if (xAxis.high_label) {
    blocks.push({
      block_type: 'text_box',
      x: r2(midX + 0.06), y: r2(midY - 0.09), w: r2(quadW - 0.12), h: 0.18,
      text: xAxis.high_label,
      font_family: axisFont, font_size: axisFs, bold: false,
      color: axisTextColor, align: 'right', valign: 'middle'
    })
  }

  // ****** Outer axis labels suppressed ******************************************************************************************************************************
  // xAxis.label and yAxis.label are metadata only.
  // low_label / high_label are self-descriptive and carry the axis name **” no outer label needed.

  // ****** 4b. Quadrant labels **” anchored at the far corner from the centre crosshair
  // q1 top-left  â†’ outer corner = top-left   â†’ stack downward, left-aligned
  // q2 top-right â†’ outer corner = top-right  â†’ stack downward, right-aligned
  // q3 bot-left  â†’ outer corner = bottom-left â†’ stack upward from bottom, left-aligned
  // q4 bot-right â†’ outer corner = bottom-rightâ†’ stack upward from bottom, right-aligned
  const qPad    = 0.12   // inner margin from the outer corner
  const titleFs = ms.quadrant_title_font_size || 11
  const msgFs   = Math.max(7, titleFs - 2)    // 2pt smaller than title
  const titleH  = 0.22
  const msgH    = 0.18
  const txtGap  = 0.04   // gap between title and message
  const txtW    = r2(quadW - qPad * 2)

  const quadLabelLayout = {
    q1: { align: 'left',  titleY: q => r2(q.y + qPad),                              msgY: q => r2(q.y + qPad + titleH + txtGap) },
    q2: { align: 'right', titleY: q => r2(q.y + qPad),                              msgY: q => r2(q.y + qPad + titleH + txtGap) },
    q3: { align: 'left',  titleY: q => r2(q.y + quadH - qPad - msgH - txtGap - titleH), msgY: q => r2(q.y + quadH - qPad - msgH) },
    q4: { align: 'right', titleY: q => r2(q.y + quadH - qPad - msgH - txtGap - titleH), msgY: q => r2(q.y + quadH - qPad - msgH) }
  }

  quadDefs.forEach((def, idx) => {
    const q    = quadMap[def.id] || quadrants[idx] || {}
    const tone = String(q.tone || 'neutral').toLowerCase()
    const tc   = toneTitleColor(tone)
    const bc   = toneBodyColor(tone)
    const layout = quadLabelLayout[def.id] || quadLabelLayout.q1
    const txtX = r2(def.x + qPad)

    // Title **” bold, at the far corner
    blocks.push({
      block_type: 'text_box',
      x: txtX, y: layout.titleY(def), w: txtW, h: titleH,
      text: q.name || '',
      font_family: titleFont, font_size: titleFs, bold: true,
      color: tc, align: layout.align, valign: 'top'
    })
    // Primary message **” 2pt smaller, directly below the title
    if (q.primary_message) {
      blocks.push({
        block_type: 'text_box',
        x: txtX, y: layout.msgY(def), w: txtW, h: msgH,
        text: q.primary_message,
        font_family: bodyFont, font_size: msgFs, bold: false,
        color: bc, align: layout.align, valign: 'top'
      })
    }
  })

  // ****** 5. Points: badge (circle or pill) + label bubble ***************************************************************
  // ALL badges use IDENTICAL size and font regardless of emphasis:
  //   circle for short_label â‰¤ 2 chars; pill (rounded rect) for 3+ chars
  const BADGE_H  = 0.22          // fixed height for every badge
  const BADGE_FS = 9             // fixed font size for every badge label
  const BADGE_CR = Math.round(BADGE_H * 36)  // corner_radius â†’ fully rounded ends

  points.slice(0, 6).forEach(pt => {
    const xRatio = Math.min(Math.max((typeof pt.x === 'number' ? pt.x : 50) / 100, 0.02), 0.98)
    const yRatio = Math.min(Math.max((typeof pt.y === 'number' ? pt.y : 50) / 100, 0.02), 0.98)
    const px = r2(gridX + gridW * xRatio)
    const py = r2(gridY + gridH * (1 - yRatio))

    const ptQId  = pt.quadrant_id
      ? String(pt.quadrant_id).toLowerCase()
      : (xRatio < 0.5 && yRatio >= 0.5 ? 'q1' : xRatio >= 0.5 && yRatio >= 0.5 ? 'q2' : xRatio < 0.5 && yRatio < 0.5 ? 'q3' : 'q4')
    const ptQ    = quadMap[ptQId] || {}
    const ptTone = String(ptQ.tone || 'neutral').toLowerCase()
    const dotFill= tonePointFill(ptTone)

    const lbl  = String(pt.label || '')
    const sLbl = String(pt.short_label || (() => {
      const words = lbl.trim().split(/\s+/)
      return words.length >= 2 ? (words[0][0] + words[1][0]).toUpperCase() : lbl.slice(0,2).toUpperCase()
    })())

    const useCircle = sLbl.length <= 2
    const bW = useCircle ? BADGE_H : r2(Math.max(BADGE_H * 1.5, sLbl.length * 0.095 + 0.14))

    if (useCircle) {
      blocks.push({
        block_type: 'circle',
        x: r2(px - BADGE_H / 2), y: r2(py - BADGE_H / 2), w: BADGE_H, h: BADGE_H,
        fill_color: dotFill, font_color: '#FFFFFF',
        text: _truncateText(sLbl, 2)
      })
    } else {
      blocks.push({
        block_type: 'rect',
        x: r2(px - bW / 2), y: r2(py - BADGE_H / 2), w: bW, h: BADGE_H,
        fill_color: dotFill, border_color: null, border_width: 0,
        corner_radius: BADGE_CR
      })
      blocks.push({
        block_type: 'text_box',
        x: r2(px - bW / 2 + 0.04), y: r2(py - BADGE_H / 2),
        w: r2(bW - 0.08), h: BADGE_H,
        text: _truncateText(sLbl, 3),
        font_family: ms.point_label_font_family || bt.body_font_family || 'Arial',
        font_size: BADGE_FS, bold: true,
        color: '#FFFFFF', align: 'center', valign: 'middle'
      })
    }

    // Label bubble below the badge
    const bubbleW = r2(Math.min(1.2, Math.max(0.52, lbl.length * 0.075 + 0.20)))
    const bubbleH = 0.24
    const bubbleY = r2(py + BADGE_H / 2 + 0.05)
    let   bubbleX = r2(px - bubbleW / 2)
    bubbleX = r2(Math.max(gridX + 0.04, Math.min(bubbleX, gridX + gridW - bubbleW - 0.04)))
    const clampedBubbleY = r2(Math.min(bubbleY, gridY + gridH - bubbleH - 0.04))
    blocks.push({
      block_type: 'rect',
      x: bubbleX, y: clampedBubbleY, w: bubbleW, h: bubbleH,
      fill_color: '#FFFFFF', border_color: dotFill, border_width: 0.8, corner_radius: 10
    })
    blocks.push({
      block_type: 'text_box',
      x: r2(bubbleX + 0.05), y: clampedBubbleY,
      w: r2(bubbleW - 0.10), h: bubbleH,
      text: _truncateText(lbl, 18),
      font_family: ms.point_label_font_family || bt.body_font_family || 'Arial',
      font_size: ms.point_label_font_size || 9, bold: false,
      color: dotFill, align: 'center', valign: 'middle'
    })
  })
  // No outer border rect **” grid framing is provided by the quadrant fills and dividers only.
}

function _driverTreeToBlocks(art, content_y, blocks, bt, r2) {
  const ts = art.tree_style || {}
  const root = art.root || {}
  const branches = art.branches || []
  const ax = art.x || 0
  const ay = content_y
  const aw = art.w || 0
  const ah = r2((art.y || 0) + (art.h || 0) - content_y)

  const leafCount = branches.reduce((sum, b) => sum + Math.max((b.children || []).length, 1), 0) || Math.max(branches.length, 1)
  const hasThirdLevel = branches.some(b => (b.children || []).length > 0)
  const rowY = [
    r2(ay + 0.06),
    r2(ay + ah * 0.40),
    r2(ay + ah * 0.73)
  ]
  const rootW = r2(Math.min(Math.max(2.8, aw * 0.42), 3.9))
  const rootH = r2(Math.min(Math.max(0.9, ah * 0.20), 1.2))
  const branchW = r2(Math.min(Math.max(2.2, aw * 0.30), 3.0))
  const branchH = r2(Math.min(Math.max(0.8, ah * 0.18), 1.05))
  const leafW = r2(Math.min(Math.max(1.6, aw * 0.20), 2.2))
  const leafH = r2(Math.min(Math.max(0.7, ah * 0.16), 0.95))

  const rootX  = r2(ax + (aw - rootW) / 2)
  const rootY  = rowY[0]
  // connector_width is in pts â†’ convert to inches; 0.02" â‰ˆ 1.44pt (default thin connector)
  const connW  = r2(ts.connector_width != null ? ts.connector_width / 72 : 0.02)
  const connHW = r2(connW / 2)  // half-width for centering offsets

  const leafCenters = []
  if (leafCount === 1) {
    leafCenters.push(r2(ax + aw / 2))
  } else {
    const left = ax + leafW / 2
    const usable = Math.max(0.5, aw - leafW)
    const step = usable / (leafCount - 1)
    for (let i = 0; i < leafCount; i++) leafCenters.push(r2(left + i * step))
  }

  let cursor = 0
  const branchLayout = branches.map((branch) => {
    const childCount = Math.max((branch.children || []).length, 1)
    const branchLeafCenters = leafCenters.slice(cursor, cursor + childCount)
    cursor += childCount
    const centerX = branchLeafCenters.length
      ? r2(branchLeafCenters.reduce((s, x) => s + x, 0) / branchLeafCenters.length)
      : r2(ax + aw / 2)
    return {
      branch,
      centerX,
      children: (branch.children || []).length
        ? branch.children.map((child, ci) => ({ child, centerX: branchLeafCenters[ci] }))
        : [{ child: null, centerX }]
    }
  })

  const pushNode = (x, y, w, h, fill, label, value, isRoot) => {
    blocks.push({
      block_type: 'rect',
      x, y, w, h,
      fill_color: fill,
      border_color: ts.node_border_color || '#D7DEE8',
      border_width: ts.node_border_width != null ? ts.node_border_width : 0.6,
      corner_radius: ts.corner_radius != null ? ts.corner_radius : 6
    })
    const labelH = value ? r2(h * 0.52) : r2(h * 0.72)
    blocks.push({
      block_type: 'text_box',
      x: r2(x + 0.1), y: r2(y + 0.08), w: r2(w - 0.2), h: labelH,
      text: label || '',
      font_family: ts.label_font_family || bt.title_font_family || 'Arial',
      font_size: isRoot ? (ts.label_font_size || 11) + 1 : (ts.label_font_size || 11),
      bold: false,
      color: ts.label_color || '#111111',
      align: 'center',
      valign: 'top'
    })
    if (value) {
      blocks.push({
        block_type: 'text_box',
        x: r2(x + 0.1), y: r2(y + h * 0.58), w: r2(w - 0.2), h: r2(h * 0.22),
        text: value,
        font_family: ts.value_font_family || bt.body_font_family || 'Arial',
        font_size: isRoot ? (ts.value_font_size || 10) + 2 : (ts.value_font_size || 10) + 1,
        bold: true,
        color: ts.value_color || bt.primary_color || '#0078AE',
        align: 'center',
        valign: 'middle'
      })
    }
  }

  const pushConnector = (x, y, w, h) => {
    blocks.push({
      block_type: 'rect',
      x: r2(x), y: r2(y), w: r2(w), h: r2(h),
      fill_color: ts.connector_color || '#7A8FA8',
      border_color: null,
      border_width: 0,
      corner_radius: 0
    })
  }

  pushNode(rootX, rootY, rootW, rootH, ts.node_fill_color || '#EAF2FB', root.label, root.value, true)

  const rootBottomX = r2(rootX + rootW / 2)
  const branchY = rowY[1]
  const branchBottomY = r2(branchY + branchH)
  const branchCenters = branchLayout.map(b => b.centerX)
  if (branchCenters.length) {
    const trunkBottomY = r2(branchY - 0.14)
    pushConnector(r2(rootBottomX - connHW), r2(rootY + rootH), connW, r2(trunkBottomY - (rootY + rootH)))
    pushConnector(Math.min(...branchCenters), r2(trunkBottomY - connHW), Math.max(connW, Math.max(...branchCenters) - Math.min(...branchCenters)), connW)
  }

  branchLayout.forEach((entry) => {
    const bx = r2(entry.centerX - branchW / 2)
    pushNode(bx, branchY, branchW, branchH, ts.node_fill_color_secondary || '#EDF7F3', entry.branch.label, entry.branch.value, false)
    pushConnector(r2(entry.centerX - connHW), r2(branchY - 0.14), connW, 0.14)

    if (!hasThirdLevel) return
    const childCenters = entry.children.map(c => c.centerX)
    const childY = rowY[2]
    const childTopY = childY
    pushConnector(r2(entry.centerX - connHW), branchBottomY, connW, r2((childY - 0.14) - branchBottomY))
    if (childCenters.length > 1) {
      pushConnector(Math.min(...childCenters), r2(childY - connHW), Math.max(connW, Math.max(...childCenters) - Math.min(...childCenters)), connW)
    }

    entry.children.forEach(({ child, centerX }) => {
      const lx = r2(centerX - leafW / 2)
      pushConnector(r2(centerX - connHW), r2(childY - 0.15), connW, 0.15)
      const label = child ? child.label : entry.branch.label
      const value = child ? child.value : entry.branch.value
      pushNode(lx, childTopY, leafW, leafH, ts.node_fill_color_leaf || '#F4F7FA', label, value, false)
    })
  })
}

// Returns rankFill if it has sufficient contrast against white (ratio >= 3),
// otherwise returns a dark fallback so the badge text is always legible.
function _badgeTextColor(rankFill, darkFallback) {
  try {
    const hex = String(rankFill || '#000000').replace(/^#/, '')
    const r = parseInt(hex.slice(0, 2), 16) / 255
    const g = parseInt(hex.slice(2, 4), 16) / 255
    const b = parseInt(hex.slice(4, 6), 16) / 255
    const lum = 0.2126 * r + 0.7152 * g + 0.0722 * b
    return (1.05 / (lum + 0.05)) >= 3.0 ? rankFill : (darkFallback || '#1F2937')
  } catch (e) { return darkFallback || '#1F2937' }
}

function _prioritizationToBlocks(art, content_y, blocks, bt, r2) {
  const ps = art.priority_style || {}
  const items = (art.items || []).slice().sort((a, b) => (+a.rank || 999) - (+b.rank || 999)).slice(0, 5)
  const ax = art.x || 0
  const ay = content_y
  const aw = art.w || 0
  const ah = r2((art.y || 0) + (art.h || 0) - content_y)
  if (!items.length || aw <= 0 || ah <= 0) return

  const estimateTextHeight = (text, widthIn, fontSizePt, lineHeight = 1.25) => {
    const textStr = String(text || '').trim()
    if (!textStr) return 0
    const usableWidth = Math.max(0.3, Number(widthIn) || 0.3)
    const fontSize = Math.max(7, Number(fontSizePt) || 10)
    const charsPerLine = Math.max(8, Math.floor((usableWidth * 72) / (fontSize * 0.52)))
    const words = textStr.split(/\s+/).filter(Boolean)
    let lines = 1
    let lineLen = 0
    for (const word of words) {
      const nextLen = lineLen === 0 ? word.length : lineLen + 1 + word.length
      if (nextLen <= charsPerLine) lineLen = nextLen
      else { lines += 1; lineLen = word.length }
    }
    return lines * (fontSize * lineHeight / 72)
  }

  const rowGap    = ps.row_gap_in != null ? ps.row_gap_in : 0.16
  const bottomPad = 0.08   // breathing room below last row before next artifact header
  const rowH      = r2((ah - bottomPad - rowGap * Math.max(0, items.length - 1)) / Math.max(items.length, 1))
  const badgeW = r2(Math.min(0.88, Math.max(0.62, aw * 0.11)))
  const rightPad = 0.14
  const cr    = ps.row_corner_radius != null ? ps.row_corner_radius : 6

  const rankPalette = ps.rank_palette || [bt.secondary_color || '#E0B324', bt.primary_color || '#0078AE']
  const qualifierPalette = ps.qualifier_value_palette || [bt.primary_color || '#0078AE']
  const baseTitleFs = ps.title_font_size || 14
  const baseDescFs = ps.description_font_size || 11
  const baseQualifierFs = ps.qualifier_label_font_size || 10
  const rankLabels = ['CRITICAL', 'HIGH', 'MEDIUM', 'LOW', 'LOW']
  const numFontSize = Math.max(10, Math.min(18, Math.round(rowH * 30)))
  const rowPad = 0.10
  const minTitleDescGap = 0.06

  // Pre-compute per-item layout data (qualifier widths drive textW)
  const itemData = items.map((item) => {
    const qualifiers = Array.isArray(item.qualifiers) ? item.qualifiers.slice(0, 2) : []
    const nonEmptyQualifiers = qualifiers.filter(q => String(q?.label || '').trim() || String(q?.value || '').trim())
    const qualifierTexts = nonEmptyQualifiers.map(q => {
      const label = String(q.label || '').trim()
      const value = String(q.value || '').trim()
      return label && value ? (label + ': ' + value) : (label || value)
    })
    const longestQualifier = qualifierTexts.reduce((maxLen, text) => Math.max(maxLen, String(text || '').length), 0)
    const qualifierAreaW = nonEmptyQualifiers.length
      ? r2(Math.min(2.55, Math.max(1.55, aw * 0.28, longestQualifier * 0.055)))
      : 0
    const textX = r2(ax + badgeW + 0.14)
    const textW = r2(Math.max(1.1, aw - (textX - ax) - qualifierAreaW - rightPad - (qualifierAreaW ? 0.14 : 0)))
    const qualifierX = qualifierAreaW ? r2(ax + aw - rightPad - qualifierAreaW) : 0
    return { nonEmptyQualifiers, qualifierTexts, qualifierAreaW, textX, textW, qualifierX }
  })

  // Global title font: same across all rows **” shrink until every title fits
  const minTextW = Math.min(...itemData.map(d => d.textW))
  const maxTitleH = Math.max(0.30, rowH * 0.42)
  let globalTitleFs = baseTitleFs
  for (const item of items) {
    while (globalTitleFs > 10 && estimateTextHeight(String(item.title || ''), minTextW, globalTitleFs, 1.22) > maxTitleH) {
      globalTitleFs--
    }
  }

  // Global desc font: same across all rows **” capped at 70% of badge number font
  const maxDescFs = Math.min(baseDescFs, Math.round(numFontSize * 0.7))
  const maxDescH = Math.max(0.20, rowH * 0.42)
  let globalDescFs = maxDescFs
  for (const item of items) {
    while (globalDescFs > 8 && estimateTextHeight(String(item.description || ''), minTextW, globalDescFs, 1.24) > maxDescH) {
      globalDescFs--
    }
  }

  items.forEach((item, idx) => {
    const rowY = r2(ay + idx * (rowH + rowGap))
    const { nonEmptyQualifiers, qualifierTexts, qualifierAreaW, textX, textW, qualifierX } = itemData[idx]
    const rankFill = rankPalette[idx % Math.max(rankPalette.length, 1)] || bt.primary_color || '#0078AE'
    const titleText = String(item.title || '')
    const descText = String(item.description || '')
    const rankLabel = rankLabels[Math.min(idx, rankLabels.length - 1)]
    const rankNum = String(item.rank != null ? item.rank : idx + 1)

    // Determine layout: vertical (title above desc) vs horizontal (title | desc side-by-side)
    const titleH = r2(Math.max(0.20, estimateTextHeight(titleText, textW, globalTitleFs, 1.22) + 0.04))
    const descH  = r2(Math.max(0.18, estimateTextHeight(descText,  textW, globalDescFs,  1.24) + 0.04))
    const contentH = titleH + minTitleDescGap + descH
    const availContentH = rowH - 2 * rowPad
    const useHorizontalLayout = descText && contentH > availContentH + 0.05

    // Row background
    blocks.push({
      block_type: 'rect',
      x: ax, y: rowY, w: aw, h: rowH,
      fill_color: ps.row_fill_color || '#FFFFFF',
      border_color: ps.row_border_color || '#D7DEE8',
      border_width: ps.row_border_width != null ? ps.row_border_width : 0.6,
      corner_radius: cr
    })

    // === BADGE - 3 steps ===
    // Step 1: rank-colored rect, rounded corners
    // Step 2: white rect shifted right by stripW, colored strip shows on left
    // Step 3: “#N” + rank label centered as a block, color contrasts with rankFill on white
    const badgeStripW   = 0.10
    const labelFontSize = Math.max(6, Math.min(9, Math.round(rowH * 13)))
    const textAreaX     = r2(ax + badgeStripW)
    const textAreaW     = r2(badgeW - badgeStripW)
    const badgeTextColor = _badgeTextColor(rankFill, bt.body_color || '#1F2937')  // FIX C: contrast

    // Step 1: rank-colored background rect
    blocks.push({
      block_type: 'rect',
      x: ax, y: rowY, w: badgeW, h: rowH,
      fill_color: rankFill, border_color: null, border_width: 0, corner_radius: cr
    })
    // Step 2: white rect - same size, shifted right so only a strip of color shows
    blocks.push({
      block_type: 'rect',
      x: textAreaX, y: rowY, w: badgeW, h: rowH,
      fill_color: '#FFFFFF', border_color: null, border_width: 0, corner_radius: cr
    })

    // FIX B: center “#N” + label as a single vertically-centered block
    const numLineH   = r2(numFontSize   / 72 * 1.35 + 0.04)
    const labelLineH = r2(labelFontSize / 72 * 1.35 + 0.02)
    const badgeInnerGap  = 0.03
    const badgeTotalH    = numLineH + badgeInnerGap + labelLineH
    const badgeTextStartY = r2(rowY + (rowH - badgeTotalH) / 2)

    // Step 3a: “#N”
    blocks.push({
      block_type: 'text_box',
      x: textAreaX, y: badgeTextStartY, w: textAreaW, h: numLineH,
      text: '#' + rankNum,
      font_family: ps.rank_font_family || bt.title_font_family || 'Arial',
      font_size: numFontSize, bold: true,
      color: badgeTextColor, align: 'center', valign: 'middle'
    })
    // Step 3b: rank label
    blocks.push({
      block_type: 'text_box',
      x: textAreaX, y: r2(badgeTextStartY + numLineH + badgeInnerGap), w: textAreaW, h: labelLineH,
      text: rankLabel,
      font_family: ps.rank_font_family || bt.title_font_family || 'Arial',
      font_size: labelFontSize, bold: false,
      color: badgeTextColor, align: 'center', valign: 'middle'
    })

    // FIX A: Title + Description layout
    if (useHorizontalLayout) {
      // Side-by-side: title (~40%) | gap | description (~rest), both vertically centred in row
      const hGap      = 0.12
      const titlePartW = r2(textW * 0.40)
      const descPartW  = r2(textW - titlePartW - hGap)
      const titleHoriz = r2(Math.max(0.20, estimateTextHeight(titleText, titlePartW, globalTitleFs, 1.22) + 0.04))
      const descHoriz  = r2(Math.max(0.18, estimateTextHeight(descText,  descPartW,  globalDescFs,  1.24) + 0.04))
      blocks.push({
        block_type: 'text_box',
        x: textX, y: r2(rowY + (rowH - titleHoriz) / 2), w: titlePartW, h: titleHoriz,
        text: titleText,
        font_family: ps.title_font_family || bt.title_font_family || 'Arial',
        font_size: globalTitleFs, bold: true,
        color: ps.title_color || '#1F2937',
        align: 'left', valign: 'middle'
      })
      blocks.push({
        block_type: 'text_box',
        x: r2(textX + titlePartW + hGap), y: r2(rowY + (rowH - descHoriz) / 2), w: descPartW, h: descHoriz,
        text: descText,
        font_family: ps.description_font_family || bt.body_font_family || 'Arial',
        font_size: globalDescFs, bold: false,
        color: ps.description_color || '#374151',
        align: 'left', valign: 'middle'
      })
    } else {
      // Vertical: title above description, block centred in row
      const contentStartY = r2(rowY + rowPad + Math.max(0, (availContentH - contentH) / 2))
      const titleY     = contentStartY
      const descY      = r2(contentStartY + titleH + minTitleDescGap)
      const descAvailH = r2(Math.max(descH, rowY + rowH - rowPad - descY))
      blocks.push({
        block_type: 'text_box',
        x: textX, y: titleY, w: textW, h: titleH,
        text: titleText,
        font_family: ps.title_font_family || bt.title_font_family || 'Arial',
        font_size: globalTitleFs, bold: true,
        color: ps.title_color || '#1F2937',
        align: 'left', valign: 'top'
      })
      if (descText) {
        blocks.push({
          block_type: 'text_box',
          x: textX, y: descY, w: textW, h: descAvailH,
          text: descText,
          font_family: ps.description_font_family || bt.body_font_family || 'Arial',
          font_size: globalDescFs, bold: false,
          color: ps.description_color || '#374151',
          align: 'left', valign: 'top'
        })
      }
    }

    // Qualifier pills: try stacked first; fall back to side-by-side if overflow
    if (nonEmptyQualifiers.length) {
      const pillGap = 0.08
      const pillPad = 0.12

      // Compute stacked pill heights
      const stackedPillData = nonEmptyQualifiers.map((_, qi) => {
        const pillTextW = Math.max(0.4, qualifierAreaW - 0.16)
        let fs = baseQualifierFs
        while (fs > 8 && estimateTextHeight(qualifierTexts[qi], pillTextW, fs, 1.2) > Math.max(0.36, rowH * 0.26)) fs--
        const pillH = r2(Math.max(0.26, estimateTextHeight(qualifierTexts[qi], pillTextW, fs, 1.2) + 0.10))
        return { fs, pillH }
      })
      const stackedTotalH = stackedPillData.reduce((s, d) => s + d.pillH, 0) + pillGap * Math.max(0, nonEmptyQualifiers.length - 1)
      const useSideBySide = nonEmptyQualifiers.length === 2 && stackedTotalH > (rowH - 2 * pillPad)

      if (useSideBySide) {
        // Side-by-side: two pills arranged horizontally, vertically centered in row
        const halfW = r2((qualifierAreaW - pillGap) / 2)
        const pillTextW = Math.max(0.3, halfW - 0.12)
        const sidePillH = r2(Math.max(0.24, rowH * 0.48))
        const pillY = r2(rowY + (rowH - sidePillH) / 2)
        nonEmptyQualifiers.forEach((_, qi) => {
          const pillX = r2(qualifierX + qi * (halfW + pillGap))
          let fs = baseQualifierFs
          while (fs > 7 && estimateTextHeight(qualifierTexts[qi], pillTextW, fs, 1.2) > sidePillH - 0.08) fs--
          const valueColor = qualifierPalette[qi % Math.max(qualifierPalette.length, 1)] || bt.primary_color || '#0078AE'
          blocks.push({
            block_type: 'rect',
            x: pillX, y: pillY, w: halfW, h: sidePillH,
            fill_color: valueColor, border_color: null, border_width: 0, corner_radius: 4
          })
          blocks.push({
            block_type: 'text_box',
            x: r2(pillX + 0.06), y: pillY, w: r2(halfW - 0.12), h: sidePillH,
            text: qualifierTexts[qi],
            font_family: ps.qualifier_label_font_family || bt.body_font_family || 'Arial',
            font_size: fs, bold: false,
            color: ps.qualifier_text_color || '#1F2937',
            align: 'center', valign: 'middle'
          })
        })
      } else {
        // Stacked: pills arranged vertically, starting from top padding
        let pillCursorY = r2(rowY + pillPad)
        nonEmptyQualifiers.forEach((_, qi) => {
          const { fs, pillH } = stackedPillData[qi]
          const valueColor = qualifierPalette[qi % Math.max(qualifierPalette.length, 1)] || bt.primary_color || '#0078AE'
          const maxPillBottom = rowY + rowH - pillPad
          const pillY = r2(Math.min(pillCursorY, Math.max(rowY + pillPad, maxPillBottom - pillH)))
          blocks.push({
            block_type: 'rect',
            x: qualifierX, y: pillY, w: qualifierAreaW, h: pillH,
            fill_color: valueColor, border_color: null, border_width: 0, corner_radius: 4
          })
          blocks.push({
            block_type: 'text_box',
            x: r2(qualifierX + 0.08), y: pillY, w: r2(qualifierAreaW - 0.16), h: pillH,
            text: qualifierTexts[qi],
            font_family: ps.qualifier_label_font_family || bt.body_font_family || 'Arial',
            font_size: fs, bold: false,
            color: ps.qualifier_text_color || '#1F2937',
            align: 'center', valign: 'middle'
          })
          pillCursorY = r2(pillY + pillH + pillGap)
        })
      }
    }
  })
}

function _estimateLegendTextWidth(label, fontSizePt) {
  const text = String(label || '')
  return Math.max(0.40, Math.min(2.20, text.length * Math.max(fontSizePt, 8) * 0.0105))
}

function _chartLegendEntries(chartType, categories, seriesData, seriesStyles, palette, allowFallback, secondarySeriesData) {
  const entries = []
  // pie, donut, and group_pie: legend represents SLICES (categories), colored per series_style[i]
  if (chartType === 'pie' || chartType === 'donut' || chartType === 'group_pie') {
    ;(categories || []).forEach((category, i) => {
      const style = i < (seriesStyles || []).length ? seriesStyles[i] : {}
      let color = style.fill_color
      if (!color && allowFallback) color = palette[i % Math.max(palette.length, 1)]
      entries.push({ label: String(category || ''), color: color || '#666666' })
    })
    return entries
  }
  // For combo charts, include both primary (bar) and secondary (line) series
  const allSeries = (chartType === 'combo' && secondarySeriesData?.length)
    ? [...(seriesData || []), ...secondarySeriesData]
    : (seriesData || [])
  allSeries.forEach((series, i) => {
    const style = i < (seriesStyles || []).length ? seriesStyles[i] : {}
    // Secondary series (lines) use line_color; primary series (bars) use fill_color
    const isPrimary = i < (seriesData || []).length
    let color = isPrimary ? style.fill_color : (style.line_color || style.fill_color)
    if (!color && allowFallback) color = palette[i % Math.max(palette.length, 1)]
    entries.push({ label: String(series?.name || ('Series ' + (i + 1))), color: color || '#666666' })
  })
  return entries
}

function _computeChartLegendLayout(x, y, w, h, legendPosition, legendEntries, fontSizePt) {
  if (!legendEntries.length || !['top', 'right'].includes(String(legendPosition || ''))) {
    return { chartRect: { x, y, w, h }, legendBox: null }
  }

  const swatch = 0.14
  const textGap = 0.06
  const itemGapX = 0.18
  const rowGap = 0.06
  const lineH = Math.max(0.20, fontSizePt * 0.022)
  const padX = 0.04
  const padY = 0.03

  if (legendPosition === 'top') {
    const rows = []
    let current = []
    let usedW = 0
    const maxRowW = Math.max(0.5, w - 0.04)

    legendEntries.forEach(entry => {
      const itemW = swatch + textGap + _estimateLegendTextWidth(entry.label, fontSizePt)
      const proposed = current.length === 0 ? itemW : usedW + itemGapX + itemW
      const item = { ...entry, item_w: itemW }
      if (current.length && proposed > maxRowW) {
        rows.push(current)
        current = [item]
        usedW = itemW
      } else {
        current.push(item)
        usedW = proposed
      }
    })
    if (current.length) rows.push(current)

    const legendH = Math.min(Math.max(padY * 2 + rows.length * lineH + Math.max(0, rows.length - 1) * rowGap, 0.28), h * 0.28)
    return {
      chartRect: { x, y: r2(y + legendH + 0.05), w, h: r2(Math.max(1.0, h - legendH - 0.05)) },
      legendBox: { position: 'top', x, y, w, h: legendH, rows, line_h: lineH, pad_x: padX, pad_y: padY, swatch, text_gap: textGap, item_gap_x: itemGapX, row_gap: rowGap }
    }
  }

  const items = legendEntries.map(entry => {
    const itemW = swatch + textGap + _estimateLegendTextWidth(entry.label, fontSizePt)
    return { ...entry, item_w: itemW }
  })
  const maxItemW = items.reduce((maxW, item) => Math.max(maxW, item.item_w), 0)
  const legendW = Math.min(Math.max(maxItemW + padX * 2, 1.05), w * 0.38)
  const chartW = Math.max(1.0, w - legendW - 0.08)
  return {
    chartRect: { x, y, w: r2(chartW), h },
    legendBox: { position: 'right', x: r2(x + chartW + 0.08), y, w: r2(legendW), h, items, line_h: lineH, pad_x: padX, pad_y: padY, swatch, text_gap: textGap, item_gap_x: itemGapX, row_gap: rowGap }
  }
}

function _chartLegendToBlocks(legendBox, fontFamily, fontSizePt, colorHex, blocks, r2) {
  if (!legendBox) return
  if (legendBox.position === 'top') {
    let curY = r2(legendBox.y + legendBox.pad_y)
    ;(legendBox.rows || []).forEach(row => {
      const rowW = row.reduce((sum, item) => sum + item.item_w, 0) + Math.max(0, row.length - 1) * legendBox.item_gap_x
      let curX = r2(legendBox.x + Math.max(legendBox.pad_x, (legendBox.w - rowW) / 2))
      row.forEach(item => {
        blocks.push({
          block_type: 'rect',
          x: curX,
          y: r2(curY + Math.max(0, (legendBox.line_h - legendBox.swatch) / 2)),
          w: legendBox.swatch,
          h: legendBox.swatch,
          fill_color: item.color,
          border_color: null,
          border_width: 0,
          corner_radius: 0
        })
        blocks.push({
          block_type: 'text_box',
          x: r2(curX + legendBox.swatch + legendBox.text_gap),
          y: curY,
          w: r2(Math.max(0.35, item.item_w - legendBox.swatch - legendBox.text_gap)),
          h: legendBox.line_h,
          text: item.label,
          font_family: fontFamily,
          font_size: fontSizePt,
          bold: false,
          color: colorHex,
          align: 'left',
          valign: 'middle'
        })
        curX = r2(curX + item.item_w + legendBox.item_gap_x)
      })
      curY = r2(curY + legendBox.line_h + legendBox.row_gap)
    })
    return
  }

  let curY = r2(legendBox.y + legendBox.pad_y)
  ;(legendBox.items || []).forEach(item => {
    blocks.push({
      block_type: 'rect',
      x: r2(legendBox.x + legendBox.pad_x),
      y: r2(curY + Math.max(0, (legendBox.line_h - legendBox.swatch) / 2)),
      w: legendBox.swatch,
      h: legendBox.swatch,
      fill_color: item.color,
      border_color: null,
      border_width: 0,
      corner_radius: 0
    })
    blocks.push({
      block_type: 'text_box',
      x: r2(legendBox.x + legendBox.pad_x + legendBox.swatch + legendBox.text_gap),
      y: curY,
      w: r2(Math.max(0.35, legendBox.w - legendBox.pad_x * 2 - legendBox.swatch - legendBox.text_gap)),
      h: legendBox.line_h,
      text: item.label,
      font_family: fontFamily,
      font_size: fontSizePt,
      bold: false,
      color: colorHex,
      align: 'left',
      valign: 'middle'
    })
    curY = r2(curY + legendBox.line_h + legendBox.row_gap)
  })
}

// workflow renders all four subtypes:
//   process_flow **” horizontal linear sequence, value above/description below nodes
//   timeline     **” same layout as process_flow + horizontal baseline with dot markers
//   hierarchy    **” top_down_branching tree, level-based node fills, description below
//   decomposition**” top_down_branching breakdown, level-based node fills, description below
// All connector segments emit arrowheads when conn.type === 'arrow'.
function _workflowToBlocks(art, content_y, blocks, bt, r2) {
  const ws    = art.workflow_style || {}
  const nodes = Array.isArray(art.nodes)       ? art.nodes       : []
  const conns = Array.isArray(art.connections) ? art.connections : []
  if (!nodes.length) return

  // ****** Brand tokens ************************************************************************************************************************************************************************************
  const titleFont        = ws.node_title_font_family || bt.title_font_family || 'Arial'
  const valueFont        = ws.node_value_font_family || bt.body_font_family  || 'Arial'
  const nodeFill         = ws.node_fill_color           || bt.primary_color  || '#0078AE'
  const nodeFillSecond   = ws.node_fill_color_secondary || bt.secondary_color || '#3A6EA5'
  const nodeFillLeaf     = ws.node_fill_color_leaf      || '#EAF2FB'
  const nodeBorder       = ws.node_border_color    || '#FFFFFF'
  const nodeBorderWidth  = ws.node_border_width    != null ? ws.node_border_width    : 1
  const nodeCornerRadius = ws.node_corner_radius   != null ? ws.node_corner_radius   : 4
  const titleColorDark   = ws.node_title_color     || '#FFFFFF'    // for dark fills (level 1 & 2)
  const titleColorLeaf   = ws.node_title_color_leaf || bt.body_color || '#111111'  // for light fills (level 3+)
  const valueColor       = ws.node_value_color     || bt.body_color || '#111111'
  const connColor        = ws.connector_color      || bt.primary_color || '#0078AE'
  const connWidth        = ws.connector_width      != null ? ws.connector_width      : 0.5
  const innerPad         = ws.node_inner_padding   != null ? ws.node_inner_padding   : 0.08
  const externalGap      = ws.external_label_gap   != null ? ws.external_label_gap   : 0.08
  const titleFs          = ws.node_title_font_size || 10
  const valueFs          = ws.node_value_font_size || 9

  // ****** Helpers ***************************************************************************************************************************************************************************************************
  const estimateTextHeight = (text, widthIn, fontSizePt, lineHeight = 1.2) => {
    const textStr = String(text || '').trim()
    if (!textStr) return 0
    const usableWidth = Math.max(0.3, Number(widthIn) || 0.3)
    const fontSize    = Math.max(7, Number(fontSizePt) || 10)
    const charsPerLine= Math.max(8, Math.floor((usableWidth * 72) / (fontSize * 0.52)))
    const words = textStr.split(/\s+/).filter(Boolean)
    let lines = 1, lineLen = 0
    for (const word of words) {
      const nextLen = lineLen === 0 ? word.length : lineLen + 1 + word.length
      if (nextLen <= charsPerLine) lineLen = nextLen
      else { lines += 1; lineLen = word.length }
    }
    return lines * (fontSize * lineHeight / 72)
  }

  // Resolve per-level node fill and text color
  const nodeStyle = level => {
    if (level <= 1) return { fill: nodeFill,       text: titleColorDark }
    if (level === 2) return { fill: nodeFillSecond, text: titleColorDark }
    return                 { fill: nodeFillLeaf,    text: titleColorLeaf }
  }

  // ****** Flow type detection ***************************************************************************************************************************************************************
  const flowDir     = String(art.flow_direction || '').toLowerCase()
  const wType       = String(art.workflow_type  || '').toLowerCase()
  const isHorizFlow = flowDir === 'left_to_right' || flowDir === 'horizontal'
    || wType === 'timeline' || wType === 'roadmap' || wType === 'process_flow'
  const isVertFlow  = flowDir === 'top_to_bottom' || flowDir === 'bottom_up'
  const isBranching = flowDir === 'top_down_branching'
    || wType === 'hierarchy' || wType === 'decomposition'
  const isTimeline  = wType === 'timeline'

  // ****** Connector segments (drawn BEFORE nodes so nodes render on top) ******************************
  conns.forEach(conn => {
    const path      = Array.isArray(conn.path) ? conn.path : []
    const isArrow   = String(conn.type || 'arrow').toLowerCase() === 'arrow'
    // Emit one `line` block per path segment; only the LAST segment gets the arrowhead
    for (let i = 0; i < path.length - 1; i++) {
      const p1 = path[i], p2 = path[i + 1]
      if (p1?.x == null || p1?.y == null || p2?.x == null || p2?.y == null) continue
      const isLastSeg = (i === path.length - 2)
      blocks.push({
        block_type: 'line',
        x1: r2(p1.x), y1: r2(p1.y), x2: r2(p2.x), y2: r2(p2.y),
        x:  r2(Math.min(p1.x, p2.x)),
        y:  r2(Math.min(p1.y, p2.y)),
        w:  r2(Math.max(Math.abs(p2.x - p1.x), 0.02)),
        h:  r2(Math.max(Math.abs(p2.y - p1.y), 0.02)),
        color:      connColor,
        line_width: connWidth,
        arrowhead:  isArrow && isLastSeg   // arrowhead only on the final segment
      })
    }
  })

  // ****** Timeline baseline (before nodes so nodes render on top) ***************************************************
  // Draw a horizontal bar at node-bottom height, spanning all phase nodes,
  // with a small filled dot at each node center and an arrowhead at the right end.
  if (isTimeline && nodes.length >= 2) {
    const validNodes = nodes.filter(n => (n.w || 0) > 0 && (n.h || 0) > 0)
    if (validNodes.length >= 2) {
      const baselineY = r2(validNodes[0].y + (validNodes[0].h || 0.6))
      const lineX1    = r2(Math.min(...validNodes.map(n => n.x || 0)))
      const lineX2    = r2(Math.max(...validNodes.map(n => (n.x || 0) + (n.w || 0.8))))
      const lineColor = ws.timeline_line_color || connColor
      blocks.push({
        block_type: 'line',
        x1: lineX1, y1: baselineY, x2: lineX2, y2: baselineY,
        x: lineX1, y: baselineY, w: r2(lineX2 - lineX1), h: 0.02,
        color: lineColor, line_width: connWidth + 0.3, arrowhead: true
      })
      // Dot markers at each phase midpoint on the baseline
      const dotR = 0.06
      validNodes.forEach(n => {
        const dotX = r2((n.x || 0) + (n.w || 0.8) / 2 - dotR)
        blocks.push({
          block_type: 'circle',
          x: dotX, y: r2(baselineY - dotR), w: r2(dotR * 2), h: r2(dotR * 2),
          fill_color: lineColor, text: ''
        })
      })
    }
  }

  // ****** Node rendering ******************************************************************************************************************************************************************************
  nodes.forEach(node => {
    const nx    = r2(node.x || 0)
    const ny    = r2(node.y || content_y)
    const nw    = r2(node.w || 0.8)
    const nh    = r2(node.h || 0.6)
    const level = node.level != null ? Number(node.level) : 1
    const innerW= r2(Math.max(0.3, nw - innerPad * 2))

    const titleText = String(node.label || node.title || node.id || '')
    const valueText = String(node.value || '').trim()
    const descText  = String(node.description || '').trim()

    // Level-based fill for hierarchy/decomposition; flat fill for process_flow/timeline
    const ns = (isBranching) ? nodeStyle(level) : { fill: nodeFill, text: titleColorDark }

    // Node box
    blocks.push({
      block_type:    'rect',
      x: nx, y: ny, w: nw, h: nh,
      fill_color:    ns.fill,
      border_color:  nodeBorder,
      border_width:  nodeBorderWidth,
      corner_radius: nodeCornerRadius
    })

    // Label inside node (always)
    blocks.push({
      block_type:  'text_box',
      x: r2(nx + innerPad), y: r2(ny + innerPad), w: innerW, h: r2(nh - innerPad * 2),
      text:        titleText,
      font_family: titleFont, font_size: titleFs, bold: true,
      color:       ns.text, align: 'center', valign: 'middle'
    })

    if (isHorizFlow) {
      // process_flow / timeline: value ABOVE, description BELOW
      if (valueText) {
        const valueH = r2(Math.max(0.16, estimateTextHeight(valueText, nw, valueFs, 1.15) + 0.04))
        blocks.push({
          block_type:  'text_box',
          x: nx, y: r2(ny - valueH - externalGap), w: nw, h: valueH,
          text:        valueText,
          font_family: valueFont, font_size: valueFs, bold: false,
          color:       valueColor, align: 'center', valign: 'bottom'
        })
      }
      if (descText) {
        const descH = r2(Math.max(0.18, estimateTextHeight(descText, nw, valueFs, 1.18) + 0.04))
        blocks.push({
          block_type:  'text_box',
          x: nx, y: r2(ny + nh + externalGap), w: nw, h: descH,
          text:        descText,
          font_family: valueFont, font_size: valueFs, bold: false,
          color:       valueColor, align: 'center', valign: 'top'
        })
      }
    } else if (isVertFlow) {
      // top_to_bottom / bottom_up: description to the RIGHT of the box
      if (descText) {
        const ax    = art.x || 0
        const aw    = art.w || 0
        const rightX= r2(nx + nw + externalGap)
        const rightW= r2(Math.max(0.3, ax + aw - rightX))
        const descH = r2(Math.max(0.18, estimateTextHeight(descText, rightW, valueFs, 1.18) + 0.04))
        blocks.push({
          block_type:  'text_box',
          x: rightX, y: r2(ny + (nh - descH) / 2), w: rightW, h: descH,
          text:        descText,
          font_family: valueFont, font_size: valueFs, bold: false,
          color:       valueColor, align: 'left', valign: 'middle'
        })
      }
    } else {
      // hierarchy / decomposition (top_down_branching): description below the node box
      if (descText) {
        const descH = r2(Math.max(0.18, estimateTextHeight(descText, nw, valueFs, 1.18) + 0.04))
        blocks.push({
          block_type:  'text_box',
          x: nx, y: r2(ny + nh + externalGap), w: nw, h: descH,
          text:        descText,
          font_family: valueFont, font_size: valueFs, bold: false,
          color:       valueColor, align: 'center', valign: 'top'
        })
      }
      // Also render value text inside the node as a secondary line (hierarchy uses node space)
      if (valueText && nh > 0.5) {
        blocks.push({
          block_type:  'text_box',
          x: r2(nx + innerPad), y: r2(ny + nh / 2), w: innerW, h: r2(nh / 2 - innerPad),
          text:        valueText,
          font_family: valueFont, font_size: Math.max(7, valueFs - 1), bold: false,
          color:       ns.text, align: 'center', valign: 'middle'
        })
      }
    }
  })
}

function _comparisonTableToBlocks(art, content_y, blocks, bt, r2) {
  const cs = art.comparison_style || art.table_style || {}
  const columns = Array.isArray(art.columns) ? art.columns : []
  const rows    = Array.isArray(art.rows)    ? art.rows    : []
  const ax = art.x || 0
  const ay = content_y
  const aw = art.w || 0
  const ah = r2((art.y || 0) + (art.h || 0) - content_y)
  if (!columns.length || !rows.length || aw <= 0 || ah <= 0) return

  const numCols  = columns.length
  const numRows  = rows.length
  const titleFont = bt.title_font_family || 'Arial'
  const bodyFont  = bt.body_font_family  || 'Arial'

  // Column widths: first col ~28% (option name), remainder split equally
  const rowGap      = 0.04
  const colPad      = 0.10
  const col0Frac    = cs.first_col_frac || 0.28
  const col0W       = r2(aw * col0Frac)
  const dataCols    = Math.max(1, numCols - 1)
  const dataColW    = r2((aw - col0W) / dataCols)
  const headerH     = cs.header_height || 0.34
  const rowH        = r2((ah - headerH - rowGap * Math.max(0, numRows - 1)) / Math.max(numRows, 1))
  const iconSize    = Math.min(0.26, rowH * 0.65)

  const colX = [ax]
  for (let ci = 1; ci < numCols; ci++) colX.push(r2(ax + col0W + (ci - 1) * dataColW))
  const colW = [col0W, ...Array(dataCols).fill(dataColW)]

  // Tone → pill colours
  const toneClr = {
    positive: { fill: cs.positive_fill || '#D1FAE5', text: cs.positive_text || '#065F46' },
    negative: { fill: cs.negative_fill || '#FEE2E2', text: cs.negative_text || '#991B1B' },
    warning:  { fill: cs.warning_fill  || '#FEF3C7', text: cs.warning_text  || '#92400E' },
    neutral:  { fill: cs.neutral_fill  || '#F3F4F6', text: cs.neutral_text  || '#374151' },
  }
  const getToneClr = t => toneClr[t] || toneClr.neutral
  const iconClr = {
    positive: { fill: '#D1FAE5', icon: '#065F46' },
    negative: { fill: '#FEE2E2', icon: '#991B1B' },
    warning:  { fill: '#FEF3C7', icon: '#92400E' },
    neutral:  { fill: '#F3F4F6', icon: '#374151' },
  }
  const getIconClr = t => iconClr[t] || iconClr.neutral

  // ── Header row ────────────────────────────────────────────────
  const headerFill     = cs.header_fill      || bt.primary_color || '#0078AE'
  const headerTextClr  = cs.header_text_color || '#FFFFFF'
  const headerFs       = cs.header_font_size  || 10
  blocks.push({ block_type: 'rect', x: ax, y: ay, w: aw, h: headerH,
    fill_color: headerFill, border_color: null, border_width: 0, corner_radius: 6 })
  columns.forEach((col, ci) => {
    blocks.push({ block_type: 'text_box',
      x: r2(colX[ci] + (ci === 0 ? colPad : colPad * 0.5)), y: ay,
      w: r2(colW[ci] - colPad), h: headerH,
      text: String(col || ''), font_family: titleFont,
      font_size: headerFs, bold: true, color: headerTextClr,
      align: ci === 0 ? 'left' : 'center', valign: 'middle' })
  })

  // ── Data rows ─────────────────────────────────────────────────
  const recBorderClr = cs.recommended_border_color || bt.secondary_color || '#E0B324'
  const recFill      = cs.recommended_fill         || '#FFFBEA'
  const dividerClr   = cs.divider_color            || '#E5E7EB'
  const rowFill      = cs.row_fill                 || '#FFFFFF'
  const altRowFill   = cs.alt_row_fill             || '#F8FAFC'

  rows.forEach((row, ri) => {
    const rowY  = r2(ay + headerH + ri * (rowH + rowGap))
    const isRec = !!row.is_recommended
    blocks.push({ block_type: 'rect', x: ax, y: rowY, w: aw, h: rowH,
      fill_color: isRec ? recFill : (ri % 2 === 0 ? rowFill : altRowFill),
      border_color: isRec ? recBorderClr : dividerClr,
      border_width: isRec ? 1.2 : 0.5, corner_radius: 4 })
    // Recommended accent strip on left edge
    if (isRec) {
      blocks.push({ block_type: 'rect', x: ax, y: rowY, w: 0.05, h: rowH,
        fill_color: recBorderClr, border_color: null, border_width: 0, corner_radius: 4 })
    }

    const cells = Array.isArray(row.cells) ? row.cells : []
    cells.forEach((cell, ci) => {
      if (ci >= numCols) return
      const cx      = colX[ci]
      const cw      = colW[ci]
      const tone    = cell.tone || 'neutral'
      const value   = cell.value != null ? String(cell.value) : ''
      const iconType = String(cell.icon_type || '')
      const subtext  = String(cell.subtext   || '')

      if (ci === 0) {
        // Option name — bold, no pill
        const nameFs = cs.option_name_font_size || 11
        const nameH  = r2(rowH > 0.5 ? rowH * 0.52 : rowH)
        const nameY  = r2(rowY + (rowH - nameH) / 2)
        blocks.push({ block_type: 'text_box',
          x: r2(cx + colPad), y: nameY, w: r2(cw - colPad - 0.04), h: nameH,
          text: value, font_family: titleFont,
          font_size: nameFs, bold: true,
          color: cs.option_name_color || '#1F2937', align: 'left', valign: 'middle' })
        if (subtext) {
          blocks.push({ block_type: 'text_box',
            x: r2(cx + colPad), y: r2(nameY + nameH * 0.55), w: r2(cw - colPad - 0.04), h: r2(nameH * 0.45),
            text: subtext, font_family: bodyFont,
            font_size: cs.subtext_font_size || 9, bold: false,
            color: '#6B7280', align: 'left', valign: 'top' })
        }
      } else if (iconType) {
        // Icon badge
        const ic  = getIconClr(tone)
        const bx  = r2(cx + (cw - iconSize) / 2)
        const by  = r2(rowY + (rowH - iconSize) / 2 - (subtext ? 0.10 : 0))
        blocks.push({ block_type: 'icon_badge', x: bx, y: by, w: iconSize, h: iconSize,
          icon: iconType, fill_color: ic.fill, icon_color: ic.icon })
        if (subtext) {
          blocks.push({ block_type: 'text_box',
            x: r2(cx + colPad * 0.5), y: r2(by + iconSize + 0.02), w: r2(cw - colPad), h: 0.16,
            text: subtext, font_family: bodyFont,
            font_size: cs.subtext_font_size || 8, bold: false,
            color: ic.icon, align: 'center', valign: 'top' })
        }
      } else if (value) {
        // Value pill
        const tc     = getToneClr(tone)
        const pillW  = r2(Math.min(cw - colPad, Math.max(0.50, value.length * 0.068 + 0.20)))
        const pillH  = Math.min(0.26, rowH * 0.55)
        const pillX  = r2(cx + (cw - pillW) / 2)
        const pillY  = r2(rowY + (rowH - pillH) / 2 - (subtext ? 0.10 : 0))
        const valueFs = cs.value_font_size || 10
        blocks.push({ block_type: 'rect', x: pillX, y: pillY, w: pillW, h: pillH,
          fill_color: tc.fill, border_color: null, border_width: 0, corner_radius: 10 })
        blocks.push({ block_type: 'text_box',
          x: r2(pillX + 0.06), y: pillY, w: r2(pillW - 0.12), h: pillH,
          text: value, font_family: bodyFont,
          font_size: valueFs, bold: true, color: tc.text,
          align: 'center', valign: 'middle' })
        if (subtext) {
          blocks.push({ block_type: 'text_box',
            x: r2(cx + colPad * 0.5), y: r2(pillY + pillH + 0.03), w: r2(cw - colPad), h: 0.16,
            text: subtext, font_family: bodyFont,
            font_size: cs.subtext_font_size || 8, bold: false,
            color: tc.text, align: 'center', valign: 'top' })
        }
      }
    })
  })
}

function _initiativeMapToBlocks(art, content_y, blocks, bt, r2) {
  const ms = art.initiative_style || art.table_style || {}
  const columnHeaders = Array.isArray(art.column_headers) ? art.column_headers : []
  const rows          = Array.isArray(art.rows)           ? art.rows           : []
  const ax = art.x || 0
  const ay = content_y
  const aw = art.w || 0
  const ah = r2((art.y || 0) + (art.h || 0) - content_y)
  if (!columnHeaders.length || !rows.length || aw <= 0 || ah <= 0) return

  const numCols   = columnHeaders.length
  const numRows   = rows.length
  const titleFont = bt.title_font_family || 'Arial'
  const bodyFont  = bt.body_font_family  || 'Arial'
  const primaryClr   = bt.primary_color   || '#0078AE'
  const secondaryClr = bt.secondary_color || '#E0B324'

  const rowGap   = 0.04
  const colPad   = 0.10
  const headerH  = ms.header_height  || 0.34
  const rowH     = r2((ah - headerH - rowGap * Math.max(0, numRows - 1)) / Math.max(numRows, 1))

  // First column ~28%, rest equal
  const col0W    = r2(aw * (ms.first_col_frac || 0.28))
  const dataCols = Math.max(1, numCols - 1)
  const dataColW = r2((aw - col0W) / dataCols)
  const colX = [ax]
  for (let ci = 1; ci < numCols; ci++) colX.push(r2(ax + col0W + (ci - 1) * dataColW))
  const colW = [col0W, ...Array(dataCols).fill(dataColW)]

  // Tag chip colours (per-tag tone)
  const tagFills = { primary: primaryClr, secondary: secondaryClr, neutral: '#E5E7EB' }
  const tagTexts = { primary: '#FFFFFF',  secondary: '#1F2937',    neutral: '#374151' }

  // Cell background tint for non-neutral cell_tone
  const cellToneFills = {
    primary:   ms.primary_cell_fill   || '#EFF6FF',
    secondary: ms.secondary_cell_fill || '#FFFBEA',
  }

  // ── Header row ────────────────────────────────────────────────
  const headerFill    = ms.header_fill       || primaryClr
  const headerTextClr = ms.header_text_color || '#FFFFFF'
  const headerFs      = ms.header_font_size  || 10
  blocks.push({ block_type: 'rect', x: ax, y: ay, w: aw, h: headerH,
    fill_color: headerFill, border_color: null, border_width: 0, corner_radius: 6 })
  columnHeaders.forEach((col, ci) => {
    blocks.push({ block_type: 'text_box',
      x: r2(colX[ci] + (ci === 0 ? colPad : colPad * 0.5)), y: ay,
      w: r2(colW[ci] - colPad), h: headerH,
      text: col.label || '', font_family: titleFont,
      font_size: headerFs, bold: true, color: headerTextClr,
      align: 'left', valign: 'middle' })
  })

  // ── Data rows ─────────────────────────────────────────────────
  rows.forEach((row, ri) => {
    const rowY = r2(ay + headerH + ri * (rowH + rowGap))

    // Row background
    blocks.push({ block_type: 'rect', x: ax, y: rowY, w: aw, h: rowH,
      fill_color: ri % 2 === 0 ? (ms.row_fill || '#FFFFFF') : (ms.alt_row_fill || '#F9FAFB'),
      border_color: ms.row_border_color || '#E5E7EB', border_width: 0.5, corner_radius: 4 })

    // First cell: initiative_name + subtitle
    const name     = String(row.initiative_name     || '')
    const subtitle = String(row.initiative_subtitle || '')
    const nameFs   = ms.name_font_size     || 11
    const subFs    = ms.subtitle_font_size || 9
    if (name) {
      const hasSubtitle = !!subtitle
      blocks.push({ block_type: 'text_box',
        x: r2(ax + colPad), y: r2(rowY + (hasSubtitle ? rowH * 0.12 : 0)),
        w: r2(col0W - colPad * 1.5), h: r2(hasSubtitle ? rowH * 0.52 : rowH),
        text: name, font_family: titleFont, font_size: nameFs, bold: true,
        color: ms.name_color || '#1F2937', align: 'left',
        valign: hasSubtitle ? 'top' : 'middle' })
    }
    if (subtitle) {
      blocks.push({ block_type: 'text_box',
        x: r2(ax + colPad), y: r2(rowY + rowH * 0.55),
        w: r2(col0W - colPad * 1.5), h: r2(rowH * 0.38),
        text: subtitle, font_family: bodyFont, font_size: subFs, bold: false,
        color: ms.subtitle_color || '#6B7280', align: 'left', valign: 'top' })
    }

    // Build cell lookup by column_id
    const cellMap = {}
    ;(Array.isArray(row.cells) ? row.cells : []).forEach(c => { if (c?.column_id) cellMap[c.column_id] = c })

    // Data columns (ci >= 1)
    columnHeaders.slice(1).forEach((col, di) => {
      const ci    = di + 1
      const cx    = colX[ci]
      const cw    = colW[ci]
      const cell  = cellMap[col.id] || {}
      const tags  = Array.isArray(cell.tags) ? cell.tags.filter(t => String(t?.label || '').trim()) : []
      const primMsg = String(cell.primary_message   || '')
      const secMsg  = String(cell.secondary_message || '')
      const cellTone = cell.cell_tone || 'neutral'

      // Optional cell tint
      if (cellTone !== 'neutral' && cellToneFills[cellTone]) {
        blocks.push({ block_type: 'rect',
          x: r2(cx + 0.04), y: r2(rowY + 0.04), w: r2(cw - 0.08), h: r2(rowH - 0.08),
          fill_color: cellToneFills[cellTone], border_color: null, border_width: 0, corner_radius: 4 })
      }

      if (tags.length) {
        // Tags as coloured chips (primary visual signal)
        const chipH   = Math.min(0.22, rowH * 0.38)
        const chipGap = 0.06
        const chipPadX = 0.08
        const tagFs   = ms.tag_font_size || 8
        const maxTags = Math.min(3, tags.length)
        const chipWidths = tags.slice(0, maxTags).map(t =>
          r2(Math.min(cw - 0.18, Math.max(0.42, String(t.label || '').length * 0.065 + 0.18))))
        const totalW  = chipWidths.reduce((s, w) => s + w, 0) + chipGap * Math.max(0, maxTags - 1)
        let chipCurX  = r2(cx + Math.max(chipPadX, (cw - totalW) / 2))
        const chipsY  = r2(rowY + (secMsg ? rowH * 0.14 : (rowH - chipH) / 2))

        tags.slice(0, maxTags).forEach((tag, ti) => {
          const chipW = chipWidths[ti]
          const cFill = tagFills[tag.tone] || tagFills.neutral
          const cText = tagTexts[tag.tone] || tagTexts.neutral
          blocks.push({ block_type: 'rect',
            x: chipCurX, y: chipsY, w: chipW, h: chipH,
            fill_color: cFill, border_color: null, border_width: 0, corner_radius: 10 })
          blocks.push({ block_type: 'text_box',
            x: r2(chipCurX + chipPadX * 0.5), y: chipsY, w: r2(chipW - chipPadX), h: chipH,
            text: String(tag.label || ''), font_family: bodyFont,
            font_size: tagFs, bold: false, color: cText,
            align: 'center', valign: 'middle' })
          chipCurX = r2(chipCurX + chipW + chipGap)
        })

        // secondary_message below chips
        if (secMsg) {
          blocks.push({ block_type: 'text_box',
            x: r2(cx + colPad * 0.5), y: r2(chipsY + chipH + 0.04), w: r2(cw - colPad), h: r2(rowH * 0.36),
            text: secMsg, font_family: bodyFont,
            font_size: ms.secondary_font_size || 8, bold: false,
            color: ms.secondary_color || '#374151', align: 'left', valign: 'top' })
        }
      } else {
        // No tags — primary_message headline + optional secondary
        const primFs = ms.primary_font_size || 10
        const secFs  = ms.secondary_font_size || 8
        const hasBoth = primMsg && secMsg
        const primH   = hasBoth ? r2(rowH * 0.50) : rowH
        const primY   = hasBoth ? r2(rowY + rowH * 0.10) : rowY
        if (primMsg) {
          blocks.push({ block_type: 'text_box',
            x: r2(cx + colPad * 0.5), y: primY, w: r2(cw - colPad), h: primH,
            text: primMsg, font_family: bodyFont, font_size: primFs, bold: !!hasBoth,
            color: ms.primary_color_text || '#1F2937',
            align: 'left', valign: hasBoth ? 'top' : 'middle' })
        }
        if (secMsg) {
          blocks.push({ block_type: 'text_box',
            x: r2(cx + colPad * 0.5), y: r2(rowY + rowH * 0.54), w: r2(cw - colPad), h: r2(rowH * 0.38),
            text: secMsg, font_family: bodyFont, font_size: secFs, bold: false,
            color: ms.secondary_color || '#6B7280', align: 'left', valign: 'top' })
        }
      }
    })
  })
}

function _profileCardSetToBlocks(art, content_y, blocks, bt, r2) {
  const ps       = art.profile_style || {}
  const profiles = Array.isArray(art.profiles) ? art.profiles : []
  const ax = art.x || 0
  const ay = content_y
  const aw = art.w || 0
  const ah = r2((art.y || 0) + (art.h || 0) - content_y)
  if (!profiles.length || aw <= 0 || ah <= 0) return

  const titleFont = bt.title_font_family || 'Arial'
  const bodyFont  = bt.body_font_family  || 'Arial'
  const layoutDir = art.layout_direction || 'horizontal'
  const count     = profiles.length
  const cardGap   = ps.card_gap || 0.12

  // Card dimensions
  let cardW, cardH, cols
  if (layoutDir === 'grid') {
    cols  = Math.min(3, Math.ceil(Math.sqrt(count)))
    const gridRows = Math.ceil(count / cols)
    cardW = r2((aw - cardGap * (cols - 1)) / cols)
    cardH = r2((ah - cardGap * (gridRows - 1)) / gridRows)
  } else {
    cols  = count
    cardW = r2((aw - cardGap * Math.max(0, count - 1)) / Math.max(count, 1))
    cardH = ah
  }

  const cardFill  = ps.card_fill         || '#FFFFFF'
  const cardBorder = ps.card_border_color || '#E5E7EB'
  const headerFill = ps.header_fill       || bt.primary_color || '#0078AE'
  const headerH   = ps.header_height      || Math.min(0.60, cardH * 0.30)
  const cr        = ps.corner_radius      || 8

  const sentClr = {
    positive: { fill: '#D1FAE5', text: '#065F46' },
    negative: { fill: '#FEE2E2', text: '#991B1B' },
    warning:  { fill: '#FEF3C7', text: '#92400E' },
    neutral:  { fill: '#F3F4F6', text: '#374151' },
  }
  const getSentClr = s => sentClr[String(s || 'neutral').toLowerCase()] || sentClr.neutral

  profiles.forEach((profile, pi) => {
    const col = pi % cols
    const row = Math.floor(pi / cols)
    const cx  = r2(ax + col * (cardW + cardGap))
    const cy  = r2(ay + row * (cardH + cardGap))

    // Card background
    blocks.push({ block_type: 'rect', x: cx, y: cy, w: cardW, h: cardH,
      fill_color: cardFill, border_color: cardBorder, border_width: 0.5, corner_radius: cr })

    // Header strip (coloured top section)
    blocks.push({ block_type: 'rect', x: cx, y: cy, w: cardW, h: headerH,
      fill_color: headerFill, border_color: null, border_width: 0, corner_radius: cr })
    // Square off the bottom of the header so only top corners are rounded
    blocks.push({ block_type: 'rect',
      x: cx, y: r2(cy + headerH * 0.5), w: cardW, h: r2(headerH * 0.55),
      fill_color: headerFill, border_color: null, border_width: 0, corner_radius: 0 })

    const namePadX  = 0.14
    const badgeText = String(profile.badge_text || '')
    const nameW     = r2(cardW - namePadX * 2 - (badgeText ? 0.50 : 0))
    const nameFs    = ps.name_font_size || Math.max(10, Math.min(14, Math.round(headerH * 22)))
    const hasSubtitle = !!(profile.subtitle)
    const nameH     = hasSubtitle ? r2(headerH * 0.55) : headerH

    blocks.push({ block_type: 'text_box',
      x: r2(cx + namePadX), y: cy, w: nameW, h: nameH,
      text: String(profile.entity_name || ''),
      font_family: titleFont, font_size: nameFs, bold: true,
      color: ps.name_color || '#FFFFFF', align: 'left', valign: hasSubtitle ? 'bottom' : 'middle' })

    if (hasSubtitle) {
      const subFs = ps.subtitle_font_size || Math.max(8, nameFs - 3)
      blocks.push({ block_type: 'text_box',
        x: r2(cx + namePadX), y: r2(cy + nameH), w: r2(cardW - namePadX * 2), h: r2(headerH - nameH),
        text: String(profile.subtitle || ''),
        font_family: bodyFont, font_size: subFs, bold: false,
        color: ps.subtitle_color || '#D1D5DB', align: 'left', valign: 'top' })
    }

    // Badge pill at top-right
    if (badgeText) {
      const badgeW = r2(Math.min(0.80, Math.max(0.40, badgeText.length * 0.075 + 0.18)))
      const badgeH = 0.22
      const badgeX = r2(cx + cardW - badgeW - 0.10)
      const badgeY = r2(cy + 0.08)
      blocks.push({ block_type: 'rect', x: badgeX, y: badgeY, w: badgeW, h: badgeH,
        fill_color: ps.badge_fill || bt.secondary_color || '#E0B324',
        border_color: null, border_width: 0, corner_radius: 10 })
      blocks.push({ block_type: 'text_box',
        x: r2(badgeX + 0.06), y: badgeY, w: r2(badgeW - 0.12), h: badgeH,
        text: badgeText, font_family: bodyFont,
        font_size: ps.badge_font_size || 8, bold: true,
        color: ps.badge_text_color || '#1F2937', align: 'center', valign: 'middle' })
    }

    // Body: secondary_items (or fallback attributes[])
    const rawItems = Array.isArray(profile.secondary_items) ? profile.secondary_items
      : (Array.isArray(profile.attributes) ? profile.attributes.map(a => ({
          label: a.key || '', value: a.value || '',
          representation_type: 'text', sentiment: a.sentiment || 'neutral'
        })) : [])

    if (rawItems.length) {
      const bodyTop  = r2(cy + headerH + 0.08)
      const bodyH    = r2(cardH - headerH - 0.10)
      const itemH    = r2(bodyH / rawItems.length)
      const labelFs  = ps.label_font_size || Math.max(7, Math.min(9, Math.round(itemH * 18)))
      const valueFs  = ps.value_font_size || Math.max(8, Math.min(10, Math.round(itemH * 20)))
      const padX     = 0.14
      const labelW   = r2(cardW * 0.38)
      const valueX   = r2(cx + padX + labelW + 0.06)
      const valueW   = r2(cardW - padX * 2 - labelW - 0.06)

      rawItems.forEach((item, ii) => {
        const iy       = r2(bodyTop + ii * itemH)
        const label    = String(item.label || '')
        const rawVal   = item.value
        const valueStr = Array.isArray(rawVal) ? rawVal.join(', ') : String(rawVal || '')
        const repType  = item.representation_type || 'text'
        const sc       = getSentClr(item.sentiment)

        // Row divider (skip first)
        if (ii > 0) {
          blocks.push({ block_type: 'rule',
            x: r2(cx + 0.10), y: r2(iy - 0.02), w: r2(cardW - 0.20), h: 0.005,
            color: '#E5E7EB', line_width: 0.5 })
        }

        // Label (left)
        blocks.push({ block_type: 'text_box',
          x: r2(cx + padX), y: r2(iy + 0.04), w: labelW, h: r2(itemH - 0.04),
          text: label, font_family: bodyFont, font_size: labelFs, bold: false,
          color: ps.label_color || '#6B7280', align: 'left', valign: 'middle' })

        // Value (right) — varies by representation_type
        if (repType === 'pill') {
          const pillH = Math.min(0.22, itemH * 0.62)
          const pillW = r2(Math.min(valueW, Math.max(0.40, valueStr.length * 0.065 + 0.18)))
          const pillX = r2(valueX)
          const pillY = r2(iy + (itemH - pillH) / 2)
          blocks.push({ block_type: 'rect', x: pillX, y: pillY, w: pillW, h: pillH,
            fill_color: sc.fill, border_color: null, border_width: 0, corner_radius: 10 })
          blocks.push({ block_type: 'text_box',
            x: r2(pillX + 0.06), y: pillY, w: r2(pillW - 0.12), h: pillH,
            text: valueStr, font_family: bodyFont,
            font_size: Math.max(7, valueFs - 1), bold: true, color: sc.text,
            align: 'center', valign: 'middle' })
        } else if (repType === 'chip_list') {
          const chips   = Array.isArray(rawVal) ? rawVal : valueStr.split(',').map(v => v.trim()).filter(Boolean)
          const chipH   = Math.min(0.20, itemH * 0.55)
          const chipGap = 0.05
          let chipCurX  = r2(valueX)
          chips.slice(0, 4).forEach(v => {
            const chipW = r2(Math.min(1.2, Math.max(0.36, v.length * 0.065 + 0.14)))
            if (chipCurX + chipW > cx + cardW - 0.10) return
            blocks.push({ block_type: 'rect',
              x: chipCurX, y: r2(iy + (itemH - chipH) / 2), w: chipW, h: chipH,
              fill_color: sc.fill, border_color: null, border_width: 0, corner_radius: 8 })
            blocks.push({ block_type: 'text_box',
              x: r2(chipCurX + 0.05), y: r2(iy + (itemH - chipH) / 2), w: r2(chipW - 0.10), h: chipH,
              text: v, font_family: bodyFont,
              font_size: Math.max(6, labelFs - 1), bold: false, color: sc.text,
              align: 'center', valign: 'middle' })
            chipCurX = r2(chipCurX + chipW + chipGap)
          })
        } else {
          // plain text
          blocks.push({ block_type: 'text_box',
            x: valueX, y: r2(iy + 0.04), w: valueW, h: r2(itemH - 0.04),
            text: valueStr, font_family: bodyFont, font_size: valueFs,
            bold: item.sentiment && item.sentiment !== 'neutral',
            color: item.sentiment && item.sentiment !== 'neutral' ? sc.text : (ps.value_color || '#1F2937'),
            align: 'left', valign: 'middle' })
        }
      })
    }
  })
}

function _artifactToBlocks(art, blocks, bt, r2, fontSizeFloor) {
  const ax = art.x || 0
  const ay = art.y || 0
  const aw = art.w || 0
  const ah = art.h || 0
  const blockStart = blocks.length

  // ****** Artifact header band (if present) ************************************************************************************************************
  // header_block sits above the artifact body, already has its own x/y/w/h
  const hb          = art.header_block || null
  let   content_y   = ay   // top of the body area (after header_block)

  if (hb && hb.text) {
    const hx  = hb.x  != null ? hb.x  : ax
    const hy  = hb.y  != null ? hb.y  : ay
    const hw  = hb.w  != null ? hb.w  : aw
    const hfs = hb.font_size || 11
    const estimatedH = estimateHeaderBlockHeight(hb.text, hw, hfs)
    const hh  = Math.max(hb.h != null ? hb.h : 0.30, estimatedH)
    const headerStyle = hb.style || 'underline'
    const headerRuleH = 0.005
    const headerGapBelow = 0.06
    content_y = r2(hy + hh + (headerStyle === 'underline' ? (headerRuleH + headerGapBelow) : headerGapBelow))

    if (headerStyle === 'brand_fill') {
      // Filled header band
      blocks.push({
        block_type:    'rect',
        x: hx, y: hy, w: hw, h: hh,
        fill_color:    hb.fill_color   || bt.primary_color || '#1A3C8F',
        border_color:  null,
        border_width:  0,
        corner_radius: hb.corner_radius || 0
      })
      blocks.push({
        block_type:  'text_box',
        x: r2(hx + 0.08), y: hy, w: r2(hw - 0.16), h: hh,
        text:        hb.text,
        font_family: hb.font_family || bt.title_font_family || 'Arial',
        font_size:   hfs,
        bold:        true,
        color:       hb.text_color || '#FFFFFF',
        align:       'left',
        valign:      'middle'
      })
    } else {
      // Underline header
      blocks.push({
        block_type:  'text_box',
        x: hx, y: hy, w: hw, h: hh,
        text:        hb.text,
        font_family: hb.font_family || bt.title_font_family || 'Arial',
        font_size:   hfs,
        bold:        true,
        color:       hb.color || bt.primary_color || '#1A3C8F',
        align:       'left',
        valign:      'top'
      })
      blocks.push({
        block_type:  'rule',
        x: hx, y: r2(hy + hh), w: hw, h: 0.005,
        color:       hb.rule_color || bt.primary_color || '#1A3C8F',
        line_width:  0.5
      })
    }
  }

  // ****** Artifact body ***************************************************************************************************************************************************************************
  const headerEnd = blocks.length
  switch (art.type) {

    case 'stat_bar': {
      _statBarToBlocks(art, content_y, blocks, bt, r2)
      break
    }

    case 'chart': {
      const computed = art._computed || {}
      const chartStyle = art.chart_style || {}
      const legendPos = computed.legend_position || chartStyle.legend_position || 'none'
      const legendFontSize = chartStyle.legend_font_size || 9
      const allowLegendFallback = (art.fallback_policy || {}).allow_renderer_fallback !== false
      const legendEntries = art.show_legend
        ? _chartLegendEntries(
            art.chart_type,
            art.categories || [],
            art.series || [],
            art.series_style || [],
            bt.chart_palette || [],
            allowLegendFallback,
            art.secondary_series || []
          )
        : []
      const legendLayout = _computeChartLegendLayout(ax, content_y, aw, r2(ay + ah - content_y), legendPos, legendEntries, legendFontSize)
      const chartRect = legendLayout.chartRect || { x: ax, y: content_y, w: aw, h: r2(ay + ah - content_y) }

      blocks.push({
        block_type:              'chart',
        x: chartRect.x, y: chartRect.y, w: chartRect.w, h: chartRect.h,
        chart_type:              art.chart_type,
        chart_header:            art.chart_header || art.artifact_header || '',
        chart_title:             art.chart_title  || '',
        categories:              art.categories   || [],
        series:                  art.series       || [],
        dual_axis:               art.dual_axis    || false,
        secondary_series:        art.secondary_series || [],
        show_data_labels:        art.show_data_labels !== false,
        show_legend:             false,
        x_label:                 art.x_label || '',
        y_label:                 art.y_label || '',
        secondary_y_label:       art.secondary_y_label || '',
        chart_style:             {
          ...chartStyle,
          legend_position: 'none'
        },
        series_style:            art.series_style  || [],
        // Pre-computed by computeArtifactInternals **” renderer reads these directly
        legend_position:         computed.legend_position        || 'none',
        data_label_size:         computed.data_label_size        || 9,
        category_label_rotation: computed.category_label_rotation || 0
      })
      const legendStart = blocks.length
      _chartLegendToBlocks(
        legendLayout.legendBox,
        chartStyle.legend_font_family || bt.body_font_family || 'Arial',
        legendFontSize,
        chartStyle.legend_color || bt.body_color || '#111111',
        blocks,
        r2
      )
      decorateArtifactBlocks(blocks, legendStart, blocks.length, art, 'artifact_body')
      break
    }

    case 'insight_text': {
      if (art.insight_mode === 'grouped') {
        _groupedInsightToBlocks(art, content_y, blocks, bt, r2, fontSizeFloor)
      } else {
        _standardInsightToBlocks(art, content_y, blocks, bt, r2, fontSizeFloor)
      }
      break
    }

    case 'table': {
      const tableY = content_y
      const tableH = r2(ay + ah - content_y)
      const colWs = art.column_widths || []
      const dataRowHs = art.row_heights || []
      const headerH = art.header_row_height || 0.35

      let curX = ax
      const columnXPositions = colWs.map(cw => {
        const x = r2(curX)
        curX += (+cw || 0)
        return x
      })

      let curY = tableY
      const rowYPositions = [r2(curY)]
      curY += (+headerH || 0)
      for (const rh of dataRowHs) {
        rowYPositions.push(r2(curY))
        curY += (+rh || 0)
      }

      const headerCellFrames = colWs.map((cw, ci) => ({
        col_index: ci,
        x: r2(columnXPositions[ci] || ax),
        y: r2(tableY),
        w: r2(+cw || 0),
        h: r2(+headerH || 0)
      }))

      const bodyCellFrames = dataRowHs.map((rh, ri) =>
        colWs.map((cw, ci) => ({
          row_index: ri,
          col_index: ci,
          x: r2(columnXPositions[ci] || ax),
          y: r2(rowYPositions[ri + 1] || tableY),
          w: r2(+cw || 0),
          h: r2(+rh || 0)
        }))
      )

      blocks.push({
        block_type:         'table',
        x: ax, y: tableY, w: aw, h: tableH,
        headers:            art.headers            || [],
        rows:               art.rows               || [],
        column_widths:      colWs,
        column_x_positions: columnXPositions,
        column_types:       art.column_types        || [],
        column_alignments:  art.column_alignments   || [],
        row_heights:        dataRowHs,
        header_row_height:  headerH,
        row_y_positions:    rowYPositions,
        header_cell_frames: headerCellFrames,
        body_cell_frames:   bodyCellFrames,
        table_style:        art.table_style         || {},
        table_fit_failed:   !!art._table_fit_failed
      })
      break
    }

    case 'comparison_table': {
      _comparisonTableToBlocks(art, content_y, blocks, bt, r2)
      break
    }

    case 'initiative_map': {
      _initiativeMapToBlocks(art, content_y, blocks, bt, r2)
      break
    }

    case 'profile_card_set': {
      _profileCardSetToBlocks(art, content_y, blocks, bt, r2)
      break
    }

    case 'risk_register': {
      _riskRegisterToBlocks(art, content_y, blocks, bt, r2)
      break
    }

    case 'matrix': {
      _matrixToBlocks(art, content_y, blocks, bt, r2)
      break
    }

    case 'driver_tree': {
      _driverTreeToBlocks(art, content_y, blocks, bt, r2)
      break
    }

    case 'prioritization': {
      _prioritizationToBlocks(art, content_y, blocks, bt, r2)
      break
    }

    case 'cards': {
      _cardsToBlocks(art, content_y, blocks, bt, r2)
      break
    }

    case 'workflow': {
      _workflowToBlocks(art, content_y, blocks, bt, r2)
      break
    }

    default:
      break
  }
  decorateArtifactBlocks(blocks, blockStart, headerEnd, art, 'artifact_header')
  decorateArtifactBlocks(blocks, headerEnd, blocks.length, art, 'artifact_body')
}

// Compute the standalone bullet font size an insight artifact would naturally use,
// without actually emitting blocks. Used for cross-artifact harmonisation.
function _computeInsightFontSize(art, content_y) {
  const r2    = x => Math.round(x * 100) / 100
  const mode  = art.insight_mode || 'standard'
  const ay    = art.y || 0
  const ah    = art.h || 0
  const aw    = art.w || 0

  if (mode !== 'grouped') {
    // standard — no background box ever
    const body_y     = r2(content_y)
    const body_h     = r2(Math.max(0.3, ay + ah - body_y))
    const padH   = 0.04
    const TOP_GAP = 0.02
    const innerH     = Math.max(0.2, body_h - TOP_GAP)
    const points     = art.points || []
    const nPoints    = Math.max(1, points.length)
    const st         = art.body_style || {}
    const avgChars   = points.reduce((s, p) => s + String(p?.text || p || '').length, 0) / nPoints
    const areaW      = Math.max(0.3, aw - 2 * padH)
    let fontSize = st.font_size || 10
    for (let tryFs = 18; tryFs >= Math.max(9, fontSize); tryFs--) {
      const linesEach = Math.max(1, Math.ceil(avgChars / Math.max(1, areaW * 72 / (tryFs * 0.56))))
      const lineH     = (tryFs / 72) * 1.3
      const nPoints2  = nPoints
      const estH      = nPoints2 * linesEach * lineH + (nPoints2 - 1) * 0.04
      if (estH <= innerH * 0.82) { fontSize = tryFs; break }
    }
    return fontSize
  } else {
    // grouped **” return the bullet font size (not the header font)
    const groups      = art.groups || []
    if (!groups.length) return 10
    const total_content_h = r2(ay + ah - content_y)
    const gLayout     = art.group_layout || 'rows'
    const g_gap       = art.group_gap_in || 0.08
    const hb_gap      = art.header_to_box_gap_in || 0.05
    const ghs         = art.group_header_style || {}
    const bsty        = art.bullet_style || {}
    const gbs         = art.group_bullet_box_style || {}
    const n           = groups.length

    let minBulletFs = 18
    if (gLayout === 'rows') {
      const h_w           = ghs.w || 1.2
      const box_w         = Math.max(0.3, aw - h_w - hb_gap)
      const total_bullets = Math.max(1, groups.reduce((s, g) => s + (g.bullets || []).length, 0))
      const total_rh      = Math.max(0.2, total_content_h - (n - 1) * g_gap)
      for (const g of groups) {
        const nb    = Math.max(1, (g.bullets || []).length)
        const row_h = r2(Math.max(0.25, total_rh * (nb / total_bullets)))
        const bPadV = (gbs.padding && gbs.padding.top)  || 0.08
        const bPadH = (gbs.padding && gbs.padding.left) || 0.10
        const bAreaW = Math.max(0.3, box_w - 2 * bPadH)
        const bAreaH = Math.max(0.1, row_h - 2 * bPadV)
        const bullets = g.bullets || []
        const avgC  = bullets.reduce((s, b) => s + String(b?.text || b || '').length, 0) / Math.max(1, bullets.length)
        let fs = bsty.font_size || 10
        for (let tryFs = 18; tryFs >= Math.max(9, fs); tryFs--) {
          const lines = Math.max(1, Math.ceil(avgC / Math.max(1, bAreaW * 72 / (tryFs * 0.56))))
          const estH  = bullets.length * lines * (tryFs / 72) * 1.3 + (bullets.length - 1) * 0.04
          if (estH <= bAreaH * 0.82) { fs = tryFs; break }
        }
        minBulletFs = Math.min(minBulletFs, fs)
      }
    } else {
      const col_w = r2((aw - (n - 1) * g_gap) / Math.max(n, 1))
      const h_h   = ghs.h || 0.28
      const box_h = r2(total_content_h - h_h - hb_gap)
      for (const g of groups) {
        const bPadV  = (gbs.padding && gbs.padding.top)  || 0.08
        const bPadH  = (gbs.padding && gbs.padding.left) || 0.10
        const bAreaW = Math.max(0.3, col_w - 2 * bPadH)
        const bAreaH = Math.max(0.1, box_h - 2 * bPadV)
        const bullets = g.bullets || []
        const avgC  = bullets.reduce((s, b) => s + String(b?.text || b || '').length, 0) / Math.max(1, bullets.length)
        let fs = bsty.font_size || 10
        for (let tryFs = 18; tryFs >= Math.max(9, fs); tryFs--) {
          const lines = Math.max(1, Math.ceil(avgC / Math.max(1, bAreaW * 72 / (tryFs * 0.56))))
          const estH  = bullets.length * lines * (tryFs / 72) * 1.3 + (bullets.length - 1) * 0.04
          if (estH <= bAreaH * 0.82) { fs = tryFs; break }
        }
        minBulletFs = Math.min(minBulletFs, fs)
      }
    }
    return minBulletFs
  }
}

function _standardInsightToBlocks(art, content_y, blocks, bt, r2, fontSizeFloor) {
  const ax = art.x || 0
  const ay = art.y || 0
  const aw = art.w || 0
  const ah = art.h || 0
  const st  = art.body_style || {}

  // No background box — render bullets directly on the slide background
  const body_y = r2(content_y)
  const body_h = r2(Math.max(0.3, ay + ah - body_y))

  const padH    = 0.04
  const TOP_GAP = 0.02   // tight gap between artifact header rule and first bullet
  const bulletPadding = { top: 0, bottom: 0, left: padH, right: 0 }

  // ****** Dynamic font size ****
  const points      = art.points || []
  const nPoints     = Math.max(1, points.length)
  const innerH      = Math.max(0.2, body_h - TOP_GAP)
  const lineSpacing = st.line_spacing || 1.3
  const avgChars    = points.reduce((s, p) => s + String(p?.text || p || '').length, 0) / nPoints
  const charsPerInch   = (fs) => Math.max(1, aw * 72 / (fs * 0.56))
  const linesPerBullet = (fs) => Math.max(1, Math.ceil(avgChars / charsPerInch(fs)))
  const lineHIn        = (fs) => (fs / 72) * lineSpacing
  const estimatedH     = (fs) => nPoints * linesPerBullet(fs) * lineHIn(fs) + (nPoints - 1) * 0.04

  let fontSize = st.font_size || 10
  for (let tryFs = 18; tryFs >= Math.max(9, fontSize); tryFs--) {
    if (estimatedH(tryFs) <= innerH * 0.82) { fontSize = tryFs; break }
  }
  if (fontSizeFloor && fontSizeFloor < fontSize) fontSize = fontSizeFloor

  // Brand primary color for bullet text
  const bulletColor = bt.primary_color || st.color || '#111111'

  const list_y = r2(body_y + TOP_GAP)
  const list_h = r2(Math.max(0.2, ay + ah - list_y))

  blocks.push({
    block_type:  'bullet_list',
    x: ax, y: list_y, w: aw, h: list_h,
    points,
    body_style:  { ...st, font_size: fontSize, color: bulletColor },
    padding:     bulletPadding,
    sentiment:   art.sentiment || 'neutral'
  })
}

function _groupedInsightToBlocks(art, content_y, blocks, bt, r2, fontSizeFloor) {
  const ax      = art.x || 0
  const ay      = art.y || 0
  const aw      = art.w || 0
  const ah      = art.h || 0
  const groups  = art.groups || []
  const n       = groups.length
  if (n === 0) return

  const ghs     = art.group_header_style    || {}
  const gbs     = art.group_bullet_box_style || {}
  const bsty    = art.bullet_style          || {}
  const g_gap   = art.group_gap_in          || 0.08
  const hb_gap  = art.header_to_box_gap_in  || 0.05
  const gLayout = art.group_layout          || 'rows'
  const isBadge = ghs.shape === 'circle_badge'

  const h_fill  = ghs.fill_color   || bt.primary_color || '#1A3C8F'
  const h_cr    = ghs.corner_radius || 4

  const total_content_h = r2(ay + ah - content_y)

  // ****** Shared bullet-size estimator ***************************************************************************************************************************
  // Returns the font size (pt) that makes bullets fill ~80% of available height
  const _bulletFontSize = (bullets, areaW, areaH, styleFs) => {
    const pts = Array.isArray(bullets) ? bullets : []
    const n   = Math.max(1, pts.length)
    const avgChars = pts.reduce((s, p) => s + String(p?.text || p || '').length, 0) / n
    const lineH    = (fs) => (fs / 72) * 1.3
    const linesEach = (fs) => Math.max(1, Math.ceil(avgChars / Math.max(1, areaW * 72 / (fs * 0.56))))
    const totalH   = (fs) => n * linesEach(fs) * lineH(fs) + (n - 1) * 0.04
    let fs = styleFs || 10
    for (let tryFs = 18; tryFs >= Math.max(9, fs); tryFs--) {
      if (totalH(tryFs) <= areaH * 0.82) { fs = tryFs; break }
    }
    return fs
  }

  // Vertical-center helper: offset + height for a bullet_list to sit in the middle of zoneH
  const _centerBullets = (bullets, areaW, areaH, fs, padV) => {
    const n = Math.max(1, (Array.isArray(bullets) ? bullets : []).length)
    const lineH = (fs / 72) * 1.3
    const linesEach = Math.max(1, Math.ceil(
      (bullets.reduce((s, p) => s + String(p?.text || p || '').length, 0) / n) /
      Math.max(1, areaW * 72 / (fs * 0.56))
    ))
    const contentH = n * linesEach * lineH + (n - 1) * 0.04 + 2 * padV
    const clipped  = Math.min(contentH, areaH)
    const offset   = Math.max(0, (areaH - clipped) / 2)
    return { offset, h: clipped }
  }

  if (gLayout === 'rows') {
    const h_w           = ghs.w || 1.2
    const box_x         = r2(ax + h_w + hb_gap)
    const box_w         = r2(aw - h_w - hb_gap)
    const total_bullets = Math.max(1, groups.reduce((s, g) => s + (g.bullets || []).length, 0))
    const total_rh      = Math.max(0.2, total_content_h - (n - 1) * g_gap)

    let cur_y = content_y
    for (let gi = 0; gi < groups.length; gi++) {
      const g        = groups[gi]
      const nbullets = Math.max(1, (g.bullets || []).length)
      const row_h    = r2(Math.max(0.25, total_rh * (nbullets / total_bullets)))

      // Dynamic header font: bounded by the narrower dimension (h_w for text wrap, row_h for height)
      const hdrFs = ghs.font_size || Math.max(9, Math.min(14, Math.min(h_w * 13, row_h * 10)))

      if (isBadge) {
        const dia     = ghs.h || 0.3
        const badge_y = r2(cur_y + (row_h - dia) / 2)
        blocks.push({
          block_type: 'circle',
          x: ax, y: badge_y, w: dia, h: dia,
          fill_color: h_fill,
          text: String(gi + 1),
          font_family: ghs.font_family || bt.title_font_family || 'Arial',
          font_size: hdrFs,
          font_color: ghs.text_color || '#FFFFFF'
        })
      } else {
        blocks.push({
          block_type: 'rect',
          x: ax, y: r2(cur_y), w: h_w, h: row_h,
          fill_color: h_fill, border_color: null, border_width: 0, corner_radius: h_cr
        })
        blocks.push({
          block_type:  'text_box',
          x: r2(ax + 0.06), y: r2(cur_y), w: r2(h_w - 0.12), h: row_h,
          text:        String(g.header || ''),
          font_family: ghs.font_family || bt.title_font_family || 'Arial',
          font_size:   hdrFs,
          bold:        true,
          color:       ghs.text_color || '#FFFFFF',
          align: 'center', valign: 'middle'
        })
      }

      // Bullet box background
      if (gbs.fill_color || gbs.border_color) {
        blocks.push({
          block_type:    'rect',
          x: box_x, y: r2(cur_y), w: box_w, h: row_h,
          fill_color:    gbs.fill_color   || null,
          border_color:  gbs.border_color || null,
          border_width:  gbs.border_width || 0.75,
          corner_radius: gbs.corner_radius || 4
        })
      }

      // Dynamic bullet font + vertical centering within this row.
      // bAreaH (text-only area, padding excluded) is used for font sizing.
      // _centerBullets receives the full row_h so that `h` it returns includes
      // the top/bottom padding the renderer will consume — preventing the renderer
      // from double-subtracting padding that was already baked into a smaller h.
      const bPadV    = (gbs.padding && gbs.padding.top)  || 0.08
      const bPadH    = (gbs.padding && gbs.padding.left) || 0.10
      const bAreaW   = Math.max(0.3, box_w - 2 * bPadH)
      const bAreaH   = Math.max(0.1, row_h - 2 * bPadV)
      const bFs      = Math.min(_bulletFontSize(g.bullets || [], bAreaW, bAreaH, bsty.font_size), fontSizeFloor || Infinity)
      const { offset: bOffset, h: bH } = _centerBullets(g.bullets || [], bAreaW, row_h, bFs, bPadV)
      blocks.push({
        block_type: 'bullet_list',
        x: box_x, y: r2(cur_y + bOffset), w: box_w, h: r2(bH),
        points:     g.bullets || [],
        body_style: { ...bsty, font_size: bFs },
        padding:    gbs.padding || {},
        sentiment:  art.sentiment || 'neutral'
      })

      cur_y = r2(cur_y + row_h + g_gap)
    }

  } else {
    // columns layout
    const col_w = r2((aw - (n - 1) * g_gap) / Math.max(n, 1))
    const h_h   = ghs.h || 0.28
    const box_h = r2(total_content_h - h_h - hb_gap)

    // Dynamic header font for columns: bounded by header bar height and column width
    const hdrFs = ghs.font_size || Math.max(9, Math.min(14, Math.min(col_w * 10, h_h * 55)))

    let cur_x = ax
    for (let gi = 0; gi < groups.length; gi++) {
      const g = groups[gi]

      if (isBadge) {
        const dia     = h_h
        const badge_x = r2(cur_x + (col_w - dia) / 2)
        blocks.push({
          block_type: 'circle',
          x: badge_x, y: content_y, w: dia, h: dia,
          fill_color: h_fill,
          text: String(gi + 1),
          font_family: ghs.font_family || bt.title_font_family || 'Arial',
          font_size: hdrFs,
          font_color: ghs.text_color || '#FFFFFF'
        })
      } else {
        blocks.push({
          block_type: 'rect',
          x: r2(cur_x), y: content_y, w: col_w, h: h_h,
          fill_color: h_fill, border_color: null, border_width: 0, corner_radius: h_cr
        })
        blocks.push({
          block_type:  'text_box',
          x: r2(cur_x + 0.05), y: content_y, w: r2(col_w - 0.10), h: h_h,
          text:        String(g.header || ''),
          font_family: ghs.font_family || bt.title_font_family || 'Arial',
          font_size:   hdrFs,
          bold:        true,
          color:       ghs.text_color || '#FFFFFF',
          align: 'center', valign: 'middle'
        })
      }

      const bullet_y = r2(content_y + h_h + hb_gap)

      if (gbs.fill_color || gbs.border_color) {
        blocks.push({
          block_type:    'rect',
          x: r2(cur_x), y: bullet_y, w: col_w, h: box_h,
          fill_color:    gbs.fill_color   || null,
          border_color:  gbs.border_color || null,
          border_width:  gbs.border_width || 0.75,
          corner_radius: gbs.corner_radius || 4
        })
      }

      // Dynamic bullet font + vertical centering within box_h.
      // bAreaH (text-only area, padding excluded) is used for font sizing.
      // _centerBullets receives the full box_h so that `h` it returns includes
      // the top/bottom padding the renderer will consume — preventing double-subtraction.
      const bPadV  = (gbs.padding && gbs.padding.top)  || 0.08
      const bPadH  = (gbs.padding && gbs.padding.left) || 0.10
      const bAreaW = Math.max(0.3, col_w - 2 * bPadH)
      const bAreaH = Math.max(0.1, box_h - 2 * bPadV)
      const bFs    = Math.min(_bulletFontSize(g.bullets || [], bAreaW, bAreaH, bsty.font_size), fontSizeFloor || Infinity)
      const { offset: bOffset, h: bH } = _centerBullets(g.bullets || [], bAreaW, box_h, bFs, bPadV)
      blocks.push({
        block_type: 'bullet_list',
        x: r2(cur_x), y: r2(bullet_y + bOffset), w: col_w, h: r2(bH),
        points:     g.bullets || [],
        body_style: { ...bsty, font_size: bFs },
        padding:    gbs.padding || {},
        sentiment:  art.sentiment || 'neutral'
      })

      cur_x = r2(cur_x + col_w + g_gap)
    }
  }
}

function _cardsToBlocks(art, content_y, blocks, bt, r2) {
  const cards = art.cards || []
  const count = cards.length
  if (!count) return

  const cs       = art.card_style || {}
  const ts       = art.title_style || {}
  const subs     = art.subtitle_style || {}
  const bs       = art.body_style || {}
  const pad      = cs.internal_padding || 0.12
  const accentW  = 0.07
  const accentGap = 0.08
  const gap      = cs.gap || 0.12

  const ax = art.x || 0
  const aw = art.w || 0
  const ab = (art.y || 0) + (art.h || 0)   // bottom of art zone

  // ****** Recompute card frames from content_y ***************************************************************************************************
  // Leave a gap between the artifact header rule and the first card
  const headerGap = 0.12
  const cardsTop  = r2(content_y + headerGap)
  const availH    = r2(Math.max(0.2, ab - cardsTop))
  const layout    = String(art.cards_layout || 'column').toLowerCase()

  const frames = []
  if (layout === 'row') {
    const cw = r2((aw - gap * (count - 1)) / Math.max(count, 1))
    for (let i = 0; i < count; i++) {
      frames.push({ x: r2(ax + i * (cw + gap)), y: cardsTop, w: cw, h: availH })
    }
  } else if (layout === 'column') {
    const ch = r2((availH - gap * (count - 1)) / Math.max(count, 1))
    for (let i = 0; i < count; i++) {
      frames.push({ x: ax, y: r2(cardsTop + i * (ch + gap)), w: aw, h: ch })
    }
  } else {
    // grid (2 columns)
    const cols = count > 1 ? 2 : 1
    const rows = Math.ceil(count / cols)
    const cw   = r2((aw - gap * (cols - 1)) / Math.max(cols, 1))
    const ch   = r2((availH - gap * (rows - 1)) / Math.max(rows, 1))
    for (let i = 0; i < count; i++) {
      frames.push({
        x: r2(ax + (i % cols) * (cw + gap)),
        y: r2(cardsTop + Math.floor(i / cols) * (ch + gap)),
        w: cw, h: ch
      })
    }
  }

  const sentimentAccent = {
    positive: bt.secondary_color || '#2D8A4E',
    negative: '#C0392B',
    neutral: bt.primary_color || '#1A3C8F'
  }
  const paletteBase = [
    bt.primary_color,
    bt.secondary_color,
    ...((bt.accent_colors || []).length ? bt.accent_colors : []),
    ...((bt.chart_palette || []).length ? bt.chart_palette : [])
  ].filter(Boolean)
  const accentPalette = [...new Set(paletteBase)]

  for (let i = 0; i < count; i++) {
    const card = cards[i]
    const fr   = frames[i]
    const fx = fr.x, fy = fr.y, fw = fr.w, fh = fr.h

    const accentColor = count > 1
      ? (accentPalette[i % Math.max(accentPalette.length, 1)] || sentimentAccent[card.sentiment] || '#1A3C8F')
      : (sentimentAccent[card.sentiment] || accentPalette[0] || '#1A3C8F')

    // Card background
    blocks.push({
      block_type: 'rect',
      x: fx, y: fy, w: fw, h: fh,
      fill_color:   cs.fill_color   || '#F5F5F5',
      border_color: cs.border_color || '#DDDDDD',
      border_width: cs.border_width || 0.75,
      corner_radius: 0
    })

    // Accent strip
    if (accentColor) {
      blocks.push({
        block_type: 'rect',
        x: fx, y: fy, w: accentW, h: fh,
        fill_color: accentColor, border_color: null, border_width: 0, corner_radius: 0
      })
    }

    // ****** Inner layout ************************************************************************************************************************************************************************
    const inner_x = r2(fx + pad + accentW + accentGap)
    const inner_y = r2(fy + pad)
    const inner_w = r2(Math.max(0.3, fw - (pad * 2) - accentW - accentGap))
    const inner_h = r2(fh - 2 * pad)

    // Zone proportions: title 22% | subtitle 40% | gap | body rest
    const title_h = r2(inner_h * 0.22)
    const sub_h   = r2(inner_h * 0.40)
    const body_h  = r2(Math.max(0.16, inner_h - title_h - sub_h - 0.10))

    const titleY    = inner_y
    const subtitleY = r2(titleY + title_h + 0.04)
    const bodyY     = r2(subtitleY + sub_h + 0.06)

    // ****** Dynamic font sizes ******************************************************************************************************************************************************
    // 1. Subtitle (centre message) sized first **” it is the primary element
    const subtitleFontSize = Math.max(18, Math.min(38, sub_h * 58))
    // 2. Title and body scale from their own zones, capped relative to subtitle
    const titleFontSize    = Math.max(10, Math.min(subtitleFontSize * 0.45, title_h * 55))
    const bodyFontSize     = Math.max(8,  Math.min(subtitleFontSize * 0.38, body_h  * 42))

    if (card.title) {
      blocks.push({
        block_type: 'text_box',
        x: inner_x, y: titleY, w: inner_w, h: title_h,
        text:        card.title,
        font_family: ts.font_family || bt.title_font_family || 'Arial',
        font_size:   ts.font_size   || titleFontSize,
        bold:        true,
        color:       ts.color || accentColor,
        align: 'left', valign: 'top'
      })
    }
    if (card.subtitle) {
      blocks.push({
        block_type: 'text_box',
        x: inner_x, y: subtitleY, w: inner_w, h: sub_h,
        text:        card.subtitle,
        font_family: subs.font_family || bt.body_font_family || 'Arial',
        font_size:   subs.font_size   || subtitleFontSize,
        bold:        true,
        color:       subs.color || '#111111',
        align: 'left', valign: 'middle'
      })
    }
    if (card.body) {
      blocks.push({
        block_type: 'text_box',
        x: inner_x, y: bodyY, w: inner_w, h: body_h,
        text:        card.body,
        font_family: bs.font_family || bt.body_font_family || 'Arial',
        font_size:   bs.font_size   || bodyFontSize,
        bold:        false,
        color:       bs.color || '#333333',
        align: 'left', valign: 'top'
      })
    }
  }
}

function flattenToBlocks(slideSpec, brandTokens) {
  const bt     = brandTokens || {}
  const blocks = []
  const r2     = x => Math.round(x * 100) / 100

  // ****** 1. Title block ************************************************************************************************************************************************************************
  const tb = slideSpec.title_block || {}
  if (tb.text) {
    blocks.push({
      block_type:  'title',
      artifact_type: 'slide',
      artifact_subtype: 'title',
      block_role: 'slide_header',
      artifact_header_text: tb.text,
      fallback_policy: {
        allow_renderer_fallback: false,
        fallback_mode: 'none',
        artifact_type: 'slide',
        artifact_subtype: 'title',
        block_role: 'slide_header',
        fallback_key: 'slide:title'
      },
      x:           tb.x           != null ? tb.x           : 0.4,
      y:           tb.y           != null ? tb.y           : 0.15,
      w:           tb.w           != null ? tb.w           : 9.2,
      h:           tb.h           != null ? tb.h           : 0.7,
      text:        tb.text,
      font_family: tb.font_family || bt.title_font_family || 'Arial',
      font_size:   tb.font_size   || 20,
      bold:        ['bold','semibold'].includes(String(tb.font_weight || 'bold').toLowerCase()),
      color:       tb.color       || bt.title_color || '#1A3C8F',
      align:       tb.align       || 'left',
      valign:      'top'
    })
  }

  // ****** 2. Subtitle block ***************************************************************************************************************************************************************
  const sb = slideSpec.subtitle_block || {}
  if (sb.text) {
    blocks.push({
      block_type:  'subtitle',
      artifact_type: 'slide',
      artifact_subtype: 'subtitle',
      block_role: 'slide_subheader',
      artifact_header_text: sb.text,
      fallback_policy: {
        allow_renderer_fallback: false,
        fallback_mode: 'none',
        artifact_type: 'slide',
        artifact_subtype: 'subtitle',
        block_role: 'slide_subheader',
        fallback_key: 'slide:subtitle'
      },
      x:           sb.x           != null ? sb.x           : 0.4,
      y:           sb.y           != null ? sb.y           : 0.9,
      w:           sb.w           != null ? sb.w           : 9.2,
      h:           sb.h           != null ? sb.h           : 0.45,
      text:        sb.text,
      font_family: sb.font_family || bt.body_font_family || 'Arial',
      font_size:   sb.font_size   || 14,
      bold:        ['bold','semibold'].includes(String(sb.font_weight || '').toLowerCase()),
      color:       sb.color       || bt.body_color || '#333333',
      align:       sb.align       || 'left',
      valign:      'middle'
    })
  }

  // ****** 3. Zones â†’ Artifacts ******************************************************************************************************************************************************
  // Pre-pass: compute each insight artifact's standalone font size, then
  // harmonise all insight bullet text on this slide to the minimum found.
  // Group headers are excluded **” only bullet/body text is harmonised.
  const allArts = (slideSpec.zones || []).flatMap(z => z.artifacts || [])
  const insightArts = allArts.filter(a => a.type === 'insight_text')
  let slideFontSizeFloor = null
  if (insightArts.length > 1) {
    const sizes = insightArts.map(a => {
      // Approximate content_y: art.y (header handled inside compute fn)
      const approxContentY = (a.header_block && a.header_block.text)
        ? (a.y || 0) + (a.header_block.h || 0.30) + 0.07
        : (a.y || 0)
      return _computeInsightFontSize(a, approxContentY)
    })
    slideFontSizeFloor = Math.min(...sizes)
  }

  // Suppress artifact header on 1-zone-1-artifact slides — the slide title already acts as the header
  const slideZones = slideSpec.zones || []
  const isSingleArtifactSlide = slideZones.length === 1 && (slideZones[0].artifacts || []).length === 1

  for (const zone of slideZones) {
    for (const art of (zone.artifacts || [])) {
      const floor = art.type === 'insight_text' ? slideFontSizeFloor : null
      if (isSingleArtifactSlide && art.header_block) {
        const artWithoutHeader = { ...art, header_block: null, artifact_header: '' }
        _artifactToBlocks(artWithoutHeader, blocks, bt, r2, floor)
      } else {
        _artifactToBlocks(art, blocks, bt, r2, floor)
      }
    }
  }

  // ****** 4. Global elements ************************************************************************************************************************************************************
  const ge = slideSpec.global_elements || {}

  // Logo is intentionally not included in blocks[]:
  // Logo is not rendered **” template mode always active; master layout carries logo automatically.

  if (ge.footer && ge.footer.text) {
    const ft = ge.footer
    blocks.push({
      block_type:  'footer',
      artifact_type: 'global_element',
      artifact_subtype: 'footer',
      block_role: 'global_element',
      artifact_header_text: ft.text,
      fallback_policy: {
        allow_renderer_fallback: false,
        fallback_mode: 'none',
        artifact_type: 'global_element',
        artifact_subtype: 'footer',
        block_role: 'global_element',
        fallback_key: 'global_element:footer'
      },
      x:           ft.x != null ? ft.x : 0.4,
      y:           ft.y != null ? ft.y : 7.3,
      w:           ft.w || 5.0,
      h:           ft.h || 0.22,
      text:        ft.text,
      font_family: ft.font_family || bt.body_font_family || 'Arial',
      font_size:   ft.font_size   || 8,
      color:       ft.color       || '#AAAAAA',
      align:       ft.align       || 'left',
      valign:      'middle'
    })
  }
  if (ge.page_number && ge.page_number.text) {
    const pn = ge.page_number
    blocks.push({
      block_type:  'page_number',
      artifact_type: 'global_element',
      artifact_subtype: 'page_number',
      block_role: 'global_element',
      artifact_header_text: pn.text || '',
      fallback_policy: {
        allow_renderer_fallback: false,
        fallback_mode: 'none',
        artifact_type: 'global_element',
        artifact_subtype: 'page_number',
        block_role: 'global_element',
        fallback_key: 'global_element:page_number'
      },
      x:           pn.x != null ? pn.x : 9.4,
      y:           pn.y != null ? pn.y : 7.3,
      w:           pn.w || 0.8,
      h:           pn.h || 0.22,
      text:        pn.text,
      font_family: pn.font_family || bt.body_font_family || 'Arial',
      font_size:   pn.font_size   || 8,
      color:       pn.color       || '#AAAAAA',
      align:       'right',
      valign:      'middle'
    })
  }

  return blocks
}

// ****** groupBlocksByArtifact ***************************************************************************************************************************************************************
// Converts a flat blocks[] into artifact_groups[] where artifact-level metadata
// (artifact_id, artifact_type, artifact_subtype, artifact_header_text,
//  fallback_policy) is hoisted to the group and removed from each block.
// Blocks without artifact_id (title, subtitle, footer, page_number) each become
// their own single-block group keyed by block_type.
// generate_pptx.py / Agent 6 call flattenArtifactGroups() to restore flat blocks[].
function groupBlocksByArtifact(blocks) {
  const groups   = []
  const indexMap = new Map()  // groupKey â†’ index in groups

  for (const block of (blocks || [])) {
    const aid      = block.artifact_id
    const groupKey = aid != null ? 'id:' + aid : 'bt:' + block.block_type

    if (!indexMap.has(groupKey)) {
      const entry = {}
      if (aid != null)                             entry.artifact_id          = aid
      if (block.artifact_type)                     entry.artifact_type        = block.artifact_type
      if (block.artifact_subtype)                  entry.artifact_subtype     = block.artifact_subtype
      if (block.artifact_header_text != null)      entry.artifact_header_text = block.artifact_header_text
      if (block.fallback_policy)                   entry.fallback_policy      = block.fallback_policy
      entry.blocks = []
      indexMap.set(groupKey, groups.length)
      groups.push(entry)
    }

    const slim = Object.assign({}, block)
    delete slim.artifact_id
    delete slim.artifact_type
    delete slim.artifact_subtype
    delete slim.artifact_header_text
    delete slim.fallback_policy
    groups[indexMap.get(groupKey)].blocks.push(slim)
  }

  return groups
}


function mergeContentIntoZones(designedZones, manifestZones, brandTokens) {
  if (!designedZones || !manifestZones) return designedZones || []

  const bt = brandTokens || {}

  const result = designedZones.map((dZone, zi) => {
    // Zone matching: zone_id takes priority over index so reordered zones still merge correctly.
    // Agent 5's LLM may reorder zones vs Agent 4's spec **” index-first matching would then
    // pair initiative_map with insight_text data (or vice-versa), causing empty content.
    const mZoneById    = manifestZones.find(z => z.zone_id && z.zone_id === dZone.zone_id)
    const mZoneByIndex = manifestZones[zi]
    const mZone = mZoneById || mZoneByIndex
    if (!mZone) return dZone

    const mergedArtifacts = (dZone.artifacts || []).map((dArt, ai) => {
      const dType = normalizeArtifactType(dArt?.type, dArt?.chart_type)

      // Artifact matching: try by position within the matched zone first.
      // If the types don't match (zone mis-match or reorder), search all manifest zones
      // for a zone whose first artifact has the same type as the designed artifact.
      let mArt = (mZone.artifacts || [])[ai]
      let mType = normalizeArtifactType(mArt?.type, mArt?.chart_type)
      if (!mArt || mType !== dType) {
        // Fallback: scan all manifest zones for a type-compatible artifact
        for (const mz of manifestZones) {
          const candidate = (mz.artifacts || [])[0]
          if (normalizeArtifactType(candidate?.type, candidate?.chart_type) === dType) {
            mArt = candidate
            mType = dType
            break
          }
        }
      }
      if (!mArt || mType !== dType) return dArt

      const t = dType

      if (t === 'insight_text') {
        // ****** Determine mode: manifest (Agent 4) is authoritative for content structure ******
        const mGroups = mArt.groups && mArt.groups.length > 0 ? mArt.groups : null
        const mPoints = mArt.points && mArt.points.length > 0 ? mArt.points : null
        const resolvedMode = mArt.insight_mode
          || (mGroups ? 'grouped' : mPoints ? 'standard' : dArt.insight_mode || 'standard')

        const heading        = mArt.heading || getArtifactHeader(mArt) || dArt.heading || 'Key Insight'
        const insight_header = getArtifactHeader(mArt) || dArt.insight_header || dArt.artifact_header || heading || 'Key Insight'
        const sentiment      = mArt.sentiment      || dArt.sentiment      || 'neutral'

        // ****** FLOW 2: Grouped ******************************************************************************************************************************************************************
        if (resolvedMode === 'grouped') {
          const primary   = bt.primary_color      || '#1A3C8F'
          const titleFont = bt.title_font_family  || 'Arial'
          const bodyFont  = bt.body_font_family   || 'Arial'
          const artW = dArt.w || 4
          const artH = dArt.h || 4
          const n    = (mGroups || []).length || 1

          // Layout direction: columns = headers above boxes (horizontal groups);
          //                   rows    = headers left of boxes (vertical groups)
          const gLayout = dArt.group_layout
            || (artW > artH && n <= 3 ? 'columns' : 'rows')

          // Re-use Agent 5 grouped styling when it exists; fill gaps otherwise
          const agentHasGrouped = dArt.insight_mode === 'grouped' && !!dArt.group_header_style
          const f = agentHasGrouped ? ((dArt.bullet_style || {}).font_size || 10) : 10

          // ****** Content-aware header dimension calculation *********************************************************************
          // Estimate the minimum w (rows) or h (columns) needed to render each
          // group header text without character-level wrapping.
          // Uses the same approximation the renderer will use: avg char width â‰ˆ
          // font_size Ã— 0.58 / 72 inches. Target: each header fits in â‰¤ 2 lines.
          const hFontSize   = 10   // header font size (pts)
          const charWIn     = hFontSize * 0.58 / 72   // avg char width in inches
          const lineHIn     = hFontSize * 1.4  / 72   // line height in inches
          const headerTexts = (mGroups || []).map(g => String(g.header || ''))

          // rows layout: fix header WIDTH so longest header fits in â‰¤ 2 lines
          const _minRowHeaderW = (() => {
            const maxLen = Math.max(...headerTexts.map(t => t.length), 1)
            const minW   = Math.ceil(maxLen / 2) * charWIn + 0.12  // 2-line target
            return Math.min(minW, artW * 0.35)   // cap at 35% of artifact width
          })()

          // columns layout: fix header HEIGHT so longest header (at col_w) fits
          const _minColHeaderH = (() => {
            const colW        = artW / Math.max(n, 1)   // approximate col width
            const charsPerLn  = Math.max(1, (colW - 0.10) / charWIn)
            const maxLines    = Math.max(...headerTexts.map(t =>
              Math.ceil(t.length / charsPerLn)
            ), 1)
            return maxLines * lineHIn + 0.08   // 0.08" top+bottom padding
          })()

          const group_header_style = dArt.group_header_style || {
            shape:        'rounded_rect',
            fill_color:   primary,
            text_color:   '#FFFFFF',
            font_family:  titleFont,
            font_size:    hFontSize,
            font_weight:  'bold',
            corner_radius: 4,
            w: gLayout === 'rows'
              ? r2(Math.max(_minRowHeaderW, Math.min(1.5, artW * 0.30)))
              : artW,
            h: gLayout === 'columns'
              ? r2(Math.max(_minColHeaderH, hFontSize * 1.8 / 72, artH * 0.06))
              : r2(Math.max(hFontSize * 1.8 / 72, artH * 0.06))
          }

          // If Agent 5 already set group_header_style but w/h may still be too small,
          // enforce the content-based floor on the existing values too.
          if (dArt.group_header_style) {
            if (gLayout === 'rows' && (dArt.group_header_style.w || 0) < _minRowHeaderW) {
              group_header_style.w = r2(_minRowHeaderW)
            }
            if (gLayout === 'columns' && (dArt.group_header_style.h || 0) < _minColHeaderH) {
              group_header_style.h = r2(_minColHeaderH)
            }
          }

          const group_bullet_box_style = dArt.group_bullet_box_style || {
            fill_color:   null,
            border_color: '#CCCCCC',
            border_width:  0.75,
            corner_radius: group_header_style.corner_radius || 4,
            padding: {
              top:    r2(Math.max(f * 0.8 / 72, 0.05)),
              right:  r2(Math.max(f * 1.0 / 72, 0.07)),
              bottom: r2(Math.max(f * 0.8 / 72, 0.05)),
              left:   r2(Math.max(f * 1.0 / 72, 0.07))
            }
          }

          const bullet_style = dArt.bullet_style || {
            font_family:     bodyFont,
            font_size:       f,
            font_weight:     'regular',
            color:           bt.body_color || '#111111',
            line_spacing:    1.35,
            indent_inches:   0.10,
            space_before_pt: r2(Math.max(f * 0.4, 2)),
            char:            '**¢'
          }

          const dimForGap = gLayout === 'rows' ? artH : artW
          const group_gap_in         = dArt.group_gap_in         || r2(Math.max(dimForGap * 0.015, 0.05))
          const header_to_box_gap_in = dArt.header_to_box_gap_in || r2(Math.max(f * 0.5 / 72, 0.03))

          return syncArtifactHeaderBlock({
            ...dArt,
            insight_mode:          'grouped',
            artifact_header:       insight_header,
            heading,
            insight_header,
            sentiment,
            groups:                mGroups,
            points:                [],
            group_layout:          gLayout,
            group_header_style,
            group_bullet_box_style,
            bullet_style:          { ...bullet_style },
            group_gap_in:          r2(group_gap_in),
            header_to_box_gap_in:  r2(header_to_box_gap_in)
          }, insight_header)
        }

        // ****** FLOW 1: Standard (points) ************************************************************************************************************************************
        const body_style = dArt.body_style || {
          font_family:           bt.body_font_family || 'Arial',
          font_size:             10,
          font_weight:           'regular',
          color:                 bt.body_color || '#000000',
          line_spacing:          1.3,
          indent_inches:         0.15,
          list_style:            'bullet',
          space_before_pt:       6,
          vertical_distribution: 'spread'
        }
        return syncArtifactHeaderBlock({
          ...dArt,
          artifact_coverage_hint: mArt.artifact_coverage_hint != null ? mArt.artifact_coverage_hint : dArt.artifact_coverage_hint,
          artifact_header: insight_header,
          insight_mode:   'standard',
          heading,
          insight_header,
          sentiment,
          points:  mPoints || dArt.points || [],
          groups:  undefined,
          body_style
        }, insight_header)
      }

      if (t === 'chart') {
        const mergedChartType = mArt.chart_type || dArt.chart_type || 'bar'
        const artifactHeader = getArtifactHeader(mArt) || dArt.chart_header || dArt.artifact_header || ''
        return syncArtifactHeaderBlock({
          ...dArt,
          artifact_coverage_hint: mArt.artifact_coverage_hint != null ? mArt.artifact_coverage_hint : dArt.artifact_coverage_hint,
          artifact_header: artifactHeader,
          chart_type:       mergedChartType,
          chart_header:     mArt.chart_header || artifactHeader || mArt.stat_header || dArt.chart_header || '',
          chart_title:      mArt.chart_title      || dArt.chart_title      || '',
          chart_insight:    mArt.chart_insight    || mArt.stat_decision || dArt.chart_insight    || '',
          rows:             mArt.rows             || dArt.rows             || [],
          annotation_style: mArt.annotation_style || dArt.annotation_style || 'trailing',
          x_label:          mArt.x_label          || dArt.x_label          || '',
          y_label:          mArt.y_label          || dArt.y_label          || '',
          categories:        mArt.categories        || dArt.categories        || [],
          series:            mArt.series            || dArt.series            || [],
          secondary_series:  mArt.secondary_series  || dArt.secondary_series  || [],
          dual_axis:         mArt.dual_axis          != null ? mArt.dual_axis  : (dArt.dual_axis || false),
          secondary_y_label: mArt.secondary_y_label || dArt.secondary_y_label || '',
          show_data_labels:  mArt.show_data_labels !== undefined
                              ? mArt.show_data_labels : (dArt.show_data_labels !== false),
          show_legend:       mArt.show_legend      !== undefined
                              ? mArt.show_legend      : (mergedChartType === 'group_pie' ? true : !!dArt.show_legend)
        }, artifactHeader)
      }

      if (t === 'stat_bar') {
        const artifactHeader = getArtifactHeader(mArt) || dArt.stat_header || dArt.artifact_header || ''
        return syncArtifactHeaderBlock({
          ...dArt,
          artifact_coverage_hint: mArt.artifact_coverage_hint != null ? mArt.artifact_coverage_hint : dArt.artifact_coverage_hint,
          artifact_header: artifactHeader,
          stat_header: mArt.stat_header || artifactHeader || dArt.stat_header || '',
          stat_decision: mArt.stat_decision || dArt.stat_decision || '',
          column_headers: mArt.column_headers || dArt.column_headers || {},
          rows: mArt.rows || dArt.rows || [],
          annotation_style: mArt.annotation_style || dArt.annotation_style || 'trailing'
        }, artifactHeader)
      }

      if (t === 'comparison_table') {
        const normalized = normalizeComparisonTableManifest(mArt)
        const artifactHeader = getArtifactHeader(mArt) || dArt.comparison_header || dArt.artifact_header || ''
        return ensureArtifactHeaderBlock({
          ...dArt,
          artifact_coverage_hint: mArt.artifact_coverage_hint != null ? mArt.artifact_coverage_hint : dArt.artifact_coverage_hint,
          artifact_header: artifactHeader,
          comparison_header: mArt.comparison_header || artifactHeader || dArt.comparison_header || mArt.table_header || '',
          _label_column_header: Array.isArray(mArt?.columns) && mArt.columns.length
            ? String(mArt.columns[0] || 'Option')
            : String((mArt?.column_headers || [])[0]?.label || dArt._label_column_header || 'Option'),
          criteria: normalized.criteria.length ? normalized.criteria : (dArt.criteria || []),
          options: normalized.options.length ? normalized.options : (dArt.options || []),
          recommended_option: normalized.recommended_option || dArt.recommended_option || ''
        }, artifactHeader, bt)
      }

      if (t === 'initiative_map') {
        const normalized = normalizeInitiativeMapManifest(mArt)
        const artifactHeader = getArtifactHeader(mArt) || dArt.initiative_header || dArt.artifact_header || ''
        return ensureArtifactHeaderBlock({
          ...dArt,
          artifact_coverage_hint: mArt.artifact_coverage_hint != null ? mArt.artifact_coverage_hint : dArt.artifact_coverage_hint,
          artifact_header: artifactHeader,
          initiative_header: mArt.initiative_header || artifactHeader || dArt.initiative_header || mArt.table_header || '',
          dimension_labels: normalized.dimension_labels.length ? normalized.dimension_labels : (dArt.dimension_labels || []),
          initiatives: normalized.initiatives.length ? normalized.initiatives : (dArt.initiatives || [])
        }, artifactHeader, bt)
      }

      if (t === 'profile_card_set') {
        const artifactHeader = getArtifactHeader(mArt) || dArt.profile_header || dArt.artifact_header || ''
        return ensureArtifactHeaderBlock({
          ...dArt,
          artifact_coverage_hint: mArt.artifact_coverage_hint != null ? mArt.artifact_coverage_hint : dArt.artifact_coverage_hint,
          artifact_header: artifactHeader,
          profile_header: mArt.profile_header || artifactHeader || dArt.profile_header || mArt.heading || '',
          profiles: mArt.profiles || dArt.profiles || [],
          layout_direction: mArt.layout_direction || dArt.layout_direction || 'horizontal'
        }, artifactHeader, bt)
      }

      if (t === 'risk_register') {
        const normalized = normalizeRiskRegisterManifest(mArt)
        const artifactHeader = getArtifactHeader(mArt) || dArt.risk_header || dArt.artifact_header || ''
        return {
          ...dArt,
          artifact_coverage_hint: mArt.artifact_coverage_hint != null ? mArt.artifact_coverage_hint : dArt.artifact_coverage_hint,
          artifact_header: artifactHeader || dArt.artifact_header || '',
          header_block: makeHeaderBlockFromManifestArtifact(mArt, bt) || dArt.header_block || null,
          risk_header: mArt.risk_header || dArt.risk_header || mArt.table_header || '',
          severity_levels: normalized.severity_levels.length ? normalized.severity_levels : (dArt.severity_levels || [])
        }
      }

      if (t === 'cards') {
        return {
          ...dArt,
          artifact_coverage_hint: mArt.artifact_coverage_hint != null ? mArt.artifact_coverage_hint : dArt.artifact_coverage_hint,
          cards: mArt.cards || dArt.cards || []
        }
      }

      if (t === 'workflow') {
        // Merge node content (label, value, description, level) into designed nodes
        // Designed nodes have x/y/w/h; manifest nodes have the text content.
        const designedNodes   = dArt.nodes   || []
        const manifestNodes   = normalizeWorkflowNodes(mArt.nodes)
        const mergedNodes = designedNodes.length > 0
          ? designedNodes.map((dn, ni) => {
              const mn = manifestNodes.find(n => n.id === dn.id) || manifestNodes[ni]
              if (!mn) return dn
              return {
                ...dn,                        // keep x, y, w, h from Agent 5
                label:       mn.node_label || mn.label || dn.label || dn.id,
                value:       mn.primary_message || mn.value || dn.value || '',
                description: mn.secondary_message || mn.description || dn.description || '',
                level:       mn.level       !== undefined ? mn.level : (dn.level || 1)
              }
            })
          : manifestNodes.map((mn, ni) => ({
              id:          mn.id || `n${ni + 1}`,
              label:       mn.node_label || mn.label || mn.id || `Step ${ni + 1}`,
              value:       mn.primary_message || mn.value || '',
              description: mn.secondary_message || mn.description || '',
              level:       mn.level !== undefined ? mn.level : 1
            }))

        // Merge connection from/to/type from manifest into designed connections
        // Designed connections have path[] waypoints; manifest has from/to/type.
        const designedConns  = dArt.connections || []
        const manifestConns  = mArt.connections || []
        const mergedConns = designedConns.length > 0
          ? designedConns.map((dc, ci) => {
              const mc = manifestConns[ci]
                      || manifestConns.find(c => c.from === dc.from && c.to === dc.to)
              if (!mc) return dc
              return {
                ...dc,                        // keep path[] from Agent 5
                from: mc.from || dc.from,
                to:   mc.to   || dc.to,
                type: mc.type || dc.type || 'arrow'
              }
            })
          : manifestConns.map((mc, ci) => ({
              from: mc.from || mergedNodes[ci]?.id || '',
              to:   mc.to   || mergedNodes[ci + 1]?.id || '',
              type: mc.type || 'arrow'
            }))

        const artifactHeader = getArtifactHeader(mArt) || dArt.workflow_header || dArt.artifact_header || ''
        return syncArtifactHeaderBlock({
          ...dArt,
          artifact_coverage_hint: mArt.artifact_coverage_hint != null ? mArt.artifact_coverage_hint : dArt.artifact_coverage_hint,
          artifact_header: artifactHeader,
          workflow_type:    mArt.workflow_type    || dArt.workflow_type    || 'process_flow',
          workflow_header:  mArt.workflow_header  || artifactHeader || dArt.workflow_header  || '',
          flow_direction:   mArt.flow_direction   || dArt.flow_direction   || 'left_to_right',
          workflow_title:   mArt.workflow_title   || dArt.workflow_title   || '',
          workflow_insight: mArt.workflow_insight || dArt.workflow_insight || '',
          nodes:            mergedNodes,
          connections:      mergedConns
        }, artifactHeader)
      }

      if (t === 'table') {
        const artifactHeader = getArtifactHeader(mArt) || dArt.table_header || dArt.artifact_header || ''
        return syncArtifactHeaderBlock({
          ...dArt,
          artifact_coverage_hint: mArt.artifact_coverage_hint != null ? mArt.artifact_coverage_hint : dArt.artifact_coverage_hint,
          artifact_header: artifactHeader,
          table_header:   mArt.table_header   || artifactHeader || dArt.table_header   || '',
          title:          mArt.title          || dArt.title          || '',
          headers:        mArt.headers        || dArt.headers        || [],
          rows:           mArt.rows           || dArt.rows           || [],
          highlight_rows: mArt.highlight_rows || dArt.highlight_rows || [],
          note:           mArt.note           || dArt.note           || ''
        }, artifactHeader)
      }

      if (t === 'matrix') {
        const normalized = normalizeMatrixManifest(mArt)
        const artifactHeader = getArtifactHeader(mArt) || dArt.matrix_header || dArt.artifact_header || ''
        return syncArtifactHeaderBlock({
          ...dArt,
          artifact_coverage_hint: mArt.artifact_coverage_hint != null ? mArt.artifact_coverage_hint : dArt.artifact_coverage_hint,
          artifact_header: artifactHeader,
          matrix_type:   mArt.matrix_type   || dArt.matrix_type   || '2x2',
          matrix_header: mArt.matrix_header || artifactHeader || dArt.matrix_header || '',
          x_axis:        mArt.x_axis        || dArt.x_axis        || { label: '', low_label: '', high_label: '' },
          y_axis:        mArt.y_axis        || dArt.y_axis        || { label: '', low_label: '', high_label: '' },
          quadrants:     normalized.quadrants.length ? normalized.quadrants : (dArt.quadrants || []),
          points:        normalized.points.length ? normalized.points : (dArt.points || [])
        }, artifactHeader)
      }

      if (t === 'driver_tree') {
        const normalized = normalizeDriverTreeManifest(mArt)
        const artifactHeader = getArtifactHeader(mArt) || dArt.tree_header || dArt.artifact_header || ''
        return syncArtifactHeaderBlock({
          ...dArt,
          artifact_coverage_hint: mArt.artifact_coverage_hint != null ? mArt.artifact_coverage_hint : dArt.artifact_coverage_hint,
          artifact_header: artifactHeader,
          tree_header: mArt.tree_header || artifactHeader || dArt.tree_header || '',
          root:        normalized.root.label || normalized.root.value ? normalized.root : (dArt.root || { label: '', value: '' }),
          branches:    normalized.branches.length ? normalized.branches : (dArt.branches || [])
        }, artifactHeader)
      }

      if (t === 'prioritization') {
        const artifactHeader = getArtifactHeader(mArt) || dArt.priority_header || dArt.artifact_header || ''
        return syncArtifactHeaderBlock({
          ...dArt,
          artifact_coverage_hint: mArt.artifact_coverage_hint != null ? mArt.artifact_coverage_hint : dArt.artifact_coverage_hint,
          artifact_header: artifactHeader,
          priority_header: mArt.priority_header || artifactHeader || dArt.priority_header || '',
          items: (mArt.items || dArt.items || []).map(it => ({
            rank: it.rank,
            title: it.title || '',
            description: it.description || '',
            qualifiers: Array.isArray(it.qualifiers)
              ? it.qualifiers.slice(0, 2).map(q => ({ label: q?.label || '', value: q?.value || '' }))
              : [{ label: '', value: '' }, { label: '', value: '' }]
          }))
        }, artifactHeader)
      }

      return dArt
    })

    return {
      ...dZone,
      zone_split:
        dZone.zone_split ||
        (dZone.layout_hint || {}).split ||
        mZone.zone_split ||
        (mZone.layout_hint || {}).split ||
        'full',
      layout_hint: dZone.layout_hint || mZone.layout_hint || null,
      artifact_arrangement:
        dZone.artifact_arrangement ||
        (dZone.layout_hint || {}).artifact_arrangement ||
        mZone.artifact_arrangement ||
        (mZone.layout_hint || {}).artifact_arrangement ||
        null,
      split_hint:
        (Array.isArray(dZone.split_hint) ? dZone.split_hint : null) ||
        (Array.isArray((dZone.layout_hint || {}).split_hint) ? (dZone.layout_hint || {}).split_hint : null) ||
        (Array.isArray(mZone.split_hint) ? mZone.split_hint : null) ||
        (Array.isArray((mZone.layout_hint || {}).split_hint) ? (mZone.layout_hint || {}).split_hint : null) ||
        null,
      artifacts: mergedArtifacts
    }
  })

  // Recovery: append any manifest zones not matched by any designed zone.
  // Track matched manifest zones by zone_id (and by the artifact-type scan used above).
  {
    // Collect the zone_ids and artifact types that were actually matched
    const matchedManifestZoneIds = new Set()
    const matchedManifestArtTypes = new Set()
    result.forEach(rz => {
      if (rz.zone_id) matchedManifestZoneIds.add(rz.zone_id)
      ;(rz.artifacts || []).forEach(a => matchedManifestArtTypes.add(normalizeArtifactType(a?.type, a?.chart_type)))
    })

    manifestZones.forEach((mZone, mi) => {
      // A manifest zone is covered if its zone_id appears in the merged result OR
      // its primary artifact type was matched by the type-scan fallback above.
      const coveredById = mZone.zone_id && matchedManifestZoneIds.has(mZone.zone_id)
      const primaryMType = normalizeArtifactType((mZone.artifacts || [])[0]?.type, (mZone.artifacts || [])[0]?.chart_type)
      const coveredByType = primaryMType && matchedManifestArtTypes.has(primaryMType)
      const coveredByIndex = mi < result.length
      if (coveredById || coveredByType || coveredByIndex) return

      const recoveredArts = (mZone.artifacts || []).map(a => buildSafeArtifactShell(a, bt))
      result.push({
        ...mZone,
        frame: null,  // geometry assigned by buildScratchZoneFrames
        artifacts: recoveredArts
      })
      console.warn('Agent 5 **” zone recovery: manifest zone', mZone.zone_id || mi, 'not matched; re-injected from manifest')
    })
  }

  return result
}


// *********************************************************************************************************************************************************************************************************************************************
// NORMALISER
// 1. Validates the designed slide structure
// 2. Merges Agent 4 artifact content into Agent 5 layout artifacts
// 3. Carries all manifest metadata through for Agent 5.1 and Agent 6
// *********************************************************************************************************************************************************************************************************************************************

// Fills zone frames and artifact placeholder_idx from the layout's ordered content_areas.
// Called AFTER mergeContentIntoZones so artifact content comes from Agent 4, not Agent 5.
function applyLayoutZoneFrames(zones, layoutName, brand) {
  if (!layoutName) return zones
  const layouts = brand.slide_layouts || []
  const layout = layouts.find(l => (l.name || '').toLowerCase() === layoutName.toLowerCase())
    || layouts.find(l => (l.name || '').toLowerCase().includes(layoutName.toLowerCase()))
  if (!layout) return zones

  // Build ordered content areas on-the-fly (same logic as extractBrandTokens)
  const contentAreas = (layout.placeholders || [])
    .filter(p => p.type === 'body' && (p.h_in || 0) > 0.5)
    .sort((a, b) => {
      const rowA = Math.round((a.y_in || 0) * 2)
      const rowB = Math.round((b.y_in || 0) * 2)
      if (rowA !== rowB) return rowA - rowB
      return (a.x_in || 0) - (b.x_in || 0)
    })

  if (contentAreas.length === 0) return null

  const zoneCount = (zones || []).length
  if (contentAreas.length < Math.max(zoneCount, 1)) {
    console.warn('Agent 5 applyLayoutZoneFrames: layout "' + layoutName + '" has only ' + contentAreas.length + ' content area(s) but spec has ' + zoneCount + ' zone(s) **” falling back to scratch mode for this slide')
    return null
  }

  // Small body placeholders (h â‰¤ 0.5") are header labels, not content areas
  const headerPhs = (layout.placeholders || [])
    .filter(p => p.type === 'body' && (p.h_in || 0) <= 0.5)

  return zones.map((zone, zi) => {
    const ca = contentAreas[zi] || contentAreas[contentAreas.length - 1]
    if (!ca || (ca.w_in || 0) <= 0.1 || (ca.h_in || 0) <= 0.1) return null
    const frame = {
      x: ca.x_in, y: ca.y_in, w: ca.w_in, h: ca.h_in,
      padding: { top: 0.08, right: 0.08, bottom: 0.08, left: 0.08 }
    }
    const inner = {
      x: r2(frame.x + frame.padding.left),
      y: r2(frame.y + frame.padding.top),
      w: r2(Math.max(0.1, frame.w - frame.padding.left - frame.padding.right)),
      h: r2(Math.max(0.1, frame.h - frame.padding.top - frame.padding.bottom))
    }
    // Find paired header placeholder: same x-column as content area, positioned just above it
    const headerPh = headerPhs.find(p =>
      Math.abs((p.x_in || 0) - (ca.x_in || 0)) < 0.15 && (p.y_in || 0) < (ca.y_in || 0)
    )
    const zoneArtifacts = zone.artifacts || []
    const singleArtifact = zoneArtifacts.length === 1
    const artifacts = zoneArtifacts.map(a => {
      const base = { ...a, placeholder_idx: ca.idx }
      if (!singleArtifact) {
        // Force downstream stacking to recompute within the resolved placeholder frame.
        return {
          ...base,
          x: null, y: null, w: null, h: null
        }
      }

      const rebound = {
        ...base,
        x: inner.x,
        y: inner.y,
        w: inner.w,
        h: inner.h
      }

      // Layout-dependent internals must be recomputed against the actual placeholder frame.
      if (rebound.type === 'cards') {
        rebound.container = { x: inner.x, y: inner.y, w: inner.w, h: inner.h }
        rebound.card_frames = []
      } else if (rebound.type === 'workflow') {
        rebound.container = { x: inner.x, y: inner.y, w: inner.w, h: inner.h }
      } else if (rebound.type === 'table') {
        rebound.column_widths = []
        rebound.row_heights = []
        rebound.header_row_height = null
      }

      return rebound
    })
    return {
      ...zone,
      frame,
      header_ph_idx: headerPh ? headerPh.idx : null,
      artifacts
    }
  }).filter(Boolean)
}

function isValidFrame(rect) {
  return !!rect
    && Number.isFinite(+rect.x)
    && Number.isFinite(+rect.y)
    && Number.isFinite(+rect.w)
    && Number.isFinite(+rect.h)
    && (+rect.w) > 0.05
    && (+rect.h) > 0.05
}

function zonesHaveValidFrames(zones) {
  return Array.isArray(zones) && zones.length > 0 && zones.every(zone => {
    if (!isValidFrame(zone.frame)) return false
    return (zone.artifacts || []).every(art => {
      if ((zone.artifacts || []).length > 1) return true
      return isValidFrame({ x: art.x, y: art.y, w: art.w, h: art.h })
    })
  })
}

function deriveScratchContentBounds(slideSpec) {
  const canvas = slideSpec.canvas || {}
  const margin = canvas.margin || {}
  const width = +canvas.width_in || 13.33
  const height = +canvas.height_in || 7.50
  const left = +margin.left || 0.4
  const right = +margin.right || 0.4
  const topMargin = +margin.top || 0.15
  const bottom = +margin.bottom || 0.3
  const tb = slideSpec.title_block || {}
  const sb = slideSpec.subtitle_block || {}

  const HEADER_CONTENT_GAP     = 0.20   // visible breathing room below the title

  const r2sc = v => Math.round(v * 100) / 100
  const _blockBottom = (block, fallbackY, fallbackH) => {
    if (!block || !block.text) return fallbackY
    const y = block.y != null ? +block.y : fallbackY
    const h = block.h != null ? +block.h : fallbackH
    return r2sc(y + h)
  }

  const defaultTitleBottom = r2sc(topMargin + 0.6)
  const titleBottom = _blockBottom(tb, topMargin, 0.6) || defaultTitleBottom
  const subtitleBottom = sb.text
    ? _blockBottom(sb, titleBottom, 0.35)
    : titleBottom

  // In template-backed slides, Agent 5 should not infer extra title clearance
  // from font metrics. Agent 6 places the actual title/subtitle and performs the
  // only authoritative post-placement shift when the real template header is taller.
  const top = Math.max(topMargin, subtitleBottom + HEADER_CONTENT_GAP)
  return {
    x: left,
    y: top,
    w: Math.max(0.5, width - left - right),
    h: Math.max(0.5, height - top - bottom)
  }
}

function chooseScratchSplitOrientation(zones) {
  if (!Array.isArray(zones) || zones.length !== 2) return 'vertical'
  // Respect explicit artifact_arrangement hint from Agent 4 before falling back to type inference
  const explicitArrangement = (zones[0]?.layout_hint?.artifact_arrangement || zones[1]?.layout_hint?.artifact_arrangement || '').toLowerCase()
  if (explicitArrangement === 'horizontal') return 'horizontal'
  if (explicitArrangement === 'vertical')   return 'vertical'
  const primaryType = (((zones[0]?.artifacts || [])[0] || {}).type || '').toLowerCase()
  const secondaryType = (((zones[1]?.artifacts || [])[0] || {}).type || '').toLowerCase()
  if (['workflow', 'prioritization', 'matrix', 'driver_tree'].includes(primaryType)) return 'vertical'
  if (primaryType === 'chart' && secondaryType === 'insight_text') return 'horizontal'
  return 'vertical'
}

function parseScratchSplitToken(split) {
  // Handle array format [pct0, pct1] e.g. [60, 40] **” treat as left/top split
  if (Array.isArray(split) && split.length >= 2) {
    const pct = Math.max(1, Math.min(99, parseFloat(split[0]) || 50))
    return { side: 'left', pct, frac: pct / 100, orientation: 'horizontal' }
  }
  const s = String(split || '').trim().toLowerCase()
  const m = s.match(/^(left|right|top|bottom)_(\d{1,3})$/)
  if (!m) return null
  const side = m[1]
  const pct = Math.max(1, Math.min(99, parseInt(m[2], 10) || 0))
  const orientation = (side === 'left' || side === 'right') ? 'horizontal' : 'vertical'
  return { side, pct, frac: pct / 100, orientation }
}

function buildScratchZoneFrames(zones, slideSpec) {
  if (!Array.isArray(zones) || zones.length === 0) return zones
  const r2 = x => Math.round(x * 100) / 100
  const bounds = deriveScratchContentBounds(slideSpec)
  const gap = 0.18
  const framed = zones.map(z => ({ ...z, artifacts: (z.artifacts || []).map(a => ({ ...a })) }))

  if (framed.length === 1) {
    framed[0].frame = {
      x: r2(bounds.x), y: r2(bounds.y), w: r2(bounds.w), h: r2(bounds.h),
      padding: { top: 0.08, right: 0.08, bottom: 0.08, left: 0.08 }
    }
    return framed
  }

  if (framed.length === 2) {
    const z0Split = parseScratchSplitToken(framed[0]?.zone_split || framed[0]?.layout_hint?.split || framed[0]?.split_hint)
    const z1Split = parseScratchSplitToken(framed[1]?.zone_split || framed[1]?.layout_hint?.split || framed[1]?.split_hint)
    const explicit = z0Split || z1Split
    // artifact_arrangement always wins over the array-derived orientation default
    const explicitArr = (framed[0]?.layout_hint?.artifact_arrangement || framed[1]?.layout_hint?.artifact_arrangement || '').toLowerCase()
    const orientation = (explicitArr === 'horizontal' || explicitArr === 'vertical')
      ? explicitArr
      : (explicit?.orientation || chooseScratchSplitOrientation(framed))
    const primaryFrac = explicit?.frac || (String((framed[0].narrative_weight || '')).toLowerCase() === 'primary' ? 0.58 : 0.50)
    if (orientation === 'horizontal') {
      const availW = bounds.w - gap
      const leftFrac = explicit
        ? (explicit.side === 'left' ? explicit.frac : 1 - explicit.frac)
        : primaryFrac
      const leftW = r2(availW * leftFrac)
      const rightW = r2(availW - leftW)
      framed[0].frame = {
        x: r2(bounds.x), y: r2(bounds.y), w: leftW, h: r2(bounds.h),
        padding: { top: 0.08, right: 0.08, bottom: 0.08, left: 0.08 }
      }
      framed[1].frame = {
        x: r2(bounds.x + leftW + gap), y: r2(bounds.y), w: rightW, h: r2(bounds.h),
        padding: { top: 0.08, right: 0.08, bottom: 0.08, left: 0.08 }
      }
    } else {
      const availH = bounds.h - gap
      const topFrac = explicit
        ? (explicit.side === 'top' ? explicit.frac : 1 - explicit.frac)
        : primaryFrac
      const topH = r2(availH * topFrac)
      const bottomH = r2(availH - topH)
      framed[0].frame = {
        x: r2(bounds.x), y: r2(bounds.y), w: r2(bounds.w), h: topH,
        padding: { top: 0.08, right: 0.08, bottom: 0.08, left: 0.08 }
      }
      framed[1].frame = {
        x: r2(bounds.x), y: r2(bounds.y + topH + gap), w: r2(bounds.w), h: bottomH,
        padding: { top: 0.08, right: 0.08, bottom: 0.08, left: 0.08 }
      }
    }
    return framed
  }

  const eachH = r2((bounds.h - gap * (framed.length - 1)) / framed.length)
  framed.forEach((zone, zi) => {
    zone.frame = {
      x: r2(bounds.x),
      y: r2(bounds.y + zi * (eachH + gap)),
      w: r2(bounds.w),
      h: eachH,
      padding: { top: 0.08, right: 0.08, bottom: 0.08, left: 0.08 }
    }
  })
  return framed
}

function sanitizeBlocks(blocks, slideSpec) {
  const r2 = x => Math.round(x * 100) / 100
  const bounds = deriveScratchContentBounds(slideSpec)
  return (blocks || []).filter(Boolean).map(block => {
    if (block.x == null) block.x = bounds.x
    if (block.y == null) block.y = bounds.y
    if (block.w == null || block.w <= 0) block.w = Math.max(0.4, bounds.w)
    if (block.h == null || block.h <= 0) block.h = Math.max(0.12, Math.min(0.8, bounds.h * 0.25))
    block.x = r2(Math.max(0, +block.x || 0))
    block.y = r2(Math.max(0, +block.y || 0))
    block.w = r2(Math.max(0.1, +block.w || 0.1))
    block.h = r2(Math.max(0.1, +block.h || 0.1))
    return block
  })
}

function normaliseDesignedSlide(designed, manifestSlide, brand) {
  if (!designed || typeof designed !== 'object') return null  // caller handles null -> fallback

  const branded = applyBrandGuidelineOverrides(designed, manifestSlide, brand)
  branded.zones = (branded.zones || []).map(zone => ({
    ...zone,
    artifacts: (zone.artifacts || []).map(normalizeArtifactDefinition)
  }))

  // Derive bt from the brand rulebook **” authoritative source, never depends on
  // what Claude returned per-slide (brand_tokens is no longer in Claude output).
  const bt = {
    primary_color:      (brand.primary_colors    || [])[0] || '#1A3C8F',
    secondary_color:    (brand.secondary_colors  || [])[0] || '#E8A020',
    title_color:        (brand.primary_colors    || [])[0] || '#1A3C8F',
    body_color:         (brand.text_colors       || [])[0] || '#111111',
    caption_color:      '#888888',
    title_font_family:  (brand.title_font   || {}).family  || 'Arial',
    body_font_family:   (brand.body_font    || {}).family  || 'Arial',
    caption_font_family:(brand.caption_font || {}).family  || 'Arial',
    accent_colors:      brand.accent_colors        || [],
    chart_palette:      brand.chart_color_sequence || brand.chart_colors || [],
    uses_template:      brand.uses_template        || false,
    slide_width_inches:  brand.slide_width_inches  || 13.33,
    slide_height_inches: brand.slide_height_inches || 7.50
  }
  // Keep brand_tokens on the slide object for internal processing only **”
  // it is stripped from every slide before runAgent5 returns.
  branded.brand_tokens = bt

  const inputIssues = validateDesignedSlide(branded)
  if (inputIssues.length > 0) {
    console.warn('Agent 5 -- S' + (branded.slide_number || '?') + ' input issues:', inputIssues.join('; '))
  }

  // Merge Agent 4 content into Agent 5 layout zones
  const mergedZones = mergeContentIntoZones(
    branded.zones || [],
    manifestSlide.zones || [],
    bt
  )

  // Layout mode: fill zone frames + artifact placeholder_idx from the layout's content_areas.
  // This runs after merge so Agent 4's artifact content is already in place.
  const layoutName = manifestSlide.selected_layout_name || designed.selected_layout_name || ''
  const manifestSlideType = String(manifestSlide.slide_type || designed.slide_type || '').toLowerCase()
  const isTemplateNonContent = manifestSlideType === 'title' || manifestSlideType === 'divider'
  // Content slides are layout-mode only when there is an actual named layout.
  // Prevent impossible states like layout_mode:true with selected_layout_name:""
  // which make Agent 5 previews diverge from Agent 6 rendering.
  const isLayoutMode = isTemplateNonContent ? !!(designed.layout_mode || layoutName) : !!layoutName
  const brandedWithLayoutTitle = isLayoutMode && layoutName
    ? applyLayoutTitleFrames(branded, layoutName, brand)
    : branded
  let finalZones = isLayoutMode && layoutName
    ? applyLayoutZoneFrames(mergedZones, layoutName, brand)
    : mergedZones
  if (!zonesHaveValidFrames(finalZones) || (Array.isArray(finalZones) && finalZones.length !== mergedZones.length)) {
    console.warn('Agent 5 -- S' + (manifestSlide.slide_number || '?') + ' switching to scratch framing for invalid/incompatible layout:', layoutName)
    finalZones = buildScratchZoneFrames(mergedZones, brandedWithLayoutTitle)
  }

  // ****** Enforce minimum gap between slide title and content zones *********************************************
  // MUST run BEFORE computeArtifactInternals: that function reads zone.frame.y
  // to compute art.y for artifacts without explicit positions.  Running after
  // would leave those positions computed from the un-enforced frame.
  // We also shift explicit art.y / header_block.y / container.y on artifacts
  // that already have absolute positions set by Agent 5, because computeArtifactInternals
  // will NOT override those (it only fills nulls).
  {
    const r2 = v => Math.round(v * 100) / 100
    const HEADER_CONTENT_GAP = 0.20
    const tb = brandedWithLayoutTitle.title_block || {}
    const sb = brandedWithLayoutTitle.subtitle_block || {}
    const topMargin = +(brandedWithLayoutTitle.canvas && brandedWithLayoutTitle.canvas.margin && brandedWithLayoutTitle.canvas.margin.top) || 0.30

    const _blockBottom = (block, fallbackY, fallbackH) => {
      if (!block || !block.text) return fallbackY
      const y = block.y != null ? +block.y : fallbackY
      const h = block.h != null ? +block.h : fallbackH
      return r2(y + h)
    }

    const titleBottom = _blockBottom(tb, topMargin, 0.60)
    const subtitleBottom = sb.text
      ? _blockBottom(sb, titleBottom, 0.35)
      : titleBottom
    const minContentY = r2(Math.max(topMargin, subtitleBottom + HEADER_CONTENT_GAP))

    finalZones.forEach(zone => {
      const frame = zone.frame
      if (!frame || frame.y == null) return
      const fy = +frame.y
      if (fy < minContentY) {
        const shift = r2(minContentY - fy)
        frame.y = minContentY
        if (frame.h != null) frame.h = r2(Math.max(0.20, +frame.h - shift))
        // Shift artifact absolute positions so they stay in sync.
        // computeArtifactInternals only fills nulls **” it won't correct explicit values.
        ;(zone.artifacts || []).forEach(art => {
          if (art.y != null) art.y = r2(+art.y + shift)
          if (art.header_block && art.header_block.y != null) {
            art.header_block.y = r2(+art.header_block.y + shift)
          }
          if (art.container && art.container.y != null) {
            art.container.y = r2(+art.container.y + shift)
          }
        })
      }
    })
  }

  // ****** Canvas overflow correction (scratch-mode slides) ************************************************************************
  // Fixes cases where Claude emits an oversized title_block.h, pushing all zones
  // below the canvas bottom.  Runs AFTER gap enforcement, BEFORE computeArtifactInternals.
  if (!isLayoutMode) {
    const _r2 = v => Math.round(v * 100) / 100
    const cv = brandedWithLayoutTitle.canvas || {}
    const canvasH = +cv.height_in || bt.slide_height_inches || 7.50
    const mBottom = +(cv.margin && cv.margin.bottom) || 0.37
    const canvasBottom = _r2(canvasH - mBottom)

    // Step 1: Cap oversized title block height (> 1.2" is always wrong for a title)
    const _tb = brandedWithLayoutTitle.title_block
    if (_tb && _tb.h != null && +_tb.h > 1.2) {
      _tb.h = 1.2
    }

    // Step 2: Shift + scale zones that overflow canvas bottom
    const _activeZones = finalZones.filter(z => z.frame && z.frame.y != null && z.frame.h != null)
    if (_activeZones.length > 0) {
      const _topY    = Math.min(..._activeZones.map(z => +z.frame.y))
      const _bottomY = Math.max(..._activeZones.map(z => _r2(+z.frame.y + +z.frame.h)))
      if (_bottomY > canvasBottom + 0.01) {
        const _titleBottom = (_tb && _tb.y != null && _tb.h != null) ? _r2(+_tb.y + +_tb.h) : 0.9
        const _idealTop    = _r2(_titleBottom + 0.20)
        const _contentH    = _r2(_bottomY - _topY)
        const _availH      = _r2(canvasBottom - _idealTop)
        const _scale       = (_availH > 0.1 && _contentH > 0) ? Math.min(1.0, _r2(_availH / _contentH)) : 1.0
        const _yShift      = _r2(_idealTop - _topY)
        _activeZones.forEach(zone => {
          const f    = zone.frame
          const newY = _r2(_topY + (+f.y - _topY + _yShift) * _scale)
          f.h = _r2(+f.h * _scale)
          f.y = newY
          ;(zone.artifacts || []).forEach(art => {
            if (art.y != null && art.h != null) {
              art.y = _r2(_topY + (+art.y - _topY + _yShift) * _scale)
              art.h = _r2(+art.h * _scale)
            }
            if (art.header_block && art.header_block.y != null) {
              art.header_block.y = _r2(_topY + (+art.header_block.y - _topY + _yShift) * _scale)
              if (art.header_block.h != null) art.header_block.h = _r2(+art.header_block.h * _scale)
            }
            if (art.container && art.container.y != null) {
              art.container.y = _r2(_topY + (+art.container.y - _topY + _yShift) * _scale)
              if (art.container.h != null) art.container.h = _r2(+art.container.h * _scale)
            }
          })
        })
        console.warn('Agent 5 -- canvas overflow corrected: scale=' + _scale + ' yShift=' + _yShift +
          ' (zones were ' + _r2(_bottomY) + '" > canvas ' + canvasBottom + '")')
      }
    }
  }

  // Post-process: fill computed layout/sizing fields (stacking, chart, table, cards, font scaling)
  // so that generate_pptx.py can act as a pure renderer reading pre-computed values.
  computeArtifactInternals(finalZones, branded.canvas || {}, bt)
  normalizeArtifactHeaderBands(finalZones)

  finalZones.forEach((zone, zi) => {
    ;(zone.artifacts || []).forEach((art, ai) => {
      if (!art._artifact_id) art._artifact_id = 's' + (manifestSlide.slide_number || '?') + '_z' + zi + '_a' + ai
    })
  })

  const finalArtifactIssues = validateDesignedSlide({
    ...brandedWithLayoutTitle,
    zones: finalZones
  })
  if (finalArtifactIssues.length > 0) {
    console.warn('Agent 5 -- S' + (manifestSlide.slide_number || '?') + ' final spec issues:', finalArtifactIssues.join('; '))
  }

  // Flatten to blocks[] **” ordered, self-contained render units.
  // Validation runs on the raw flat blocks; output is grouped into artifact_groups[]
  // to eliminate repeated artifact metadata (artifact_id, artifact_type, etc.) on
  // every block. generate_pptx.py / Agent 6 call flattenArtifactGroups() to restore.
  const rawBlocks = sanitizeBlocks(flattenToBlocks(
    { ...brandedWithLayoutTitle, zones: finalZones },
    bt
  ), brandedWithLayoutTitle)
  // Temporarily attach flat blocks for validateRenderCompleteness (which reads slide.blocks)
  brandedWithLayoutTitle.blocks = rawBlocks
  const renderIssues = validateRenderCompleteness({ ...brandedWithLayoutTitle, zones: finalZones })
  // Replace flat blocks with grouped form before output
  brandedWithLayoutTitle.artifact_groups = groupBlocksByArtifact(rawBlocks)
  delete brandedWithLayoutTitle.blocks
  if (renderIssues.length > 0) {
    console.warn('Agent 5 -- S' + (manifestSlide.slide_number || '?') + ' render issues:', renderIssues.join('; '))
  }
  const criticalRenderIssues = renderIssues.filter(i =>
    i.includes('workflow missing nodes') ||
    i.includes('line block missing endpoints') ||
    i.includes('missing chart_style') ||
    i.includes('missing series_style') ||
    i.includes('missing table_style') ||
    i.includes('missing card_frames') ||
    i.includes('missing heading_style') ||
    i.includes('missing body_style') ||
    i.includes('artifact blocks overlap')
  )
  if (criticalRenderIssues.length > 0) {
    console.warn('Agent 5 -- S' + (manifestSlide.slide_number || '?') + ' rejected after block flattening:', criticalRenderIssues.join('; '))
    return null
  }

  // Log merge summary
  const contentCounts = { insight_text: 0, chart: 0, cards: 0, workflow: 0, table: 0 }
  finalZones.forEach(z => (z.artifacts || []).forEach(a => {
    if (contentCounts[a.type] !== undefined) contentCounts[a.type]++
  }))
  console.log('  S' + manifestSlide.slide_number + ' merged content:',
    Object.entries(contentCounts).filter(([,n]) => n > 0).map(([t,n]) => t + ':' + n).join(' ') || 'none')

  const {
    zones: _ignoredZones,
    title_block: _ignoredTitleBlock,
    subtitle_block: _ignoredSubtitleBlock,
    ...brandedWithoutZones
  } = brandedWithLayoutTitle

  return {
    ...brandedWithoutZones,
    // Always override with manifest ground truth (Claude may drift on slide_number etc.)
    slide_number:          manifestSlide.slide_number,
    slide_type:            manifestSlide.slide_type            || designed.slide_type,
    // Layout mode fields **” ground truth from Agent 4 manifest
    layout_mode:           isLayoutMode,
    selected_layout_name:  manifestSlide.selected_layout_name  || designed.selected_layout_name || '',
    // Slide-level content metadata
    title:            (brandedWithLayoutTitle.title_block || {}).text || manifestSlide.title    || '',
    subtitle:         (brandedWithLayoutTitle.subtitle_block || {}).text || manifestSlide.subtitle || '',
    key_message:      manifestSlide.key_message      || '',
    speaker_note:     manifestSlide.speaker_note     || '',
    // Condensed structural summary for Agent 5.1 review/debug; final render contract is blocks[].
    zones_summary:    finalZones.map(z => ({
      zone_id:          z.zone_id,
      zone_role:        z.zone_role,
      narrative_weight: z.narrative_weight,
      artifact_types:   (z.artifacts || []).map(a => artifactSignatureType(a))
    })),
    _validation_issues: finalArtifactIssues.length > 0 ? finalArtifactIssues : undefined,
    _source_validation_issues: inputIssues.length > 0 ? inputIssues : undefined,
    _render_validation_issues: renderIssues.length > 0 ? renderIssues : undefined
  }
}

function inferLayoutName(manifestSlide, brand) {
  const st   = manifestSlide.slide_type      || 'content'
  const _NON_CONTENT_TYPES = new Set(['title', 'sechead', 'blank'])
  const isNonContent = (l) => {
    const t = (l.type || '').toLowerCase()
    const n = (l.name || '').toLowerCase()
    return _NON_CONTENT_TYPES.has(t) ||
      /^blank$/i.test(n) ||
      /thank[\s_-]*you|end[\s_-]*slide|closing[\s_-]*slide|section[\s_-]*header|^section$|divider|title slide/i.test(n)
  }
  const allLayouts = brand.slide_layouts || []
  const avail = st === 'content' ? allLayouts.filter(l => !isNonContent(l)) : allLayouts
  const find = (kws) => avail.find(l => kws.some(k => (l.name || '').toLowerCase().includes(k.toLowerCase())))
  const zones = manifestSlide.zones || []
  const artifacts = zones.flatMap(z => z.artifacts || [])
  const artifactTypes = artifacts.map(a => String(a.type || '').toLowerCase())
  const zoneCount = zones.length
  const hasReasoning = artifactTypes.some(t => ['matrix', 'driver_tree', 'prioritization'].includes(t))
  const hasStructuredDisplay = artifactTypes.some(t => ['comparison_table', 'initiative_map', 'profile_card_set', 'risk_register'].includes(t))
  const hasWorkflow = artifactTypes.includes('workflow')
  const hasWideWorkflow = artifacts.some(a => {
    const t = String(a.type || '').toLowerCase()
    const dir = String(a.flow_direction || '').toLowerCase()
    return t === 'workflow' && (dir === 'left_to_right' || dir === 'timeline')
  })
  const hasChart = artifactTypes.includes('chart')
  const hasCards = artifactTypes.includes('cards')
  const hasOnlyInsight = artifactTypes.length > 0 && artifactTypes.every(t => t === 'insight_text')
  const selectedLayout = String(manifestSlide.selected_layout_name || '').trim()

  if (st === 'title')   return (find(['Title Slide', 'title'])                     || {}).name || 'Title Slide'
  if (st === 'divider') return (find(['Section', 'Divider', 'section header'])     || {}).name || 'Section Divider'
  if (selectedLayout) return selectedLayout
  if (hasReasoning)
    return (find(['Body Text', '1 Across', 'body text', '1 across', 'single'])     || {}).name || 'Body Text'
  if (hasWideWorkflow)
    return (find(['Body Text', '1 Across', 'body text', '1 across'])               || {}).name || 'Body Text'
  if (hasWorkflow && zoneCount >= 2)
    return (find(['Body Text', '1 Across', 'body text', '1 across'])               || {}).name || 'Body Text'
  if (zoneCount >= 3 && (hasChart || hasCards || hasStructuredDisplay))
    return (find(['1 Across', '2 Across', '1 across', '2 across'])                 || {}).name || '1 Across'
  if (zoneCount === 2 && !hasOnlyInsight)
    return (find(['2 Across', '2 across', '2 Column', '2 column', '1 on 1'])      || {}).name || '2 Across'
  if (hasOnlyInsight)
    return (find(['Body Text', '1 Across', 'body text', '1 across'])               || {}).name || 'Body Text'
  return (find(['1 Across','Body Text','1 across','body text','2 Column','2 column']) || {}).name || (avail[0] || {}).name || 'Body Text'
}


// *********************************************************************************************************************************************************************************************************************************************
// MAIN RUNNER
// *********************************************************************************************************************************************************************************************************************************************

async function runAgent5(state) {
  const manifest = state.slideManifest
  const brand    = state.brandRulebook

  if (!manifest || !manifest.length) {
    console.error('Agent 5 -- slideManifest is empty')
    return []
  }

  const tokens = extractBrandTokens(brand)
  console.log('Agent 5 starting -- slides:', manifest.length)
  console.log('  Primary color:', (tokens.primary_colors || [])[0] || 'none')
  console.log('  Slide size:', tokens.slide_width_inches + '" x ' + tokens.slide_height_inches + '"')
  console.log('  Title font:', (tokens.title_font || {}).family || 'default')

  // Batch rules:
  //   - Repaired slides (flagged by Agent 4 repair) â†’ solo batch of 1 (complex content, avoid overflow)
  //   - Structural slides (title / divider / thank_you) â†’ batch of 3 (light, text-only)
  //   - Regular content slides â†’ batch of 2
  const STRUCTURAL_TYPES = new Set(['title', 'divider', 'thank_you', 'thankyou'])
  const repairedSlides   = manifest.filter(s => s._was_repaired)
  const structuralSlides = manifest.filter(s => !s._was_repaired && STRUCTURAL_TYPES.has((s.slide_type || '').toLowerCase()))
  const contentSlides    = manifest.filter(s => !s._was_repaired && !STRUCTURAL_TYPES.has((s.slide_type || '').toLowerCase()))
  if (repairedSlides.length > 0) {
    console.log('  Slides from Agent 4 repair (will be processed solo):', repairedSlides.map(s => s.slide_number).join(', '))
  }

  const batches = []
  // Structural slides: batch of 3
  for (let i = 0; i < structuralSlides.length; i += 3) batches.push(structuralSlides.slice(i, i + 3))
  // Content slides: batch of 2
  for (let i = 0; i < contentSlides.length;   i += 2) batches.push(contentSlides.slice(i, i + 2))
  // Repaired slides: solo, inserted in slide_number order
  for (const rs of repairedSlides) {
    let insertIdx = batches.findIndex(b => b.some(s => s.slide_number > rs.slide_number))
    if (insertIdx === -1) insertIdx = batches.length
    batches.splice(insertIdx, 0, [rs])
  }
  // Re-sort all batches so they run in slide_number order
  batches.sort((a, b) => (a[0]?.slide_number || 0) - (b[0]?.slide_number || 0))
  console.log('  Batches:', batches.length,
    '| structural(3):', Math.ceil(structuralSlides.length / 3),
    '| content(2):', Math.ceil(contentSlides.length / 2),
    '| repaired(1):', repairedSlides.length)

  const allDesigned = []
  const manifestBySlide = new Map((manifest || []).map(s => [s.slide_number, s]))

  // Minimum interval between batch API calls (ms). Adaptive: if the batch itself took
  // longer than this threshold (e.g. due to a slow Claude response), no extra sleep needed.
  const BATCH_MIN_INTERVAL_MS = 45000

  let batchStartTime = 0
  for (let b = 0; b < batches.length; b++) {
    if (b > 0) {
      const elapsed = Date.now() - batchStartTime
      const remaining = BATCH_MIN_INTERVAL_MS - elapsed
      if (remaining > 500) {
        console.log('Agent 5 -- rate limit pause: waiting', Math.ceil(remaining / 1000) + 's before batch', b + 1, '...')
        await new Promise(r => setTimeout(r, remaining))
      } else {
        console.log('Agent 5 -- batch', b + 1, 'starting immediately (prior batch used ' + Math.ceil(elapsed / 1000) + 's)')
      }
    }
    batchStartTime = Date.now()
    const batch  = batches[b]
    const result = await designSlideBatch(batch, brand, b + 1)

    if (!result) {
      // Entire batch failed to parse **” brief pause then fall back per slide via Claude.
      // Without a pause, consecutive Claude calls after a truncation/rate-limit failure
      // would immediately hit the same limit and produce minimal-safe-slides.
      console.warn('Agent 5 -- batch', b + 1, 'failed entirely, pausing 5s before per-slide fallbacks')
      await new Promise(r => setTimeout(r, 5000))
      for (const ms of batch) {
        const fb = await buildFallbackDesign(ms, brand)
        allDesigned.push(normaliseDesignedSlide(fb, ms, brand) || buildMinimalSafeSlide(ms, tokens))
      }
      continue
    }

    // Match each manifest slide to the returned result
    for (let i = 0; i < batch.length; i++) {
      const mSlide = batch[i]
      const match  = result.find(r => r.slide_number === mSlide.slide_number)
                  || result[i]
                  || null

      if (!match) {
        console.warn('Agent 5 -- no match for S' + mSlide.slide_number + ', running fallback')
        const fb = await buildFallbackDesign(mSlide, brand)
        allDesigned.push(normaliseDesignedSlide(fb, mSlide, brand) || buildMinimalSafeSlide(mSlide, tokens))
        continue
      }

      const issues = validateDesignedSlide(match)
      const normalized = normaliseDesignedSlide(match, mSlide, brand)
      if (normalized) {
        const finalIssues = []
          .concat(normalized._validation_issues || [])
          .concat(normalized._render_validation_issues || [])
        const fatalFinal = finalIssues.filter(i =>
          i.includes('missing canvas') ||
          i.includes('no zones') ||
          i.includes('no artifacts') ||
          i.includes('no blocks')
        )
        if (fatalFinal.length === 0) {
          if (issues.length > 0) {
            console.warn('Agent 5 -- S' + mSlide.slide_number + ' repaired during normalization:', issues.join('; '))
          }
          allDesigned.push(normalized)
          continue
        }
        console.warn('Agent 5 -- S' + mSlide.slide_number + ' still has fatal post-normalization issues, running fallback:', fatalFinal.join('; '))
      } else if (issues.length > 0) {
        console.warn('Agent 5 -- S' + mSlide.slide_number + ' rejected during normalization after issues:', issues.join('; '))
      } else {
        console.warn('Agent 5 -- S' + mSlide.slide_number + ' rejected during normalization, running fallback')
      }

      const fb = await buildFallbackDesign(mSlide, brand)
      allDesigned.push(normaliseDesignedSlide(fb, mSlide, brand) || buildMinimalSafeSlide(mSlide, tokens))
    }
  }

  // Summary
  const withIssues    = allDesigned.filter(s => s._validation_issues?.length > 0)
  const withFallback  = allDesigned.filter(s => s._fallback)
  const typeCounts    = {}
  allDesigned.forEach(s =>
    (s.zones_summary || []).forEach(z =>
      (z.artifact_types || []).forEach(t => { typeCounts[t] = (typeCounts[t] || 0) + 1 })
    )
  )

  console.log('Agent 5 complete')
  console.log('  Slides:', allDesigned.length,
    '| fallback:', withFallback.length,
    '| with issues:', withIssues.length)
  console.log('  Artifact types:', JSON.stringify(typeCounts))

  const finalDesigned = allDesigned.map(slide => {
    if (slide && Array.isArray(slide.artifact_groups) && slide.artifact_groups.length > 0) return slide
    const manifestSlide = manifestBySlide.get(slide?.slide_number)
    if (!manifestSlide) return slide
    console.warn('Agent 5 -- S' + manifestSlide.slide_number + ' missing blocks at final handoff, forcing minimal safe render spec')
    return buildMinimalSafeSlide(manifestSlide, tokens)
  })

  // Hoist brand_tokens to the top level **” derived authoritatively from the brand
  // rulebook (same source used throughout Agent 5), not from slide output which
  // may be absent (e.g. single-slide decks or fallback paths that skip brand_tokens).
  const hoistedBrandTokens = {
    title_font_family:  (tokens.title_font  || {}).family || 'Arial',
    body_font_family:   (tokens.body_font   || {}).family || 'Arial',
    caption_font_family:(tokens.caption_font|| {}).family || 'Arial',
    primary_color:      (tokens.primary_colors   || [])[0] || '#1A3C8F',
    secondary_color:    (tokens.secondary_colors || [])[0] || '#E8A020',
    title_color:        (tokens.primary_colors   || [])[0] || '#1A3C8F',
    body_color:         (tokens.text_colors      || [])[0] || '#111111',
    caption_color:      '#888888',
    accent_colors:      tokens.accent_colors        || [],
    chart_palette:      tokens.chart_color_sequence || tokens.chart_colors || [],
    uses_template:      tokens.uses_template        || false
  }

  // Sort slides into the exact sequence defined by Agent 4 (slide_number ascending).
  // Batches can complete in insertion order but repaired-slide interleaving or any
  // future parallelism could disturb position **” sort here so the renderer always gets
  // title â†’ content â†’ dividers â†’ closing in the right order.
  const manifestOrder = new Map((manifest || []).map((s, i) => [s.slide_number, i]))
  finalDesigned.sort((a, b) => {
    const ia = manifestOrder.has(a?.slide_number) ? manifestOrder.get(a.slide_number) : 9999
    const ib = manifestOrder.has(b?.slide_number) ? manifestOrder.get(b.slide_number) : 9999
    return ia - ib
  })

  // Verify sequence and warn on any gaps
  const slideNums = finalDesigned.map(s => s?.slide_number).filter(n => n != null)
  const missing = (manifest || []).map(s => s.slide_number).filter(n => !slideNums.includes(n))
  if (missing.length > 0) {
    console.warn('Agent 5 -- missing slide numbers in output:', missing.join(', '))
  }
  console.log('Agent 5 -- final slide sequence:', slideNums.join(', '))

  // Strip internal-only fields from every slide before handing off to Agent 6 / renderer.
  // brand_tokens: renderer reads from the top-level key.
  // zones_summary, _*_validation_issues: debug metadata **” not read by Agent 6 or generate_pptx.
  const slides = finalDesigned.map(slide => {
    if (!slide) return slide
    const out = Object.assign({}, slide)
    delete out.brand_tokens
    delete out.zones_summary
    delete out._validation_issues
    delete out._source_validation_issues
    delete out._render_validation_issues
    return out
  })

  console.log('Agent 5 -- brand_tokens hoisted to top level; removed from', slides.length, 'slides')
  return { brand_tokens: hoistedBrandTokens, slides }
}
