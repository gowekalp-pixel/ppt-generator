// ─── AGENT 5 — DESIGN EXECUTOR ───────────────────────────────────────────────
// Input:  state.slideManifest  — from Agent 4 (zones, artifacts, archetypes)
//         state.brandRulebook  — from Agent 2 (colors, fonts, layouts, placeholders)
//         state.outline        — from Agent 3 (brief, doc type)
//
// Output: designedSpec — flat JSON array, one fully specified slide per entry
//         Each slide has elements[] — every element with exact x,y,w,h + styling
//
// Agent 5 does NOT review or critique.
// Agent 5 translates messaging structure into positioned visual elements.
// Agent 5.1 handles review and revision.

// ═══════════════════════════════════════════════════════════════════════════════
// PART A — BRAND & LAYOUT DETECTION
// ═══════════════════════════════════════════════════════════════════════════════

function detectLayoutMode(brand) {
  const layouts = brand.slide_layouts || []

  // Find the primary content layout (1 across or body text)
  const contentLayout = layouts.find(l =>
    l.name && (l.name.toLowerCase().includes('1 across') || l.name.toLowerCase().includes('body text'))
  )

  if (contentLayout && contentLayout.placeholders && contentLayout.placeholders.length > 0) {
    const phs = contentLayout.placeholders

    // Sidebar — narrow tall placeholder on left (x<1.5, w<2.5, h>3)
    const sidebar = phs.find(p => p.x_in !== undefined && p.x_in < 1.5 && p.w_in < 2.5 && p.h_in > 3)

    // Main content area — wide placeholder (x>2, w>5, h>3)
    const contentPH = phs.find(p => p.x_in > 2 && p.w_in > 5 && p.h_in > 3)

    // Full-width title — wide placeholder at top (y<2, w>8)
    const titlePH = phs.find(p => p.w_in > 8 && p.y_in < 2 && p.h_in < 1.5)

    if (sidebar && contentPH) {
      return {
        mode:       'sidebar',
        sidebar:    { x: sidebar.x_in,   y: sidebar.y_in,   w: sidebar.w_in,   h: sidebar.h_in },
        content:    { x: contentPH.x_in, y: contentPH.y_in, w: contentPH.w_in, h: contentPH.h_in },
        has_coords: true
      }
    }

    if (titlePH && contentPH) {
      return {
        mode:       'top_title',
        title:      { x: titlePH.x_in,   y: titlePH.y_in,   w: titlePH.w_in,   h: titlePH.h_in },
        content:    { x: contentPH.x_in, y: contentPH.y_in, w: contentPH.w_in, h: contentPH.h_in },
        has_coords: true
      }
    }
  }

  // Standard grid fallback
  return { mode: 'standard_grid', has_coords: false }
}

function getSlideFooterLayout(brand, layouts) {
  // Find footer placeholder
  const allPHs = layouts.flatMap(l => l.placeholders || [])
  const datePH = allPHs.find(p => p.type === 'dt')
  const footerPH = allPHs.find(p => p.type === 'ftr')
  const slideNumPH = allPHs.find(p => p.default_text === 'Slide number' || (p.x_in > 9 && p.y_in > 7))

  return {
    date:       datePH    ? { x: datePH.x_in,    y: datePH.y_in,    w: datePH.w_in,    h: datePH.h_in }    : null,
    footer:     footerPH  ? { x: footerPH.x_in,  y: footerPH.y_in,  w: footerPH.w_in,  h: footerPH.h_in }  : null,
    slide_num:  slideNumPH? { x: slideNumPH.x_in, y: slideNumPH.y_in,w: slideNumPH.w_in, h: slideNumPH.h_in }: null
  }
}

function getDimensions(brand) {
  return { w: brand.slide_width_inches || 11.02, h: brand.slide_height_inches || 8.29 }
}

function getBrandColors(brand) {
  return {
    primary:    (brand.primary_colors    || ['#0F2FB5'])[0],
    secondary:  (brand.secondary_colors  || ['#FF8E00'])[0],
    bg:         (brand.background_colors || ['#FFFFFF'])[0],
    text:       (brand.text_colors       || ['#000000'])[0],
    positive:   '#2D962D',
    warning:    '#D60202',
    neutral:    '#555555',
    chart:      brand.chart_colors || ['#0F2FB5','#FF8E00','#2D962D','#D60202','#2E046C','#0092D0'],
    titleFont:  (brand.title_font || {}).family || 'Arial',
    bodyFont:   (brand.body_font  || {}).family || 'Arial'
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// PART B — ZONE COORDINATE CALCULATOR
// ═══════════════════════════════════════════════════════════════════════════════

// Converts a split string into {x, y, w, h} within the given content area

function zoneRect(split, contentArea) {
  const { x, y, w, h } = contentArea
  const GAP = 0.15  // gap between zones in inches

  const rects = {
    // 1 zone
    'full':               { x, y, w, h },

    // 2 side-by-side
    'left_50':            { x,             y, w: w/2 - GAP/2,       h },
    'right_50':           { x: x+w/2+GAP/2, y, w: w/2 - GAP/2,     h },
    'left_60':            { x,             y, w: w*0.60 - GAP/2,     h },
    'right_40':           { x: x+w*0.60+GAP/2, y, w: w*0.40-GAP/2,  h },
    'left_40':            { x,             y, w: w*0.40 - GAP/2,     h },
    'right_60':           { x: x+w*0.40+GAP/2, y, w: w*0.60-GAP/2,  h },

    // 2 stacked
    'top_30':             { x, y,             w, h: h*0.30 - GAP/2    },
    'bottom_70':          { x, y: y+h*0.30+GAP/2, w, h: h*0.70-GAP/2 },
    'top_40':             { x, y,             w, h: h*0.40 - GAP/2    },
    'bottom_60':          { x, y: y+h*0.40+GAP/2, w, h: h*0.60-GAP/2 },
    'top_50':             { x, y,             w, h: h*0.50 - GAP/2    },
    'bottom_50':          { x, y: y+h*0.50+GAP/2, w, h: h*0.50-GAP/2 },

    // 3 zones
    'top_left_50':        { x,             y,             w: w/2-GAP/2,   h: h/2-GAP/2 },
    'top_right_50':       { x: x+w/2+GAP/2, y,            w: w/2-GAP/2,   h: h/2-GAP/2 },
    'bottom_full':        { x, y: y+h/2+GAP/2,             w,             h: h/2-GAP/2 },
    'left_full_50':       { x,             y,             w: w/2-GAP/2,   h            },
    'top_right_50_h':     { x: x+w/2+GAP/2, y,            w: w/2-GAP/2,   h: h/2-GAP/2 },
    'bottom_right_50_h':  { x: x+w/2+GAP/2, y: y+h/2+GAP/2, w: w/2-GAP/2,h: h/2-GAP/2 },

    // 4 zones
    'tl':                 { x,             y,             w: w/2-GAP/2,   h: h/2-GAP/2 },
    'tr':                 { x: x+w/2+GAP/2, y,            w: w/2-GAP/2,   h: h/2-GAP/2 },
    'bl':                 { x,             y: y+h/2+GAP/2, w: w/2-GAP/2,  h: h/2-GAP/2 },
    'br':                 { x: x+w/2+GAP/2, y: y+h/2+GAP/2, w: w/2-GAP/2,h: h/2-GAP/2 }
  }

  return rects[split] || rects['full']
}

// ═══════════════════════════════════════════════════════════════════════════════
// PART C — ELEMENT PRIMITIVE BUILDERS
// ═══════════════════════════════════════════════════════════════════════════════

const R = (n) => Math.round(n * 1000) / 1000  // round to 3dp

function elShape(id, x, y, w, h, fill, borderColor, borderPt) {
  return { id, type: 'shape', x:R(x), y:R(y), w:R(w), h:R(h), fill_color: fill||'#CCCCCC', border_color: borderColor||null, border_pt: borderPt||0 }
}

function elText(id, x, y, w, h, text, font, size, bold, color, align, valign, italic, rotation) {
  const el = { id, type: 'text_box', x:R(x), y:R(y), w:R(w), h:R(h), text: String(text||''), font: font||'Arial', size: size||12, bold:!!bold, italic:!!italic||false, color: color||'#000000', align: align||'left', valign: valign||'top', wrap: true }
  if (rotation) el.rotation = rotation
  return el
}

function elChart(id, x, y, w, h, chartType, chartTitle, categories, series, colors, showLabels, showLegend, xLabel, yLabel) {
  return {
    id, type: 'chart', chart_type: chartType||'bar',
    x:R(x), y:R(y), w:R(w), h:R(h),
    chart_title:      chartTitle   || '',
    x_label:          xLabel       || '',
    y_label:          yLabel       || '',
    categories:       categories   || [],
    series:           series       || [],
    colors:           colors       || [],
    show_data_labels: showLabels   !== false,
    show_legend:      !!showLegend
  }
}

function elTable(id, x, y, w, h, headers, rows, headerBg, headerText, altBg, font, fontSize, highlightRows) {
  return {
    id, type: 'table', x:R(x), y:R(y), w:R(w), h:R(h),
    headers:        headers      || [],
    rows:           rows         || [],
    header_bg:      headerBg     || '#0F2FB5',
    header_text:    headerText   || '#FFFFFF',
    alt_row_bg:     altBg        || '#F5F5F5',
    font:           font         || 'Arial',
    font_size:      fontSize     || 11,
    header_font_size: (fontSize||11) + 1,
    highlight_rows: highlightRows|| []
  }
}

function elWorkflow(id, x, y, w, h, workflowType, flowDirection, nodes, connections, colors, title) {
  return {
    id, type: 'workflow',
    workflow_type:  workflowType  || 'process_flow',
    flow_direction: flowDirection || 'left_to_right',
    x:R(x), y:R(y), w:R(w), h:R(h),
    title:          title         || '',
    nodes:          nodes         || [],
    connections:    connections   || [],
    node_fill:      colors ? colors[0] : '#0F2FB5',
    node_text:      '#FFFFFF',
    arrow_color:    colors ? colors[1] : '#FF8E00'
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// PART D — ARTIFACT RENDERERS
// Returns elements[] array from an artifact object + its bounding rect
// ═══════════════════════════════════════════════════════════════════════════════

function renderArtifact(artifact, rect, colors, zoneId, artifactIdx, narrativeWeight) {
  const { x, y, w, h } = rect
  const id_prefix = `${zoneId}_a${artifactIdx}_`
  const t = (artifact.type || '').toLowerCase()

  // ── insight_text ──────────────────────────────────────────────────────────
  if (t === 'insight_text') {
    const elements = []
    const sentColor = artifact.sentiment === 'positive' ? colors.positive
                    : artifact.sentiment === 'warning'  ? colors.warning
                    : colors.primary

    // Background box
    elements.push(elShape(id_prefix+'bg', x, y, w, h, sentColor+'0D', sentColor, 0.5))

    // Left accent bar
    elements.push(elShape(id_prefix+'bar', x, y, 0.06, h, sentColor))

    // Heading
    const headingY = y + 0.1
    elements.push(elText(id_prefix+'hd', x+0.18, headingY, w-0.28, 0.35,
      artifact.heading || 'Key Insight',
      colors.titleFont, 11, true, sentColor, 'left', 'middle'))

    // Points
    const points = (artifact.points || []).slice(0, 4)
    const rowH   = Math.min(0.72, (h - 0.5) / Math.max(points.length, 1))

    points.forEach((pt, i) => {
      const py = headingY + 0.38 + i * rowH
      elements.push(elShape(id_prefix+'dot'+i, x+0.18, py+rowH*0.4, 0.07, 0.07, sentColor))
      elements.push(elText(id_prefix+'pt'+i, x+0.32, py, w-0.42, rowH,
        pt, colors.bodyFont, 11, false, colors.text, 'left', 'middle'))
    })

    return elements
  }

  // ── chart ─────────────────────────────────────────────────────────────────
  if (t === 'chart') {
    const series = (artifact.series || []).map(s => ({
      name:   s.name   || '',
      values: (s.values || []).map(v => typeof v === 'number' ? v : parseFloat(v) || 0),
      types:  s.types  || null
    }))

    return [elChart(
      id_prefix+'chart', x, y, w, h,
      artifact.chart_type    || 'bar',
      artifact.chart_title   || '',
      artifact.categories    || [],
      series,
      colors.chart,
      artifact.show_data_labels !== false,
      !!artifact.show_legend,
      artifact.x_label || '',
      artifact.y_label || ''
    )]
  }

  // ── cards ─────────────────────────────────────────────────────────────────
  if (t === 'cards') {
    const elements = []
    const cards    = (artifact.cards || []).slice(0, 4)
    const count    = cards.length
    if (!count) return elements

    const GAP   = 0.15
    const cardW = (w - GAP * (count + 1)) / count

    cards.forEach((card, i) => {
      const cx = x + GAP + i * (cardW + GAP)
      const sentColor = card.sentiment === 'positive' ? colors.positive
                      : card.sentiment === 'negative' ? colors.warning
                      : colors.primary

      elements.push(elShape(id_prefix+'card_bg'+i,  cx, y,      cardW, h,    '#F8FAFF', '#E0E4F0', 0.5))
      elements.push(elShape(id_prefix+'card_top'+i, cx, y,      cardW, 0.07, sentColor))

      // Title
      elements.push(elText(id_prefix+'card_ttl'+i, cx+0.12, y+0.12, cardW-0.24, 0.4,
        card.title || '', colors.titleFont, 12, true, sentColor, 'left', 'top'))

      // Subtitle
      if (card.subtitle) {
        elements.push(elText(id_prefix+'card_sub'+i, cx+0.12, y+0.54, cardW-0.24, 0.3,
          card.subtitle, colors.bodyFont, 10, false, '#888888', 'left', 'top'))
      }

      // Body
      const bodyY = card.subtitle ? y+0.86 : y+0.58
      elements.push(elText(id_prefix+'card_bdy'+i, cx+0.12, bodyY, cardW-0.24, h-bodyY+y-0.15,
        card.body || '', colors.bodyFont, 11, false, colors.text, 'left', 'top'))
    })

    return elements
  }

  // ── workflow ──────────────────────────────────────────────────────────────
  if (t === 'workflow') {
    return [elWorkflow(
      id_prefix+'wf', x, y, w, h,
      artifact.workflow_type  || 'process_flow',
      artifact.flow_direction || 'left_to_right',
      artifact.nodes       || [],
      artifact.connections || [],
      colors.chart,
      artifact.workflow_title || ''
    )]
  }

  // ── table ─────────────────────────────────────────────────────────────────
  if (t === 'table') {
    return [elTable(
      id_prefix+'table', x, y, w, h,
      artifact.headers     || [],
      artifact.rows        || [],
      colors.primary,
      '#FFFFFF',
      '#F5F5F5',
      colors.bodyFont,
      11,
      artifact.highlight_rows || []
    )]
  }

  return []
}

// ═══════════════════════════════════════════════════════════════════════════════
// PART E — ARTIFACT ZONE SPLITTER
// When a zone has 2 artifacts, split its bounding rect between them
// ═══════════════════════════════════════════════════════════════════════════════

function splitArtifactRects(zoneRect, artifact1, artifact2) {
  const { x, y, w, h } = zoneRect
  const GAP = 0.12
  const t1  = (artifact1.type || '').toLowerCase()
  const t2  = (artifact2.type || '').toLowerCase()

  // chart/workflow is primary — gets 60–65% of space
  // insight_text is secondary — gets 35–40%

  if (['chart','workflow'].includes(t1) && t2 === 'insight_text') {
    // Chart tall, insight below
    const split = 0.62
    return [
      { x, y,                   w, h: h*split - GAP/2 },
      { x, y: y+h*split+GAP/2,  w, h: h*(1-split) - GAP/2 }
    ]
  }

  if (t1 === 'insight_text' && ['chart','workflow'].includes(t2)) {
    // Insight short on top, chart below
    const split = 0.35
    return [
      { x, y,                   w, h: h*split - GAP/2 },
      { x, y: y+h*split+GAP/2,  w, h: h*(1-split) - GAP/2 }
    ]
  }

  if (['chart','workflow'].includes(t1) && t2 === 'table') {
    // Chart left, table right
    return [
      { x,             y, w: w*0.58 - GAP/2, h },
      { x: x+w*0.58+GAP/2, y, w: w*0.42-GAP/2, h }
    ]
  }

  if (t1 === 'table' && t2 === 'insight_text') {
    // Table takes more space, insight below
    const split = 0.65
    return [
      { x, y,                   w, h: h*split - GAP/2 },
      { x, y: y+h*split+GAP/2,  w, h: h*(1-split) - GAP/2 }
    ]
  }

  // Default — equal split horizontal
  return [
    { x,             y, w: w/2 - GAP/2, h },
    { x: x+w/2+GAP/2, y, w: w/2-GAP/2, h }
  ]
}

// ═══════════════════════════════════════════════════════════════════════════════
// PART F — FULL SLIDE BUILDER
// ═══════════════════════════════════════════════════════════════════════════════

function buildTitleSlide(slide, dim, colors) {
  const elements = []

  // Full background
  elements.push(elShape('bg', 0, 0, dim.w, dim.h, colors.primary))

  // Accent bar
  elements.push(elShape('accent', 0, dim.h*0.57, dim.w, 0.07, colors.secondary))

  // Title
  elements.push(elText('title', 0.7, dim.h*0.19, dim.w-1.4, dim.h*0.32,
    slide.title || '', colors.titleFont, 34, true, '#FFFFFF', 'left', 'middle'))

  // Subtitle
  if (slide.subtitle) {
    elements.push(elText('subtitle', 0.7, dim.h*0.52, dim.w-1.4, 0.55,
      slide.subtitle, colors.bodyFont, 15, false, '#DDDDDD', 'left', 'top'))
  }

  // Date
  elements.push(elText('date', dim.w-3.5, dim.h*0.62, 3.2, 0.32,
    new Date().toLocaleDateString('en-IN', { month: 'long', year: 'numeric' }),
    colors.bodyFont, 10, false, '#AAAAAA', 'right', 'middle'))

  return elements
}

function buildDividerSlide(slide, dim, colors) {
  const elements = []

  elements.push(elShape('bg',       0, 0, dim.w, dim.h, colors.primary))
  elements.push(elShape('left_bar', 0, 0, 0.12,  dim.h, colors.secondary))

  // Section label
  elements.push(elText('sec_lbl', 0.35, dim.h*0.31, dim.w-0.7, 0.35,
    'SECTION', colors.titleFont, 11, true, colors.secondary, 'left', 'middle'))

  // Section name
  elements.push(elText('sec_name', 0.35, dim.h*0.39, dim.w-0.7, dim.h*0.22,
    slide.title || '', colors.titleFont, 30, true, '#FFFFFF', 'left', 'top'))

  // Descriptor from key_message
  if (slide.key_message) {
    elements.push(elText('sec_desc', 0.35, dim.h*0.64, dim.w-0.7, 0.55,
      slide.key_message, colors.bodyFont, 13, false, '#CCCCCC', 'left', 'top'))
  }

  return elements
}

function buildContentSlide(slide, dim, colors, layoutMode, footerLayout) {
  const elements = []
  const zones    = slide.zones || []

  // ── Frame elements ─────────────────────────────────────────────────────────
  // Top accent bar
  elements.push(elShape('top_bar', 0, 0, dim.w, 0.08, colors.primary))

  // Title placement — depends on layout mode
  if (layoutMode.mode === 'sidebar' && layoutMode.has_coords) {
    const sb = layoutMode.sidebar
    // Sidebar background
    elements.push(elShape('sidebar_bg', sb.x, sb.y, sb.w, sb.h, colors.primary+'12'))
    // Vertical title in sidebar
    elements.push({
      id: 'sidebar_title', type: 'text_box_rotated',
      x: R(sb.x), y: R(sb.y), w: R(sb.w), h: R(sb.h),
      text: slide.title || '', font: colors.titleFont, size: 12, bold: true,
      color: colors.primary, align: 'center', valign: 'middle', rotation: 270
    })
  } else if (layoutMode.mode === 'top_title' && layoutMode.has_coords) {
    const tt = layoutMode.title
    elements.push(elText('title', tt.x, tt.y, tt.w, Math.min(tt.h, 0.75),
      slide.title || '', colors.titleFont, 20, true, colors.primary, 'left', 'middle'))
    elements.push(elShape('title_line', tt.x, tt.y + Math.min(tt.h, 0.75) + 0.03, tt.w, 0.025, '#E5E7EB'))
  } else {
    // Standard grid
    elements.push(elText('title', 0.4, 0.15, dim.w-0.8, 0.72,
      slide.title || '', colors.titleFont, 20, true, colors.primary, 'left', 'middle'))
    elements.push(elShape('title_line', 0.4, 0.92, dim.w-0.8, 0.025, '#E5E7EB'))
  }

  // Footer elements
  if (footerLayout.slide_num) {
    const sn = footerLayout.slide_num
    elements.push(elText('slide_num', sn.x, sn.y, sn.w, sn.h,
      String(slide.slide_number||''), colors.bodyFont, 9, false, '#AAAAAA', 'right', 'middle'))
  } else {
    elements.push(elText('slide_num', dim.w-0.7, dim.h-0.28, 0.55, 0.22,
      String(slide.slide_number||''), colors.bodyFont, 9, false, '#AAAAAA', 'right', 'middle'))
  }

  if (footerLayout.footer) {
    const f = footerLayout.footer
    elements.push(elText('footer', f.x, f.y, f.w, f.h,
      'Confidential', colors.bodyFont, 8, false, '#AAAAAA', 'left', 'middle'))
  }

  // ── Content area ───────────────────────────────────────────────────────────
  let contentArea
  if (layoutMode.mode === 'sidebar' && layoutMode.has_coords) {
    const ct = layoutMode.content
    contentArea = { x: ct.x, y: ct.y, w: ct.w, h: ct.h }
  } else if (layoutMode.mode === 'top_title' && layoutMode.has_coords) {
    const ct = layoutMode.content
    contentArea = { x: ct.x, y: ct.y, w: ct.w, h: ct.h }
  } else {
    // Standard grid fallback
    contentArea = { x: 0.4, y: 1.05, w: dim.w-0.8, h: dim.h-1.45 }
  }

  // ── Render each zone ───────────────────────────────────────────────────────
  zones.forEach(zone => {
    const split      = (zone.layout_hint || {}).split || 'full'
    const zoneR      = zoneRect(split, contentArea)
    const artifacts  = (zone.artifacts || []).slice(0, 2)

    if (!artifacts.length) return

    if (artifacts.length === 1) {
      // Single artifact fills the whole zone rect
      const artElements = renderArtifact(artifacts[0], zoneR, colors, zone.zone_id, 0, zone.narrative_weight)
      elements.push(...artElements)
    } else {
      // Two artifacts — split the zone rect between them
      const rects = splitArtifactRects(zoneR, artifacts[0], artifacts[1])
      const art0  = renderArtifact(artifacts[0], rects[0], colors, zone.zone_id, 0, zone.narrative_weight)
      const art1  = renderArtifact(artifacts[1], rects[1], colors, zone.zone_id, 1, zone.narrative_weight)
      elements.push(...art0, ...art1)
    }
  })

  return elements
}

// ═══════════════════════════════════════════════════════════════════════════════
// PART G — LAYOUT NAME SELECTOR
// ═══════════════════════════════════════════════════════════════════════════════

function selectLayoutName(slide, availableLayouts) {
  const st   = slide.slide_type
  const arch = slide.slide_archetype || ''
  const find = (keywords) => availableLayouts.find(l => keywords.some(k => (l.name||'').toLowerCase().includes(k)))

  if (st === 'title')   return (find(['title slide'])||{}).name  || 'Title slide'
  if (st === 'divider') return (find(['section','divider'])||{}).name || 'Section divider'

  // Match archetype to layout
  if (['recommendation','process','roadmap'].includes(arch)) return (find(['3 across','body text'])||{}).name || 'Body text KM'
  if (['dashboard','summary'].includes(arch))               return (find(['2 across','1 across'])||{}).name  || '1 across KM'
  if (['comparison','trend','proof'].includes(arch))        return (find(['1 across','body text'])||{}).name  || '1 across KM'

  return (find(['body text','1 across'])||{}).name || 'Body text KM'
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN RUNNER
// ═══════════════════════════════════════════════════════════════════════════════

async function runAgent5(state) {
  const manifest = state.slideManifest
  const brand    = state.brandRulebook
  const brief    = state.outline

  console.log('Agent 5 starting')
  console.log('  Slides:', manifest.length)

  const dim        = getDimensions(brand)
  const colors     = getBrandColors(brand)
  const layoutMode = detectLayoutMode(brand)
  const footerLyt  = getSlideFooterLayout(brand, brand.slide_layouts || [])

  console.log('  Slide size:', dim.w + '" × ' + dim.h + '"')
  console.log('  Layout mode:', layoutMode.mode, '| has_coords:', layoutMode.has_coords)
  console.log('  Primary color:', colors.primary)
  console.log('  Title font:', colors.titleFont)

  const designedSpec = manifest.map(slide => {
    let elements = []

    if (slide.slide_type === 'title') {
      elements = buildTitleSlide(slide, dim, colors)
    } else if (slide.slide_type === 'divider') {
      elements = buildDividerSlide(slide, dim, colors)
    } else {
      elements = buildContentSlide(slide, dim, colors, layoutMode, footerLyt)
    }

    const totalArtifacts = (slide.zones||[]).reduce((s,z)=>s+(z.artifacts||[]).length, 0)
    console.log('  S'+slide.slide_number, slide.slide_type, slide.slide_archetype||'—',
      '| zones:', (slide.zones||[]).length,
      '| artifacts:', totalArtifacts,
      '| elements:', elements.length)

    return {
      slide_number:      slide.slide_number,
      slide_type:        slide.slide_type,
      slide_archetype:   slide.slide_archetype   || '',
      section_name:      slide.section_name      || '',
      section_type:      slide.section_type      || '',
      layout_name:       selectLayoutName(slide, brand.slide_layouts || []),
      slide_width:       dim.w,
      slide_height:      dim.h,
      background_color:  (slide.slide_type === 'title' || slide.slide_type === 'divider') ? colors.primary : colors.bg,
      title:             slide.title             || '',
      key_message:       slide.key_message       || '',
      visual_flow_hint:  slide.visual_flow_hint  || '',
      zones_summary:     (slide.zones||[]).map(z => ({
        zone_id:          z.zone_id,
        zone_role:        z.zone_role,
        narrative_weight: z.narrative_weight,
        split:            (z.layout_hint||{}).split || 'full',
        artifact_types:   (z.artifacts||[]).map(a=>a.type)
      })),
      elements:          elements,
      speaker_note:      slide.speaker_note      || ''
    }
  })

  // Summary
  const totalElements = designedSpec.reduce((s,sl)=>s+sl.elements.length, 0)
  const elTypes = {}
  designedSpec.forEach(sl => sl.elements.forEach(el => { elTypes[el.type]=(elTypes[el.type]||0)+1 }))

  console.log('Agent 5 complete')
  console.log('  Total elements:', totalElements)
  console.log('  Element types:', JSON.stringify(elTypes))

  return designedSpec
}
