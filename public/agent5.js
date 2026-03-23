// ─── AGENT 5 — DESIGN DIRECTOR ────────────────────────────────────────────────
// Input:  state.slideManifest  — from Agent 4 (with primary + secondary content)
//         state.brandRulebook  — from Agent 2 (colors, fonts, layouts, placeholders)
//         state.outline        — from Agent 3 (brief, document type, tone)
//
// Output: finalSpec — fully positioned element array per slide
//
// Level 1  — Layout detection + zone allocation
// Level 2.1— Chart type finalisation
// Level 2.2— Bullet formatting with emphasis
// Partner  — Senior partner review loop (critique + fix pass)

// ═══════════════════════════════════════════════════════════════════════════════
// PART A — LAYOUT DETECTION
// ═══════════════════════════════════════════════════════════════════════════════

function detectLayoutType(brand) {
  const layouts = brand.slide_layouts || []

  // Try to find a content layout with sidebar characteristics
  const contentLayout = layouts.find(l =>
    l.name && (l.name.toLowerCase().includes('1 across') || l.name.toLowerCase().includes('body text'))
  )

  if (contentLayout && contentLayout.placeholders && contentLayout.placeholders.length > 0) {
    // Check for sidebar — narrow placeholder on left (x < 1.5, w < 2.5)
    const sidebar = contentLayout.placeholders.find(p =>
      p.x_in !== undefined && p.x_in < 1.5 && p.w_in < 2.5 && p.h_in > 3
    )

    // Check for title placeholder
    const titlePH = contentLayout.placeholders.find(p =>
      p.type === 'title' || p.type === 'body' && p.w_in > 5 && p.y_in < 2
    )

    // Check for main content area
    const contentPH = contentLayout.placeholders.find(p =>
      p.x_in > 2 && p.w_in > 5 && p.h_in > 3
    )

    if (sidebar) {
      return {
        type:       'sidebar',
        sidebar:    { x: sidebar.x_in, y: sidebar.y_in, w: sidebar.w_in, h: sidebar.h_in },
        content:    contentPH ? { x: contentPH.x_in, y: contentPH.y_in, w: contentPH.w_in, h: contentPH.h_in } : null,
        has_coords: true
      }
    }

    if (titlePH) {
      return {
        type:       'top_title',
        title:      { x: titlePH.x_in, y: titlePH.y_in, w: titlePH.w_in, h: titlePH.h_in },
        content:    contentPH ? { x: contentPH.x_in, y: contentPH.y_in, w: contentPH.w_in, h: contentPH.h_in } : null,
        has_coords: true
      }
    }
  }

  // Fallback — standard grid
  return { type: 'standard_grid', has_coords: false }
}

// ═══════════════════════════════════════════════════════════════════════════════
// PART B — ZONE CALCULATOR
// ═══════════════════════════════════════════════════════════════════════════════

function calculateZones(slide, layoutInfo, dim) {
  const sw = dim.w
  const sh = dim.h

  if (slide.slide_type === 'title') {
    return {
      type: 'title',
      background: { x: 0, y: 0, w: sw, h: sh },
      title_zone: { x: 0.7, y: sh * 0.2, w: sw - 1.4, h: sh * 0.32 },
      subtitle_zone: { x: 0.7, y: sh * 0.53, w: sw - 1.4, h: 0.6 },
      accent_bar: { x: 0, y: sh * 0.58, w: sw, h: 0.07 },
      footer_zone: { x: sw - 3.5, y: sh * 0.63, w: 3.2, h: 0.3 }
    }
  }

  if (slide.slide_type === 'divider') {
    return {
      type: 'divider',
      background: { x: 0, y: 0, w: sw, h: sh },
      left_bar: { x: 0, y: 0, w: 0.12, h: sh },
      label_zone: { x: 0.4, y: sh * 0.32, w: sw - 0.8, h: 0.4 },
      title_zone: { x: 0.4, y: sh * 0.42, w: sw - 0.8, h: 1.6 },
      descriptor_zone: { x: 0.4, y: sh * 0.67, w: sw - 0.8, h: 0.6 }
    }
  }

  // Content slide — depends on layout type
  if (layoutInfo.type === 'sidebar' && layoutInfo.has_coords) {
    const sb = layoutInfo.sidebar
    const ct = layoutInfo.content || { x: sb.x + sb.w + 0.1, y: sb.y, w: sw - sb.x - sb.w - 0.5, h: sb.h }

    return {
      type: 'sidebar',
      top_bar: { x: 0, y: 0, w: sw, h: 0.08 },
      sidebar_zone: { x: sb.x, y: sb.y, w: sb.w, h: sb.h },
      content_zone: { x: ct.x, y: ct.y, w: ct.w, h: ct.h },
      footer_zone:  { x: 0.1, y: sh - 0.35, w: sw - 0.2, h: 0.25 },
      slide_num:    { x: sw - 0.7, y: sh - 0.3, w: 0.5, h: 0.22 }
    }
  }

  if (layoutInfo.type === 'top_title' && layoutInfo.has_coords) {
    const tt = layoutInfo.title
    const ct = layoutInfo.content || { x: 0.4, y: tt.y + tt.h + 0.1, w: sw - 0.8, h: sh - tt.y - tt.h - 0.5 }

    return {
      type: 'top_title',
      top_bar: { x: 0, y: 0, w: sw, h: 0.08 },
      title_zone: { x: tt.x, y: tt.y, w: tt.w, h: Math.min(tt.h, 0.8) },
      underline:  { x: tt.x, y: tt.y + Math.min(tt.h, 0.8) + 0.02, w: tt.w, h: 0.02 },
      content_zone: { x: ct.x, y: ct.y, w: ct.w, h: ct.h },
      slide_num: { x: sw - 0.7, y: sh - 0.3, w: 0.5, h: 0.22 }
    }
  }

  // Standard grid fallback
  return {
    type: 'standard_grid',
    top_bar:      { x: 0,    y: 0,    w: sw,      h: 0.08 },
    title_zone:   { x: 0.4,  y: 0.15, w: sw - 0.8, h: 0.75 },
    underline:    { x: 0.4,  y: 0.95, w: sw - 0.8, h: 0.02 },
    content_zone: { x: 0.4,  y: 1.05, w: sw - 0.8, h: sh - 1.5 },
    slide_num:    { x: sw - 0.7, y: sh - 0.3, w: 0.5, h: 0.22 }
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// PART C — CONTENT ZONE SPLITTER (for mixed slides)
// ═══════════════════════════════════════════════════════════════════════════════

function splitContentZone(contentZone, primaryType, secondaryType) {
  const { x, y, w, h } = contentZone

  // Full zone — no split needed
  if (!secondaryType) {
    return { primary: contentZone, secondary: null }
  }

  // Chart + insight/bullets → 65% top / 35% bottom
  if (['chart'].includes(primaryType) && ['bullets', 'insight_box'].includes(secondaryType)) {
    const primaryH = h * 0.63
    const gap      = 0.12
    return {
      primary:   { x, y,                  w, h: primaryH },
      secondary: { x, y: y + primaryH + gap, w, h: h - primaryH - gap }
    }
  }

  // Stat callout + bullets → 40% left / 60% right
  if (primaryType === 'stat_callout' && ['bullets', 'insight_box'].includes(secondaryType)) {
    const primaryW = w * 0.38
    const gap      = 0.2
    return {
      primary:   { x,                  y, w: primaryW, h },
      secondary: { x: x + primaryW + gap, y, w: w - primaryW - gap, h }
    }
  }

  // Stat callout + chart → 40% top / 60% bottom
  if (primaryType === 'stat_callout' && secondaryType === 'chart') {
    const primaryH = h * 0.38
    const gap      = 0.15
    return {
      primary:   { x, y,                  w, h: primaryH },
      secondary: { x, y: y + primaryH + gap, w, h: h - primaryH - gap }
    }
  }

  // Cards + bullets → 65% top / 35% bottom
  if (primaryType === 'cards' && secondaryType === 'bullets') {
    const primaryH = h * 0.62
    const gap      = 0.12
    return {
      primary:   { x, y,                  w, h: primaryH },
      secondary: { x, y: y + primaryH + gap, w, h: h - primaryH - gap }
    }
  }

  // Default split — 65/35 vertical
  const primaryH = h * 0.65
  return {
    primary:   { x, y, w, h: primaryH },
    secondary: { x, y: y + primaryH + 0.12, w, h: h - primaryH - 0.12 }
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// PART D — ELEMENT BUILDERS
// ═══════════════════════════════════════════════════════════════════════════════

function getBrandColors(brand) {
  return {
    primary:   (brand.primary_colors    || ['#0F2FB5'])[0],
    secondary: (brand.secondary_colors  || ['#FF8E00'])[0],
    bg:        (brand.background_colors || ['#FFFFFF'])[0],
    text:      (brand.text_colors       || ['#000000'])[0],
    positive:  (brand.accent_colors     || [])[0] || '#2D962D',
    warning:   (brand.all_colors        || {}).accent4 || '#D60202',
    chart:     brand.chart_colors       || ['#0F2FB5','#FF8E00','#2D962D','#D60202','#2E046C','#0092D0'],
    titleFont: (brand.title_font        || {}).family || 'Arial',
    bodyFont:  (brand.body_font         || {}).family || 'Arial'
  }
}

function el_shape(id, x, y, w, h, fill, border_color, border_pt) {
  return { id, type: 'shape', x, y, w, h, fill_color: fill || '#CCCCCC', border_color: border_color || null, border_pt: border_pt || 0 }
}

function el_text(id, x, y, w, h, text, font, size, bold, color, align, valign, italic) {
  return { id, type: 'text_box', x, y, w, h, text: String(text || ''), font: font || 'Arial', size: size || 14, bold: !!bold, italic: !!italic || false, color: color || '#000000', align: align || 'left', valign: valign || 'top', wrap: true }
}

function el_chart(id, x, y, w, h, chart_type, chart_title, categories, series, colors, show_labels, show_legend, x_label, y_label) {
  return { id, type: 'chart', chart_type: chart_type || 'bar', x, y, w, h, chart_title: chart_title || '', x_label: x_label || '', y_label: y_label || '', categories: categories || [], series: series || [], colors: colors || [], show_data_labels: show_labels !== false, show_legend: !!show_legend }
}

function el_table(id, x, y, w, h, headers, rows, header_bg, header_text, row_bg_alt, font, font_size) {
  return { id, type: 'table', x, y, w, h, headers: headers || [], rows: rows || [], header_bg: header_bg || '#0F2FB5', header_text: header_text || '#FFFFFF', row_bg_alt: row_bg_alt || '#F5F5F5', font: font || 'Arial', font_size: font_size || 11, header_font_size: 12 }
}

function el_rich_bullets(id, x, y, w, h, bullets, font, size, text_color, accent_color, positive_color, warning_color) {
  return {
    id, type: 'rich_bullets', x, y, w, h,
    font:           font       || 'Arial',
    size:           size       || 13,
    text_color:     text_color || '#000000',
    accent_color:   accent_color  || '#FF8E00',
    positive_color: positive_color || '#2D962D',
    warning_color:  warning_color  || '#D60202',
    bullets:        (bullets || []).map(b => normaliseBullet(b))
  }
}

function normaliseBullet(b) {
  if (typeof b === 'string') {
    return { text: b, emphasis: autoEmphasis(b), sentiment: autoSentiment(b) }
  }
  return {
    text:      b.text || '',
    emphasis:  b.emphasis || autoEmphasis(b.text || ''),
    sentiment: b.sentiment || autoSentiment(b.text || '')
  }
}

function autoEmphasis(text) {
  const emphasis = []
  // Numbers and percentages
  const numMatches = text.match(/[\₹\$\€]?[\d,]+\.?\d*[%CrLKM]*/g) || []
  if (numMatches.length > 0) {
    emphasis.push({ text: numMatches[0], style: 'bold', color: null })
  }
  return emphasis
}

function autoSentiment(text) {
  const t = text.toLowerCase()
  if (/risk|decline|fall|breach|exceed|critical|concern|dangerous|concentrat/.test(t)) return 'warning'
  if (/grew|growth|strong|record|exceed.*target|improv|positive/.test(t)) return 'positive'
  return 'neutral'
}

// ═══════════════════════════════════════════════════════════════════════════════
// PART E — CONTENT RENDERERS (builds elements from content objects)
// ═══════════════════════════════════════════════════════════════════════════════

function renderPrimaryContent(pc, zone, colors, idPrefix) {
  if (!pc) return []
  const type = (pc.type || '').toLowerCase()

  switch(type) {
    case 'chart':         return renderChart(pc, zone, colors, idPrefix)
    case 'stat_callout':  return renderStatCallout(pc, zone, colors, idPrefix)
    case 'data_table':    return renderDataTable(pc, zone, colors, idPrefix)
    case 'three_column':  return renderThreeColumn(pc, zone, colors, idPrefix)
    case 'two_column':    return renderTwoColumn(pc, zone, colors, idPrefix)
    case 'cards':         return renderCards(pc, zone, colors, idPrefix)
    case 'process_flow':  return renderProcessFlow(pc, zone, colors, idPrefix)
    case 'bullets':       return renderBullets(pc, zone, colors, idPrefix)
    default:              return renderBullets(pc, zone, colors, idPrefix)
  }
}

function renderSecondaryContent(sc, zone, colors, idPrefix) {
  if (!sc) return []
  const type = (sc.type || '').toLowerCase()

  switch(type) {
    case 'bullets':      return renderBullets(sc, zone, colors, idPrefix)
    case 'insight_box':  return renderInsightBox(sc, zone, colors, idPrefix)
    case 'stat_callout': return renderStatCallout(sc, zone, colors, idPrefix)
    default:             return renderBullets(sc, zone, colors, idPrefix)
  }
}

function renderChart(pc, zone, colors, prefix) {
  const { x, y, w, h } = zone
  const series = (pc.series || []).map(s => ({
    name:   s.name   || '',
    values: (s.values || []).map(v => typeof v === 'number' ? v : parseFloat(v) || 0),
    types:  s.types  || null
  }))

  return [el_chart(
    prefix + 'chart', x, y, w, h,
    pc.chart_type || 'bar',
    pc.chart_title || '',
    pc.categories  || [],
    series,
    colors.chart,
    true,
    series.length > 1,
    pc.x_label || '',
    pc.y_label || ''
  )]
}

function renderStatCallout(pc, zone, colors, prefix) {
  const { x, y, w, h } = zone
  const stats   = (pc.stats || []).slice(0, 4)
  const count   = stats.length
  const cols    = count <= 2 ? count : (count === 3 ? 3 : 2)
  const rows    = Math.ceil(count / cols)
  const padX    = 0.15
  const padY    = 0.15
  const cellW   = (w - padX * (cols + 1)) / cols
  const cellH   = (h - padY * (rows + 1)) / rows
  const elements= []

  stats.forEach((stat, i) => {
    const col  = i % cols
    const row  = Math.floor(i / cols)
    const cx   = x + padX + col * (cellW + padX)
    const cy   = y + padY + row * (cellH + padY)

    const sentimentColor = stat.sentiment === 'positive' ? colors.positive
                         : stat.sentiment === 'negative' ? colors.warning
                         : colors.primary

    elements.push(el_shape(prefix + 'sc_bg_' + i,     cx, cy, cellW, cellH, '#F5F7FF', '#E0E4F0', 0.5))
    elements.push(el_shape(prefix + 'sc_top_' + i,    cx, cy, cellW, 0.06,  sentimentColor))
    elements.push(el_text( prefix + 'sc_val_' + i,    cx + 0.18, cy + 0.15, cellW - 0.36, cellH * 0.42, stat.value || '—', colors.titleFont, 26, true,  sentimentColor, 'left', 'middle'))
    elements.push(el_text( prefix + 'sc_lbl_' + i,    cx + 0.18, cy + cellH * 0.55, cellW - 0.36, cellH * 0.22, stat.label  || '', colors.bodyFont,  11, false, '#555555', 'left', 'top'))
    if (stat.change) {
      elements.push(el_text(prefix + 'sc_chg_' + i,   cx + 0.18, cy + cellH * 0.76, cellW - 0.36, cellH * 0.18, stat.change, colors.bodyFont, 10, false, colors.secondary, 'left', 'top'))
    }
  })

  return elements
}

function renderBullets(pc, zone, colors, prefix) {
  const { x, y, w, h } = zone
  const bullets = (pc.bullets || [])
  if (!bullets.length) return []

  return [el_rich_bullets(
    prefix + 'bullets', x, y, w, h,
    bullets,
    colors.bodyFont, 13,
    colors.text, colors.secondary, colors.positive, colors.warning
  )]
}

function renderInsightBox(sc, zone, colors, prefix) {
  const { x, y, w, h } = zone
  const elements = []

  const sentColor = sc.sentiment === 'positive' ? colors.positive
                  : sc.sentiment === 'warning'  ? colors.warning
                  : colors.primary

  elements.push(el_shape(prefix + 'ib_bg',  x, y, w, h, sentColor + '15', sentColor, 0.8))
  elements.push(el_shape(prefix + 'ib_bar', x, y, 0.06, h, sentColor))

  if (sc.heading) {
    elements.push(el_text(prefix + 'ib_head', x + 0.2, y + 0.1, w - 0.3, 0.3, sc.heading, colors.titleFont, 11, true, sentColor, 'left', 'top'))
  }
  elements.push(el_text(prefix + 'ib_text', x + 0.2, y + (sc.heading ? 0.42 : 0.15), w - 0.3, h - 0.25, sc.text || '', colors.bodyFont, 12, false, colors.text, 'left', 'top'))

  return elements
}

function renderDataTable(pc, zone, colors, prefix) {
  const { x, y, w, h } = zone
  return [el_table(prefix + 'table', x, y, w, h, pc.headers || [], pc.rows || [], colors.primary, '#FFFFFF', '#F5F5F5', colors.bodyFont, 11)]
}

function renderThreeColumn(pc, zone, colors, prefix) {
  const { x, y, w, h } = zone
  const columns  = (pc.columns || []).slice(0, 3)
  const padX     = 0.15
  const colW     = (w - padX * 4) / 3
  const elements = []

  columns.forEach((col, i) => {
    const cx = x + padX + i * (colW + padX)
    const sentimentColor = col.sentiment === 'positive' ? colors.positive
                         : col.sentiment === 'warning'  ? colors.warning
                         : colors.primary

    elements.push(el_shape(prefix + 'col_bg_'  + i, cx, y,        colW, h,    '#F9FAFB', '#E5E7EB', 0.5))
    elements.push(el_shape(prefix + 'col_top_' + i, cx, y,        colW, 0.07, sentimentColor))
    elements.push(el_text( prefix + 'col_num_' + i, cx+0.15, y+0.12, 0.5, 0.45, String(i+1).padStart(2,'0'), colors.titleFont, 18, true, sentimentColor, 'left', 'top'))
    elements.push(el_text( prefix + 'col_hdr_' + i, cx+0.15, y+0.62, colW-0.3, 0.45, col.header || '', colors.titleFont, 12, true, sentimentColor, 'left', 'top'))
    elements.push(el_text( prefix + 'col_bdy_' + i, cx+0.15, y+1.12, colW-0.3, h-1.3, col.body || '', colors.bodyFont, 11, false, colors.text, 'left', 'top'))
  })

  return elements
}

function renderTwoColumn(pc, zone, colors, prefix) {
  const { x, y, w, h } = zone
  const colW     = (w - 0.25) / 2
  const elements = []

  // Left
  elements.push(el_text(prefix + 'lh', x, y, colW, 0.38, pc.left_header || '', colors.titleFont, 13, true, colors.primary, 'left', 'top'))
  elements.push(el_shape(prefix + 'ld', x, y + 0.42, colW, 0.03, colors.primary))
  ;(pc.left_points || []).forEach((pt, i) => {
    elements.push(el_shape(prefix + 'ldot_' + i, x, y + 0.58 + i * 0.72, 0.09, 0.09, colors.secondary))
    elements.push(el_text( prefix + 'lpt_'  + i, x + 0.2, y + 0.54 + i * 0.72, colW - 0.2, 0.68, pt, colors.bodyFont, 12, false, colors.text, 'left', 'middle'))
  })

  // Divider
  elements.push(el_shape(prefix + 'div', x + colW + 0.08, y, 0.02, h, '#E5E7EB'))

  // Right
  const rx = x + colW + 0.25
  elements.push(el_text(prefix + 'rh', rx, y, colW, 0.38, pc.right_header || '', colors.titleFont, 13, true, colors.primary, 'left', 'top'))
  elements.push(el_shape(prefix + 'rd', rx, y + 0.42, colW, 0.03, colors.primary))
  ;(pc.right_points || []).forEach((pt, i) => {
    elements.push(el_shape(prefix + 'rdot_' + i, rx, y + 0.58 + i * 0.72, 0.09, 0.09, colors.secondary))
    elements.push(el_text( prefix + 'rpt_'  + i, rx + 0.2, y + 0.54 + i * 0.72, colW - 0.2, 0.68, pt, colors.bodyFont, 12, false, colors.text, 'left', 'middle'))
  })

  return elements
}

function renderCards(pc, zone, colors, prefix) {
  const { x, y, w, h } = zone
  const cards    = (pc.cards || []).slice(0, 4)
  const count    = cards.length
  const padX     = 0.15
  const cardW    = (w - padX * (count + 1)) / count
  const elements = []

  cards.forEach((card, i) => {
    const cx = x + padX + i * (cardW + padX)
    const sentimentColor = card.sentiment === 'positive' ? colors.positive
                         : card.sentiment === 'warning'  ? colors.warning
                         : colors.primary

    elements.push(el_shape(prefix + 'card_bg_'  + i, cx, y, cardW, h, '#F9FAFB', '#E5E7EB', 0.5))
    elements.push(el_shape(prefix + 'card_top_' + i, cx, y, cardW, 0.07, sentimentColor))
    elements.push(el_text( prefix + 'card_num_' + i, cx + cardW/2 - 0.25, y + 0.12, 0.5, 0.45, String(i+1), colors.titleFont, 16, true, sentimentColor, 'center', 'middle'))
    elements.push(el_text( prefix + 'card_hdr_' + i, cx + 0.12, y + 0.65, cardW - 0.24, 0.45, card.header || '', colors.titleFont, 12, true, sentimentColor, 'left', 'top'))
    elements.push(el_text( prefix + 'card_bdy_' + i, cx + 0.12, y + 1.15, cardW - 0.24, h - 1.35, card.body || '', colors.bodyFont, 11, false, colors.text, 'left', 'top'))
  })

  return elements
}

function renderProcessFlow(pc, zone, colors, prefix) {
  const { x, y, w, h } = zone
  const steps    = (pc.steps || []).slice(0, 5)
  const count    = steps.length
  const padX     = 0.12
  const stepW    = (w - padX * (count + 1)) / count
  const elements = []

  steps.forEach((step, i) => {
    const sx       = x + padX + i * (stepW + padX)
    const isFirst  = i === 0
    const fillColor= isFirst ? colors.primary : colors.primary + '18'
    const textColor= isFirst ? '#FFFFFF' : colors.primary

    elements.push(el_shape(prefix + 'step_bg_' + i,  sx, y + 0.5, stepW, h - 0.5, fillColor, colors.primary, 0.5))
    elements.push(el_shape(prefix + 'step_num_bg_' + i, sx + stepW/2 - 0.28, y, 0.56, 0.56, colors.primary))
    elements.push(el_text( prefix + 'step_num_' + i, sx + stepW/2 - 0.28, y, 0.56, 0.56, String(step.step_number || i+1), colors.titleFont, 14, true, '#FFFFFF', 'center', 'middle'))

    if (i < count - 1) {
      elements.push(el_text(prefix + 'arrow_' + i, sx + stepW + padX/2 - 0.12, y + h/2 - 0.15, 0.24, 0.3, '▶', colors.bodyFont, 11, false, colors.secondary, 'center', 'middle'))
    }

    elements.push(el_text(prefix + 'step_ttl_' + i, sx + 0.1, y + 0.62, stepW - 0.2, 0.48, step.title || '', colors.titleFont, 11, true, textColor, 'center', 'top'))
    elements.push(el_text(prefix + 'step_dsc_' + i, sx + 0.1, y + 1.15, stepW - 0.2, h - 1.35, step.description || '', colors.bodyFont, 10, false, isFirst ? '#EEEEEE' : colors.text, 'left', 'top'))
  })

  return elements
}

// ═══════════════════════════════════════════════════════════════════════════════
// PART F — FULL SLIDE BUILDER
// ═══════════════════════════════════════════════════════════════════════════════

function buildSlideSpec(slide, brand, layoutInfo, dim) {
  const colors  = getBrandColors(brand)
  const zones   = calculateZones(slide, layoutInfo, dim)
  const elements= []

  // ── Title slide ─────────────────────────────────────────────────────────────
  if (slide.slide_type === 'title') {
    elements.push(el_shape('bg', 0, 0, dim.w, dim.h, colors.primary))
    elements.push(el_shape('accent_bar', 0, zones.accent_bar.y, dim.w, zones.accent_bar.h, colors.secondary))

    const pc = slide.primary_content || {}
    elements.push(el_text('main_title', zones.title_zone.x, zones.title_zone.y, zones.title_zone.w, zones.title_zone.h, pc.title || slide.title || '', colors.titleFont, 34, true, '#FFFFFF', 'left', 'middle'))
    if (pc.subtitle || slide.subtitle) {
      elements.push(el_text('subtitle', zones.subtitle_zone.x, zones.subtitle_zone.y, zones.subtitle_zone.w, zones.subtitle_zone.h, pc.subtitle || slide.subtitle || '', colors.bodyFont, 16, false, '#DDDDDD', 'left', 'top'))
    }
    elements.push(el_text('date', zones.footer_zone.x, zones.footer_zone.y, zones.footer_zone.w, zones.footer_zone.h, pc.date || new Date().toLocaleDateString('en-IN', { month: 'long', year: 'numeric' }), colors.bodyFont, 10, false, '#AAAAAA', 'right', 'middle'))
    return elements
  }

  // ── Divider slide ───────────────────────────────────────────────────────────
  if (slide.slide_type === 'divider') {
    const pc = slide.primary_content || {}
    elements.push(el_shape('bg',       0, 0, dim.w, dim.h, colors.primary))
    elements.push(el_shape('left_bar', 0, 0, zones.left_bar.w, dim.h, colors.secondary))
    elements.push(el_text('sec_label', zones.label_zone.x, zones.label_zone.y, zones.label_zone.w, zones.label_zone.h, 'SECTION', colors.titleFont, 12, true, colors.secondary, 'left', 'middle'))
    elements.push(el_text('sec_name',  zones.title_zone.x, zones.title_zone.y, zones.title_zone.w, zones.title_zone.h, pc.section_name || slide.title || '', colors.titleFont, 30, true, '#FFFFFF', 'left', 'top'))
    if (pc.section_descriptor) {
      elements.push(el_text('sec_desc', zones.descriptor_zone.x, zones.descriptor_zone.y, zones.descriptor_zone.w, zones.descriptor_zone.h, pc.section_descriptor, colors.bodyFont, 14, false, '#CCCCCC', 'left', 'top'))
    }
    return elements
  }

  // ── Content slide ───────────────────────────────────────────────────────────
  const zType = zones.type

  // Always add top bar
  if (zones.top_bar) {
    elements.push(el_shape('top_bar', 0, 0, dim.w, zones.top_bar.h, colors.primary))
  }

  // Sidebar layout — title in sidebar
  if (zType === 'sidebar') {
    // Sidebar background
    elements.push(el_shape('sidebar_bg', zones.sidebar_zone.x, zones.sidebar_zone.y, zones.sidebar_zone.w, zones.sidebar_zone.h, colors.primary + '12'))
    // Vertical title in sidebar
    elements.push({
      id: 'sidebar_title', type: 'text_box_rotated',
      x: zones.sidebar_zone.x, y: zones.sidebar_zone.y,
      w: zones.sidebar_zone.w, h: zones.sidebar_zone.h,
      text: slide.title || '', font: colors.titleFont, size: 13, bold: true,
      color: colors.primary, align: 'center', valign: 'middle', rotation: 270
    })
    // Footer elements
    elements.push(el_text('footer', 0.1, dim.h - 0.35, dim.w * 0.55, 0.22, 'Confidential', colors.bodyFont, 8, false, '#AAAAAA', 'left', 'middle'))
    elements.push(el_text('slide_num', zones.slide_num.x, zones.slide_num.y, zones.slide_num.w, zones.slide_num.h, String(slide.slide_number || ''), colors.bodyFont, 9, false, '#AAAAAA', 'right', 'middle'))

    // Content zone — split if mixed
    const contentZone = zones.content_zone
    const pc = slide.primary_content
    const sc = slide.secondary_content

    if (slide.is_mixed && sc) {
      const split = splitContentZone(contentZone, pc ? pc.type : 'bullets', sc.type)
      elements.push(...renderPrimaryContent(pc, split.primary, colors, 'p_'))
      elements.push(...renderSecondaryContent(sc, split.secondary, colors, 's_'))
    } else {
      elements.push(...renderPrimaryContent(pc, contentZone, colors, 'p_'))
    }

  } else {
    // Top-title or standard grid layout
    // Title
    if (zones.title_zone) {
      elements.push(el_text('title', zones.title_zone.x, zones.title_zone.y, zones.title_zone.w, zones.title_zone.h, slide.title || '', colors.titleFont, 20, true, colors.primary, 'left', 'middle'))
    }
    if (zones.underline) {
      elements.push(el_shape('underline', zones.underline.x, zones.underline.y, zones.underline.w, zones.underline.h, '#E5E7EB'))
    }
    elements.push(el_text('slide_num', zones.slide_num.x, zones.slide_num.y, zones.slide_num.w, zones.slide_num.h, String(slide.slide_number || ''), colors.bodyFont, 9, false, '#AAAAAA', 'right', 'middle'))

    // Content
    const contentZone = zones.content_zone
    const pc = slide.primary_content
    const sc = slide.secondary_content

    if (slide.is_mixed && sc) {
      const split = splitContentZone(contentZone, pc ? pc.type : 'bullets', sc.type)
      elements.push(...renderPrimaryContent(pc, split.primary, colors, 'p_'))
      elements.push(...renderSecondaryContent(sc, split.secondary, colors, 's_'))
    } else {
      elements.push(...renderPrimaryContent(pc, contentZone, colors, 'p_'))
    }
  }

  return elements
}

// ═══════════════════════════════════════════════════════════════════════════════
// PART G — PARTNER REVIEW LOOP
// ═══════════════════════════════════════════════════════════════════════════════

const PARTNER_REVIEW_SYSTEM = `You are a senior partner at a top management consulting firm.
You are reviewing a presentation spec built by your analyst team.

Review the spec against these five criteria and return a structured critique:

1. NARRATIVE — Do key_messages tell a coherent story? Is each one insight-led not descriptive?
2. TITLES — Is every content slide title specific and insight-driven? Not "Revenue Analysis" but "Revenue grew 18%"
3. VISUALS — Are numbers always visualised? Is there visual variety? No 5+ bullet slides in a row?
4. MIXED CONTENT — Where data and narrative coexist, is is_mixed: true being used?
5. CLUTTER — Does any slide try to do too much? Is the key message immediately obvious?

Return ONLY a valid JSON object:
{
  "overall_rating": "approved" | "needs_revision",
  "approved": true | false,
  "issues": [
    {
      "slide_number": number,
      "criterion": "narrative" | "title" | "visual" | "mixed" | "clutter",
      "issue": "string — what is wrong",
      "fix": "string — exact correction to make"
    }
  ],
  "narrative_gaps": ["string — any flow issues between sections"],
  "summary": "string — 2 sentence overall assessment"
}`

async function partnerReview(slideManifest, brief) {
  console.log('Agent 5 — Partner review starting...')

  // Send a condensed version (no full content, just structure and messages)
  const condensed = slideManifest.map(s => ({
    slide_number: s.slide_number,
    slide_type:   s.slide_type,
    section_type: s.section_type,
    title:        s.title,
    key_message:  s.key_message,
    is_mixed:     s.is_mixed,
    primary_type: s.primary_content ? s.primary_content.type : null,
    secondary_type: s.secondary_content ? s.secondary_content.type : null
  }))

  const messages = [{
    role: 'user',
    content: `PRESENTATION TYPE: ${(brief || {}).document_type || '—'}
GOVERNING THOUGHT: ${(brief || {}).governing_thought || '—'}

SLIDE MANIFEST TO REVIEW:
${JSON.stringify(condensed, null, 2)}

Review this presentation spec as a senior partner. Identify issues and return structured critique JSON.`
  }]

  try {
    const raw      = await callClaude(PARTNER_REVIEW_SYSTEM, messages, 1500)
    const critique = safeParseJSON(raw, null)

    if (!critique) {
      console.warn('Agent 5 — Partner review parse failed')
      return null
    }

    console.log('Agent 5 — Partner review:', critique.overall_rating, '—', (critique.issues || []).length, 'issues found')
    return critique

  } catch(e) {
    console.warn('Agent 5 — Partner review failed:', e.message)
    return null
  }
}

function applyPartnerFixes(slideManifest, critique) {
  if (!critique || !critique.issues || critique.issues.length === 0) return slideManifest

  const fixed = slideManifest.map(slide => {
    const issues = critique.issues.filter(i => i.slide_number === slide.slide_number)
    if (!issues.length) return slide

    let updated = { ...slide }

    issues.forEach(issue => {
      console.log('Agent 5 — Applying fix for slide', slide.slide_number, ':', issue.criterion, '-', issue.issue)

      if (issue.criterion === 'title' && issue.fix) {
        // Extract new title from fix instruction
        const titleMatch = issue.fix.match(/["']([^"']{10,80})["']/) || issue.fix.match(/Change to:?\s*(.+)/)
        if (titleMatch) updated.title = titleMatch[1].trim()
      }

      if (issue.criterion === 'mixed' && issue.fix && issue.fix.toLowerCase().includes('mixed')) {
        updated.is_mixed = true
        // Build a basic insight box as secondary if none exists
        if (!updated.secondary_content && updated.key_message) {
          updated.secondary_content = {
            type: 'insight_box',
            heading: 'Key Insight',
            text: updated.key_message,
            sentiment: 'neutral'
          }
        }
      }

      if (issue.criterion === 'visual' && issue.fix && issue.fix.toLowerCase().includes('stat')) {
        const pc = updated.primary_content
        if (pc && pc.type === 'bullets') {
          updated.primary_content = {
            type: 'stat_callout',
            stats: pc.bullets ? pc.bullets.slice(0, 3).map(b => ({
              value:  (b.text || b).match(/[\₹\d,]+\.?\d*[%CrLKM]*/)?.[0] || '—',
              label:  (b.text || b).replace(/[\₹\d,]+\.?\d*[%CrLKM]*/g, '').trim().slice(0, 40),
              change: '',
              sentiment: 'neutral'
            })) : []
          }
        }
      }
    })

    return updated
  })

  return fixed
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
  console.log('  Slide size:', (brand.slide_width_inches || 11.02) + '" x ' + (brand.slide_height_inches || 8.29) + '"')

  const dim        = { w: brand.slide_width_inches || 11.02, h: brand.slide_height_inches || 8.29 }
  const layoutInfo = detectLayoutType(brand)
  const colors     = getBrandColors(brand)

  console.log('  Layout type detected:', layoutInfo.type, '(has_coords:', layoutInfo.has_coords + ')')
  console.log('  Primary color:', colors.primary)
  console.log('  Title font:', colors.titleFont)

  // ── Partner Review Pass 1 ──────────────────────────────────────────────────
  console.log('Agent 5 — Round 1: Analyst build...')
  let reviewedManifest = manifest

  const critique = await partnerReview(manifest, brief)
  if (critique && !critique.approved && (critique.issues || []).length > 0) {
    console.log('Agent 5 — Partner requested', critique.issues.length, 'fixes')
    reviewedManifest = applyPartnerFixes(manifest, critique)
  } else {
    console.log('Agent 5 — Partner approved or review unavailable')
  }

  // ── Build element-level spec ───────────────────────────────────────────────
  console.log('Agent 5 — Building element specs...')

  const finalSpec = reviewedManifest.map(slide => {
    const elements = buildSlideSpec(slide, brand, layoutInfo, dim)

    console.log('  S' + slide.slide_number, slide.slide_type, (slide.primary_content || {}).type || '—',
      slide.is_mixed ? '[MIXED+' + ((slide.secondary_content || {}).type || '?') + ']' : '',
      '→', elements.length, 'elements')

    return {
      slide_number:      slide.slide_number,
      slide_type:        slide.slide_type,
      section_name:      slide.section_name  || '',
      section_type:      slide.section_type  || '',
      layout_name:       selectLayoutName(slide, brand.slide_layouts || [], layoutInfo),
      slide_width:       dim.w,
      slide_height:      dim.h,
      background_color:  (slide.slide_type === 'title' || slide.slide_type === 'divider') ? colors.primary : colors.bg,
      title:             slide.title         || '',
      key_message:       slide.key_message   || '',
      is_mixed:          slide.is_mixed      || false,
      visual_type:       (slide.primary_content || {}).type || 'bullets',
      secondary_type:    slide.is_mixed ? ((slide.secondary_content || {}).type || null) : null,
      elements:          elements,
      speaker_note:      slide.speaker_note  || '',
      partner_review:    critique ? { rating: critique.overall_rating, issues_on_slide: (critique.issues || []).filter(i => i.slide_number === slide.slide_number) } : null
    }
  })

  // Log summary
  const mixedCount   = finalSpec.filter(s => s.is_mixed).length
  const totalElements= finalSpec.reduce((s, sl) => s + sl.elements.length, 0)
  const vtBreakdown  = {}
  finalSpec.forEach(s => { vtBreakdown[s.visual_type] = (vtBreakdown[s.visual_type] || 0) + 1 })

  console.log('Agent 5 complete')
  console.log('  Total elements:', totalElements)
  console.log('  Mixed slides:', mixedCount)
  console.log('  Visual breakdown:', JSON.stringify(vtBreakdown))
  console.log('  Partner review:', critique ? critique.overall_rating : 'skipped')

  return finalSpec
}

// ─── LAYOUT NAME SELECTOR ────────────────────────────────────────────────────
function selectLayoutName(slide, availableLayouts, layoutInfo) {
  const st = (slide.slide_type || '').toLowerCase()
  const pt = (slide.primary_content || {}).type || ''

  const find = (keywords) => availableLayouts.find(l => keywords.some(k => (l.name || '').toLowerCase().includes(k)))

  if (st === 'title')   return (find(['title slide']) || {}).name || 'Title slide'
  if (st === 'divider') return (find(['section', 'divider']) || {}).name || 'Section divider'

  if (pt === 'three_column' || pt === 'cards') return (find(['3 across']) || {}).name || '3 across KM'
  if (pt === 'two_column') return (find(['2 across']) || {}).name || '2 across KM'
  if (pt === 'chart' || pt === 'data_table') return (find(['1 across']) || {}).name || '1 across KM'

  return (find(['body text', '1 across']) || {}).name || 'Body text KM'
}
