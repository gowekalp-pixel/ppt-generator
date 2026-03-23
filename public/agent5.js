// ─── AGENT 5 — DESIGN DIRECTOR ────────────────────────────────────────────────
// Input:  state.slideManifest  — slide content from Agent 4
//         state.brandRulebook  — colors, fonts, layouts, dimensions from Agent 2
//         state.outline        — presentationBrief from Agent 3
//
// Output: finalSpec — flat JSON array, one fully specified slide object per slide
//         Each slide contains an elements[] array with precise coordinates,
//         dimensions, content, and styling for every element on the slide.
//         Agent 6 reads this and builds the PPTX mechanically — no guessing.
//
// Coordinate system: inches from top-left (0,0)
// Slide dimensions: from brandRulebook (e.g. 11.02" × 8.29" for A4 landscape)

// ─── PRESENTATION TYPE DETECTION ────────────────────────────────────────────
const PRES_TYPE_KEYWORDS = {
  financial:       ['financial','finance','revenue','profit','loss','balance','cash','ebitda','margin','cost','budget','forecast','portfolio','disburs','outstanding','npa','provision'],
  strategic:       ['strategy','strategic','growth','market entry','competitive','position','vision','transformation','initiative','roadmap','merger','acquisition'],
  market_research: ['market','research','competition','competitor','customer','industry','trend','survey','segment','cagr','share'],
  operational:     ['operational','operations','process','efficiency','kpi','performance','metric','production','supply','quality','throughput']
}

function detectPresentationType(brief) {
  if (!brief) return 'financial'
  const text = ((brief.document_type || '') + ' ' + (brief.narrative_flow || '') + ' ' + (brief.governing_thought || '')).toLowerCase()
  for (const [type, keywords] of Object.entries(PRES_TYPE_KEYWORDS)) {
    if (keywords.some(k => text.includes(k))) return type
  }
  return 'financial'
}

// ─── SLIDE DIMENSIONS ────────────────────────────────────────────────────────
function getSlideDimensions(brand) {
  return {
    w: brand.slide_width_inches  || 11.02,
    h: brand.slide_height_inches || 8.29
  }
}

// ─── BRAND COLOR SHORTCUTS ───────────────────────────────────────────────────
function getBrandColors(brand) {
  return {
    primary:    (brand.primary_colors    || ['#0F2FB5'])[0],
    secondary:  (brand.secondary_colors  || ['#FF8E00'])[0],
    bg:         (brand.background_colors || ['#FFFFFF'])[0],
    text:       (brand.text_colors       || ['#000000'])[0],
    accent3:    (brand.accent_colors     || ['#2D962D'])[0],
    accent4:    (brand.all_colors        || {}).accent4 || '#D60202',
    chart:      brand.chart_colors       || ['#0F2FB5','#FF8E00','#2D962D','#D60202','#2E046C','#0092D0'],
    titleFont:  (brand.title_font        || {}).family || 'Arial',
    bodyFont:   (brand.body_font         || {}).family || 'Arial'
  }
}

// ─── ELEMENT BUILDERS ────────────────────────────────────────────────────────
// Each function returns one element object with full positioning and styling

function el_shape(id, x, y, w, h, fill, border_color, border_pt) {
  return { id, type: 'shape', x, y, w, h, fill_color: fill, border_color: border_color || null, border_pt: border_pt || 0 }
}

function el_text(id, x, y, w, h, text, font, size, bold, color, align, valign, italic, wrap) {
  return { id, type: 'text_box', x, y, w, h, text: String(text || ''), font: font || 'Arial', size: size || 14, bold: !!bold, italic: !!italic, color: color || '#000000', align: align || 'left', valign: valign || 'top', wrap: wrap !== false }
}

function el_chart(id, x, y, w, h, chart_type, chart_title, categories, series_list, colors, show_labels, show_legend, x_label, y_label) {
  return {
    id, type: 'chart', chart_type, x, y, w, h,
    chart_title:      chart_title || '',
    x_label:          x_label    || '',
    y_label:          y_label    || '',
    categories:       categories  || [],
    series:           series_list || [],
    colors:           colors      || [],
    show_data_labels: show_labels !== false,
    show_legend:      show_legend || false
  }
}

function el_table(id, x, y, w, h, headers, rows, header_bg, header_text, row_bg_alt, font, font_size, header_size) {
  return {
    id, type: 'table', x, y, w, h,
    headers:        headers     || [],
    rows:           rows        || [],
    header_bg:      header_bg   || '#0F2FB5',
    header_text:    header_text || '#FFFFFF',
    row_bg_alt:     row_bg_alt  || '#F5F5F5',
    font:           font        || 'Arial',
    font_size:      font_size   || 11,
    header_font_size: header_size || 12
  }
}

// ─── STANDARD SLIDE FRAME ────────────────────────────────────────────────────
// Every content slide gets a top accent bar, title, and slide number
function standardFrame(slide, dim, colors, brand) {
  const elements = []

  // Top accent bar
  elements.push(el_shape('top_bar', 0, 0, dim.w, 0.08, colors.primary))

  // Slide title
  elements.push(el_text(
    'title',
    0.4, 0.15, dim.w - 0.8, 0.75,
    slide.title || '',
    colors.titleFont, 22, true, colors.primary,
    'left', 'middle'
  ))

  // Thin underline below title
  elements.push(el_shape('title_underline', 0.4, 0.95, dim.w - 0.8, 0.02, '#E5E7EB'))

  // Slide number bottom right
  elements.push(el_text(
    'slide_number',
    dim.w - 0.8, dim.h - 0.3, 0.6, 0.25,
    String(slide.slide_number || ''),
    colors.bodyFont, 9, false, '#AAAAAA', 'right', 'middle'
  ))

  return elements
}

// ─── CONTENT AREA BOUNDS ─────────────────────────────────────────────────────
function contentArea(dim) {
  return {
    x: 0.4,
    y: 1.05,
    w: dim.w - 0.8,
    h: dim.h - 1.45   // leaves space for title + footer
  }
}

// ─── LAYOUT BUILDERS BY VISUAL TYPE ─────────────────────────────────────────

function buildTitleSlide(slide, dim, colors, brand) {
  const elements = []
  const content  = slide.content || {}

  // Full background
  elements.push(el_shape('bg', 0, 0, dim.w, dim.h, colors.primary))

  // Accent bar (secondary color) at 58% height
  elements.push(el_shape('accent_bar', 0, dim.h * 0.58, dim.w, 0.07, colors.secondary))

  // Main title — large, white, left aligned, upper half
  elements.push(el_text(
    'title', 0.7, dim.h * 0.2, dim.w - 1.4, dim.h * 0.32,
    content.title || slide.title || 'Presentation',
    colors.titleFont, 36, true, '#FFFFFF', 'left', 'middle'
  ))

  // Subtitle
  if (content.subtitle || slide.subtitle) {
    elements.push(el_text(
      'subtitle', 0.7, dim.h * 0.53, dim.w - 1.4, 0.5,
      content.subtitle || slide.subtitle || '',
      colors.bodyFont, 16, false, '#DDDDDD', 'left', 'top'
    ))
  }

  // Date bottom right
  const dateStr = content.date || new Date().toLocaleDateString('en-IN', { month: 'long', year: 'numeric' })
  elements.push(el_text(
    'date', dim.w - 3.5, dim.h * 0.63, 3.2, 0.35,
    dateStr,
    colors.bodyFont, 11, false, '#AAAAAA', 'right', 'middle'
  ))

  return elements
}

function buildDividerSlide(slide, dim, colors, brand) {
  const elements = []
  const content  = slide.content || {}

  // Full background
  elements.push(el_shape('bg', 0, 0, dim.w, dim.h, colors.primary))

  // Left accent bar
  elements.push(el_shape('left_bar', 0, 0, 0.12, dim.h, colors.secondary))

  // Section label (small caps above title)
  elements.push(el_text(
    'section_label', 0.4, dim.h * 0.32, dim.w - 0.8, 0.4,
    'SECTION',
    colors.titleFont, 12, true, colors.secondary, 'left', 'middle'
  ))

  // Section name — large white
  elements.push(el_text(
    'section_name', 0.4, dim.h * 0.4, dim.w - 0.8, 1.4,
    content.section_name || slide.title || '',
    colors.titleFont, 32, true, '#FFFFFF', 'left', 'top'
  ))

  // Descriptor
  if (content.section_descriptor) {
    elements.push(el_text(
      'descriptor', 0.4, dim.h * 0.67, dim.w - 0.8, 0.6,
      content.section_descriptor,
      colors.bodyFont, 14, false, '#CCCCCC', 'left', 'top'
    ))
  }

  return elements
}

function buildStatCallout(slide, dim, colors, brand) {
  const elements = standardFrame(slide, dim, colors, brand)
  const ca       = contentArea(dim)
  const stats    = (slide.content || {}).stats || []

  const count    = Math.min(stats.length, 4)
  const cols     = count <= 2 ? count : (count === 3 ? 3 : 2)
  const rows     = Math.ceil(count / cols)
  const padX     = 0.2
  const padY     = 0.2
  const cellW    = (ca.w - padX * (cols + 1)) / cols
  const cellH    = (ca.h - padY * (rows + 1)) / rows

  stats.slice(0, 4).forEach((stat, i) => {
    const col  = i % cols
    const row  = Math.floor(i / cols)
    const x    = ca.x + padX + col * (cellW + padX)
    const y    = ca.y + padY + row * (cellH + padY)

    // Card background
    elements.push(el_shape('stat_card_' + i, x, y, cellW, cellH, '#F9FAFB', '#E5E7EB', 0.5))

    // Top accent line on card
    elements.push(el_shape('stat_accent_' + i, x, y, cellW, 0.06, colors.primary))

    // Value — large
    elements.push(el_text(
      'stat_value_' + i, x + 0.2, y + 0.2, cellW - 0.4, cellH * 0.45,
      stat.value || '—',
      colors.titleFont, 28, true, colors.primary, 'left', 'middle'
    ))

    // Label
    elements.push(el_text(
      'stat_label_' + i, x + 0.2, y + cellH * 0.55, cellW - 0.4, cellH * 0.22,
      stat.label || '',
      colors.bodyFont, 12, false, '#555555', 'left', 'top'
    ))

    // Change/sub-label
    if (stat.change) {
      elements.push(el_text(
        'stat_change_' + i, x + 0.2, y + cellH * 0.76, cellW - 0.4, cellH * 0.18,
        stat.change,
        colors.bodyFont, 10, false, colors.secondary, 'left', 'top'
      ))
    }
  })

  return elements
}

function buildBulletList(slide, dim, colors, brand) {
  const elements = standardFrame(slide, dim, colors, brand)
  const ca       = contentArea(dim)
  const bullets  = (slide.content || {}).bullets || []

  const rowH     = Math.min(0.85, ca.h / Math.max(bullets.length, 1))

  bullets.forEach((bullet, i) => {
    const y = ca.y + i * rowH

    // Bullet dot (accent color square)
    elements.push(el_shape('dot_' + i, ca.x, y + rowH * 0.35, 0.1, 0.1, colors.secondary))

    // Bullet text
    elements.push(el_text(
      'bullet_' + i, ca.x + 0.22, y, ca.w - 0.22, rowH,
      bullet,
      colors.bodyFont, 14, false, colors.text, 'left', 'middle'
    ))
  })

  return elements
}

function buildChartBar(slide, dim, colors, brand) {
  const elements = standardFrame(slide, dim, colors, brand)
  const ca       = contentArea(dim)
  const content  = slide.content || {}
  const series   = content.series || []

  // Extract categories and series values
  let categories = []
  let seriesList = []

  if (series.length > 0) {
    categories = (series[0].data || []).map(d => d.label || '')
    seriesList = series.map(s => ({
      name:   s.name || '',
      values: (s.data || []).map(d => typeof d.value === 'number' ? d.value : parseFloat(d.value) || 0)
    }))
  }

  elements.push(el_chart(
    'chart', ca.x, ca.y, ca.w, ca.h,
    'bar',
    content.chart_title || '',
    categories,
    seriesList,
    colors.chart,
    true, series.length > 1,
    content.x_label || '', content.y_label || ''
  ))

  return elements
}

function buildChartLine(slide, dim, colors, brand) {
  const elements = standardFrame(slide, dim, colors, brand)
  const ca       = contentArea(dim)
  const content  = slide.content || {}
  const series   = content.series || []

  let categories = []
  let seriesList = []

  if (series.length > 0) {
    categories = (series[0].data || []).map(d => d.label || '')
    seriesList = series.map(s => ({
      name:   s.name || '',
      values: (s.data || []).map(d => typeof d.value === 'number' ? d.value : parseFloat(d.value) || 0)
    }))
  }

  elements.push(el_chart(
    'chart', ca.x, ca.y, ca.w, ca.h,
    'line',
    content.chart_title || '',
    categories,
    seriesList,
    colors.chart,
    false, series.length > 1,
    content.x_label || '', content.y_label || ''
  ))

  return elements
}

function buildChartWaterfall(slide, dim, colors, brand) {
  const elements = standardFrame(slide, dim, colors, brand)
  const ca       = contentArea(dim)
  const content  = slide.content || {}
  const items    = content.items || []

  const categories = items.map(i => i.label || '')
  const values     = items.map(i => typeof i.value === 'number' ? i.value : parseFloat(i.value) || 0)
  const types      = items.map(i => i.type || 'positive')

  // Encode type in series name for python-pptx builder
  elements.push({
    id: 'chart', type: 'chart', chart_type: 'waterfall',
    x: ca.x, y: ca.y, w: ca.w, h: ca.h,
    chart_title: content.chart_title || '',
    categories,
    series: [{ name: 'Value', values, types }],
    colors: [colors.primary, colors.accent4, colors.secondary],
    show_data_labels: true,
    show_legend: false
  })

  return elements
}

function buildDataTable(slide, dim, colors, brand) {
  const elements = standardFrame(slide, dim, colors, brand)
  const ca       = contentArea(dim)
  const content  = slide.content || {}

  elements.push(el_table(
    'table', ca.x, ca.y, ca.w, ca.h,
    content.headers || [],
    content.rows    || [],
    colors.primary, '#FFFFFF', '#F5F5F5',
    colors.bodyFont, 11, 12
  ))

  return elements
}

function buildThreeColumn(slide, dim, colors, brand) {
  const elements = standardFrame(slide, dim, colors, brand)
  const ca       = contentArea(dim)
  const columns  = (slide.content || {}).columns || []

  const padX  = 0.2
  const colW  = (ca.w - padX * 4) / 3

  columns.slice(0, 3).forEach((col, i) => {
    const x = ca.x + padX + i * (colW + padX)

    // Card background
    elements.push(el_shape('col_bg_' + i, x, ca.y, colW, ca.h, '#F9FAFB', '#E5E7EB', 0.5))

    // Top accent
    elements.push(el_shape('col_accent_' + i, x, ca.y, colW, 0.07, colors.secondary))

    // Number
    elements.push(el_text(
      'col_num_' + i, x + 0.15, ca.y + 0.12, 0.5, 0.5,
      String(i + 1).padStart(2, '0'),
      colors.titleFont, 20, true, colors.primary, 'left', 'top'
    ))

    // Header
    elements.push(el_text(
      'col_header_' + i, x + 0.15, ca.y + 0.65, colW - 0.3, 0.5,
      col.header || '',
      colors.titleFont, 14, true, colors.primary, 'left', 'top'
    ))

    // Body
    elements.push(el_text(
      'col_body_' + i, x + 0.15, ca.y + 1.2, colW - 0.3, ca.h - 1.4,
      col.body || '',
      colors.bodyFont, 12, false, colors.text, 'left', 'top'
    ))
  })

  return elements
}

function buildTwoColumn(slide, dim, colors, brand) {
  const elements = standardFrame(slide, dim, colors, brand)
  const ca       = contentArea(dim)
  const content  = slide.content || {}
  const colW     = (ca.w - 0.3) / 2

  // Left column
  elements.push(el_text(
    'left_header', ca.x, ca.y, colW, 0.4,
    content.left_header || '',
    colors.titleFont, 14, true, colors.primary, 'left', 'top'
  ))
  elements.push(el_shape('left_divider', ca.x, ca.y + 0.45, colW, 0.03, colors.primary))
  ;(content.left_points || []).forEach((pt, i) => {
    elements.push(el_shape('ldot_' + i, ca.x, ca.y + 0.6 + i * 0.7, 0.1, 0.1, colors.secondary))
    elements.push(el_text('lpt_' + i, ca.x + 0.2, ca.y + 0.55 + i * 0.7, colW - 0.2, 0.65, pt, colors.bodyFont, 12, false, colors.text, 'left', 'middle'))
  })

  // Vertical divider
  elements.push(el_shape('col_divider', ca.x + colW + 0.1, ca.y, 0.02, ca.h, '#E5E7EB'))

  // Right column
  const rx = ca.x + colW + 0.3
  elements.push(el_text(
    'right_header', rx, ca.y, colW, 0.4,
    content.right_header || '',
    colors.titleFont, 14, true, colors.primary, 'left', 'top'
  ))
  elements.push(el_shape('right_divider', rx, ca.y + 0.45, colW, 0.03, colors.primary))
  ;(content.right_points || []).forEach((pt, i) => {
    elements.push(el_shape('rdot_' + i, rx, ca.y + 0.6 + i * 0.7, 0.1, 0.1, colors.secondary))
    elements.push(el_text('rpt_' + i, rx + 0.2, ca.y + 0.55 + i * 0.7, colW - 0.2, 0.65, pt, colors.bodyFont, 12, false, colors.text, 'left', 'middle'))
  })

  return elements
}

function buildQuoteCallout(slide, dim, colors, brand) {
  const elements = standardFrame(slide, dim, colors, brand)
  const ca       = contentArea(dim)
  const content  = slide.content || {}

  // Opening quote mark
  elements.push(el_text('quote_mark', ca.x, ca.y, 0.8, 1.2, '\u201C', colors.titleFont, 72, true, colors.secondary, 'left', 'top'))

  // Quote text
  elements.push(el_text(
    'quote_text', ca.x + 0.7, ca.y + 0.3, ca.w - 0.7, ca.h * 0.55,
    content.quote || '',
    colors.titleFont, 20, false, colors.text, 'left', 'top'
  ))

  // Accent line
  elements.push(el_shape('quote_line', ca.x + 0.7, ca.y + ca.h * 0.62, 3.5, 0.06, colors.secondary))

  // Attribution
  if (content.attribution) {
    elements.push(el_text(
      'attribution', ca.x + 0.7, ca.y + ca.h * 0.68, ca.w - 0.7, 0.4,
      content.attribution,
      colors.bodyFont, 12, false, '#666666', 'left', 'top'
    ))
  }

  return elements
}

function buildIconCards(slide, dim, colors, brand) {
  const elements = standardFrame(slide, dim, colors, brand)
  const ca       = contentArea(dim)
  const cards    = (slide.content || {}).cards || []

  const count    = Math.min(cards.length, 4)
  const padX     = 0.2
  const cardW    = (ca.w - padX * (count + 1)) / count

  cards.slice(0, 4).forEach((card, i) => {
    const x = ca.x + padX + i * (cardW + padX)

    elements.push(el_shape('card_bg_' + i, x, ca.y, cardW, ca.h, '#F9FAFB', '#E5E7EB', 0.5))
    elements.push(el_shape('card_top_' + i, x, ca.y, cardW, 0.06, colors.primary))

    // Icon placeholder (circle)
    const cx = x + cardW / 2 - 0.35
    elements.push(el_shape('icon_bg_' + i, cx, ca.y + 0.2, 0.7, 0.7, colors.primary + '22', colors.primary, 0.5))
    elements.push(el_text('icon_lbl_' + i, cx, ca.y + 0.2, 0.7, 0.7, (i+1).toString(), colors.titleFont, 16, true, colors.primary, 'center', 'middle'))

    elements.push(el_text('card_header_' + i, x + 0.15, ca.y + 1.1, cardW - 0.3, 0.5, card.header || '', colors.titleFont, 13, true, colors.primary, 'center', 'top'))
    elements.push(el_text('card_desc_' + i, x + 0.15, ca.y + 1.65, cardW - 0.3, ca.h - 1.85, card.description || '', colors.bodyFont, 11, false, colors.text, 'left', 'top'))
  })

  return elements
}

function buildProcessFlow(slide, dim, colors, brand) {
  const elements = standardFrame(slide, dim, colors, brand)
  const ca       = contentArea(dim)
  const steps    = (slide.content || {}).steps || []

  const count    = Math.min(steps.length, 5)
  const padX     = 0.15
  const stepW    = (ca.w - padX * (count + 1)) / count

  steps.slice(0, 5).forEach((step, i) => {
    const x = ca.x + padX + i * (stepW + padX)

    // Step box
    elements.push(el_shape('step_bg_' + i, x, ca.y + 0.5, stepW, ca.h - 0.5, colors.primary + (i === 0 ? 'FF' : '22'), colors.primary, 0.5))

    // Step number circle
    elements.push(el_shape('step_num_bg_' + i, x + stepW/2 - 0.3, ca.y, 0.6, 0.6, colors.primary))
    elements.push(el_text('step_num_' + i, x + stepW/2 - 0.3, ca.y, 0.6, 0.6, String(step.step_number || i+1), colors.titleFont, 14, true, '#FFFFFF', 'center', 'middle'))

    // Arrow between steps
    if (i < count - 1) {
      const arrowX = x + stepW + padX/2
      elements.push(el_text('arrow_' + i, arrowX - 0.15, ca.y + ca.h/2 - 0.15, 0.3, 0.3, '▶', colors.bodyFont, 12, false, colors.secondary, 'center', 'middle'))
    }

    // Step title
    elements.push(el_text('step_title_' + i, x + 0.1, ca.y + 0.65, stepW - 0.2, 0.5, step.title || '', colors.titleFont, 12, true, i === 0 ? '#FFFFFF' : colors.primary, 'center', 'top'))

    // Step description
    elements.push(el_text('step_desc_' + i, x + 0.1, ca.y + 1.2, stepW - 0.2, ca.h - 1.4, step.description || '', colors.bodyFont, 10, false, i === 0 ? '#EEEEEE' : colors.text, 'left', 'top'))
  })

  return elements
}

// ─── LAYOUT SELECTOR ─────────────────────────────────────────────────────────
function selectLayoutName(vt, availableLayouts) {
  const layouts = availableLayouts || []

  const find = (keywords) => layouts.find(l => keywords.some(k => l.name.toLowerCase().includes(k)))

  if (vt === 'title_slide')    return (find(['title slide']) || layouts[0] || {}).name || 'Title slide'
  if (vt === 'divider_slide')  return (find(['section', 'divider']) || layouts[1] || {}).name || 'Section divider'
  if (vt === 'three_column')   return (find(['3 across', 'three']) || find(['2 across']) || {}).name || '3 across KM'
  if (vt === 'two_column')     return (find(['2 across', '1 on 2', '2 on 1']) || {}).name || '2 across KM'
  if (vt === 'data_table')     return (find(['body text', '1 across']) || {}).name || 'Body text KM'
  if (vt === 'chart_bar')      return (find(['1 across', 'body text']) || {}).name || '1 across KM'
  if (vt === 'chart_line')     return (find(['1 across', 'body text']) || {}).name || '1 across KM'
  if (vt === 'chart_waterfall') return (find(['1 across', 'body text']) || {}).name || '1 across KM'
  if (vt === 'stat_callout')   return (find(['2 across', '1 across', 'title only']) || {}).name || '2 across KM'
  if (vt === 'quote_callout')  return (find(['1 across', 'body text']) || {}).name || '1 across KM'
  if (vt === 'icon_cards')     return (find(['3 across', '2 across']) || {}).name || '3 across KM'
  if (vt === 'process_flow')   return (find(['1 across', 'body text']) || {}).name || '1 across KM'
  if (vt === 'bullet_list')    return (find(['body text', '1 across']) || {}).name || 'Body text KM'

  return (find(['body text', '1 across']) || {}).name || 'Body text KM'
}

// ─── VISUAL TYPE REVIEWER ────────────────────────────────────────────────────
function reviewVisualType(slide, presentationType) {
  const vt      = (slide.visual_type || '').toLowerCase()
  const content = slide.content || {}
  const st      = (slide.slide_type || '').toLowerCase()

  if (st === 'title')   return 'title_slide'
  if (st === 'divider') return 'divider_slide'

  // Table too small — downgrade
  if (vt === 'data_table') {
    const rows = content.rows || []
    if (rows.length < 3) return 'stat_callout'
  }

  // Three column needs exactly 3
  if (vt === 'three_column') {
    const cols = content.columns || []
    if (cols.length < 2) return 'bullet_list'
  }

  // Numbers in bullets for financial — upgrade to stat callout
  if (vt === 'bullet_list' && presentationType === 'financial') {
    const bullets = content.bullets || []
    if (bullets.length <= 3 && bullets.some(b => /₹|%|\d+/.test(b))) {
      // Only upgrade if ALL bullets are short stat-like
      const avgLen = bullets.reduce((s, b) => s + b.length, 0) / (bullets.length || 1)
      if (avgLen < 60) return 'stat_callout'
    }
  }

  return vt
}

// ─── MAIN ELEMENT BUILDER ────────────────────────────────────────────────────
function buildSlideElements(slide, dim, colors, brand) {
  const vt = slide.visual_type || 'bullet_list'

  switch (vt) {
    case 'title_slide':     return buildTitleSlide(slide, dim, colors, brand)
    case 'divider_slide':   return buildDividerSlide(slide, dim, colors, brand)
    case 'stat_callout':    return buildStatCallout(slide, dim, colors, brand)
    case 'bullet_list':     return buildBulletList(slide, dim, colors, brand)
    case 'chart_bar':       return buildChartBar(slide, dim, colors, brand)
    case 'chart_line':      return buildChartLine(slide, dim, colors, brand)
    case 'chart_waterfall': return buildChartWaterfall(slide, dim, colors, brand)
    case 'data_table':      return buildDataTable(slide, dim, colors, brand)
    case 'three_column':    return buildThreeColumn(slide, dim, colors, brand)
    case 'two_column':      return buildTwoColumn(slide, dim, colors, brand)
    case 'quote_callout':   return buildQuoteCallout(slide, dim, colors, brand)
    case 'icon_cards':      return buildIconCards(slide, dim, colors, brand)
    case 'process_flow':    return buildProcessFlow(slide, dim, colors, brand)
    default:                return buildBulletList(slide, dim, colors, brand)
  }
}

// ─── MAIN RUNNER ─────────────────────────────────────────────────────────────
async function runAgent5(state) {
  const manifest = state.slideManifest
  const brand    = state.brandRulebook
  const brief    = state.outline

  console.log('Agent 5 starting')
  console.log('  Slides:', manifest.length)
  console.log('  Slide size:', (brand.slide_width_inches || 11.02) + '" x ' + (brand.slide_height_inches || 8.29) + '"')
  console.log('  Available layouts:', (brand.slide_layouts || []).length)

  const presentationType = detectPresentationType(brief)
  const dim              = getSlideDimensions(brand)
  const colors           = getBrandColors(brand)

  console.log('  Presentation type:', presentationType)
  console.log('  Primary color:', colors.primary)
  console.log('  Title font:', colors.titleFont)

  // Step 1 — Review and finalise visual types
  const reviewed = manifest.map(slide => ({
    ...slide,
    visual_type: reviewVisualType(slide, presentationType)
  }))

  // Step 2 — Build element-level spec for every slide
  const finalSpec = reviewed.map(slide => {
    const vt         = slide.visual_type
    const layoutName = selectLayoutName(vt, brand.slide_layouts || [])
    const elements   = buildSlideElements(slide, dim, colors, brand)

    console.log('  Slide', slide.slide_number, '—', vt, '→', layoutName, '→', elements.length, 'elements')

    return {
      slide_number:     slide.slide_number,
      slide_type:       slide.slide_type,
      section_name:     slide.section_name  || '',
      section_type:     slide.section_type  || '',
      layout_name:      layoutName,
      slide_width:      dim.w,
      slide_height:     dim.h,
      background_color: (slide.slide_type === 'title' || slide.slide_type === 'divider')
                          ? colors.primary
                          : colors.bg,
      title:            slide.title         || '',
      subtitle:         slide.subtitle      || '',
      key_message:      slide.key_message   || '',
      visual_type:      vt,
      elements:         elements,
      speaker_note:     slide.speaker_note  || '',
      presentation_type: presentationType
    }
  })

  // Log summary
  const vtBreakdown = {}
  finalSpec.forEach(s => { vtBreakdown[s.visual_type] = (vtBreakdown[s.visual_type] || 0) + 1 })
  console.log('Agent 5 complete')
  console.log('  Visual types:', JSON.stringify(vtBreakdown))
  console.log('  Total elements:', finalSpec.reduce((s, sl) => s + sl.elements.length, 0))

  return finalSpec
}
