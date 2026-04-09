// ─── AGENT 2 — BRAND GUIDELINE PARSER ────────────────────────────────────────
// Input:  state (brandB64, brandExt, brandFile)
// Output: brandRulebook object with accurate colors, fonts, sizes, layouts
//
// Flow:
//   Step A — Call /api/extract-brand (Python)
//            Reads PPTX directly: theme colors, fonts, slide size, layout structures
//            This is 100% accurate — no guessing
//
//   Step B — Call Claude to enrich with usage guidance and visual style
//
//   Step C — If file is PDF or image (not PPTX)
//            Skip Step A, use Claude vision to extract brand rules

// ─── PROMPT TEXT ──────────────────────────────────────────────────────────────
// Loaded via <script> tags (see index.html):
//   prompts/agent2/P1-enrich-vision.js → _A2_ENRICH, _A2_VISION

const AGENT2_ENRICH_SYSTEM = _A2_ENRICH
const AGENT2_VISION_SYSTEM = _A2_VISION

function buildLayoutBlueprints(layouts) {
  return (layouts || []).map(l => ({
    name: l.name || 'Unknown',
    type: l.type || 'custom',
    structure: l.structure || 'Content layout',
    master_name: l.master_name || '',
    ph_count: l.ph_count || 0,
    grid_summary: l.grid_summary || { rows: 0, columns: 0, content_blocks: 0 },
    title_placeholder: l.title_placeholder || null,
    body_placeholder: l.body_placeholder || null,
    usage_guidance: l.usage_guidance || ''
  }))
}

function buildMasterBlueprints(masters) {
  return (masters || []).map(m => ({
    name: m.name || 'Master',
    background_color: m.background_color || null,
    placeholder_count: m.placeholder_count || 0,
    grid_summary: m.grid_summary || { rows: 0, columns: 0, content_blocks: 0 },
    text_style_summary: m.text_style_summary || {},
    regions: m.regions || {},
    layout_names: m.layout_names || [],
    media_refs: m.media_refs || []
  }))
}

function summarizeForClaudeEnrichment(extracted) {
  const d = extracted || {}
  return {
    color_scheme_name: d.color_scheme_name || '',
    font_scheme_name: d.font_scheme_name || '',
    slide_width_inches: d.slide_width_inches || 0,
    slide_height_inches: d.slide_height_inches || 0,
    primary_colors: (d.primary_colors || []).slice(0, 3),
    secondary_colors: (d.secondary_colors || []).slice(0, 3),
    background_colors: (d.background_colors || []).slice(0, 3),
    text_colors: (d.text_colors || []).slice(0, 3),
    accent_colors: (d.accent_colors || []).slice(0, 6),
    chart_colors: (d.chart_colors || []).slice(0, 6),
    title_font: d.title_font || {},
    body_font: d.body_font || {},
    caption_font: d.caption_font || {},
    slide_masters: (d.slide_masters || []).slice(0, 6).map(m => ({
      name: m.name || '',
      background_color: m.background_color || null,
      placeholder_count: m.placeholder_count || 0,
      grid_summary: m.grid_summary || { rows: 0, columns: 0, content_blocks: 0 },
      text_style_summary: m.text_style_summary || {},
      regions: m.regions || {},
      layout_names: (m.layout_names || []).slice(0, 12),
      media_refs: (m.media_refs || []).slice(0, 6).map(r => ({ name: r.name || '', target: r.target || '' }))
    })),
    slide_layouts: (d.slide_layouts || []).slice(0, 18).map(l => ({
      name: l.name || '',
      type: l.type || 'custom',
      structure: l.structure || '',
      ph_count: l.ph_count || 0,
      grid_summary: l.grid_summary || { rows: 0, columns: 0, content_blocks: 0 },
      master_name: l.master_name || '',
      title_placeholder: l.title_placeholder ? {
        x_in: l.title_placeholder.x_in,
        y_in: l.title_placeholder.y_in,
        w_in: l.title_placeholder.w_in,
        h_in: l.title_placeholder.h_in
      } : null,
      body_placeholder: l.body_placeholder ? {
        x_in: l.body_placeholder.x_in,
        y_in: l.body_placeholder.y_in,
        w_in: l.body_placeholder.w_in,
        h_in: l.body_placeholder.h_in
      } : null
    }))
  }
}


async function enrichAgent2(extracted) {
  const summary = summarizeForClaudeEnrichment(extracted)

  const input = {
    color_scheme_name:  summary.color_scheme_name,
    font_scheme_name:   summary.font_scheme_name,
    slide_width_inches: summary.slide_width_inches,
    slide_height_inches: summary.slide_height_inches,
    primary_colors:     summary.primary_colors,
    secondary_colors:   summary.secondary_colors,
    background_colors:  summary.background_colors,
    text_colors:        summary.text_colors,
    accent_colors:      summary.accent_colors,
    chart_colors:       summary.chart_colors,
    title_font:         summary.title_font,
    body_font:          summary.body_font,
    caption_font:       summary.caption_font,
    slide_masters: summary.slide_masters.map(m => ({
      name: m.name,
      background_color: m.background_color,
      grid_summary: m.grid_summary,
      text_style_summary: m.text_style_summary,
      regions: m.regions,
      layout_names: m.layout_names
    })),
    slide_layouts: summary.slide_layouts.map(l => ({
      name: l.name,
      type: l.type,
      structure: l.structure,
      master_name: l.master_name,
      grid_summary: l.grid_summary
    }))
  }

  console.log('Agent 2 Step B — single enrichment call (' + summary.slide_layouts.length + ' layouts)')
  const raw = await callClaude(AGENT2_ENRICH_SYSTEM, [{
    role: 'user',
    content: 'Infer brand guidance from this extracted metadata:\n' + JSON.stringify(input, null, 2)
  }], 1200)

  const parsed = safeParseJSON(raw, null)
  if (!parsed) return {}

  const usageByName = {}
  if (Array.isArray(parsed.layout_usage)) {
    parsed.layout_usage.forEach(item => {
      if (item && item.name && item.usage_guidance) usageByName[item.name] = item.usage_guidance
    })
  }

  return {
    ...parsed,
    slide_layouts: (extracted.slide_layouts || []).map(l => ({
      ...l,
      usage_guidance: usageByName[l.name] || l.usage_guidance || ''
    }))
  }
}

function hasUsablePptxTemplateExtraction(extracted) {
  const layouts = extracted?.slide_layouts || []
  const masters = extracted?.slide_masters || []
  if (!layouts.length || !masters.length) return false
  const contentLike = layouts.filter(l => {
    const t = String(l?.type || '').toLowerCase()
    const n = String(l?.name || '').toLowerCase()
    const ph = +l?.ph_count || 0
    return !['title', 'sechead', 'blank'].includes(t) &&
      !/thank[\s_-]*you|end[\s_-]*slide|closing[\s_-]*slide/i.test(n) &&
      (ph > 0 || /content|topic|image|two|column|body/i.test(n))
  })
  return contentLike.length > 0
}

async function runAgent2(state, brandContent) {
  console.log('Agent 2 starting — file type:', state.brandExt)

  const isPptTemplate = state.brandExt === 'pptx' || state.brandExt === 'ppt'
  let extracted = null

  // ── STEP A: Python extraction for PPTX files ──────────────────────────────
  if (isPptTemplate) {
    try {
      console.log('Agent 2 Step A — calling /api/extract-brand...')

      const res = await fetch('/api/extract-brand', {
        method:  'POST',
        headers: { 'Content-Type': 'application/json' },
        body:    JSON.stringify({
          pptxBase64: state.brandB64,
          fileType:   state.brandExt
        })
      })

      let data
      try {
        data = await res.json()
      } catch (_) {
        const raw = await res.text().catch(() => '')
        const hint = /request entity too large|payload too large/i.test(raw)
          ? 'Request payload too large for /api/extract-brand'
          : ('Non-JSON response from /api/extract-brand: ' + (raw || ('status ' + res.status)).slice(0, 120))
        throw new Error(hint)
      }

      if (res.ok && data.success && data.extracted) {
        extracted = data.extracted
        console.log('Agent 2 Step A — extraction successful')
        console.log('  Colors:', Object.keys(extracted.all_colors || {}).length, 'found')
        console.log('  Layouts:', (extracted.slide_layouts || []).length, 'found')
        console.log('  Masters:', (extracted.slide_masters || []).length, 'found')
        console.log('  Fonts:', JSON.stringify(extracted.raw_fonts))
        console.log('  Slide size:', extracted.slide_width_inches + '" x ' + extracted.slide_height_inches + '"')
        if (!hasUsablePptxTemplateExtraction(extracted)) {
          console.warn('Agent 2 Step A â€” PPTX extraction returned no usable content layouts')
        }
      } else {
        console.warn('Agent 2 Step A — failed:', data.error)
      }

    } catch (e) {
      console.warn('Agent 2 Step A — network error:', e.message)
    }
  }

  // ── STEP B: Claude enrichment if we have extracted data ───────────────────
  if (extracted) {
    if (isPptTemplate && !hasUsablePptxTemplateExtraction(extracted)) {
      throw new Error('PPTX extraction completed but did not return usable slide masters/content layouts. Check /api/extract-brand or run the app through vercel dev so the Python endpoint is available.')
    }
    try {
      console.log('Agent 2 Step B — enriching with Claude...')
      const enriched = await enrichAgent2(extracted)

      if (enriched && (enriched.primary_colors || enriched.slide_layouts)) {
        console.log('Agent 2 Step B — enrichment successful')
        return buildRulebook({ ...extracted, ...enriched })
      } else {
        console.warn('Agent 2 Step B — parse failed, using raw extraction only')
        return buildRulebook(extracted)
      }

    } catch (e) {
      console.warn('Agent 2 Step B — Claude error:', e.message)
      return buildRulebook(extracted)
    }
  }

  // ── STEP C: Claude vision for PDF / image / fallback ─────────────────────
  if (isPptTemplate) {
    throw new Error('PPTX brand extraction failed or returned no usable slide masters/content layouts. Agent 2 will not fall back to vision for PPTX templates because that loses layout geometry.')
  }

  console.log('Agent 2 Step C — using Claude vision...')

  let messages

  if (brandContent === '__PDF__') {
    messages = [{
      role: 'user',
      content: [
        { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: state.brandB64 } },
        { type: 'text', text: 'Extract all brand design rules from this PDF. Return JSON only.' }
      ]
    }]
  } else if (brandContent === '__IMAGE__') {
    const mime = state.brandExt === 'png' ? 'image/png' : 'image/jpeg'
    messages = [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: mime, data: state.brandB64 } },
        { type: 'text', text: 'Extract all brand design rules from this image. Return JSON only.' }
      ]
    }]
  } else {
    messages = [{
      role: 'user',
      content: 'Brand guideline content:\n\n' + brandContent + '\n\nReturn brand rules as JSON only.'
    }]
  }

  const raw    = await callClaude(AGENT2_VISION_SYSTEM, messages, 2000)
  const result = safeParseJSON(raw, null)

  if (result) {
    console.log('Agent 2 Step C — vision extraction successful')
    return buildRulebook(result)
  }

  console.warn('Agent 2 — all methods failed, using safe defaults')
  return buildRulebook(null)
}


// ─── BUILD NORMALISED RULEBOOK ────────────────────────────────────────────────
function buildRulebook(data) {
  const d = data || {}

  const primary    = d.primary_colors    || []
  const secondary  = d.secondary_colors  || []
  const accent     = d.accent_colors     || []
  const background = d.background_colors || ['#FFFFFF']
  const text       = d.text_colors       || ['#000000']
  const chart      = d.chart_colors      || [...primary, ...secondary, ...accent].slice(0, 6)

  // Extract actual font sizes from master placeholder data when available
  const typo = d.typography_hierarchy || {}
  const masterTitleStyle = ((d.slide_masters || [])[0] || {}).text_style_summary || {}
  const masterTitlePt = (masterTitleStyle.title && masterTitleStyle.title.font_size_pt) || null
  const masterBodyPt  = (masterTitleStyle.body  && masterTitleStyle.body.font_size_pt)  || null

  const titleFontSize = typo.title_size_pt ? (typo.title_size_pt + 'pt')
    : masterTitlePt ? (masterTitlePt + 'pt') : '24pt'
  const bodyFontSize  = typo.body_size_pt  ? (typo.body_size_pt  + 'pt')
    : masterBodyPt  ? (masterBodyPt  + 'pt') : '12pt'

  const titleFont = d.title_font || {
    family: (d.raw_fonts && d.raw_fonts.major) || 'Arial',
    size:   titleFontSize,
    weight: 'bold',
    color:  typo.title_color || primary[0] || '#000000'
  }
  // Override size even if title_font already exists but has a hardcoded default
  if (d.title_font && d.title_font.size === '28pt' && masterTitlePt) {
    titleFont.size = masterTitlePt + 'pt'
  }

  const bodyFont = d.body_font || {
    family: (d.raw_fonts && d.raw_fonts.minor) || 'Arial',
    size:   bodyFontSize,
    weight: 'regular',
    color:  typo.body_color || text[0] || '#000000'
  }
  if (d.body_font && d.body_font.size === '14pt' && masterBodyPt) {
    bodyFont.size = masterBodyPt + 'pt'
  }

  const captionFont = d.caption_font || {
    family: bodyFont.family,
    size:   typo.caption_size_pt ? (typo.caption_size_pt + 'pt') : '9pt',
    weight: 'regular',
    color:  '#666666'
  }

  const widthIn  = d.slide_width_inches  || 10
  const heightIn = d.slide_height_inches || 7.5

  const layouts = (d.slide_layouts || []).map(l => ({
    name:           l.name           || 'Unknown',
    type:           l.type           || 'custom',
    structure:      l.structure      || 'Content layout',
    usage_guidance: l.usage_guidance || '',
    ph_count:       l.ph_count       || 0,
    grid_summary:   l.grid_summary   || { rows: 0, columns: 0, content_blocks: 0 },
    master_name:    l.master_name    || '',
    master_summary: l.master_summary || null,
    title_placeholder: l.title_placeholder || null,
    body_placeholder:  l.body_placeholder  || null,
    placeholders:   l.placeholders   || []
  }))

  // ── Identify special-purpose vs content layouts ────────────────────────────
  // Priority 1: OOXML type attribute (authoritative)
  // Priority 2: Name heuristics (fallback for PDF/image-sourced brands)
  const _isNonContentLayout = (l) => {
    const t = (l.type || '').toLowerCase()
    const n = (l.name || '').toLowerCase()
    return (
      t === 'title'   || t === 'sechead' || t === 'blank' ||
      /^blank$/i.test(n) ||
      /thank[\s_-]*you|end[\s_-]*slide|closing[\s_-]*slide/i.test(n)
    )
  }

  const titleLayout = layouts.find(l => l.type === 'title') ||
    layouts.find(l => /title[\s_-]*slide|^title$/i.test(l.name))

  const dividerLayout = layouts.find(l => l.type === 'secHead' || l.type === 'sechead') ||
    layouts.find(l => /section[\s_-]*header|^section$|divider/i.test(l.name))

  const thankYouLayout = layouts.find(l =>
    /thank[\s_-]*you|end[\s_-]*slide|closing[\s_-]*slide/i.test(l.name))

  const contentLayouts = layouts.filter(l => !_isNonContentLayout(l))
  const slideMasters = (d.slide_masters || []).map(m => ({
    name: m.name || 'Master',
    path: m.path || '',
    background_color: m.background_color || null,
    placeholder_count: m.placeholder_count || 0,
    grid_summary: m.grid_summary || { rows: 0, columns: 0, content_blocks: 0 },
    text_style_summary: m.text_style_summary || {},
    regions: m.regions || {},
    media_refs: m.media_refs || [],
    layout_paths: m.layout_paths || [],
    layout_names: m.layout_names || [],
    title_placeholder: m.title_placeholder || null,
    body_placeholder: m.body_placeholder || null
  }))
  // Build distinct chart color sequence — use enriched sequence if available,
  // otherwise fall back to accent1–6 from all_colors, ensuring 6 distinct entries
  const allC = d.all_colors || {}
  const enrichedChartSeq = d.chart_color_sequence || []
  const accentSeq = ['accent1','accent2','accent3','accent4','accent5','accent6']
    .map(k => allC[k]).filter(Boolean)
  const chartColorSeq = enrichedChartSeq.length >= 3 ? enrichedChartSeq : accentSeq.length >= 2 ? accentSeq : chart

  const bulletStyle = d.bullet_style || { char: '•', indent_inches: 0.12, space_before_pt: 4, space_after_pt: 0 }
  const insightBoxStyle = d.insight_box_style || { fill_color: null, border_color: null, corner_radius: 4 }
  const typographyHierarchy = d.typography_hierarchy || {
    title_size_pt:    parseInt(titleFont.size) || 24,
    subtitle_size_pt: parseInt(titleFont.size) ? parseInt(titleFont.size) - 6 : 18,
    body_size_pt:     parseInt(bodyFont.size)  || 12,
    caption_size_pt:  9,
    title_color:      titleFont.color,
    body_color:       bodyFont.color
  }

  const rulebook = {
    color_scheme_name:    d.color_scheme_name  || 'Brand',
    font_scheme_name:     d.font_scheme_name   || 'Brand',
    visual_style:         d.visual_style       || 'corporate',
    primary_colors:       primary,
    secondary_colors:     secondary,
    background_colors:    background,
    text_colors:          text,
    accent_colors:        accent,
    chart_colors:         chart,
    chart_color_sequence: chartColorSeq,
    all_colors:           allC,
    title_font:           titleFont,
    body_font:            bodyFont,
    caption_font:         captionFont,
    typography_hierarchy: typographyHierarchy,
    bullet_style:         bulletStyle,
    insight_box_style:    insightBoxStyle,
    slide_width_inches:   widthIn,
    slide_height_inches:  heightIn,
    slide_width_pt:       d.slide_width_pt    || Math.round(widthIn  * 72),
    slide_height_pt:      d.slide_height_pt   || Math.round(heightIn * 72),
    slide_layouts:        layouts,
    slide_masters:        slideMasters,
    layout_blueprints:    buildLayoutBlueprints(layouts),
    master_blueprints:    buildMasterBlueprints(slideMasters),
    spacing_notes:        d.spacing_notes     || '0.5 inch margins',
    extraction_source:    d.source            || 'agent2',
    // True when the brand was extracted from a real PPTX master.
    // Agent 5 uses this to skip background/logo/footer (the master handles them).
    uses_template:        slideMasters.length > 0 && layouts.length > 0,

    // ── Layout classification (used by Agent 4, 5, and 6) ──────────────────
    // Identifies which named layouts serve which structural role.
    // content_layout_names excludes title, section header, blank, and thank-you
    // layouts so Agent 4's "5+ layouts → use layout mode" threshold only counts
    // layouts that are actually valid for content slides.
    title_layout_name:     titleLayout    ? titleLayout.name    : null,
    divider_layout_name:   dividerLayout  ? dividerLayout.name  : null,
    thank_you_layout_name: thankYouLayout ? thankYouLayout.name : null,
    content_layout_names:  contentLayouts.map(l => l.name)
  }

  console.log('Agent 2 — rulebook complete:')
  console.log('  Primary colors:', rulebook.primary_colors)
  console.log('  Title font:', rulebook.title_font.family, rulebook.title_font.size)
  console.log('  Slide:', rulebook.slide_width_inches + '" x ' + rulebook.slide_height_inches + '"')
  console.log('  Layouts:', rulebook.slide_layouts.length)

  return rulebook
}
