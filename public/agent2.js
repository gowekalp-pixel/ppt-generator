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

const AGENT2_CLAUDE_SYSTEM = `You are a senior brand designer reviewing extracted brand guidelines.

You will receive structured data extracted directly from a PowerPoint brand template.
Your job is to:
1. Add a clear "visual_style" description (e.g. "clean corporate", "bold financial", "minimal consulting")
2. Review slide masters and linked slide layouts together, not independently
3. For each slide layout, add a "usage_guidance" field — one sentence on when to use it
4. Add "spacing_notes" based on the slide dimensions, master regions, and layout patterns
5. Add "logo_position" based on layout analysis, especially title/master slides

Return the enriched data as valid JSON only. Keep all existing fields exactly as-is.
Add these new fields at the top level:
- visual_style: string
- spacing_notes: string
- logo_position: string

Add this field to each layout object:
- usage_guidance: string

Return ONLY valid JSON. No explanation. No markdown fences.`

const AGENT2_GLOBAL_ENRICH_SYSTEM = `You are a senior brand designer reviewing extracted PowerPoint brand metadata.
Use the slide masters and layout summary to infer deck-level design guidance.
Return ONLY a valid JSON object with exactly these fields:
{
  "visual_style": "string",
  "spacing_notes": "string",
  "logo_position": "string"
}
No markdown. No explanation.`

const AGENT2_LAYOUT_ENRICH_SYSTEM = `You are a senior brand designer reviewing PowerPoint slide layouts.
You will receive a small batch of layouts plus a summary of the relevant slide masters.
Return ONLY a valid JSON array. For each input layout return:
{
  "name": "original layout name",
  "usage_guidance": "one sentence on when to use this layout"
}
Keep layout names exactly unchanged. No markdown. No explanation.`

function cachePrimaryLogoLocally(primaryLogo) {
  if (!primaryLogo || !primaryLogo.base64) return null
  const localRef = 'brand-logo-' + Date.now()
  try {
    localStorage.setItem(localRef, JSON.stringify({
      name: primaryLogo.name || 'logo',
      mime_type: primaryLogo.mime_type || 'image/png',
      base64: primaryLogo.base64
    }))
    return localRef
  } catch (e) {
    console.warn('Agent 2 — could not cache logo locally:', e.message)
    return null
  }
}

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
    logo_candidates: (d.logos || []).slice(0, 3).map(l => ({
      name: l.name || '',
      width_px: l.width_px || 0,
      height_px: l.height_px || 0,
      usage_score: l.usage_score || 0,
      score: l.score || 0
    })),
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

function chunkArray(arr, size) {
  const out = []
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size))
  return out
}

async function enrichAgent2InBatches(extracted) {
  const summary = summarizeForClaudeEnrichment(extracted)

  let globalEnrichment = null
  try {
    const globalInput = {
      color_scheme_name: summary.color_scheme_name,
      font_scheme_name: summary.font_scheme_name,
      slide_width_inches: summary.slide_width_inches,
      slide_height_inches: summary.slide_height_inches,
      primary_colors: summary.primary_colors,
      secondary_colors: summary.secondary_colors,
      background_colors: summary.background_colors,
      text_colors: summary.text_colors,
      accent_colors: summary.accent_colors,
      chart_colors: summary.chart_colors,
      title_font: summary.title_font,
      body_font: summary.body_font,
      caption_font: summary.caption_font,
      logo_candidates: summary.logo_candidates,
      slide_masters: summary.slide_masters.map(m => ({
        name: m.name,
        background_color: m.background_color,
        grid_summary: m.grid_summary,
        text_style_summary: m.text_style_summary,
        regions: m.regions,
        layout_names: m.layout_names
      })),
      layout_overview: summary.slide_layouts.map(l => ({
        name: l.name,
        type: l.type,
        structure: l.structure,
        master_name: l.master_name,
        grid_summary: l.grid_summary
      }))
    }

    const raw = await callClaude(AGENT2_GLOBAL_ENRICH_SYSTEM, [{
      role: 'user',
      content: 'Infer deck-level brand guidance from this extracted metadata:\n' +
        JSON.stringify(globalInput, null, 2)
    }], 300)
    globalEnrichment = safeParseJSON(raw, null)
  } catch (e) {
    console.warn('Agent 2 Step B — global enrichment failed:', e.message)
  }

  const layoutBatches = chunkArray(summary.slide_layouts || [], 6)
  const usageByName = {}

  for (let i = 0; i < layoutBatches.length; i++) {
    const batch = layoutBatches[i]
    try {
      const masterNames = [...new Set(batch.map(l => l.master_name).filter(Boolean))]
      const relevantMasters = (summary.slide_masters || []).filter(m => masterNames.includes(m.name))
      const batchInput = { slide_masters: relevantMasters, slide_layouts: batch }
      console.log('Agent 2 Step B — layout batch', i + 1, 'of', layoutBatches.length, '| layouts:', batch.length)
      const raw = await callClaude(AGENT2_LAYOUT_ENRICH_SYSTEM, [{
        role: 'user',
        content: 'Add usage guidance for this layout batch:\n' + JSON.stringify(batchInput, null, 2)
      }], 450)
      const parsed = safeParseJSON(raw, null)
      if (Array.isArray(parsed)) {
        parsed.forEach(item => {
          if (item && item.name && item.usage_guidance) usageByName[item.name] = item.usage_guidance
        })
      }
    } catch (e) {
      console.warn('Agent 2 Step B — layout batch', i + 1, 'failed:', e.message)
    }
  }

  return {
    ...(globalEnrichment || {}),
    slide_layouts: (extracted.slide_layouts || []).map(l => ({
      ...l,
      usage_guidance: usageByName[l.name] || l.usage_guidance || ''
    }))
  }
}

const AGENT2_VISION_SYSTEM = `You are an expert brand designer analyzing a brand guideline document.
Extract ALL design rules and return as a single valid JSON object with these exact fields:
{
  "color_scheme_name": "scheme name",
  "primary_colors": ["#hex"],
  "secondary_colors": ["#hex"],
  "background_colors": ["#hex"],
  "text_colors": ["#hex"],
  "accent_colors": ["#hex"],
  "chart_colors": ["#hex"],
  "all_colors": { "accent1": "#hex", "accent2": "#hex" },
  "title_font": { "family": "font name", "size": "28pt", "weight": "bold", "color": "#hex" },
  "body_font": { "family": "font name", "size": "14pt", "weight": "regular", "color": "#hex" },
  "caption_font": { "family": "font name", "size": "9pt", "weight": "regular", "color": "#hex" },
  "slide_width_inches": 10,
  "slide_height_inches": 7.5,
  "visual_style": "corporate",
  "spacing_notes": "0.5 inch margins",
  "logo_position": "top-right",
  "slide_layouts": [
    { "name": "Title slide", "structure": "Full-page title", "usage_guidance": "Use for opening" }
  ]
}
Return ONLY valid JSON. No explanation. No markdown fences.`

async function runAgent2(state, brandContent) {
  console.log('Agent 2 starting — file type:', state.brandExt)

  let extracted = null

  // ── STEP A: Python extraction for PPTX files ──────────────────────────────
  if (state.brandExt === 'pptx' || state.brandExt === 'ppt') {
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

      const data = await res.json()

      if (res.ok && data.success && data.extracted) {
        extracted = data.extracted
        console.log('Agent 2 Step A — extraction successful')
        console.log('  Colors:', Object.keys(extracted.all_colors || {}).length, 'found')
        console.log('  Layouts:', (extracted.slide_layouts || []).length, 'found')
        console.log('  Fonts:', JSON.stringify(extracted.raw_fonts))
        console.log('  Slide size:', extracted.slide_width_inches + '" x ' + extracted.slide_height_inches + '"')
      } else {
        console.warn('Agent 2 Step A — failed:', data.error)
      }

    } catch (e) {
      console.warn('Agent 2 Step A — network error:', e.message)
    }
  }

  // ── STEP B: Claude enrichment if we have extracted data ───────────────────
  if (extracted) {
    try {
      console.log('Agent 2 Step B — enriching with Claude...')
      const enriched = await enrichAgent2InBatches(extracted)

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

  const titleFont = d.title_font || {
    family: (d.raw_fonts && d.raw_fonts.major) || 'Arial',
    size:   '28pt',
    weight: 'bold',
    color:  primary[0] || '#000000'
  }
  const bodyFont = d.body_font || {
    family: (d.raw_fonts && d.raw_fonts.minor) || 'Arial',
    size:   '14pt',
    weight: 'regular',
    color:  text[0] || '#000000'
  }
  const captionFont = d.caption_font || {
    family: bodyFont.family,
    size:   '9pt',
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
  const primaryLogo = d.primary_logo || ((d.logos || [])[0] || null)
  const localLogoRef = cachePrimaryLogoLocally(primaryLogo)

  const rulebook = {
    color_scheme_name:   d.color_scheme_name || 'Brand',
    font_scheme_name:    d.font_scheme_name  || 'Brand',
    visual_style:        d.visual_style      || 'corporate',
    primary_colors:      primary,
    secondary_colors:    secondary,
    background_colors:   background,
    text_colors:         text,
    accent_colors:       accent,
    chart_colors:        chart,
    all_colors:          d.all_colors        || {},
    title_font:          titleFont,
    body_font:           bodyFont,
    caption_font:        captionFont,
    slide_width_inches:  widthIn,
    slide_height_inches: heightIn,
    slide_width_pt:      d.slide_width_pt    || Math.round(widthIn  * 72),
    slide_height_pt:     d.slide_height_pt   || Math.round(heightIn * 72),
    slide_layouts:       layouts,
    slide_masters:       slideMasters,
    layout_blueprints:   buildLayoutBlueprints(layouts),
    master_blueprints:   buildMasterBlueprints(slideMasters),
    logos:               d.logos             || [],
    primary_logo:        primaryLogo,
    primary_logo_local_ref: localLogoRef,
    logo_position:       d.logo_position     || 'top-right',
    spacing_notes:       d.spacing_notes     || '0.5 inch margins',
    extraction_source:   d.source            || 'agent2'
  }

  console.log('Agent 2 — rulebook complete:')
  console.log('  Primary colors:', rulebook.primary_colors)
  console.log('  Title font:', rulebook.title_font.family, rulebook.title_font.size)
  console.log('  Slide:', rulebook.slide_width_inches + '" x ' + rulebook.slide_height_inches + '"')
  console.log('  Layouts:', rulebook.slide_layouts.length)

  return rulebook
}
