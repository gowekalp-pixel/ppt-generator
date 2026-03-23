// ─── AGENT 5 — SLIDE LAYOUT & VISUAL DESIGN ENGINE ───────────────────────────
// Input:  state.slideManifest   — output from Agent 4
//         state.brandRulebook   — brand guideline JSON from Agent 2
//         state.outline         — presentation brief from Agent 3
//
// Output: designedSpec — flat JSON array, one render-ready object per slide
//
// Architecture: Claude API call per batch of 4 slides.
// Claude receives the Agent 4 manifest + brand guideline + brief and returns
// a precise layout spec: canvas, brand_tokens, title_block, subtitle_block,
// zones (with fully positioned artifacts), and global_elements.
//
// Agent 5.1 then reviews this spec and applies targeted fixes.
// Agent 6 (python-pptx) consumes the final reviewed spec.

// ═══════════════════════════════════════════════════════════════════════════════
// SYSTEM PROMPT
// ═══════════════════════════════════════════════════════════════════════════════

const AGENT5_SYSTEM = `You are a senior presentation designer and layout system architect.

You will receive:
1. A slide content manifest created by Agent 4
2. A brand guideline JSON for the current deck
3. A presentation brief that explains the narrative flow and tone

═══════════════════════════
BATCH PROCESSING RULE
═══════════════════════════

You will receive slides in batches of 4–5.
Process ONLY the slides in this batch.
Return ONLY those slides in the JSON array.
Do not infer slides before or after this batch.
Do not summarize the entire deck.
Do not renumber slides.

═══════════════════════════
ROLE
═══════════════════════════

Your task is to convert each slide into an exact render-ready layout specification for PowerPoint generation.

You are NOT rewriting content.
You are NOT changing business meaning.

You ARE responsible for:
- spatial layout
- coordinates and dimensions
- typography
- color application
- chart styling
- workflow geometry
- table styling
- card styling
- alignment and spacing

Your output must be directly usable by a PPT rendering engine.

═══════════════════════════
COORDINATE SYSTEM
═══════════════════════════

- unit: inches
- origin: (0.00, 0.00) at top-left of slide
- x increases left → right
- y increases top → bottom
- all numeric values must be decimal inches
- round all numeric values to 2 decimal places

Applies to:
- canvas
- margins
- title/subtitle
- zones
- artifacts
- workflow nodes
- workflow connectors
- tables
- cards
- global elements

═══════════════════════════
BRAND GUIDELINE AUTHORITY RULE
═══════════════════════════

The brand guideline is the primary authority for design.

If the brand guideline defines:
- slide size
- fonts
- color palette
- typography hierarchy
- layout styles
- chart styles
- footer conventions

You MUST follow it.

DO NOT override brand-defined styles.

Only if the brand guideline is missing or incomplete:
- use neutral corporate defaults
- ensure readability and hierarchy
- avoid decorative styling

═══════════════════════════
OUTPUT STRUCTURE
═══════════════════════════

Each slide must return EXACTLY:

{
  "slide_number": number,
  "slide_type": "title" | "divider" | "content",
  "slide_archetype": "summary" | "trend" | "comparison" | "breakdown" | "driver_analysis" | "process" | "recommendation" | "dashboard" | "proof" | "roadmap",
  "canvas": {
    "width_in": number,
    "height_in": number,
    "margin": { "left": number, "right": number, "top": number, "bottom": number },
    "background": { "color": "hex" }
  },
  "brand_tokens": {
    "title_font_family": "string",
    "body_font_family": "string",
    "caption_font_family": "string",
    "title_color": "hex",
    "body_color": "hex",
    "caption_color": "hex",
    "primary_color": "hex",
    "secondary_color": "hex",
    "accent_colors": ["hex"],
    "chart_palette": ["hex"]
  },
  "title_block": {
    "text": "string",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string",
    "font_size": number,
    "font_weight": "regular" | "semibold" | "bold",
    "color": "hex",
    "align": "left" | "center",
    "valign": "middle" | "top",
    "wrap": true
  },
  "subtitle_block": null or {
    "text": "string",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string",
    "font_size": number,
    "font_weight": "regular" | "semibold",
    "color": "hex",
    "align": "left" | "center",
    "valign": "middle" | "top",
    "wrap": true
  },
  "zones": [...],
  "global_elements": {
    "footer": {
      "show": true,
      "x": number, "y": number, "w": number, "h": number,
      "font_family": "string",
      "font_size": number,
      "color": "hex",
      "align": "left"
    },
    "page_number": {
      "show": true,
      "x": number, "y": number, "w": number, "h": number,
      "font_family": "string",
      "font_size": number,
      "color": "hex",
      "align": "right"
    }
  }
}

If subtitle_block is not needed set it to null.
global_elements is optional — include when appropriate.

═══════════════════════════
TITLE / SUBTITLE SIZING
═══════════════════════════

- title slides: 24–34 pt
- divider slides: 22–30 pt
- content slides: 16–22 pt
- subtitle is 9–16 pt smaller than title

═══════════════════════════
ZONES
═══════════════════════════

Each zone:

{
  "zone_id": "z1",
  "zone_role": "string",
  "message_objective": "string",
  "narrative_weight": "primary" | "secondary" | "supporting",
  "frame": {
    "x": number, "y": number, "w": number, "h": number,
    "padding": { "top": number, "right": number, "bottom": number, "left": number }
  },
  "artifacts": [...]
}

Rules:
- max 4 zones per slide
- zones must NOT overlap
- must respect layout_hint from Agent 4
- primary zones get more space than secondary
- each zone has 1–2 artifacts
- title and subtitle sit OUTSIDE zones

═══════════════════════════
ARTIFACT CONTRACT
═══════════════════════════

Every artifact must be FULLY specified. No missing fields. No placeholders.

Allowed types: insight_text | chart | cards | workflow | table

═══════════════════════════
1. INSIGHT TEXT
═══════════════════════════

{
  "type": "insight_text",
  "x": number, "y": number, "w": number, "h": number,
  "style": {
    "fill_color": "hex or null",
    "border_color": "hex or null",
    "border_width": number,
    "corner_radius": number
  },
  "heading_style": {
    "font_family": "string",
    "font_size": number,
    "font_weight": "semibold" | "bold",
    "color": "hex"
  },
  "body_style": {
    "font_family": "string",
    "font_size": number,
    "font_weight": "regular",
    "color": "hex",
    "line_spacing": number,
    "bullet_indent": number
  }
}

═══════════════════════════
2. CHART
═══════════════════════════

{
  "type": "chart",
  "x": number, "y": number, "w": number, "h": number,
  "chart_style": {
    "title_font_family": "string",
    "title_font_size": number,
    "axis_font_family": "string",
    "axis_font_size": number,
    "label_font_family": "string",
    "label_font_size": number,
    "title_color": "hex",
    "axis_color": "hex",
    "gridline_color": "hex",
    "legend_font_family": "string",
    "legend_font_size": number,
    "legend_color": "hex",
    "show_gridlines": true,
    "show_border": false,
    "border_color": null,
    "background_color": null
  },
  "series_style": [
    {
      "series_name": "string",
      "fill_color": "hex",
      "line_color": "hex",
      "line_width": number,
      "marker": "none" | "circle" | "square",
      "data_label_color": "hex"
    }
  ]
}

Chart rules:
- use brand chart_palette in sequence for series
- primary series uses primary brand color
- minimum axis font size: 8pt
- if chart + table in zone: chart takes 60–75% of zone width

═══════════════════════════
3. CARDS
═══════════════════════════

{
  "type": "cards",
  "cards_layout": "row" | "column" | "grid",
  "container": { "x": number, "y": number, "w": number, "h": number },
  "card_style": {
    "fill_color": "hex",
    "border_color": "hex or null",
    "border_width": number,
    "corner_radius": number,
    "shadow": false,
    "internal_padding": number
  },
  "title_style": {
    "font_family": "string",
    "font_size": number,
    "font_weight": "bold",
    "color": "hex"
  },
  "subtitle_style": {
    "font_family": "string",
    "font_size": number,
    "font_weight": "semibold",
    "color": "hex"
  },
  "body_style": {
    "font_family": "string",
    "font_size": number,
    "font_weight": "regular",
    "color": "hex",
    "line_spacing": number
  },
  "card_frames": [
    { "card_index": 0, "x": number, "y": number, "w": number, "h": number }
  ]
}

Card rules: max 4 cards, equal gutters. card_frames must have exact x,y,w,h per card.

═══════════════════════════
4. WORKFLOW
═══════════════════════════

{
  "type": "workflow",
  "container": { "x": number, "y": number, "w": number, "h": number },
  "workflow_style": {
    "title_font_family": "string",
    "title_font_size": number,
    "title_color": "hex",
    "node_fill_color": "hex",
    "node_border_color": "hex",
    "node_border_width": number,
    "node_corner_radius": number,
    "node_title_font_family": "string",
    "node_title_font_size": number,
    "node_title_font_weight": "semibold" | "bold",
    "node_title_color": "hex",
    "node_value_font_family": "string",
    "node_value_font_size": number,
    "node_value_color": "hex",
    "connector_color": "hex",
    "connector_width": number,
    "arrowhead_style": "triangle" | "stealth",
    "connector_label_font_size": number,
    "connector_label_color": "hex"
  },
  "nodes": [
    { "id": "n1", "x": number, "y": number, "w": number, "h": number }
  ],
  "connections": [
    {
      "from": "n1",
      "to": "n2",
      "path": [ { "x": number, "y": number }, { "x": number, "y": number } ],
      "type": "arrow"
    }
  ]
}

Workflow rules:
- max 6 nodes, max 8 connections, no crossing arrows
- follow flow_direction and workflow_type from Agent 4
- left_to_right: horizontal sequence, even spacing
- top_to_bottom: vertical stack, centered
- top_down_branching: root at top, children below, symmetric
- process_flow → linear; hierarchy → tree; decomposition → branching; timeline → equal phases
- all node x,y,w,h must be computed within container bounds
- connection path waypoints: straight or single-elbow only

═══════════════════════════
5. TABLE
═══════════════════════════

{
  "type": "table",
  "x": number, "y": number, "w": number, "h": number,
  "table_style": {
    "header_fill_color": "hex",
    "header_text_color": "hex",
    "header_font_family": "string",
    "header_font_size": number,
    "header_font_weight": "bold",
    "body_fill_color": "hex",
    "body_alt_fill_color": "hex or null",
    "body_text_color": "hex",
    "body_font_family": "string",
    "body_font_size": number,
    "body_font_weight": "regular",
    "grid_color": "hex",
    "grid_width": number,
    "highlight_fill_color": "hex or null",
    "highlight_text_color": "hex or null",
    "cell_padding": number
  },
  "column_widths": [number],
  "row_heights": [number]
}

Table rules: column_widths must sum to table width. body font min 9pt.

═══════════════════════════
INTERNAL ZONE LAYOUT (2 artifacts)
═══════════════════════════

If 2 artifacts in a zone, split zone frame between them:
- chart + insight_text: chart 65%, insight 35%
- workflow + insight_text: workflow dominant
- chart + table: chart 65%, table 35% side by side
- cards + insight_text: cards dominant
- table + insight_text: table 65%, insight 35%
Artifacts must NOT overlap.

═══════════════════════════
LAYOUT DECISION RULES
═══════════════════════════

Translate Agent 4 layout_hint to zone geometry.
Content area = canvas minus margins minus title band.
Gutter between zones: 0.15 inches.

- full: single frame fills content area
- left_60 + right_40: split width 60/40
- left_50 + right_50: equal width split
- top_30 + bottom_70: split height 30/70
- top_left_50 + top_right_50 + bottom_full: two upper + full lower
- left_full_50 + top_right_50_h + bottom_right_50_h: left full + right stacked
- tl + tr + bl + br: 2×2 grid, equal gutters

═══════════════════════════
QUALITY RULES
═══════════════════════════

- no missing fields
- no overlapping zones or artifacts
- body text ≥ 9pt; captions ≥ 8pt
- no placeholder values
- reflect narrative hierarchy in space allocation
- primary zones visually dominant

═══════════════════════════
OUTPUT RULE
═══════════════════════════

Return ONLY a valid JSON array.
No explanation. No markdown. No text outside JSON.`


// ═══════════════════════════════════════════════════════════════════════════════
// BRAND BRIEF BUILDER
// ═══════════════════════════════════════════════════════════════════════════════

function buildBrandBrief(brand, brief) {
  return `BRAND GUIDELINE:
${JSON.stringify(brand, null, 2)}

PRESENTATION BRIEF:
Document type:     ${(brief || {}).document_type     || '—'}
Governing thought: ${(brief || {}).governing_thought || '—'}
Narrative flow:    ${(brief || {}).narrative_flow    || '—'}
Tone:              ${(brief || {}).tone              || 'professional'}
Data heavy:        ${(brief || {}).data_heavy        ? 'yes' : 'no'}`
}


// ═══════════════════════════════════════════════════════════════════════════════
// BATCH WRITER
// ═══════════════════════════════════════════════════════════════════════════════

async function designSlideBatch(batchManifest, brand, brief, batchNum) {
  const slideNums = batchManifest.map(s => s.slide_number)
  console.log('Agent 5 — batch', batchNum, ': slides', slideNums.join(', '))

  const prompt = `${buildBrandBrief(brand, brief)}

SLIDE BATCH ${batchNum} — ${batchManifest.length} slides:
${JSON.stringify(batchManifest, null, 2)}

INSTRUCTIONS:
- Process ONLY these ${batchManifest.length} slides
- Apply brand guideline exactly
- Compute exact coordinates for every element
- Fully specify all artifacts — no missing style fields
- Return a valid JSON array of ${batchManifest.length} slide objects`

  const raw    = await callClaude(AGENT5_SYSTEM, [{ role: 'user', content: prompt }], 4500)
  const parsed = safeParseJSON(raw, null)

  if (!Array.isArray(parsed)) {
    console.warn('Agent 5 batch', batchNum, '— parse failed. Raw length:', raw.length)
    return null
  }

  console.log('Agent 5 batch', batchNum, '— got', parsed.length, 'slides')
  return parsed
}


// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE VALIDATOR
// ═══════════════════════════════════════════════════════════════════════════════

function validateDesignedSlide(slide) {
  const issues = []

  if (!slide.canvas)              issues.push('missing canvas')
  if (!slide.brand_tokens)        issues.push('missing brand_tokens')
  if (!slide.title_block)         issues.push('missing title_block')
  if (!slide.title_block?.text)   issues.push('empty title_block.text')

  if (slide.slide_type === 'content') {
    if (!slide.zones || slide.zones.length === 0) issues.push('no zones on content slide')
  }

  ;(slide.zones || []).forEach((z, zi) => {
    if (!z.frame)             issues.push(`z${zi}: missing frame`)
    if (!z.artifacts?.length) issues.push(`z${zi}: no artifacts`)
    ;(z.artifacts || []).forEach((a, ai) => {
      if (!a.type)                                    issues.push(`z${zi}.a${ai}: missing type`)
      if (a.type === 'chart'    && !a.chart_style)    issues.push(`z${zi}.a${ai}: chart missing chart_style`)
      if (a.type === 'workflow' && !a.nodes?.length)  issues.push(`z${zi}.a${ai}: workflow missing nodes`)
      if (a.type === 'table'    && !a.table_style)    issues.push(`z${zi}.a${ai}: table missing table_style`)
      if (a.type === 'cards'    && !a.card_frames?.length) issues.push(`z${zi}.a${ai}: cards missing card_frames`)
    })
  })

  return issues
}


// ═══════════════════════════════════════════════════════════════════════════════
// FALLBACK BUILDER
// ═══════════════════════════════════════════════════════════════════════════════

function buildFallbackDesign(manifestSlide, brand) {
  const w         = brand.slide_width_inches  || 13.33
  const h         = brand.slide_height_inches || 7.50
  const primary   = (brand.primary_colors    || ['#1A3C8F'])[0]
  const secondary = (brand.secondary_colors  || ['#E8A020'])[0]
  const bg        = (brand.background_colors || ['#FFFFFF'])[0]
  const titleFont = (brand.title_font  || {}).family || 'Calibri'
  const bodyFont  = (brand.body_font   || {}).family || 'Calibri'
  const isDark    = ['title', 'divider'].includes(manifestSlide.slide_type)

  const designed = {
    slide_number:    manifestSlide.slide_number,
    slide_type:      manifestSlide.slide_type      || 'content',
    slide_archetype: manifestSlide.slide_archetype || 'summary',
    canvas: {
      width_in: w, height_in: h,
      margin:     { left: 0.40, right: 0.40, top: 0.15, bottom: 0.30 },
      background: { color: isDark ? primary : bg }
    },
    brand_tokens: {
      title_font_family:   titleFont,
      body_font_family:    bodyFont,
      caption_font_family: bodyFont,
      title_color:         isDark ? '#FFFFFF' : primary,
      body_color:          isDark ? '#CCDDFF' : '#111111',
      caption_color:       '#888888',
      primary_color:       primary,
      secondary_color:     secondary,
      accent_colors:       brand.accent_colors || [],
      chart_palette:       brand.chart_colors  || [primary, secondary, '#2E9E5B', '#C82333']
    },
    title_block: {
      text:        manifestSlide.title || '',
      x: 0.40, y: 0.20, w: w - 0.80,
      h:           isDark ? 2.00 : 0.72,
      font_family: titleFont,
      font_size:   isDark ? 30 : 18,
      font_weight: 'bold',
      color:       isDark ? '#FFFFFF' : primary,
      align: 'left', valign: 'middle', wrap: true
    },
    subtitle_block: manifestSlide.subtitle ? {
      text:        manifestSlide.subtitle,
      x: 0.40, y: isDark ? 2.60 : 0.98,
      w: w - 0.80, h: 0.50,
      font_family: bodyFont, font_size: 14, font_weight: 'regular',
      color:       isDark ? '#BBCCFF' : '#555555',
      align: 'left', valign: 'top', wrap: true
    } : null,
    zones: [],
    global_elements: {
      footer: {
        show: true, x: 0.40, y: h - 0.28, w: 3.00, h: 0.22,
        font_family: bodyFont, font_size: 8, color: '#AAAAAA', align: 'left'
      },
      page_number: {
        show: true, x: w - 0.90, y: h - 0.28, w: 0.65, h: 0.22,
        font_family: bodyFont, font_size: 8, color: '#AAAAAA', align: 'right'
      }
    }
  }

  // Minimal content zone for content slides
  if (manifestSlide.slide_type === 'content') {
    const ct = 1.10  // content top
    designed.zones = [{
      zone_id: 'z1_fallback', zone_role: 'primary_proof',
      message_objective: manifestSlide.key_message || '',
      narrative_weight:  'primary',
      frame: { x: 0.40, y: ct, w: w - 0.80, h: h - ct - 0.45,
               padding: { top: 0.10, right: 0.10, bottom: 0.10, left: 0.10 } },
      artifacts: [{
        type: 'insight_text',
        x: 0.50, y: ct + 0.10, w: w - 1.00, h: h - ct - 0.65,
        style:          { fill_color: null, border_color: null, border_width: 0, corner_radius: 0 },
        heading_style:  { font_family: titleFont, font_size: 12, font_weight: 'bold',    color: primary },
        body_style:     { font_family: bodyFont,  font_size: 11, font_weight: 'regular', color: '#111111', line_spacing: 1.4, bullet_indent: 0.15 }
      }]
    }]
  }

  return designed
}


// ═══════════════════════════════════════════════════════════════════════════════
// NORMALISER
// Merges manifest metadata into the designed slide and adds helper fields
// ═══════════════════════════════════════════════════════════════════════════════

function normaliseDesignedSlide(designed, manifestSlide, brand) {
  if (!designed || typeof designed !== 'object') {
    return buildFallbackDesign(manifestSlide, brand)
  }

  const issues = validateDesignedSlide(designed)
  if (issues.length > 0) {
    console.warn('Agent 5 — S' + (designed.slide_number || '?') + ' validation issues:', issues.join('; '))
  }

  return {
    ...designed,
    // Always carry manifest numbers/types (Claude may drift on these)
    slide_number:      manifestSlide.slide_number,
    slide_type:        manifestSlide.slide_type      || designed.slide_type,
    slide_archetype:   manifestSlide.slide_archetype || designed.slide_archetype || 'summary',
    // Metadata for display and Agent 5.1 review
    section_name:      manifestSlide.section_name    || '',
    section_type:      manifestSlide.section_type    || '',
    title:             (designed.title_block || {}).text || manifestSlide.title       || '',
    subtitle:          (designed.subtitle_block || {}).text || manifestSlide.subtitle || '',
    key_message:       manifestSlide.key_message      || '',
    visual_flow_hint:  manifestSlide.visual_flow_hint || '',
    speaker_note:      manifestSlide.speaker_note     || '',
    layout_name:       inferLayoutName(manifestSlide, brand),
    // Condensed zone summary for Agent 5.1
    zones_summary:     (designed.zones || []).map(z => ({
      zone_id:          z.zone_id,
      zone_role:        z.zone_role,
      narrative_weight: z.narrative_weight,
      artifact_types:   (z.artifacts || []).map(a => a.type)
    })),
    _validation_issues: issues.length > 0 ? issues : undefined
  }
}

function inferLayoutName(manifestSlide, brand) {
  const st   = manifestSlide.slide_type      || 'content'
  const arch = manifestSlide.slide_archetype || ''
  const avail= brand.slide_layouts || []
  const find = (kws) => avail.find(l => kws.some(k => (l.name || '').toLowerCase().includes(k.toLowerCase())))

  if (st === 'title')   return (find(['Title Slide', 'title'])     || {}).name || 'Title Slide'
  if (st === 'divider') return (find(['Section', 'Divider', 'section header']) || {}).name || 'Section Divider'

  if (['recommendation','process','roadmap'].includes(arch))
    return (find(['3 Across','3 across','body text']) || {}).name || 'Body Text'
  if (['dashboard','summary'].includes(arch))
    return (find(['2 Across','1 Across','2 across','1 across']) || {}).name || '1 Across'

  return (find(['1 Across','Body Text','1 across','body text']) || {}).name || 'Body Text'
}


// ═══════════════════════════════════════════════════════════════════════════════
// MAIN RUNNER
// ═══════════════════════════════════════════════════════════════════════════════

async function runAgent5(state) {
  const manifest = state.slideManifest
  const brand    = state.brandRulebook
  const brief    = state.outline || {}

  if (!manifest || !manifest.length) {
    console.error('Agent 5 — slideManifest is empty')
    return []
  }

  console.log('Agent 5 starting — slides to design:', manifest.length)
  console.log('  Brand primary:', (brand.primary_colors || [])[0] || '—')
  console.log('  Slide size:', (brand.slide_width_inches || 13.33) + '" × ' + (brand.slide_height_inches || 7.50) + '"')
  console.log('  Deck type:', brief.document_type || '—')

  // Batch into groups of 4
  const BATCH_SIZE = 4
  const batches    = []
  for (let i = 0; i < manifest.length; i += BATCH_SIZE) {
    batches.push(manifest.slice(i, i + BATCH_SIZE))
  }
  console.log('  Batches:', batches.length, '(max', BATCH_SIZE, 'slides each)')

  const allDesigned = []

  for (let b = 0; b < batches.length; b++) {
    const batch  = batches[b]
    const result = await designSlideBatch(batch, brand, brief, b + 1)

    if (!result) {
      console.warn('Agent 5 — batch', b + 1, 'failed entirely, using fallbacks')
      batch.forEach(ms => allDesigned.push(buildFallbackDesign(ms, brand)))
      continue
    }

    batch.forEach((mSlide, idx) => {
      const match = result.find(r => r.slide_number === mSlide.slide_number)
                 || result[idx]
                 || null
      allDesigned.push(normaliseDesignedSlide(match, mSlide, brand))
    })
  }

  // Summary
  const withIssues = allDesigned.filter(s => s._validation_issues?.length > 0)
  const typeCounts = {}
  allDesigned.forEach(s =>
    (s.zones || []).forEach(z =>
      (z.artifacts || []).forEach(a => { typeCounts[a.type] = (typeCounts[a.type] || 0) + 1 })
    )
  )

  console.log('Agent 5 complete')
  console.log('  Slides designed:', allDesigned.length)
  console.log('  Artifact types:', JSON.stringify(typeCounts))
  console.log('  Slides with validation issues:', withIssues.length)
  withIssues.forEach(s =>
    console.warn('  S' + s.slide_number + ':', s._validation_issues.join('; '))
  )

  return allDesigned
}
