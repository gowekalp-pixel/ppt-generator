// ─── AGENT 4 — SLIDE CONTENT WRITER ──────────────────────────────────────────
// Input:  state.outline (presentationBrief from Agent 3)
//         state.contentB64 (original PDF — fallback for missing detail)
// Output: slideManifest — flat JSON array, one object per slide
//
// Agent 4's job:
// - Expand every section from the brief into individual slides
// - Decide the BEST way to present each message (text, table, chart, cards, quote)
// - Write the actual content for every slide — title, key message, body content
// - Go back to the PDF when Agent 3's brief lacks enough detail
// - NO brand colours, NO fonts, NO layout positions — that is Agent 5's job
// - Every slide must have ONE clear key message the audience takes away

const AGENT4_SYSTEM = `You are a senior management consultant and presentation specialist.
You have two inputs:
1. A presentation brief from a document analyst (structured JSON)
2. The original source document (PDF) as a fallback for detail

Your job is to write the CONTENT for every single slide in the presentation.

You must produce a flat JSON array — one object per slide — covering the FULL presentation.

Each slide object must have EXACTLY these fields:

{
  "slide_number": 1,
  "section_name": "string — which section this slide belongs to",
  "section_type": "string — type from the brief (title/executive_summary/financial_data etc)",
  "slide_type": "string — one of: title, divider, content",
  "title": "string — the slide title (specific, not generic)",
  "subtitle": "string — subtitle for title slides, empty for others",
  "key_message": "string — THE single most important sentence this slide communicates. This is what the audience must remember. Make it specific and insight-driven, not descriptive.",
  "visual_type": "string — the best visual representation for this content",
  "content": { ... },
  "speaker_note": "string — what the presenter says that is NOT on the slide"
}

VISUAL TYPE SELECTION RULES — choose the type that best wins eyeballs for this content:

"bullet_list"    — 3 to 5 crisp bullet points. Use for qualitative analysis, observations, strategic points.
"three_column"   — exactly 3 parallel items with header + description each. Use for 3 options, 3 drivers, 3 recommendations.
"two_column"     — two contrasting or complementary groups. Use for before/after, problem/solution, pros/cons.
"data_table"     — structured rows and columns with headers. Use for financial statements, comparisons, schedules.
"stat_callout"   — 2 to 4 large numbers with labels. Use when specific metrics are the story (revenue, growth %, market share).
"chart_bar"      — bar chart with categories and values. Use for revenue by segment, performance comparison.
"chart_line"     — line chart with time series. Use for trends over time (quarterly, monthly, annual).
"chart_waterfall"— waterfall/bridge chart. Use for P&L bridges, variance analysis, cost buildup.
"quote_callout"  — one powerful insight or statement in large text. Use for executive summary opener or governing thought.
"icon_cards"     — 3 to 4 cards each with an icon description, header, and 1-line description. Use for key themes or initiatives.
"title_slide"    — full-page title. Only for the opening slide.
"divider_slide"  — section break with section name and brief descriptor.
"process_flow"   — numbered sequential steps. Use for methodology, implementation plan, process description.

CONTENT FIELD RULES — each visual_type has its own content structure:

For "bullet_list":
"content": { "bullets": ["bullet 1", "bullet 2", "bullet 3"] }

For "three_column":
"content": { "columns": [ { "header": "string", "body": "string" }, ... ] }

For "two_column":
"content": { "left_header": "string", "left_points": ["..."], "right_header": "string", "right_points": ["..."] }

For "data_table":
"content": { "headers": ["Col1", "Col2", "Col3"], "rows": [ ["val", "val", "val"], ... ] }

For "stat_callout":
"content": { "stats": [ { "value": "₹450Cr", "label": "Revenue", "change": "+18% YoY" }, ... ] }

For "chart_bar":
"content": { "chart_title": "string", "x_label": "string", "y_label": "string", "series": [ { "name": "string", "data": [ { "label": "Q1", "value": 450 }, ... ] } ] }

For "chart_line":
"content": { "chart_title": "string", "x_label": "string", "y_label": "string", "series": [ { "name": "string", "data": [ { "label": "FY22", "value": 320 }, ... ] } ] }

For "chart_waterfall":
"content": { "chart_title": "string", "items": [ { "label": "Revenue", "value": 450, "type": "positive" }, { "label": "COGS", "value": -180, "type": "negative" }, { "label": "Gross Profit", "value": 270, "type": "total" } ] }

For "quote_callout":
"content": { "quote": "string — the insight in full", "attribution": "string — source or context" }

For "icon_cards":
"content": { "cards": [ { "icon": "string — describe the icon concept e.g. growth arrow, shield, clock", "header": "string", "description": "string" }, ... ] }

For "title_slide":
"content": { "title": "string", "subtitle": "string", "date": "string" }

For "divider_slide":
"content": { "section_name": "string", "section_descriptor": "string — one line on what this section covers" }

For "process_flow":
"content": { "steps": [ { "step_number": 1, "title": "string", "description": "string" }, ... ] }

CRITICAL RULES:
1. Total slides must equal EXACTLY the number specified in the brief (total_slides field)
2. The sum of slides per section must match suggested_slide_count from the brief
3. Slide 1 is ALWAYS a title_slide
4. Section dividers get slide_type "divider" and visual_type "divider_slide"
5. key_message must be SPECIFIC — use actual numbers, names, and facts from the document
6. key_message must be an INSIGHT not a description — "Revenue grew 18% driven by product X" not "Revenue information is shown"
7. If the brief does not have enough detail for a slide, use the source PDF to fill it in
8. content must be fully populated — no placeholders, no "TBD", no "insert data here"
9. Return ONLY a valid JSON array. No explanation. No markdown fences.`


async function runAgent4(state) {
  const brief = state.outline  // presentationBrief from Agent 3
  console.log('Agent 4 starting')
  console.log('  Brief sections:', (brief.sections || []).length)
  console.log('  Target slides:', brief.total_slides || state.slideCount)
  console.log('  Data heavy:', brief.data_heavy)

  // Build the user prompt — brief + instruction to use PDF for detail
  const userPrompt = `PRESENTATION BRIEF (from document analyst):
${JSON.stringify(brief, null, 2)}

INSTRUCTIONS:
- Produce exactly ${brief.total_slides || state.slideCount} slides as a flat JSON array
- Use the presentation brief above as your primary source
- Where the brief lacks detail, draw on the source PDF document attached below
- Every slide must have a specific, insight-driven key_message
- Choose the best visual_type for each slide's content
- Fully populate the content field — no placeholders

Return ONLY the JSON array. No explanation. No markdown.`

  const messages = [{
    role: 'user',
    content: [
      {
        type: 'document',
        source: {
          type:       'base64',
          media_type: 'application/pdf',
          data:       state.contentB64
        }
      },
      {
        type: 'text',
        text: userPrompt
      }
    ]
  }]

  console.log('Agent 4 — calling Claude with brief + PDF...')
  const raw = await callClaude(AGENT4_SYSTEM, messages, 4000)
  console.log('Agent 4 — response length:', raw.length, 'chars')

  const fallback = buildFallbackManifest(brief, state.slideCount)
  let manifest   = safeParseJSON(raw, fallback)

  // ── Validate ──────────────────────────────────────────────────────────────
  if (!Array.isArray(manifest) || manifest.length === 0) {
    console.warn('Agent 4 — result not an array, using fallback')
    return fallback
  }

  const target = brief.total_slides || state.slideCount
  if (manifest.length < target) {
    console.warn('Agent 4 — got', manifest.length, 'slides, expected', target)
  }

  // ── Normalise every slide ─────────────────────────────────────────────────
  manifest = manifest.map((slide, i) => normaliseSlide(slide, i))

  console.log('Agent 4 complete — slides produced:', manifest.length)
  console.log('  Visual types used:', [...new Set(manifest.map(s => s.visual_type))].join(', '))
  console.log('  Slide types:', manifest.filter(s => s.slide_type === 'title').length, 'title,',
    manifest.filter(s => s.slide_type === 'divider').length, 'dividers,',
    manifest.filter(s => s.slide_type === 'content').length, 'content')

  return manifest
}


// ─── NORMALISE A SINGLE SLIDE ─────────────────────────────────────────────────
function normaliseSlide(slide, index) {
  const slideType = slide.slide_type || inferSlideType(slide)

  return {
    slide_number:  slide.slide_number  || (index + 1),
    section_name:  slide.section_name  || '',
    section_type:  slide.section_type  || 'content',
    slide_type:    slideType,
    title:         slide.title         || 'Slide ' + (index + 1),
    subtitle:      slide.subtitle      || '',
    key_message:   slide.key_message   || slide.speaker_note || '',
    visual_type:   slide.visual_type   || inferVisualType(slide, slideType),
    content:       normaliseContent(slide.content, slide.visual_type, slide),
    speaker_note:  slide.speaker_note  || ''
  }
}

function inferSlideType(slide) {
  const vt = (slide.visual_type || '').toLowerCase()
  if (vt === 'title_slide')   return 'title'
  if (vt === 'divider_slide') return 'divider'
  const st = (slide.section_type || '').toLowerCase()
  if (st === 'title')   return 'title'
  if (st === 'divider') return 'divider'
  return 'content'
}

function inferVisualType(slide, slideType) {
  if (slideType === 'title')   return 'title_slide'
  if (slideType === 'divider') return 'divider_slide'
  const st = (slide.section_type || '').toLowerCase()
  if (st === 'financial_data')     return 'data_table'
  if (st === 'executive_summary')  return 'stat_callout'
  if (st === 'market_analysis')    return 'bullet_list'
  if (st === 'recommendations')    return 'three_column'
  if (st === 'conclusion')         return 'process_flow'
  return 'bullet_list'
}

function normaliseContent(content, visualType, slide) {
  if (content && typeof content === 'object' && Object.keys(content).length > 0) {
    return content
  }

  // Build minimal valid content if missing
  const vt = (visualType || '').toLowerCase()

  if (vt === 'title_slide')   return { title: slide.title || '', subtitle: slide.subtitle || '', date: '' }
  if (vt === 'divider_slide') return { section_name: slide.title || '', section_descriptor: '' }
  if (vt === 'bullet_list')   return { bullets: ['Key point from this slide'] }
  if (vt === 'three_column')  return { columns: [{ header: 'Point 1', body: '' }, { header: 'Point 2', body: '' }, { header: 'Point 3', body: '' }] }
  if (vt === 'two_column')    return { left_header: 'Left', left_points: [], right_header: 'Right', right_points: [] }
  if (vt === 'data_table')    return { headers: ['Item', 'Value'], rows: [] }
  if (vt === 'stat_callout')  return { stats: [{ value: '—', label: 'Metric', change: '' }] }
  if (vt === 'chart_bar')     return { chart_title: '', x_label: '', y_label: '', series: [] }
  if (vt === 'chart_line')    return { chart_title: '', x_label: '', y_label: '', series: [] }
  if (vt === 'chart_waterfall') return { chart_title: '', items: [] }
  if (vt === 'quote_callout') return { quote: slide.key_message || '', attribution: '' }
  if (vt === 'icon_cards')    return { cards: [] }
  if (vt === 'process_flow')  return { steps: [] }

  return {}
}


// ─── FALLBACK MANIFEST ────────────────────────────────────────────────────────
function buildFallbackManifest(brief, slideCount) {
  const sections = brief.sections || []
  const manifest = []
  let slideNum   = 1

  for (const section of sections) {
    const count = section.suggested_slide_count || 1

    for (let i = 0; i < count; i++) {
      const isFirst   = slideNum === 1
      const isDivider = i === 0 && section.section_type === 'divider'

      manifest.push({
        slide_number: slideNum,
        section_name: section.section_name,
        section_type: section.section_type,
        slide_type:   isFirst ? 'title' : isDivider ? 'divider' : 'content',
        title:        isFirst ? (brief.governing_thought || 'Presentation') : section.section_name,
        subtitle:     isFirst ? (brief.audience || '') : '',
        key_message:  (section.so_what || section.purpose || ''),
        visual_type:  isFirst ? 'title_slide' : isDivider ? 'divider_slide' : inferVisualType({}, 'content'),
        content:      {},
        speaker_note: section.purpose || ''
      })
      slideNum++
      if (slideNum > slideCount) break
    }
    if (slideNum > slideCount) break
  }

  return manifest
}
