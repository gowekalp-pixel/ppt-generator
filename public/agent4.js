// ─── AGENT 4 — SLIDE CONTENT WRITER ──────────────────────────────────────────
// Input:  state.outline        — presentationBrief from Agent 3
//         state.contentB64     — original PDF
// Output: slideManifest        — flat JSON array
//
// Key responsibilities:
// 1. Write specific, data-driven content for every slide
// 2. Decide primary + secondary content for mixed slides
// 3. Select the best chart type based on data signals
// 4. Format bullets with emphasis markers for important numbers/words
// 5. Split into batches to avoid token limits

const AGENT4_SYSTEM = `You are a senior management consultant creating slide content for a board presentation.

You will receive a presentation brief and a batch of slides to populate.

Return a JSON array — one object per slide — with EXACTLY these fields:

{
  "slide_number": number,
  "section_name": "string",
  "section_type": "string",
  "slide_type": "title" | "divider" | "content",
  "title": "string — insight-led, specific. For title slides: short name 5-8 words max",
  "subtitle": "string — for title slides only",
  "key_message": "string — the single most important insight. Specific, data-driven.",
  "is_mixed": true | false,
  "primary_content": { ... see rules below ... },
  "secondary_content": null OR { ... see rules below ... },
  "speaker_note": "string"
}

═══ MIXED CONTENT RULES ═══

Set is_mixed: true when the slide has BOTH numbers AND narrative insight that together tell a stronger story.
Set is_mixed: false for: title slides, divider slides, data_table, three_column, process_flow (these need full space).

When is_mixed: true — populate BOTH primary_content and secondary_content.
When is_mixed: false — populate primary_content only, set secondary_content: null.

═══ PRIMARY CONTENT TYPES ═══

type: "chart"
{
  "type": "chart",
  "chart_type": "bar" | "line" | "pie" | "waterfall" | "clustered_bar" | "combo",
  "chart_decision": "string — why this chart type was chosen",
  "chart_title": "string",
  "x_label": "string",
  "y_label": "string",
  "categories": ["string"],
  "series": [{ "name": "string", "values": [number], "types": ["positive"|"negative"|"total"] }],
  "show_data_labels": true | false,
  "show_legend": true | false
}

type: "stat_callout"
{
  "type": "stat_callout",
  "stats": [{ "value": "string", "label": "string", "change": "string", "sentiment": "positive"|"negative"|"neutral" }]
}

type: "data_table"
{
  "type": "data_table",
  "headers": ["string"],
  "rows": [["string"]],
  "highlight_rows": [number]
}

type: "three_column"
{
  "type": "three_column",
  "columns": [{ "header": "string", "body": "string", "sentiment": "positive"|"negative"|"neutral" }]
}

type: "two_column"
{
  "type": "two_column",
  "left_header": "string",
  "left_points": ["string"],
  "right_header": "string",
  "right_points": ["string"]
}

type: "cards"
{
  "type": "cards",
  "cards": [{ "header": "string", "body": "string", "sentiment": "positive"|"negative"|"neutral" }]
}

type: "process_flow"
{
  "type": "process_flow",
  "steps": [{ "step_number": number, "title": "string", "description": "string" }]
}

type: "title_slide"
{
  "type": "title_slide",
  "title": "string — short presentation name",
  "subtitle": "string",
  "date": ""
}

type: "divider_slide"
{
  "type": "divider_slide",
  "section_name": "string",
  "section_descriptor": "string — one line on what this section covers"
}

═══ SECONDARY CONTENT TYPES ═══

type: "bullets"
{
  "type": "bullets",
  "bullets": [
    {
      "text": "string — full bullet text",
      "emphasis": [{ "text": "string — exact substring to emphasise", "style": "bold", "color": "accent"|"positive"|"warning"|null }],
      "sentiment": "positive"|"warning"|"neutral"
    }
  ]
}

type: "insight_box"
{
  "type": "insight_box",
  "heading": "string — short label like 'Key Insight' or 'So What'",
  "text": "string — 1-2 sentences",
  "sentiment": "positive"|"warning"|"neutral"
}

type: "stat_callout" — same structure as primary stat_callout but smaller (1-2 stats only)

═══ CHART TYPE DECISION RULES ═══

bar:           comparison across 3+ categories, one metric, no time dimension
line:          trend over time — labels are months/quarters/years
pie:           part of a whole — values sum to ~100%, max 5 segments
waterfall:     bridge/variance — items have types: positive/negative/total
clustered_bar: two series compared across same categories (actual vs target)
combo:         two metrics with different scales on same categories

═══ BULLET EMPHASIS RULES ═══

- Numbers and percentages: ALWAYS bold + no color change (e.g. "18%" → bold)
- Risk/warning signals (risk, decline, fall, breach, exceed, critical, concern): bold + accent color
- Positive signals (grew, growth, strong, record, exceeded): bold + positive color
- Key entity names (zone names, product names, company names): bold only
- Max ONE emphasis per bullet — the single most important element only
- Never emphasise generic words like "the", "and", "is"

═══ CRITICAL RULES ═══

1. ZERO placeholder text — every bullet, stat, number must be specific and real
2. Title slide: title is SHORT (5-8 words) — NOT the governing thought
3. Numbers in slides must match the source document
4. key_message must be an INSIGHT not a description
5. Return ONLY valid JSON array. No explanation. No markdown fences.`


// ─── CHART TYPE SELECTOR ─────────────────────────────────────────────────────
function selectChartType(content, sectionType, slideTitle) {
  const text = (slideTitle + ' ' + sectionType).toLowerCase()

  // Check for waterfall signals
  if (content && content.items && content.items.some(i => i.type)) return 'waterfall'

  // Check series count
  const seriesCount = (content && content.series) ? content.series.length : 0
  if (seriesCount >= 2) return 'clustered_bar'

  // Check categories
  const categories = content && content.categories ? content.categories : []

  // Time series detection
  const timeLabels = ['q1','q2','q3','q4','jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec','fy','h1','h2','2020','2021','2022','2023','2024','2025']
  const hasTimeLabels = categories.some(c => timeLabels.some(t => String(c).toLowerCase().includes(t)))
  if (hasTimeLabels) return 'line'

  // Percentage/composition detection
  if (text.includes('distribution') || text.includes('breakdown') || text.includes('mix') || text.includes('composition') || text.includes('share')) {
    const values = content && content.series && content.series[0] ? content.series[0].values || [] : []
    const sum = values.reduce((a, b) => a + b, 0)
    if (sum > 90 && sum < 110 && categories.length <= 5) return 'pie'
  }

  // Default comparison
  return 'bar'
}


// ─── MIXED CONTENT DETECTOR ──────────────────────────────────────────────────
function shouldBeMixed(sectionType, slideType, visualType) {
  if (slideType === 'title' || slideType === 'divider') return false
  if (['data_table', 'three_column', 'process_flow', 'two_column'].includes(visualType)) return false
  if (['executive_summary'].includes(sectionType)) return false

  // These benefit from mixed treatment
  if (['financial_data', 'strategic_analysis', 'market_analysis', 'operational_review'].includes(sectionType)) return true
  if (['recommendations'].includes(sectionType)) return true

  return false
}


// ─── PLACEHOLDER DETECTOR ────────────────────────────────────────────────────
const PLACEHOLDER_PATTERNS = [
  /key point from this slide/i,
  /insert data here/i,
  /\btbd\b/i,
  /placeholder/i,
  /sample text/i,
  /^\s*-\s*$/,
  /click to edit/i
]

function isPlaceholder(text) {
  if (!text || typeof text !== 'string') return true
  if (text.trim().length < 5) return true
  return PLACEHOLDER_PATTERNS.some(p => p.test(text))
}

function hasPlaceholderContent(slide) {
  const pc = slide.primary_content || {}

  if (pc.type === 'bullets' || pc.bullets) {
    const bullets = pc.bullets || []
    if (bullets.length === 0) return true
    if (bullets.every(b => isPlaceholder(typeof b === 'string' ? b : b.text))) return true
  }

  if (pc.stats) {
    if (pc.stats.every(s => isPlaceholder(s.value) || s.value === '—')) return true
  }

  if (pc.series) {
    if (pc.series.length === 0) return true
  }

  if (isPlaceholder(slide.key_message)) return true

  return false
}


// ─── SLIDE PLAN BUILDER ──────────────────────────────────────────────────────
function buildSlidePlan(brief, slideCount) {
  const sections = brief.sections || []
  const plan     = []
  let   slideNum = 1

  for (const section of sections) {
    const count = Math.max(1, section.suggested_slide_count || 1)

    for (let i = 0; i < count; i++) {
      if (slideNum > slideCount) break

      let slideType  = 'content'
      let visualType = inferVisualType(section.section_type, i, count)

      if (slideNum === 1) {
        slideType  = 'title'
        visualType = 'title_slide'
      } else if (section.section_type === 'divider') {
        slideType  = 'divider'
        visualType = 'divider_slide'
      }

      const mixed = shouldBeMixed(section.section_type, slideType, visualType)

      plan.push({
        slide_number:          slideNum,
        section_name:          section.section_name,
        section_type:          section.section_type,
        slide_type:            slideType,
        suggested_visual_type: visualType,
        suggested_mixed:       mixed,
        purpose:               section.purpose || '',
        key_content:           section.key_content || [],
        so_what:               section.so_what || '',
        data_available:        section.data_available || false
      })

      slideNum++
    }
    if (slideNum > slideCount) break
  }

  return plan
}

function inferVisualType(sectionType, slideIndex, totalInSection) {
  switch (sectionType) {
    case 'financial_data':       return slideIndex === 0 ? 'stat_callout' : 'chart'
    case 'executive_summary':    return 'stat_callout'
    case 'strategic_analysis':   return slideIndex % 2 === 0 ? 'chart' : 'bullets'
    case 'market_analysis':      return 'chart'
    case 'recommendations':      return totalInSection > 1 ? 'three_column' : 'cards'
    case 'conclusion':           return 'process_flow'
    case 'operational_review':   return 'data_table'
    default:                     return 'bullets'
  }
}


// ─── WRITE A BATCH OF SLIDES ─────────────────────────────────────────────────
async function writeSlideBatch(batchPlan, brief, contentB64, batchNum) {
  console.log('Agent 4 — batch', batchNum, ': slides', batchPlan[0].slide_number, '–', batchPlan[batchPlan.length-1].slide_number)

  const batchPrompt = `PRESENTATION BRIEF:
Document type: ${brief.document_type || '—'}
Governing thought: ${brief.governing_thought || '—'}
Narrative flow: ${brief.narrative_flow || '—'}
Tone: ${brief.tone || 'professional'}
Key messages: ${(brief.key_messages || []).join(' | ')}
Key data points: ${(brief.key_data_points || []).join(' | ')}
Recommendations: ${(brief.recommendations || []).join(' | ')}

SLIDES TO WRITE — batch ${batchNum} (${batchPlan.length} slides):
${JSON.stringify(batchPlan, null, 2)}

INSTRUCTIONS:
- Write full content for each slide using the brief and source document
- Where suggested_mixed is true: populate BOTH primary_content AND secondary_content
- Where suggested_mixed is false: primary_content only, secondary_content: null
- For chart types: use chart_decision rules to select the best format
- For bullets: apply emphasis rules — bold numbers, flag warnings with accent color
- Title slide: SHORT name (5-8 words max) not the governing thought
- Use REAL data from the source document — actual numbers, percentages, names
- ZERO placeholder content
- Return ONLY a JSON array for these ${batchPlan.length} slides`

  const messages = [{
    role: 'user',
    content: [
      { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: contentB64 } },
      { type: 'text', text: batchPrompt }
    ]
  }]

  const raw    = await callClaude(AGENT4_SYSTEM, messages, 4000)
  const parsed = safeParseJSON(raw, null)

  if (!Array.isArray(parsed)) {
    console.warn('Agent 4 batch', batchNum, '— parse failed')
    return null
  }

  console.log('Agent 4 batch', batchNum, '— got', parsed.length, 'slides')
  return parsed
}


// ─── REPAIR A FAILED SLIDE ───────────────────────────────────────────────────
async function repairSlide(slide, brief, contentB64) {
  console.log('Agent 4 — repairing slide', slide.slide_number)

  const repairPrompt = `Fix this slide — it has placeholder or missing content.

CONTEXT:
Document type: ${brief.document_type || '—'}
Key messages: ${(brief.key_messages || []).join(' | ')}
Key data: ${(brief.key_data_points || []).join(' | ')}

SLIDE TO FIX:
${JSON.stringify(slide, null, 2)}

Replace ALL placeholder content with specific data-driven content from the source document.
For title slides: SHORT name (5-8 words).
Return ONLY a single JSON object. No array. No explanation.`

  const messages = [{
    role: 'user',
    content: [
      { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: contentB64 } },
      { type: 'text', text: repairPrompt }
    ]
  }]

  const raw     = await callClaude(AGENT4_SYSTEM, messages, 1500)
  const cleaned = raw.replace(/```json\s*/g, '').replace(/```\s*/g, '').trim()

  try {
    const parsed = JSON.parse(cleaned)
    if (Array.isArray(parsed) && parsed.length > 0) return parsed[0]
    if (typeof parsed === 'object' && parsed.slide_number) return parsed
  } catch(e) {
    console.warn('Agent 4 repair — parse failed for slide', slide.slide_number)
  }

  return null
}


// ─── NORMALISE A SLIDE ───────────────────────────────────────────────────────
function normaliseSlide(slide, plan) {
  const slideType = slide.slide_type || plan.slide_type || 'content'
  const isMixed   = slide.is_mixed   !== undefined ? slide.is_mixed : plan.suggested_mixed || false

  // Normalise primary content
  let primaryContent = slide.primary_content || null
  if (!primaryContent) {
    // Legacy format fallback — Agent 4 might return old format
    if (slide.content) {
      primaryContent = normaliseLegacyContent(slide.content, slide.visual_type || plan.suggested_visual_type)
    } else {
      primaryContent = buildEmptyContent(plan.suggested_visual_type, slideType, plan)
    }
  }

  // Ensure chart type is decided
  if (primaryContent && primaryContent.type === 'chart' && !primaryContent.chart_type) {
    primaryContent.chart_type = selectChartType(primaryContent, plan.section_type, slide.title || '')
    primaryContent.chart_decision = primaryContent.chart_decision || 'defaulted to bar chart'
  }

  return {
    slide_number:      slide.slide_number      || plan.slide_number,
    section_name:      slide.section_name      || plan.section_name,
    section_type:      slide.section_type      || plan.section_type,
    slide_type:        slideType,
    title:             slide.title             || plan.section_name || 'Slide ' + plan.slide_number,
    subtitle:          slide.subtitle          || '',
    key_message:       slide.key_message       || plan.so_what || '',
    is_mixed:          isMixed,
    primary_content:   primaryContent,
    secondary_content: isMixed ? (slide.secondary_content || null) : null,
    speaker_note:      slide.speaker_note      || plan.purpose || ''
  }
}

function normaliseLegacyContent(content, visualType) {
  if (!content) return null
  const type = (visualType || 'bullets').replace('bullet_list', 'bullets').replace('chart_bar','chart').replace('chart_line','chart')
  return { type, ...content }
}

function buildEmptyContent(visualType, slideType, plan) {
  if (slideType === 'title')   return { type: 'title_slide',   title: plan.section_name || '', subtitle: '', date: '' }
  if (slideType === 'divider') return { type: 'divider_slide', section_name: plan.section_name || '', section_descriptor: plan.purpose || '' }

  switch (visualType) {
    case 'chart':        return { type: 'chart', chart_type: 'bar', categories: [], series: [] }
    case 'stat_callout': return { type: 'stat_callout', stats: [] }
    case 'data_table':   return { type: 'data_table', headers: [], rows: [] }
    case 'three_column': return { type: 'three_column', columns: [] }
    case 'cards':        return { type: 'cards', cards: [] }
    case 'process_flow': return { type: 'process_flow', steps: [] }
    default:             return { type: 'bullets', bullets: [] }
  }
}


// ─── MAIN RUNNER ─────────────────────────────────────────────────────────────
async function runAgent4(state) {
  const brief      = state.outline
  const contentB64 = state.contentB64
  const slideCount = brief.total_slides || state.slideCount || 12

  console.log('Agent 4 starting')
  console.log('  Target slides:', slideCount)
  console.log('  Document type:', brief.document_type || '—')

  // Step 1 — Build plan
  const slidePlan = buildSlidePlan(brief, slideCount)
  console.log('Agent 4 — plan:', slidePlan.length, 'slides,', slidePlan.filter(s => s.suggested_mixed).length, 'mixed')

  // Step 2 — Batch into groups of 5
  const BATCH_SIZE = 5
  const batches    = []
  for (let i = 0; i < slidePlan.length; i += BATCH_SIZE) {
    batches.push(slidePlan.slice(i, i + BATCH_SIZE))
  }

  // Step 3 — Write batches
  let allSlides = []

  for (let b = 0; b < batches.length; b++) {
    const batch       = batches[b]
    const batchResult = await writeSlideBatch(batch, brief, contentB64, b + 1)

    if (!batchResult) {
      batch.forEach(plan => allSlides.push(normaliseSlide({}, plan)))
    } else {
      batch.forEach((plan, idx) => {
        const result = batchResult[idx] || batchResult.find(s => s.slide_number === plan.slide_number)
        allSlides.push(normaliseSlide(result || {}, plan))
      })
    }
  }

  // Step 4 — Validate and repair
  const failedSlides = allSlides.filter(s => hasPlaceholderContent(s) && s.slide_type === 'content')
  console.log('Agent 4 — slides needing repair:', failedSlides.length)

  for (const slide of failedSlides.slice(0, 5)) {
    const repaired = await repairSlide(slide, brief, contentB64)
    if (repaired) {
      const idx = allSlides.findIndex(s => s.slide_number === slide.slide_number)
      if (idx >= 0) {
        allSlides[idx] = normaliseSlide(repaired, slidePlan.find(p => p.slide_number === slide.slide_number) || {})
        console.log('Agent 4 — repaired slide', slide.slide_number)
      }
    }
  }

  // Step 5 — Finalise chart types
  allSlides = allSlides.map(slide => {
    const pc = slide.primary_content
    if (pc && pc.type === 'chart' && !pc.chart_type) {
      pc.chart_type     = selectChartType(pc, slide.section_type, slide.title)
      pc.chart_decision = pc.chart_decision || 'auto-selected'
    }
    return slide
  })

  // Log
  const mixedCount = allSlides.filter(s => s.is_mixed).length
  const chartTypes = {}
  allSlides.forEach(s => {
    const pc = s.primary_content
    if (pc && pc.type === 'chart') chartTypes[pc.chart_type] = (chartTypes[pc.chart_type] || 0) + 1
  })

  console.log('Agent 4 complete')
  console.log('  Total slides:', allSlides.length)
  console.log('  Mixed slides:', mixedCount)
  console.log('  Chart types used:', JSON.stringify(chartTypes))

  return allSlides
}
