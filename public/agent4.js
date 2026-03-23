// ─── AGENT 4 — SLIDE CONTENT WRITER ──────────────────────────────────────────
// Input:  state.outline    — presentationBrief from Agent 3
//         state.contentB64 — original PDF (fallback for detail)
// Output: slideManifest    — flat JSON array, one object per slide
//
// Each slide object has:
//   slide_number, section_name, section_type, slide_type
//   title, subtitle, key_message
//   is_mixed         — true when slide has both data and narrative
//   primary_content  — main visual element (chart/stat/table/bullets/cards etc.)
//   secondary_content— supporting element when is_mixed = true (bullets/insight_box)
//   speaker_note

// ═══════════════════════════════════════════════════════════════════════════════
// SYSTEM PROMPT
// ═══════════════════════════════════════════════════════════════════════════════

const AGENT4_SYSTEM = `You are a senior management consultant creating slide content for a board-level presentation.

You will receive a presentation brief and a batch of slides to populate.
The source document is also attached for reference when the brief lacks detail.

Return a JSON array — one object per slide — with EXACTLY these fields:

{
  "slide_number": number,
  "section_name": "string",
  "section_type": "string",
  "slide_type": "title" | "divider" | "content",
  "title": "string",
  "subtitle": "string",
  "key_message": "string",
  "is_mixed": true | false,
  "primary_content": { ... },
  "secondary_content": null or { ... },
  "speaker_note": "string"
}

════════════════════════════════════════
TITLE RULES
════════════════════════════════════════

Title slides:
- title: SHORT presentation name, 5-8 words max. Example: "Retail Portfolio Analytics — Q3 2025"
- key_message: the governing thought (the single most important insight for the whole presentation)
- is_mixed: false
- primary_content type: "title_slide"
- secondary_content: null

Divider slides:
- title: section name only
- is_mixed: false
- primary_content type: "divider_slide"
- secondary_content: null

════════════════════════════════════════
MIXED CONTENT RULES
════════════════════════════════════════

Set is_mixed: true when the slide has BOTH numbers (charts/stats) AND narrative insight that together tell a stronger story.
Set is_mixed: false for: title slides, divider slides, data_table, three_column, process_flow (these need full space).

When is_mixed: true  → populate BOTH primary_content AND secondary_content
When is_mixed: false → populate primary_content only, set secondary_content: null

════════════════════════════════════════
PRIMARY CONTENT TYPES
════════════════════════════════════════

── title_slide ──
{
  "type": "title_slide",
  "title": "short presentation name",
  "subtitle": "audience or context",
  "date": "Month Year"
}

── divider_slide ──
{
  "type": "divider_slide",
  "section_name": "section title",
  "section_descriptor": "one line describing what this section covers"
}

── chart ──
{
  "type": "chart",
  "chart_type": "bar" | "line" | "pie" | "waterfall" | "clustered_bar",
  "chart_decision": "one line explaining why this chart type was chosen",
  "chart_title": "descriptive chart title",
  "x_label": "string or empty",
  "y_label": "string or empty",
  "categories": ["string", "string"],
  "series": [
    {
      "name": "series name",
      "values": [number, number],
      "types": ["positive" | "negative" | "total"]
    }
  ],
  "show_data_labels": true,
  "show_legend": true | false
}

── stat_callout ──
{
  "type": "stat_callout",
  "stats": [
    {
      "value": "₹351.11 Cr",
      "label": "Principal Outstanding",
      "change": "76.7% of disbursed",
      "sentiment": "positive" | "negative" | "neutral"
    }
  ]
}
Note: Use "negative" sentiment for bad/risk metrics. Max 4 stats.

── data_table ──
{
  "type": "data_table",
  "headers": ["Col1", "Col2", "Col3"],
  "rows": [["val", "val", "val"]],
  "highlight_rows": [0]
}

── three_column ──
{
  "type": "three_column",
  "columns": [
    { "header": "string", "body": "2-3 sentence description", "sentiment": "positive" | "negative" | "neutral" }
  ]
}

── two_column ──
{
  "type": "two_column",
  "left_header": "string",
  "left_points": ["point 1", "point 2"],
  "right_header": "string",
  "right_points": ["point 1", "point 2"]
}

── cards ──
{
  "type": "cards",
  "cards": [
    { "header": "string", "body": "2-3 sentences", "sentiment": "positive" | "negative" | "neutral" }
  ]
}
Note: Use for 3-4 parallel recommendations or strategic initiatives.

── process_flow ──
{
  "type": "process_flow",
  "steps": [
    { "step_number": 1, "title": "string", "description": "string" }
  ]
}

── bullets ──
{
  "type": "bullets",
  "bullets": [
    {
      "text": "full bullet text with specific data",
      "emphasis": [
        { "text": "exact substring to emphasise", "style": "bold", "color": "accent" | "positive" | "warning" | null }
      ],
      "sentiment": "positive" | "warning" | "neutral"
    }
  ]
}

════════════════════════════════════════
SECONDARY CONTENT TYPES
════════════════════════════════════════

── bullets (same structure as primary bullets above) ──

── insight_box ──
{
  "type": "insight_box",
  "heading": "Key Insight" | "So What" | "Risk Alert" | "Action Required",
  "text": "1-2 specific sentences with data",
  "sentiment": "positive" | "warning" | "neutral"
}

── stat_callout (1-2 stats only, for secondary position) ──

════════════════════════════════════════
CHART TYPE DECISION RULES
════════════════════════════════════════

bar:          comparison across 3+ categories, one metric, no time dimension
line:         trend over time — labels are months / quarters / years
pie:          part of a whole — values sum to ~100%, max 5 segments
waterfall:    bridge / variance — items have types: positive / negative / total
clustered_bar: two series compared across same categories (e.g. actual vs target)

════════════════════════════════════════
BULLET EMPHASIS RULES
════════════════════════════════════════

- Numbers and percentages: ALWAYS bold, no color change
- Risk words (risk, decline, fall, breach, exceed, critical, dangerous, concentration): bold + accent color
- Positive words (grew, growth, strong, record, exceeded target, improved): bold + positive color
- Key entity names (zone names, product names): bold only
- Max ONE emphasis per bullet — the single most important element only

════════════════════════════════════════
CRITICAL QUALITY RULES
════════════════════════════════════════

1. ZERO placeholder text — "Key point from this slide", "TBD", "Insert data" = FAILURE
2. Every number must come from the source document — no invented figures
3. Title slides: title is SHORT (5-8 words), not the governing thought
4. key_message must be an INSIGHT not a description — state the so-what
5. Titles of content slides must be insight-led: "Revenue grew 18%" not "Revenue Analysis"
6. Return ONLY a valid JSON array. No explanation. No markdown fences.`


// ═══════════════════════════════════════════════════════════════════════════════
// CHART TYPE AUTO-SELECTOR (fallback if Claude doesn't decide)
// ═══════════════════════════════════════════════════════════════════════════════

function selectChartType(pc, sectionType, slideTitle) {
  // Already decided
  if (pc && pc.chart_type) return pc.chart_type

  const text = ((slideTitle || '') + ' ' + (sectionType || '')).toLowerCase()

  // Waterfall — items have types array
  if (pc && pc.series && pc.series[0] && pc.series[0].types) return 'waterfall'

  // Two series → clustered bar
  if (pc && pc.series && pc.series.length >= 2) return 'clustered_bar'

  const cats = (pc && pc.categories) ? pc.categories : []

  // Time series
  const timeLabels = ['q1','q2','q3','q4','jan','feb','mar','apr','may','jun','jul','aug','sep','oct','nov','dec','fy','h1','h2','2020','2021','2022','2023','2024','2025','2026']
  if (cats.some(c => timeLabels.some(t => String(c).toLowerCase().includes(t)))) return 'line'

  // Pie — distribution / composition / share
  if (/distribution|breakdown|mix|composition|share|proportion/.test(text)) {
    const vals = pc && pc.series && pc.series[0] ? pc.series[0].values || [] : []
    const sum  = vals.reduce((a, b) => a + b, 0)
    if (sum > 80 && sum < 120 && cats.length <= 5) return 'pie'
  }

  return 'bar'
}


// ═══════════════════════════════════════════════════════════════════════════════
// MIXED CONTENT DECISION (whether a slide should be mixed)
// ═══════════════════════════════════════════════════════════════════════════════

function shouldBeMixed(sectionType, slideType, visualType) {
  if (slideType === 'title' || slideType === 'divider') return false
  if (['data_table', 'three_column', 'process_flow', 'two_column'].includes(visualType)) return false
  if (sectionType === 'executive_summary') return false
  return ['financial_data', 'strategic_analysis', 'market_analysis', 'operational_review', 'recommendations'].includes(sectionType)
}


// ═══════════════════════════════════════════════════════════════════════════════
// VISUAL TYPE INFERENCE (fallback when not specified)
// ═══════════════════════════════════════════════════════════════════════════════

function inferVisualType(sectionType, slideIndex, totalInSection) {
  switch (sectionType) {
    case 'financial_data':     return slideIndex === 0 ? 'stat_callout' : 'chart'
    case 'executive_summary':  return 'stat_callout'
    case 'strategic_analysis': return slideIndex % 2 === 0 ? 'chart' : 'bullets'
    case 'market_analysis':    return 'chart'
    case 'recommendations':    return totalInSection > 1 ? 'three_column' : 'cards'
    case 'conclusion':         return 'process_flow'
    case 'operational_review': return 'data_table'
    default:                   return 'bullets'
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// PLACEHOLDER DETECTOR
// ═══════════════════════════════════════════════════════════════════════════════

const PLACEHOLDER_PATTERNS = [
  /key point from this slide/i,
  /insert data here/i,
  /\btbd\b/i,
  /^placeholder$/i,
  /sample text/i,
  /click to edit/i,
  /^\s*-\s*$/
]

function isPlaceholder(text) {
  if (!text || typeof text !== 'string') return true
  if (text.trim().length < 5) return true
  return PLACEHOLDER_PATTERNS.some(p => p.test(text.trim()))
}

function hasPlaceholderContent(slide) {
  const pc = slide.primary_content || {}

  // Check bullets
  if (pc.type === 'bullets') {
    const bullets = pc.bullets || []
    if (bullets.length === 0) return true
    const texts = bullets.map(b => typeof b === 'string' ? b : (b.text || ''))
    if (texts.every(t => isPlaceholder(t))) return true
  }

  // Check stats
  if (pc.type === 'stat_callout') {
    const stats = pc.stats || []
    if (stats.length === 0) return true
    if (stats.every(s => isPlaceholder(s.value) || s.value === '—')) return true
  }

  // Check chart has data
  if (pc.type === 'chart') {
    if (!pc.categories || pc.categories.length === 0) return true
    if (!pc.series || pc.series.length === 0) return true
    if (pc.series.every(s => !s.values || s.values.length === 0)) return true
  }

  // Check cards
  if (pc.type === 'cards') {
    const cards = pc.cards || []
    if (cards.length === 0) return true
    if (cards.every(c => isPlaceholder(c.body || '') && isPlaceholder(c.header || ''))) return true
  }

  // Check key message
  if (isPlaceholder(slide.key_message)) return true

  return false
}


// ═══════════════════════════════════════════════════════════════════════════════
// NORMALISE A SINGLE SLIDE (clean up structure from Claude's response)
// ═══════════════════════════════════════════════════════════════════════════════

function normaliseSlide(slide, plan) {
  const slideType = slide.slide_type || plan.slide_type || 'content'
  const isMixed   = slide.is_mixed !== undefined ? !!slide.is_mixed : plan.suggested_mixed || false

  // Get primary content — handle both new format and legacy format
  let pc = slide.primary_content || null

  if (!pc && slide.content) {
    // Legacy format — Agent returned old-style content object
    pc = normaliseLegacyContent(slide.content, slide.visual_type || plan.suggested_visual_type)
  }

  if (!pc) {
    pc = buildEmptyContent(plan.suggested_visual_type, slideType, plan)
  }

  // Ensure type field always exists
  if (!pc.type) {
    pc.type = plan.suggested_visual_type || 'bullets'
  }

  // Auto-select chart type if missing
  if (pc.type === 'chart' && !pc.chart_type) {
    pc.chart_type     = selectChartType(pc, plan.section_type, slide.title || '')
    pc.chart_decision = pc.chart_decision || 'auto-selected based on data structure'
  }

  // Ensure chart_title is populated
  if (pc.type === 'chart' && !pc.chart_title) {
    pc.chart_title = slide.title || plan.section_name || ''
  }

  // Secondary content — only if mixed
  let sc = null
  if (isMixed) {
    sc = slide.secondary_content || null
    // If mixed but no secondary provided, build a minimal insight box
    if (!sc && slide.key_message && !isPlaceholder(slide.key_message)) {
      sc = {
        type:      'insight_box',
        heading:   'Key Insight',
        text:      slide.key_message,
        sentiment: 'neutral'
      }
    }
  }

  return {
    slide_number:      slide.slide_number      || plan.slide_number,
    section_name:      slide.section_name      || plan.section_name   || '',
    section_type:      slide.section_type      || plan.section_type   || '',
    slide_type:        slideType,
    title:             slide.title             || plan.section_name   || ('Slide ' + plan.slide_number),
    subtitle:          slide.subtitle          || '',
    key_message:       slide.key_message       || plan.so_what        || '',
    is_mixed:          isMixed,
    primary_content:   pc,
    secondary_content: sc,
    speaker_note:      slide.speaker_note      || plan.purpose        || ''
  }
}

function normaliseLegacyContent(content, visualType) {
  if (!content || typeof content !== 'object') return null
  const vt = (visualType || 'bullets')
    .replace('bullet_list', 'bullets')
    .replace('chart_bar',   'chart')
    .replace('chart_line',  'chart')
    .replace('chart_waterfall', 'chart')
    .replace('stat_callout', 'stat_callout')
  return { type: vt, ...content }
}

function buildEmptyContent(visualType, slideType, plan) {
  if (slideType === 'title')   return { type: 'title_slide',   title: plan.section_name || '', subtitle: '', date: '' }
  if (slideType === 'divider') return { type: 'divider_slide', section_name: plan.section_name || '', section_descriptor: plan.purpose || '' }

  switch (visualType) {
    case 'chart':        return { type: 'chart',        chart_type: 'bar', chart_title: '', categories: [], series: [] }
    case 'stat_callout': return { type: 'stat_callout', stats: [] }
    case 'data_table':   return { type: 'data_table',   headers: [], rows: [] }
    case 'three_column': return { type: 'three_column', columns: [] }
    case 'cards':        return { type: 'cards',        cards: [] }
    case 'process_flow': return { type: 'process_flow', steps: [] }
    default:             return { type: 'bullets',      bullets: [] }
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE PLAN BUILDER
// ═══════════════════════════════════════════════════════════════════════════════

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

      plan.push({
        slide_number:          slideNum,
        section_name:          section.section_name  || '',
        section_type:          section.section_type  || '',
        slide_type:            slideType,
        suggested_visual_type: visualType,
        suggested_mixed:       shouldBeMixed(section.section_type, slideType, visualType),
        purpose:               section.purpose       || '',
        key_content:           section.key_content   || [],
        so_what:               section.so_what       || '',
        data_available:        section.data_available || false
      })

      slideNum++
    }

    if (slideNum > slideCount) break
  }

  return plan
}


// ═══════════════════════════════════════════════════════════════════════════════
// BATCH WRITER
// ═══════════════════════════════════════════════════════════════════════════════

async function writeSlideBatch(batchPlan, brief, contentB64, batchNum) {
  console.log('Agent 4 — batch', batchNum, ': slides', batchPlan[0].slide_number, '–', batchPlan[batchPlan.length-1].slide_number)

  const prompt = `PRESENTATION BRIEF:
Document type:    ${brief.document_type || '—'}
Governing thought: ${brief.governing_thought || '—'}
Narrative flow:   ${brief.narrative_flow || '—'}
Tone:             ${brief.tone || 'professional'}
Key messages:     ${(brief.key_messages || []).join(' | ')}
Key data points:  ${(brief.key_data_points || []).join(' | ')}
Recommendations:  ${(brief.recommendations || []).join(' | ')}

SLIDES TO WRITE — batch ${batchNum} (${batchPlan.length} slides):
${JSON.stringify(batchPlan, null, 2)}

INSTRUCTIONS:
- Write full content for each slide using the brief AND the attached source document
- Where suggested_mixed is true: populate BOTH primary_content AND secondary_content
- Where suggested_mixed is false: primary_content only, secondary_content: null
- Select the best chart type using the decision rules
- Apply bullet emphasis rules — bold numbers, flag risks with accent color
- Title slide: SHORT name (5-8 words max), not the governing thought
- Use REAL data from the source document — actual numbers, percentages, names
- ZERO placeholder content — every field must be specific and meaningful
- Return ONLY a JSON array for these ${batchPlan.length} slides`

  const messages = [{
    role: 'user',
    content: [
      { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: contentB64 } },
      { type: 'text', text: prompt }
    ]
  }]

  const raw    = await callClaude(AGENT4_SYSTEM, messages, 4000)
  const parsed = safeParseJSON(raw, null)

  if (!Array.isArray(parsed)) {
    console.warn('Agent 4 batch', batchNum, '— parse failed, raw length:', raw.length)
    return null
  }

  console.log('Agent 4 batch', batchNum, '— got', parsed.length, 'slides')
  return parsed
}


// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE REPAIR (for slides with placeholder content)
// ═══════════════════════════════════════════════════════════════════════════════

async function repairSlide(slide, brief, contentB64) {
  console.log('Agent 4 — repairing slide', slide.slide_number, ':', slide.title)

  const prompt = `This slide has placeholder or missing content. Fix it with specific data from the source document.

CONTEXT:
Document type:   ${brief.document_type || '—'}
Key messages:    ${(brief.key_messages || []).join(' | ')}
Key data:        ${(brief.key_data_points || []).join(' | ')}

SLIDE TO FIX:
${JSON.stringify(slide, null, 2)}

RULES:
- Replace ALL placeholder text with specific, data-driven content
- For title slides: SHORT name (5-8 words max)
- Numbers must come from the source document
- Return ONLY a single JSON object for this one slide. No array. No explanation.`

  const messages = [{
    role: 'user',
    content: [
      { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: contentB64 } },
      { type: 'text', text: prompt }
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


// ═══════════════════════════════════════════════════════════════════════════════
// MAIN RUNNER
// ═══════════════════════════════════════════════════════════════════════════════

async function runAgent4(state) {
  const brief      = state.outline
  const contentB64 = state.contentB64
  const slideCount = (brief && brief.total_slides) || state.slideCount || 12

  console.log('Agent 4 starting')
  console.log('  Target slides:', slideCount)
  console.log('  Document type:', (brief && brief.document_type) || '—')

  // Step 1 — Build slide plan from brief
  const slidePlan = buildSlidePlan(brief, slideCount)
  const mixedCount = slidePlan.filter(s => s.suggested_mixed).length
  console.log('  Slide plan:', slidePlan.length, 'slides,', mixedCount, 'suggested as mixed')

  // Step 2 — Split into batches of 5
  const BATCH_SIZE = 5
  const batches    = []
  for (let i = 0; i < slidePlan.length; i += BATCH_SIZE) {
    batches.push(slidePlan.slice(i, i + BATCH_SIZE))
  }
  console.log('  Batches:', batches.length)

  // Step 3 — Write each batch
  let allSlides = []

  for (let b = 0; b < batches.length; b++) {
    const batch  = batches[b]
    const result = await writeSlideBatch(batch, brief, contentB64, b + 1)

    if (!result) {
      // Batch failed — create empty shells
      console.warn('Agent 4 — batch', b+1, 'failed, creating shells')
      batch.forEach(plan => allSlides.push(normaliseSlide({}, plan)))
    } else {
      // Merge results with plan
      batch.forEach((plan, idx) => {
        const match = result[idx] || result.find(s => s.slide_number === plan.slide_number)
        allSlides.push(normaliseSlide(match || {}, plan))
      })
    }
  }

  // Step 4 — Validate and repair placeholder slides
  const failed = allSlides.filter(s => s.slide_type === 'content' && hasPlaceholderContent(s))
  console.log('Agent 4 — slides needing repair:', failed.length)

  for (const slide of failed.slice(0, 5)) {
    const repaired = await repairSlide(slide, brief, contentB64)
    if (repaired) {
      const idx = allSlides.findIndex(s => s.slide_number === slide.slide_number)
      if (idx >= 0) {
        const plan = slidePlan.find(p => p.slide_number === slide.slide_number) || {}
        allSlides[idx] = normaliseSlide(repaired, plan)
        console.log('  Repaired slide', slide.slide_number)
      }
    }
  }

  // Step 5 — Final quality log
  const finalFailed = allSlides.filter(s => s.slide_type === 'content' && hasPlaceholderContent(s))
  if (finalFailed.length > 0) {
    console.warn('Agent 4 — slides still with placeholder content:', finalFailed.map(s => s.slide_number).join(', '))
  }

  // Log visual type breakdown
  const vtBreakdown  = {}
  const finalMixed   = allSlides.filter(s => s.is_mixed).length
  const chartTypes   = {}

  allSlides.forEach(s => {
    const ptype = (s.primary_content || {}).type || 'unknown'
    vtBreakdown[ptype] = (vtBreakdown[ptype] || 0) + 1
    if (ptype === 'chart') {
      const ct = (s.primary_content || {}).chart_type || 'unknown'
      chartTypes[ct] = (chartTypes[ct] || 0) + 1
    }
  })

  console.log('Agent 4 complete')
  console.log('  Total slides:', allSlides.length)
  console.log('  Mixed slides:', finalMixed)
  console.log('  Content types:', JSON.stringify(vtBreakdown))
  console.log('  Chart types:', JSON.stringify(chartTypes))
  console.log('  Placeholder slides remaining:', finalFailed.length)

  return allSlides
}
