// ─── AGENT 4 — SLIDE CONTENT WRITER ──────────────────────────────────────────
// Input:  state.outline    — presentationBrief from Agent 3
//         state.contentB64 — original PDF
// Output: slideManifest    — flat JSON array, one object per slide
//
// Each slide has:
//   slide_number, section_name, section_type, slide_type
//   title, subtitle, key_message, speaker_note
//   zones[] — array of content zones with split hints and content
//
// Agent 4 decides WHAT to show and how many zones.
// Agent 5 decides WHERE everything goes (coordinates, title placement, brand).

// ═══════════════════════════════════════════════════════════════════════════════
// ZONE SPLIT REFERENCE
// ═══════════════════════════════════════════════════════════════════════════════
//
// Each zone has a "split" field that tells Agent 5 how to divide the content area.
//
// ONE ZONE:
//   "full"             — one zone fills entire content area
//
// TWO ZONES — HORIZONTAL (side by side):
//   "left_50"          — left half
//   "right_50"         — right half
//   "left_60"          — wider left
//   "right_40"         — narrower right
//   "left_40"          — narrower left
//   "right_60"         — wider right
//
// TWO ZONES — VERTICAL (top / bottom):
//   "top_30"           — narrow top row (good for stat callouts)
//   "bottom_70"        — main content below
//   "top_40"           — equal-ish top
//   "bottom_60"        — main content below
//   "top_50"           — top half
//   "bottom_50"        — bottom half
//
// THREE ZONES:
//   "top_left_50"      — top left quadrant
//   "top_right_50"     — top right quadrant
//   "bottom_full"      — full-width bottom
//   "left_full_50"     — left half full height
//   "top_right_50_h"   — top right (stacked)
//   "bottom_right_50_h"— bottom right (stacked)
//
// FOUR ZONES (2x2 grid):
//   "tl" "tr" "bl" "br" — top-left, top-right, bottom-left, bottom-right

// ═══════════════════════════════════════════════════════════════════════════════
// SYSTEM PROMPT
// ═══════════════════════════════════════════════════════════════════════════════

const AGENT4_SYSTEM = `You are a senior management consultant creating slide content for a board-level presentation.

You will receive a presentation brief and a batch of slides to populate.
The source document is attached for reference when the brief lacks detail.

Return a JSON array — one object per slide — with EXACTLY these fields:

{
  "slide_number": number,
  "section_name": "string",
  "section_type": "string",
  "slide_type": "title" | "divider" | "content",
  "title": "string — insight-led for content slides, short name for title slides",
  "subtitle": "string — for title slides only, empty for others",
  "key_message": "string — the single most important insight, specific and data-driven",
  "zones": [ ... zone objects ... ],
  "speaker_note": "string"
}

═══════════════════════════
SLIDE TYPE RULES
═══════════════════════════

Title slides:
- title: SHORT presentation name, 5-8 words max
- key_message: the governing thought (most important insight for the whole deck)
- zones: single zone with type "title_slide"

Divider slides:
- title: section name only
- zones: single zone with type "divider_slide"

Content slides:
- title: INSIGHT-LED — state the conclusion, not the topic
  WRONG: "Revenue Analysis"  RIGHT: "Revenue grew 18% driven by Product X"
  WRONG: "Geographic Risk"   RIGHT: "North Zone concentration at 66% exceeds safe limits"
- zones: 1 to 4 zones depending on what the slide needs

═══════════════════════════
ZONE STRUCTURE
═══════════════════════════

Each zone object:
{
  "zone_id": "z1",
  "label": "short description of what this zone shows",
  "split": "full" | "left_50" | "right_50" | "left_60" | "right_40" | "left_40" | "right_60" |
           "top_30" | "bottom_70" | "top_40" | "bottom_60" | "top_50" | "bottom_50" |
           "top_left_50" | "top_right_50" | "bottom_full" |
           "left_full_50" | "top_right_50_h" | "bottom_right_50_h" |
           "tl" | "tr" | "bl" | "br",
  "content": { ... content object ... }
}

SPLIT RULES:
- All splits in a slide must together cover exactly the full content area
- For 1 zone: use "full"
- For 2 side-by-side zones: pair "left_XX" with "right_YY" where XX+YY=100
- For 2 stacked zones: pair "top_XX" with "bottom_YY" where XX+YY=100
- For 3 zones (2 top + 1 bottom): use "top_left_50" + "top_right_50" + "bottom_full"
- For 3 zones (1 left + 2 stacked right): use "left_full_50" + "top_right_50_h" + "bottom_right_50_h"
- For 4 zones (2x2 grid): use "tl" + "tr" + "bl" + "br"

MINIMUM ZONE SIZE GUIDANCE (Agent 5 will enforce, but respect in planning):
- A chart needs at least 40% of height to be readable
- A data table needs at least 0.4" per row — do not use tables for more than 6 rows in a zone
- Stat callouts work well in top_30 (narrow top strip) or tl/tr quarters
- Bullets/text work well in right_40 or bottom_40
- Full-width chart works best in "full" or "bottom_70"

═══════════════════════════
CONTENT TYPES
═══════════════════════════

── title_slide ──
{ "type": "title_slide", "title": "short name", "subtitle": "audience/context", "date": "Month Year" }

── divider_slide ──
{ "type": "divider_slide", "section_name": "string", "section_descriptor": "one line" }

── chart ──
{
  "type": "chart",
  "chart_type": "bar" | "line" | "pie" | "waterfall" | "clustered_bar",
  "chart_decision": "why this chart type was chosen",
  "chart_title": "descriptive title",
  "x_label": "string",
  "y_label": "string",
  "categories": ["string"],
  "series": [{ "name": "string", "values": [number], "types": ["positive"|"negative"|"total"] }],
  "show_data_labels": true,
  "show_legend": true | false
}

Chart type selection:
- bar:           3+ categories, one metric, no time
- line:          trend over time (months, quarters, years)
- pie:           parts of a whole summing to ~100%, max 5 segments
- waterfall:     bridge/variance — items have types positive/negative/total
- clustered_bar: two series on same categories (actual vs target, period vs period)

── stat_callout ──
{
  "type": "stat_callout",
  "stats": [{ "value": "₹351Cr", "label": "string", "change": "string", "sentiment": "positive"|"negative"|"neutral" }]
}
Max 4 stats. Use "negative" for risk/bad metrics.
In a narrow zone (top_30, tl, tr): use max 2-3 stats.

── data_table ──
{
  "type": "data_table",
  "headers": ["Col1","Col2"],
  "rows": [["val","val"]],
  "highlight_rows": [0],
  "note": "optional table footnote"
}
Max 6 rows for readability. If more data needed, use two table zones.

── bullets ──
{
  "type": "bullets",
  "bullets": [
    {
      "text": "specific insight with data",
      "emphasis": [{ "text": "exact substring", "style": "bold", "color": "accent"|"positive"|"warning"|null }],
      "sentiment": "positive"|"warning"|"neutral"
    }
  ]
}
Max 5 bullets per zone. Each must be a complete sentence with specific data.
Emphasis rules:
- Numbers/percentages: always bold, no color
- Risk words (risk, decline, breach, critical, dangerous, concentration): bold + accent
- Positive words (grew, improved, record, exceeded): bold + positive
- Max ONE emphasis per bullet

── cards ──
{
  "type": "cards",
  "cards": [{ "header": "string", "body": "2-3 sentences", "sentiment": "positive"|"negative"|"neutral" }]
}
For 3-4 parallel items. In a full zone: up to 4 cards. In left/right zone: up to 2 cards.

── insight_box ──
{
  "type": "insight_box",
  "heading": "Key Insight"|"Risk Alert"|"So What"|"Action Required",
  "text": "1-2 specific sentences with data",
  "sentiment": "positive"|"warning"|"neutral"
}
Good for small zones (right_40, bottom_30, tr, br).

── two_column ──
{
  "type": "two_column",
  "left_header": "string",
  "left_points": ["string"],
  "right_header": "string",
  "right_points": ["string"]
}
Only use in a "full" zone. Do not combine with other zones.

── three_column ──
{
  "type": "three_column",
  "columns": [{ "header": "string", "body": "string", "sentiment": "positive"|"negative"|"neutral" }]
}
Only use in a "full" zone. Do not combine with other zones.

── process_flow ──
{
  "type": "process_flow",
  "steps": [{ "step_number": 1, "title": "string", "description": "string" }]
}
Only use in a "full" zone. Max 5 steps.

═══════════════════════════
ZONE COMBINATION EXAMPLES
═══════════════════════════

Example 1 — Chart with key insight (2 zones):
zones: [
  { zone_id: "z1", split: "left_60", content: { type: "chart", ... } },
  { zone_id: "z2", split: "right_40", content: { type: "insight_box", ... } }
]

Example 2 — Stats above chart (2 zones):
zones: [
  { zone_id: "z1", split: "top_30", content: { type: "stat_callout", stats: [2-3 stats] } },
  { zone_id: "z2", split: "bottom_70", content: { type: "chart", ... } }
]

Example 3 — Two comparison charts (2 zones):
zones: [
  { zone_id: "z1", split: "left_50", content: { type: "chart", chart_title: "North Zone", ... } },
  { zone_id: "z2", split: "right_50", content: { type: "chart", chart_title: "West Zone", ... } }
]

Example 4 — Three charts (3 zones):
zones: [
  { zone_id: "z1", split: "top_left_50", content: { type: "chart", ... } },
  { zone_id: "z2", split: "top_right_50", content: { type: "chart", ... } },
  { zone_id: "z3", split: "bottom_full", content: { type: "bullets", ... } }
]

Example 5 — Main chart left, table + insight right (3 zones):
zones: [
  { zone_id: "z1", split: "left_full_50", content: { type: "chart", ... } },
  { zone_id: "z2", split: "top_right_50_h", content: { type: "data_table", ... } },
  { zone_id: "z3", split: "bottom_right_50_h", content: { type: "insight_box", ... } }
]

Example 6 — 4 stat callouts in grid (4 zones):
zones: [
  { zone_id: "z1", split: "tl", content: { type: "stat_callout", stats: [1 stat] } },
  { zone_id: "z2", split: "tr", content: { type: "stat_callout", stats: [1 stat] } },
  { zone_id: "z3", split: "bl", content: { type: "stat_callout", stats: [1 stat] } },
  { zone_id: "z4", split: "br", content: { type: "stat_callout", stats: [1 stat] } }
]

═══════════════════════════
QUALITY RULES
═══════════════════════════

1. ZERO placeholder text — every value must be real and specific
2. Numbers must come from the source document — no invented figures
3. Key messages must state the insight, not describe the slide
4. Use multiple zones when data richness justifies it
5. Do NOT force multi-zone when a single chart or table tells the story clearly
6. Return ONLY valid JSON array. No explanation. No markdown fences.`


// ═══════════════════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════════════════

function inferDefaultZones(sectionType, slideType, slideIndex) {
  if (slideType === 'title')   return [{ zone_id: 'z1', label: 'title', split: 'full', content: { type: 'title_slide', title: '', subtitle: '', date: '' } }]
  if (slideType === 'divider') return [{ zone_id: 'z1', label: 'section', split: 'full', content: { type: 'divider_slide', section_name: '', section_descriptor: '' } }]

  switch (sectionType) {
    case 'financial_data':
      return slideIndex === 0
        ? [
            { zone_id: 'z1', label: 'key metrics', split: 'top_30', content: { type: 'stat_callout', stats: [] } },
            { zone_id: 'z2', label: 'trend chart', split: 'bottom_70', content: { type: 'chart', chart_type: 'bar', categories: [], series: [] } }
          ]
        : [
            { zone_id: 'z1', label: 'chart', split: 'left_60', content: { type: 'chart', chart_type: 'bar', categories: [], series: [] } },
            { zone_id: 'z2', label: 'insight', split: 'right_40', content: { type: 'insight_box', heading: 'Key Insight', text: '', sentiment: 'neutral' } }
          ]

    case 'executive_summary':
      return [
        { zone_id: 'z1', label: 'headline metrics', split: 'top_30', content: { type: 'stat_callout', stats: [] } },
        { zone_id: 'z2', label: 'key points', split: 'bottom_70', content: { type: 'bullets', bullets: [] } }
      ]

    case 'strategic_analysis':
    case 'market_analysis':
      return [
        { zone_id: 'z1', label: 'primary chart', split: 'left_60', content: { type: 'chart', chart_type: 'bar', categories: [], series: [] } },
        { zone_id: 'z2', label: 'supporting bullets', split: 'right_40', content: { type: 'bullets', bullets: [] } }
      ]

    case 'recommendations':
      return [{ zone_id: 'z1', label: 'recommendations', split: 'full', content: { type: 'cards', cards: [] } }]

    case 'conclusion':
      return [{ zone_id: 'z1', label: 'next steps', split: 'full', content: { type: 'process_flow', steps: [] } }]

    case 'operational_review':
      return [
        { zone_id: 'z1', label: 'metrics', split: 'top_30', content: { type: 'stat_callout', stats: [] } },
        { zone_id: 'z2', label: 'detail table', split: 'bottom_70', content: { type: 'data_table', headers: [], rows: [] } }
      ]

    default:
      return [{ zone_id: 'z1', label: 'content', split: 'full', content: { type: 'bullets', bullets: [] } }]
  }
}

function isZonePlaceholder(zone) {
  const c = zone.content || {}
  const type = (c.type || '').toLowerCase()

  if (type === 'bullets') {
    const bullets = c.bullets || []
    if (!bullets.length) return true
    return bullets.every(b => {
      const text = typeof b === 'string' ? b : (b.text || '')
      return !text || text.trim().length < 5 || /key point|placeholder|tbd|insert/i.test(text)
    })
  }
  if (type === 'stat_callout') {
    const stats = c.stats || []
    if (!stats.length) return true
    return stats.every(s => !s.value || s.value === '—' || /placeholder|tbd/i.test(s.value))
  }
  if (type === 'chart') {
    return !c.categories || c.categories.length === 0 || !c.series || c.series.length === 0
  }
  if (type === 'cards') {
    return !c.cards || c.cards.length === 0
  }
  return false
}

function hasPlaceholderContent(slide) {
  if (slide.slide_type !== 'content') return false
  if (!slide.zones || slide.zones.length === 0) return true
  if (!slide.key_message || slide.key_message.trim().length < 10) return true
  return slide.zones.some(z => isZonePlaceholder(z))
}

function normaliseZone(z) {
  if (!z) return null
  const content = z.content || {}
  if (!content.type) content.type = 'bullets'

  // Ensure chart has chart_title
  if (content.type === 'chart' && !content.chart_title) {
    content.chart_title = z.label || ''
  }

  // Ensure chart type is set
  if (content.type === 'chart' && !content.chart_type) {
    content.chart_type = 'bar'
  }

  // Normalise bullets
  if (content.type === 'bullets' && content.bullets) {
    content.bullets = content.bullets.map(b => {
      if (typeof b === 'string') {
        return { text: b, emphasis: autoEmphasis(b), sentiment: autoSentiment(b) }
      }
      return {
        text:      b.text      || '',
        emphasis:  b.emphasis  || autoEmphasis(b.text || ''),
        sentiment: b.sentiment || autoSentiment(b.text || '')
      }
    })
  }

  return { zone_id: z.zone_id || 'z1', label: z.label || '', split: z.split || 'full', content }
}

function autoEmphasis(text) {
  const m = text.match(/[₹$€]?[\d,]+\.?\d*[%CrLKMB]*/g)
  if (m && m[0]) return [{ text: m[0], style: 'bold', color: null }]
  return []
}

function autoSentiment(text) {
  const t = text.toLowerCase()
  if (/risk|decline|fall|breach|exceed|critical|concern|dangerous|concentrat|below|deficit/.test(t)) return 'warning'
  if (/grew|growth|strong|record|exceed.*target|improv|positive|increase/.test(t)) return 'positive'
  return 'neutral'
}

function normaliseSlide(slide, plan) {
  const slideType = slide.slide_type || plan.slide_type || 'content'

  // Normalise zones
  let zones = []

  if (slide.zones && Array.isArray(slide.zones) && slide.zones.length > 0) {
    zones = slide.zones.map(normaliseZone).filter(Boolean)
  } else if (slide.primary_content || slide.content) {
    // Legacy format conversion
    const pc = slide.primary_content || slide.content || {}
    const sc = slide.secondary_content || null
    const pcType = (pc.type || '').replace('bullet_list','bullets').replace('chart_bar','chart').replace('chart_line','chart')

    if (!pc.type) pc.type = pcType || 'bullets'

    if (sc) {
      zones = [
        { zone_id: 'z1', label: 'primary', split: 'left_60', content: { ...pc, type: pcType || 'bullets' } },
        { zone_id: 'z2', label: 'secondary', split: 'right_40', content: sc }
      ]
    } else {
      zones = [{ zone_id: 'z1', label: 'main', split: 'full', content: { ...pc, type: pcType || 'bullets' } }]
    }
    zones = zones.map(normaliseZone).filter(Boolean)
  } else {
    zones = inferDefaultZones(plan.section_type, slideType, 0).map(normaliseZone).filter(Boolean)
  }

  return {
    slide_number:  slide.slide_number  || plan.slide_number,
    section_name:  slide.section_name  || plan.section_name  || '',
    section_type:  slide.section_type  || plan.section_type  || '',
    slide_type:    slideType,
    title:         slide.title         || plan.section_name  || ('Slide ' + plan.slide_number),
    subtitle:      slide.subtitle      || '',
    key_message:   slide.key_message   || plan.so_what       || '',
    zones:         zones,
    speaker_note:  slide.speaker_note  || plan.purpose       || ''
  }
}

function buildSlidePlan(brief, slideCount) {
  const sections = brief.sections || []
  const plan     = []
  let   num      = 1

  for (const section of sections) {
    const count = Math.max(1, section.suggested_slide_count || 1)

    for (let i = 0; i < count; i++) {
      if (num > slideCount) break

      let slideType = 'content'
      if (num === 1) slideType = 'title'
      else if (section.section_type === 'divider') slideType = 'divider'

      plan.push({
        slide_number:  num,
        section_name:  section.section_name  || '',
        section_type:  section.section_type  || '',
        slide_type:    slideType,
        purpose:       section.purpose       || '',
        key_content:   section.key_content   || [],
        so_what:       section.so_what       || '',
        data_available:section.data_available || false,
        slide_index_in_section: i
      })
      num++
    }
    if (num > slideCount) break
  }

  return plan
}

async function writeSlideBatch(batchPlan, brief, contentB64, batchNum) {
  console.log('Agent 4 — batch', batchNum, ': slides', batchPlan[0].slide_number, '–', batchPlan[batchPlan.length-1].slide_number)

  const prompt = `PRESENTATION BRIEF:
Document type:     ${brief.document_type || '—'}
Governing thought: ${brief.governing_thought || '—'}
Narrative flow:    ${brief.narrative_flow || '—'}
Tone:              ${brief.tone || 'professional'}
Key messages:      ${(brief.key_messages || []).join(' | ')}
Key data points:   ${(brief.key_data_points || []).join(' | ')}
Recommendations:   ${(brief.recommendations || []).join(' | ')}

SLIDES TO WRITE — batch ${batchNum} (${batchPlan.length} slides):
${JSON.stringify(batchPlan, null, 2)}

INSTRUCTIONS:
- Write full zones[] content for each slide using the brief AND the attached source document
- Decide how many zones each slide needs based on the richness of the data
- Use multiple zones when comparing data sets, showing trends alongside summaries, or presenting parallel insights
- Do NOT force multi-zone when a single chart or table tells the story clearly
- Zone splits must be complementary — left_60 + right_40, top_30 + bottom_70 etc.
- Title slides: SHORT name (5-8 words max), not the governing thought
- Content slide titles: INSIGHT-LED — state the conclusion
- All numbers must come from the source document
- ZERO placeholder content
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

async function repairSlide(slide, brief, contentB64) {
  console.log('Agent 4 — repairing slide', slide.slide_number, ':', slide.title)

  const prompt = `This slide has placeholder or missing content in its zones. Fix every zone with real data from the source document.

CONTEXT:
Document type:  ${brief.document_type || '—'}
Key messages:   ${(brief.key_messages || []).join(' | ')}
Key data:       ${(brief.key_data_points || []).join(' | ')}

SLIDE TO FIX:
${JSON.stringify(slide, null, 2)}

Rules:
- Replace ALL placeholder content with real, specific data-driven content
- Keep the same zones[] structure — just fill in the content
- Numbers must come from the source document
- Return ONLY a single JSON object for this one slide`

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

async function runAgent4(state) {
  const brief      = state.outline
  const contentB64 = state.contentB64
  const slideCount = (brief && brief.total_slides) || state.slideCount || 12

  console.log('Agent 4 starting — target slides:', slideCount, '| doc type:', (brief && brief.document_type) || '—')

  const slidePlan = buildSlidePlan(brief, slideCount)
  console.log('  Slide plan:', slidePlan.length, 'slides')

  // Batch into groups of 5
  const batches = []
  for (let i = 0; i < slidePlan.length; i += 5) batches.push(slidePlan.slice(i, i + 5))

  let allSlides = []

  for (let b = 0; b < batches.length; b++) {
    const batch  = batches[b]
    const result = await writeSlideBatch(batch, brief, contentB64, b + 1)

    if (!result) {
      batch.forEach(plan => allSlides.push(normaliseSlide({}, plan)))
    } else {
      batch.forEach((plan, idx) => {
        const match = result[idx] || result.find(s => s.slide_number === plan.slide_number)
        allSlides.push(normaliseSlide(match || {}, plan))
      })
    }
  }

  // Validate and repair
  const failed = allSlides.filter(s => hasPlaceholderContent(s))
  console.log('  Slides needing repair:', failed.length)

  for (const slide of failed.slice(0, 5)) {
    const repaired = await repairSlide(slide, brief, contentB64)
    if (repaired) {
      const idx = allSlides.findIndex(s => s.slide_number === slide.slide_number)
      if (idx >= 0) {
        allSlides[idx] = normaliseSlide(repaired, slidePlan.find(p => p.slide_number === slide.slide_number) || {})
        console.log('  Repaired slide', slide.slide_number)
      }
    }
  }

  // Summary log
  const zoneBreakdown = {}
  const typeBreakdown = {}
  allSlides.forEach(s => {
    const zc = (s.zones || []).length
    zoneBreakdown[zc] = (zoneBreakdown[zc] || 0) + 1
    ;(s.zones || []).forEach(z => {
      const t = (z.content || {}).type || 'unknown'
      typeBreakdown[t] = (typeBreakdown[t] || 0) + 1
    })
  })

  console.log('Agent 4 complete')
  console.log('  Total slides:', allSlides.length)
  console.log('  Zone counts:', JSON.stringify(zoneBreakdown))
  console.log('  Content types:', JSON.stringify(typeBreakdown))
  console.log('  Placeholder remaining:', allSlides.filter(s => hasPlaceholderContent(s)).length)

  return allSlides
}
