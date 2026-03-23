// ─── AGENT 4 — SLIDE CONTENT WRITER ──────────────────────────────────────────
// Input:  state.outline        — presentationBrief from Agent 3
//         state.contentB64     — original PDF (fallback for detail)
// Output: slideManifest        — flat JSON array, one object per slide
//
// Key fixes in this version:
// 1. Splits into batches of 6 slides — prevents token exhaustion
// 2. Validates every slide — re-requests any slide with placeholder content
// 3. Title slide gets a clean name, not the governing thought as title
// 4. Strong prompt enforcement — no placeholders allowed

const AGENT4_SYSTEM = `You are a senior management consultant and presentation specialist.
You will be given a presentation brief and a batch of slides to write content for.

For EACH slide in the batch, produce a fully populated slide object.

Each slide object must have EXACTLY these fields:

{
  "slide_number": number,
  "section_name": "string",
  "section_type": "string",
  "slide_type": "title" | "divider" | "content",
  "title": "string — insight-led title, specific to the content",
  "subtitle": "string — for title slides only, empty string for others",
  "key_message": "string — THE single most important insight this slide communicates. Specific, data-driven, not generic.",
  "visual_type": "string",
  "content": { ... fully populated, see rules below ... },
  "speaker_note": "string — what presenter says that is NOT on the slide"
}

TITLE SLIDE RULES:
- title: short presentation name (e.g. "Retail Portfolio Risk Review — Q3 2025") NOT the governing thought
- subtitle: audience or date
- key_message: the governing thought (the single most important insight)
- visual_type: "title_slide"
- content: { "title": same short name, "subtitle": audience/date, "date": "" }

DIVIDER SLIDE RULES:
- title: section name only
- visual_type: "divider_slide"
- content: { "section_name": "string", "section_descriptor": "one line on what this section covers" }
- key_message: what the audience will learn in this section

CONTENT SLIDE RULES — visual_type determines content structure:

"bullet_list":
content: { "bullets": ["specific insight with data", "specific insight", "specific insight"] }
— 3 to 5 bullets. Each must be a COMPLETE SENTENCE with SPECIFIC information. NO generic text.

"three_column":
content: { "columns": [ { "header": "string", "body": "2-3 sentence description" }, { "header": "string", "body": "string" }, { "header": "string", "body": "string" } ] }

"two_column":
content: { "left_header": "string", "left_points": ["point","point"], "right_header": "string", "right_points": ["point","point"] }

"data_table":
content: { "headers": ["Col1","Col2","Col3"], "rows": [ ["val","val","val"], ["val","val","val"] ] }
— Use REAL data from the document. Every cell must have actual values.

"stat_callout":
content: { "stats": [ { "value": "₹450Cr", "label": "Total Disbursed", "change": "+18% YoY" }, ... ] }
— 2 to 4 stats. Values must be REAL numbers from the document.

"chart_bar":
content: { "chart_title": "string", "x_label": "string", "y_label": "string", "series": [ { "name": "string", "data": [ { "label": "Category", "value": 123 }, ... ] } ] }
— Use REAL data points from the document.

"chart_line":
content: { "chart_title": "string", "x_label": "string", "y_label": "string", "series": [ { "name": "string", "data": [ { "label": "Period", "value": 123 }, ... ] } ] }

"chart_waterfall":
content: { "chart_title": "string", "items": [ { "label": "string", "value": number, "type": "positive|negative|total" }, ... ] }

"quote_callout":
content: { "quote": "the full insight statement", "attribution": "source or context" }

"icon_cards":
content: { "cards": [ { "icon": "icon concept", "header": "string", "description": "string" }, ... ] }

"process_flow":
content: { "steps": [ { "step_number": 1, "title": "string", "description": "string" }, ... ] }

CRITICAL RULES — these will be validated:
1. ZERO placeholder text allowed. "Key point from this slide", "TBD", "Insert data here" = FAILURE
2. Every bullet must contain SPECIFIC information from the document — names, numbers, percentages, facts
3. Every stat must contain REAL values — not "X%", not "N/A", not placeholder numbers
4. Every chart must have REAL data points from the source document
5. key_message must be specific — "66% of portfolio concentrated in North Zone" not "concentration risk exists"
6. Title slide title must be SHORT (5-8 words max) — a presentation name, NOT a paragraph

Return ONLY a valid JSON array for the slides in this batch. No explanation. No markdown fences.`


// ─── PLACEHOLDER DETECTOR ─────────────────────────────────────────────────────
const PLACEHOLDER_PATTERNS = [
  /key point from this slide/i,
  /insert data here/i,
  /tbd/i,
  /placeholder/i,
  /lorem ipsum/i,
  /sample text/i,
  /your (text|content|data) here/i,
  /\[.*?\]/,  // [anything in brackets]
  /^\s*-\s*$/  // just a dash
]

function isPlaceholder(text) {
  if (!text || typeof text !== 'string') return true
  if (text.trim().length < 5) return true
  return PLACEHOLDER_PATTERNS.some(p => p.test(text))
}

function hasPlaceholderContent(slide) {
  const content = slide.content || {}

  // Check bullets
  if (content.bullets) {
    if (content.bullets.length === 0) return true
    if (content.bullets.every(b => isPlaceholder(b))) return true
  }

  // Check stats
  if (content.stats) {
    if (content.stats.every(s => isPlaceholder(s.value) || s.value === '—')) return true
  }

  // Check columns
  if (content.columns) {
    if (content.columns.every(c => isPlaceholder(c.body))) return true
  }

  // Check chart data
  if (content.series) {
    if (content.series.length === 0) return true
    if (content.series.every(s => (s.data || []).length === 0)) return true
  }

  // Check rows
  if (content.rows) {
    if (content.rows.length === 0) return true
  }

  // Check key message
  if (isPlaceholder(slide.key_message)) return true

  return false
}


// ─── SLIDE PLAN BUILDER ──────────────────────────────────────────────────────
// Creates the slide plan from the brief — what each slide should cover

function buildSlidePlan(brief, slideCount) {
  const sections = brief.sections || []
  const plan = []
  let slideNum = 1

  for (const section of sections) {
    const count = Math.max(1, section.suggested_slide_count || 1)

    for (let i = 0; i < count; i++) {
      if (slideNum > slideCount) break

      let slideType   = 'content'
      let visualType  = inferVisualType(section.section_type, i, count)

      if (slideNum === 1) {
        slideType  = 'title'
        visualType = 'title_slide'
      } else if (section.section_type === 'divider' || section.section_type === 'title') {
        slideType  = section.section_type === 'title' ? 'title' : 'divider'
        visualType = section.section_type === 'title' ? 'title_slide' : 'divider_slide'
      }

      plan.push({
        slide_number:  slideNum,
        section_name:  section.section_name,
        section_type:  section.section_type,
        slide_type:    slideType,
        suggested_visual_type: visualType,
        purpose:       section.purpose || '',
        key_content:   section.key_content || [],
        so_what:       section.so_what || '',
        data_available: section.data_available || false
      })

      slideNum++
    }
    if (slideNum > slideCount) break
  }

  return plan
}

function inferVisualType(sectionType, slideIndex, totalInSection) {
  switch (sectionType) {
    case 'financial_data':
      return slideIndex === 0 ? 'stat_callout' : 'chart_bar'
    case 'executive_summary':
      return 'stat_callout'
    case 'strategic_analysis':
      return slideIndex === 0 ? 'bullet_list' : 'chart_bar'
    case 'market_analysis':
      return 'chart_bar'
    case 'recommendations':
      return totalInSection > 1 ? 'three_column' : 'bullet_list'
    case 'conclusion':
      return 'process_flow'
    case 'operational_review':
      return 'data_table'
    default:
      return 'bullet_list'
  }
}


// ─── WRITE A BATCH OF SLIDES ──────────────────────────────────────────────────
async function writeSlideBatch(batchPlan, brief, contentB64, batchNum) {
  console.log('Agent 4 — writing batch', batchNum, '— slides', batchPlan[0].slide_number, 'to', batchPlan[batchPlan.length-1].slide_number)

  const batchPrompt = `PRESENTATION BRIEF:
Document type: ${brief.document_type || '—'}
Governing thought: ${brief.governing_thought || '—'}
Narrative flow: ${brief.narrative_flow || '—'}
Tone: ${brief.tone || 'professional'}
Key messages: ${(brief.key_messages || []).join(' | ')}
Key data points: ${(brief.key_data_points || []).join(' | ')}
Recommendations: ${(brief.recommendations || []).join(' | ')}

SLIDES TO WRITE (batch ${batchNum}):
${JSON.stringify(batchPlan, null, 2)}

INSTRUCTIONS:
- Write full content for each of the ${batchPlan.length} slides above
- Use the key_content and so_what from each slide as your starting point
- Pull REAL data from the source document (attached) — actual numbers, names, percentages
- For financial slides: use real figures from the document
- For title slide: title must be SHORT (5-8 words) — the presentation name, not the governing thought
- ZERO placeholder content — every bullet, stat, and data point must be specific
- Return ONLY a JSON array for these ${batchPlan.length} slides`

  const messages = [{
    role: 'user',
    content: [
      {
        type: 'document',
        source: { type: 'base64', media_type: 'application/pdf', data: contentB64 }
      },
      { type: 'text', text: batchPrompt }
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


// ─── REPAIR A FAILED SLIDE ────────────────────────────────────────────────────
async function repairSlide(slide, brief, contentB64) {
  console.log('Agent 4 — repairing slide', slide.slide_number, '(placeholder content detected)')

  const repairPrompt = `This slide has placeholder content that needs to be replaced with real content.

PRESENTATION CONTEXT:
Document type: ${brief.document_type || '—'}
Governing thought: ${brief.governing_thought || '—'}
Key messages: ${(brief.key_messages || []).join(' | ')}
Key data: ${(brief.key_data_points || []).join(' | ')}

SLIDE TO FIX:
${JSON.stringify(slide, null, 2)}

Write the FULL content for this ONE slide.
- Replace all placeholder text with specific, data-driven content from the document
- For title slide: use a SHORT presentation name (5-8 words max)
- For content slides: use real numbers, facts, and insights from the source document
- Return ONLY a single JSON object for this one slide. No array. No explanation.`

  const messages = [{
    role: 'user',
    content: [
      {
        type: 'document',
        source: { type: 'base64', media_type: 'application/pdf', data: contentB64 }
      },
      { type: 'text', text: repairPrompt }
    ]
  }]

  const raw    = await callClaude(AGENT4_SYSTEM, messages, 1500)
  const cleaned = raw.replace(/```json\s*/g,'').replace(/```\s*/g,'').trim()

  // Handle both object and array response
  try {
    const parsed = JSON.parse(cleaned)
    if (Array.isArray(parsed) && parsed.length > 0) return parsed[0]
    if (typeof parsed === 'object' && parsed.slide_number) return parsed
  } catch(e) {
    console.warn('Agent 4 repair — parse failed for slide', slide.slide_number)
  }

  return null
}


// ─── NORMALISE A SLIDE ────────────────────────────────────────────────────────
function normaliseSlide(slide, plan) {
  const slideType = slide.slide_type || plan.slide_type || 'content'
  const vt        = slide.visual_type || plan.suggested_visual_type || 'bullet_list'

  return {
    slide_number:  slide.slide_number  || plan.slide_number,
    section_name:  slide.section_name  || plan.section_name,
    section_type:  slide.section_type  || plan.section_type,
    slide_type:    slideType,
    title:         slide.title         || plan.section_name || ('Slide ' + plan.slide_number),
    subtitle:      slide.subtitle      || '',
    key_message:   slide.key_message   || plan.so_what || '',
    visual_type:   vt,
    content:       normaliseContent(slide.content, vt, slide),
    speaker_note:  slide.speaker_note  || plan.purpose || ''
  }
}

function normaliseContent(content, vt, slide) {
  if (content && typeof content === 'object' && Object.keys(content).length > 0) {
    // Check if bullets are all placeholders — if so, return minimal shell for repair
    if (content.bullets && content.bullets.every(b => isPlaceholder(b))) {
      return { bullets: [] }
    }
    return content
  }

  // Minimal valid content by visual type
  switch ((vt || '').toLowerCase()) {
    case 'title_slide':   return { title: slide.title || '', subtitle: slide.subtitle || '', date: '' }
    case 'divider_slide': return { section_name: slide.title || '', section_descriptor: '' }
    case 'bullet_list':   return { bullets: [] }
    case 'three_column':  return { columns: [] }
    case 'two_column':    return { left_header: '', left_points: [], right_header: '', right_points: [] }
    case 'data_table':    return { headers: [], rows: [] }
    case 'stat_callout':  return { stats: [] }
    case 'chart_bar':     return { chart_title: '', x_label: '', y_label: '', series: [] }
    case 'chart_line':    return { chart_title: '', x_label: '', y_label: '', series: [] }
    case 'chart_waterfall': return { chart_title: '', items: [] }
    case 'quote_callout': return { quote: slide.key_message || '', attribution: '' }
    case 'icon_cards':    return { cards: [] }
    case 'process_flow':  return { steps: [] }
    default:              return {}
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

  // Step 1 — Build slide plan from brief
  const slidePlan = buildSlidePlan(brief, slideCount)
  console.log('Agent 4 — slide plan built:', slidePlan.length, 'slides')

  // Step 2 — Split into batches of 5
  const BATCH_SIZE = 5
  const batches    = []
  for (let i = 0; i < slidePlan.length; i += BATCH_SIZE) {
    batches.push(slidePlan.slice(i, i + BATCH_SIZE))
  }
  console.log('Agent 4 — split into', batches.length, 'batches of max', BATCH_SIZE)

  // Step 3 — Write each batch
  let allSlides = []

  for (let b = 0; b < batches.length; b++) {
    const batch       = batches[b]
    const batchResult = await writeSlideBatch(batch, brief, contentB64, b + 1)

    if (!batchResult) {
      // Batch failed — use plan as shell for repair
      console.warn('Agent 4 — batch', b+1, 'failed, using plan shells')
      batch.forEach(plan => {
        allSlides.push(normaliseSlide({
          slide_number: plan.slide_number,
          section_name: plan.section_name,
          section_type: plan.section_type,
          slide_type:   plan.slide_type,
          title:        plan.section_name,
          key_message:  plan.so_what,
          visual_type:  plan.suggested_visual_type,
          content:      normaliseContent({}, plan.suggested_visual_type, plan),
          speaker_note: plan.purpose
        }, plan))
      })
    } else {
      // Merge batch results with plan
      batch.forEach((plan, idx) => {
        const result = batchResult[idx] || batchResult.find(s => s.slide_number === plan.slide_number)
        if (result) {
          allSlides.push(normaliseSlide(result, plan))
        } else {
          allSlides.push(normaliseSlide({ slide_number: plan.slide_number }, plan))
        }
      })
    }
  }

  // Step 4 — Validate and repair placeholder slides
  console.log('Agent 4 — validating', allSlides.length, 'slides for placeholder content...')
  const failedSlides = allSlides.filter(s => hasPlaceholderContent(s) && s.slide_type === 'content')
  console.log('Agent 4 — slides needing repair:', failedSlides.length)

  // Repair failed slides (max 5 repairs to avoid excessive API calls)
  const toRepair = failedSlides.slice(0, 5)
  for (const slide of toRepair) {
    const repaired = await repairSlide(slide, brief, contentB64)
    if (repaired) {
      const idx = allSlides.findIndex(s => s.slide_number === slide.slide_number)
      if (idx >= 0) {
        allSlides[idx] = normaliseSlide(repaired, slidePlan.find(p => p.slide_number === slide.slide_number) || {})
        console.log('Agent 4 — repaired slide', slide.slide_number)
      }
    }
  }

  // Step 5 — Final validation log
  const finalFailed = allSlides.filter(s => hasPlaceholderContent(s) && s.slide_type === 'content')
  if (finalFailed.length > 0) {
    console.warn('Agent 4 — slides still with placeholder content after repair:', finalFailed.map(s => s.slide_number).join(', '))
  }

  // Log summary
  const vtBreakdown = {}
  allSlides.forEach(s => { vtBreakdown[s.visual_type] = (vtBreakdown[s.visual_type] || 0) + 1 })
  console.log('Agent 4 complete')
  console.log('  Total slides:', allSlides.length)
  console.log('  Visual types:', JSON.stringify(vtBreakdown))
  console.log('  Placeholder slides remaining:', finalFailed.length)

  return allSlides
}
