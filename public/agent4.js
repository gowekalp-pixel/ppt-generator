// ─── AGENT 4 — SLIDE CONTENT ARCHITECT ────────────────────────────────────────
// Input:  state.outline    — presentation brief / slide plan from Agent 3
//         state.contentB64 — original PDF or source document
// Output: slideManifest    — flat JSON array, one object per slide
//
// Agent 4 decides WHAT each slide is trying to say, WHAT zones (messaging arcs)
// it needs, and WHAT artifacts sit inside each zone.
// Agent 5 decides final layout, coordinates, styling, and rendering.
//
// Key concepts:
//   Zone    = a messaging arc — one coherent argument unit
//   Artifact= the visual expression inside a zone (chart, table, workflow, etc.)
//   Each slide: max 4 zones, title/subtitle outside zones
//   Each zone:  max 2 artifacts

// ═══════════════════════════════════════════════════════════════════════════════
// SYSTEM PROMPT (pasted from consultant-authored spec)
// ═══════════════════════════════════════════════════════════════════════════════

const AGENT4_SYSTEM = `You are a senior management consultant acting as the slide content architect for a board-level presentation.

You will receive:
1. A structured presentation brief
2. A batch of slide plans
3. A source document for reference

Your role is to define the HIGH-LEVEL CONTENT STRUCTURE for each slide.

You do NOT design the final slide.
You do NOT decide coordinates, colors, fonts, or exact visual styling.
You DO decide:
- what the slide is trying to prove
- what messaging arcs (zones) it needs
- what artifacts belong inside each zone
- what should visually dominate
- how a workflow should be structured when needed

Return ONLY a valid JSON array with one object per slide.

═══════════════════════════
OUTPUT OBJECT — REQUIRED FIELDS
═══════════════════════════

Each slide object must contain EXACTLY these top-level fields:

{
  "slide_number": number,
  "section_name": "string",
  "section_type": "string",
  "slide_type": "title" | "divider" | "content",
  "slide_archetype": "summary" | "trend" | "comparison" | "breakdown" | "driver_analysis" | "process" | "recommendation" | "dashboard" | "proof" | "roadmap",
  "title": "string",
  "subtitle": "string",
  "key_message": "string",
  "visual_flow_hint": "string",
  "context_from_previous_slide": "string",
  "zones": [ ... ],
  "speaker_note": "string"
}

═══════════════════════════
SLIDE TYPE RULES
═══════════════════════════

1. Title slide
- title: short presentation name, 4-8 words
- subtitle: audience / context / date if relevant
- key_message: governing thought of the full deck
- slide_archetype: "summary"
- zones: []

2. Divider slide
- title: section name only
- subtitle: empty
- key_message: one-line purpose of the section
- slide_archetype: "summary"
- zones: []

3. Content slide
- title must be insight-led — never generic topic titles
  WRONG: "Revenue Analysis" | RIGHT: "Premium mix drove most of the revenue uplift"
  WRONG: "Market Overview"  | RIGHT: "Market growing at 22% CAGR with untapped headroom"
  WRONG: "Geographic Risk"  | RIGHT: "North Zone concentration exceeds the safe exposure threshold"

═══════════════════════════
SLIDE ARCHETYPE RULES
═══════════════════════════

Choose ONE slide_archetype per content slide:

summary       — executive summaries, headline synthesis, recap. Often metrics + implications.
trend         — time-based movement. Often line/bar + implication.
comparison    — compare categories, products, geographies, cohorts. Often bar / clustered_bar / cards / table.
breakdown     — composition or segmentation. Often pie / bar / decomposition workflow / table.
driver_analysis — explain movement from one state to another. Often waterfall + insight.
process       — process, workflow, hierarchy, information movement. Often workflow + insight.
recommendation — actions, priorities, strategic choices. Often cards / bullets / roadmap workflow.
dashboard     — metric-heavy summary. Stats, short tables, compact insights.
proof         — validate a claim with evidence. Chart/table/workflow plus interpretation.
roadmap       — phased plan, milestones, implementation sequencing. Workflow or structured steps.

═══════════════════════════
ZONE DEFINITION
═══════════════════════════

A zone is a self-contained messaging arc within a slide.
It is a structured unit of meaning that communicates one distinct part of the slide argument.
A zone is NOT merely a visual box or layout area.

Each slide may contain a maximum of 4 zones.
Title and subtitle are OUTSIDE zones.
Each zone may contain a maximum of 2 artifacts.

Each zone object must contain:

{
  "zone_id": "z1",
  "zone_role": "primary_proof" | "supporting_evidence" | "implication" | "summary" | "comparison" | "breakdown" | "process" | "recommendation",
  "message_objective": "string — one sentence: what this zone proves or communicates",
  "narrative_weight": "primary" | "secondary" | "supporting",
  "artifacts": [ ... ],
  "layout_hint": {
    "split": "full" | "left_50" | "right_50" | "left_60" | "right_40" | "left_40" | "right_60" |
             "top_30" | "bottom_70" | "top_40" | "bottom_60" | "top_50" | "bottom_50" |
             "top_left_50" | "top_right_50" | "bottom_full" |
             "left_full_50" | "top_right_50_h" | "bottom_right_50_h" |
             "tl" | "tr" | "bl" | "br"
  }
}

Zone rules:
- max 4 zones per slide
- max 2 artifacts per zone
- at least 1 primary zone per content slide
- no more than 2 primary zones per slide
- every zone must support the slide key_message
- every zone must communicate one coherent message objective

ALLOWED SPLIT COMBINATIONS:
1 zone:  full
2 zones side by side: left_50+right_50 | left_60+right_40 | left_40+right_60
2 zones stacked:      top_30+bottom_70 | top_40+bottom_60 | top_50+bottom_50
3 zones:              top_left_50+top_right_50+bottom_full | left_full_50+top_right_50_h+bottom_right_50_h
4 zones:              tl+tr+bl+br
All zones on a slide must together cover the full content area.

═══════════════════════════
ARTIFACT TYPES
═══════════════════════════

Allowed artifact types: insight_text | chart | cards | workflow | table

Good artifact combinations inside one zone:
- chart + insight_text
- workflow + insight_text
- table + insight_text
- cards + insight_text
- chart + table

Discouraged:
- chart + chart  (use two separate zones instead)
- table + table
- workflow + workflow

═══════════════════════════
ARTIFACT 1: insight_text
═══════════════════════════

{
  "type": "insight_text",
  "heading": "Key Insight" | "So What" | "Risk Alert" | "Action Required",
  "points": ["specific insight with data", "..."],
  "sentiment": "positive" | "warning" | "neutral"
}

Rules:
- max 4 points
- each point must be SPECIFIC — include actual numbers, names, percentages
- final point should ideally state implication or action
- ZERO placeholder or generic text

═══════════════════════════
ARTIFACT 2: chart
═══════════════════════════

{
  "type": "chart",
  "chart_type": "bar" | "line" | "pie" | "waterfall" | "clustered_bar",
  "chart_decision": "one line: why this chart type was chosen",
  "chart_title": "descriptive title",
  "chart_insight": "the one-line insight the chart proves",
  "x_label": "string",
  "y_label": "string",
  "categories": ["string", "string", "string"],
  "series": [
    { "name": "string", "values": [number, number, number], "types": ["positive"|"negative"|"total"] }
  ],
  "show_data_labels": true,
  "show_legend": true | false
}

Chart type selection:
- bar:           compare 3+ categories, one series, no time
- line:          trend over time (months, quarters, years in categories)
- pie:           composition — values sum to ~100%, max 5 segments
- waterfall:     bridge or variance — series items have types: positive/negative/total
- clustered_bar: EXACTLY 2 series compared across the same 3+ categories

CRITICAL chart rules:
- bar, line, clustered_bar: MINIMUM 3 categories
- categories and values must match in count
- NO zeros-only series
- NO placeholder values
- clustered_bar: MUST have exactly 2 series — if only 1 series exists, use bar
- All numbers must be sourced from the document

═══════════════════════════
ARTIFACT 3: cards
═══════════════════════════

{
  "type": "cards",
  "cards": [
    {
      "title": "string",
      "subtitle": "string",
      "body": "string",
      "sentiment": "positive" | "negative" | "neutral"
    }
  ]
}

Rules:
- max 4 cards in full-width zones
- max 2 cards in side zones (left_X, right_X, tl, tr, bl, br)
- use for metrics, parallel messages, recommendations, priorities

═══════════════════════════
ARTIFACT 4: workflow
═══════════════════════════

{
  "type": "workflow",
  "workflow_type": "process_flow" | "hierarchy" | "decomposition" | "information_flow" | "timeline",
  "flow_direction": "left_to_right" | "top_to_bottom" | "top_down_branching" | "bottom_up",
  "workflow_title": "string",
  "workflow_insight": "string",
  "nodes": [
    {
      "id": "n1",
      "label": "string",
      "value": "string",
      "description": "string",
      "level": 1
    }
  ],
  "connections": [
    { "from": "n1", "to": "n2", "type": "arrow" }
  ]
}

Workflow type rules:
- process_flow:       linear sequence of steps, max 5 nodes
- hierarchy:          parent-child structure across levels, max 6 nodes
- decomposition:      top number split into lower-level components, max 6 nodes
- information_flow:   movement across systems/teams/stages, max 5 nodes
- timeline:           phased progression, max 5 nodes

Flow direction rules:
- left_to_right:      pipelines, sequences, timelines
- top_to_bottom:      vertical flows, approvals
- top_down_branching: decomposition and hierarchy
- bottom_up:          aggregation or roll-up logic

Node rules:
- max 6 nodes total
- label is required
- value and description are optional
- level is required for hierarchy / decomposition

Connection rules:
- directional arrows only
- no crossing connections
- max 8 connections
- keep structure simple and readable

Use workflow when you need to show:
- process steps
- hierarchy
- number or concept decomposition
- information movement
- phased roadmap

Pair workflow with insight_text when interpretation is needed.

═══════════════════════════
ARTIFACT 5: table
═══════════════════════════

{
  "type": "table",
  "title": "string",
  "headers": ["string"],
  "rows": [["string"]],
  "highlight_rows": [0],
  "note": "string"
}

Rules:
- max 6 rows
- use when precise row/column comparison is necessary
- table must support the message objective — not dump raw data
- numbers must be specific and sourced

═══════════════════════════
STORYTELLING RULES
═══════════════════════════

1. Every content slide must prove ONE thing
   title = the conclusion / key_message = the exact takeaway

2. Every content slide must contain an implication
   either via an implication zone OR via insight_text in a zone

3. Every zone must contribute meaningfully
   no decorative zones, no unrelated side content

4. Visual hierarchy:
   primary zone    = anchor proof
   secondary zone  = interpretation or key support
   supporting zone = detail only

5. Think like a consultant:
   what should the audience understand in 3 seconds?
   build the slide around that answer

═══════════════════════════
QUALITY GATES
═══════════════════════════

- No placeholder text anywhere
- No invented numbers — all figures must come from the source document
- No vague wording
- Max 4 zones per slide
- Max 2 artifacts per zone
- Insight-led titles on every content slide
- Workflows must be structurally coherent — nodes and connections must match
- Content must be board-ready and decision-oriented

Return ONLY a valid JSON array. No explanation. No markdown. No text outside the JSON.`


// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE ARCHETYPE → DEFAULT ZONE STRUCTURE
// ═══════════════════════════════════════════════════════════════════════════════

function defaultZonesForArchetype(archetype, sectionType) {
  switch (archetype) {

    case 'dashboard':
      return [
        { zone_id: 'z1', zone_role: 'summary', narrative_weight: 'primary',
          message_objective: 'Headline metrics at a glance',
          layout_hint: { split: 'top_30' },
          artifacts: [{ type: 'cards', cards: [] }] },
        { zone_id: 'z2', zone_role: 'supporting_evidence', narrative_weight: 'secondary',
          message_objective: 'Supporting detail',
          layout_hint: { split: 'bottom_70' },
          artifacts: [{ type: 'chart', chart_type: 'bar', categories: [], series: [] }] }
      ]

    case 'trend':
      return [
        { zone_id: 'z1', zone_role: 'primary_proof', narrative_weight: 'primary',
          message_objective: 'Show trend over time',
          layout_hint: { split: 'left_60' },
          artifacts: [{ type: 'chart', chart_type: 'line', categories: [], series: [] }] },
        { zone_id: 'z2', zone_role: 'implication', narrative_weight: 'secondary',
          message_objective: 'Interpret the trend',
          layout_hint: { split: 'right_40' },
          artifacts: [{ type: 'insight_text', heading: 'So What', points: [], sentiment: 'neutral' }] }
      ]

    case 'comparison':
      return [
        { zone_id: 'z1', zone_role: 'primary_proof', narrative_weight: 'primary',
          message_objective: 'Visual comparison across categories',
          layout_hint: { split: 'left_60' },
          artifacts: [{ type: 'chart', chart_type: 'bar', categories: [], series: [] }] },
        { zone_id: 'z2', zone_role: 'implication', narrative_weight: 'secondary',
          message_objective: 'What the comparison means',
          layout_hint: { split: 'right_40' },
          artifacts: [{ type: 'insight_text', heading: 'Key Insight', points: [], sentiment: 'neutral' }] }
      ]

    case 'breakdown':
      return [
        { zone_id: 'z1', zone_role: 'primary_proof', narrative_weight: 'primary',
          message_objective: 'Show composition or segmentation',
          layout_hint: { split: 'left_50' },
          artifacts: [{ type: 'chart', chart_type: 'pie', categories: [], series: [] }] },
        { zone_id: 'z2', zone_role: 'supporting_evidence', narrative_weight: 'secondary',
          message_objective: 'Detail behind the segments',
          layout_hint: { split: 'right_50' },
          artifacts: [{ type: 'insight_text', heading: 'Key Insight', points: [], sentiment: 'neutral' }] }
      ]

    case 'driver_analysis':
      return [
        { zone_id: 'z1', zone_role: 'primary_proof', narrative_weight: 'primary',
          message_objective: 'Show movement from baseline to result',
          layout_hint: { split: 'left_60' },
          artifacts: [{ type: 'chart', chart_type: 'waterfall', categories: [], series: [] }] },
        { zone_id: 'z2', zone_role: 'implication', narrative_weight: 'secondary',
          message_objective: 'Interpret the key drivers',
          layout_hint: { split: 'right_40' },
          artifacts: [{ type: 'insight_text', heading: 'Key Insight', points: [], sentiment: 'neutral' }] }
      ]

    case 'process':
    case 'roadmap':
      return [
        { zone_id: 'z1', zone_role: 'process', narrative_weight: 'primary',
          message_objective: 'Show process or roadmap structure',
          layout_hint: { split: 'top_60' },
          artifacts: [{ type: 'workflow', workflow_type: 'process_flow', flow_direction: 'left_to_right', nodes: [], connections: [] }] },
        { zone_id: 'z2', zone_role: 'implication', narrative_weight: 'secondary',
          message_objective: 'Key insight or action from the process',
          layout_hint: { split: 'bottom_40' },
          artifacts: [{ type: 'insight_text', heading: 'So What', points: [], sentiment: 'neutral' }] }
      ]

    case 'recommendation':
      return [
        { zone_id: 'z1', zone_role: 'recommendation', narrative_weight: 'primary',
          message_objective: 'Recommended actions or priorities',
          layout_hint: { split: 'full' },
          artifacts: [{ type: 'cards', cards: [] }] }
      ]

    case 'proof':
      return [
        { zone_id: 'z1', zone_role: 'primary_proof', narrative_weight: 'primary',
          message_objective: 'Evidence supporting the claim',
          layout_hint: { split: 'left_60' },
          artifacts: [{ type: 'chart', chart_type: 'bar', categories: [], series: [] }] },
        { zone_id: 'z2', zone_role: 'implication', narrative_weight: 'secondary',
          message_objective: 'Interpretation and implication of the evidence',
          layout_hint: { split: 'right_40' },
          artifacts: [{ type: 'insight_text', heading: 'So What', points: [], sentiment: 'neutral' }] }
      ]

    default: // summary
      return [
        { zone_id: 'z1', zone_role: 'summary', narrative_weight: 'primary',
          message_objective: 'Key summary of the section',
          layout_hint: { split: 'full' },
          artifacts: [{ type: 'insight_text', heading: 'Key Insight', points: [], sentiment: 'neutral' }] }
      ]
  }
}

function inferArchetype(sectionType, slideIndex) {
  switch (sectionType) {
    case 'financial_data':     return slideIndex === 0 ? 'dashboard' : 'comparison'
    case 'executive_summary':  return 'dashboard'
    case 'strategic_analysis': return slideIndex % 2 === 0 ? 'comparison' : 'breakdown'
    case 'market_analysis':    return 'comparison'
    case 'recommendations':    return 'recommendation'
    case 'conclusion':         return 'roadmap'
    case 'operational_review': return 'dashboard'
    default:                   return 'summary'
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// VALIDATION
// ═══════════════════════════════════════════════════════════════════════════════

function validateArtifact(artifact) {
  const t = (artifact.type || '').toLowerCase()

  if (t === 'chart') {
    const cats = artifact.categories || []
    const series = artifact.series || []

    // Need 2+ categories (3+ ideally but 2 is minimum after normalisation)
    if (cats.length < 2) return { valid: false, reason: 'chart needs 2+ categories, got ' + cats.length }

    if (!series.length) return { valid: false, reason: 'chart has no series' }

    // Values must match categories
    for (const s of series) {
      if ((s.values || []).length !== cats.length) {
        return { valid: false, reason: 'series values count mismatch: ' + (s.values||[]).length + ' vs ' + cats.length }
      }
    }

    // All-zero series
    for (const s of series) {
      if ((s.values || []).every(v => v === 0)) {
        return { valid: false, reason: 'chart series has all-zero values' }
      }
    }

    // clustered_bar needs 2 series
    if (artifact.chart_type === 'clustered_bar' && series.length < 2) {
      artifact.chart_type = 'bar' // auto-fix
    }

    return { valid: true }
  }

  if (t === 'insight_text') {
    const points = artifact.points || []
    if (!points.length) return { valid: false, reason: 'insight_text has no points' }
    if (points.every(p => !p || p.trim().length < 5)) return { valid: false, reason: 'insight_text has only placeholder points' }
    return { valid: true }
  }

  if (t === 'cards') {
    const cards = artifact.cards || []
    if (!cards.length) return { valid: false, reason: 'cards has no items' }
    return { valid: true }
  }

  if (t === 'workflow') {
    const nodes = artifact.nodes || []
    if (!nodes.length) return { valid: false, reason: 'workflow has no nodes' }
    return { valid: true }
  }

  if (t === 'table') {
    if (!(artifact.headers || []).length) return { valid: false, reason: 'table has no headers' }
    if (!(artifact.rows || []).length) return { valid: false, reason: 'table has no rows' }
    return { valid: true }
  }

  return { valid: true }
}

function hasPlaceholderContent(slide) {
  if (slide.slide_type !== 'content') return false
  if (!slide.zones || !slide.zones.length) return true
  if (!slide.key_message || slide.key_message.trim().length < 10) return true

  for (const zone of slide.zones) {
    for (const artifact of (zone.artifacts || [])) {
      const check = validateArtifact(artifact)
      if (!check.valid) return true
    }
  }
  return false
}


// ═══════════════════════════════════════════════════════════════════════════════
// NORMALISE
// ═══════════════════════════════════════════════════════════════════════════════

function normaliseArtifact(a) {
  if (!a || !a.type) return null
  const t = a.type.toLowerCase()

  if (t === 'chart') {
    if (!a.categories) a.categories = []
    if (!a.series) a.series = []
    if (!a.chart_type) a.chart_type = 'bar'
    if (!a.chart_title) a.chart_title = ''
    if (!a.chart_insight) a.chart_insight = ''
    if (a.show_data_labels === undefined) a.show_data_labels = true

    // Normalise series values to numbers
    a.series = a.series.map(s => ({
      name:   s.name   || '',
      values: (s.values || []).map(v => typeof v === 'number' ? v : parseFloat(String(v).replace(/[^0-9.-]/g,'')) || 0),
      types:  s.types  || null
    }))

    // Auto-fix clustered_bar with 1 series
    if (a.chart_type === 'clustered_bar' && a.series.length < 2) {
      a.chart_type = 'bar'
    }

    // Align values length to categories length
    a.series = a.series.map(s => {
      while (s.values.length < a.categories.length) s.values.push(0)
      s.values = s.values.slice(0, a.categories.length)
      return s
    })
  }

  if (t === 'insight_text') {
    if (!a.points) a.points = []
    if (!a.heading) a.heading = 'Key Insight'
    if (!a.sentiment) a.sentiment = 'neutral'
    // Normalise points — flatten if they're objects
    a.points = a.points.map(p => typeof p === 'string' ? p : (p.text || p.point || JSON.stringify(p)))
  }

  if (t === 'cards') {
    if (!a.cards) a.cards = []
    a.cards = a.cards.map(c => ({
      title:     c.title     || c.header || '',
      subtitle:  c.subtitle  || '',
      body:      c.body      || '',
      sentiment: c.sentiment || 'neutral'
    }))
  }

  if (t === 'workflow') {
    if (!a.nodes) a.nodes = []
    if (!a.connections) a.connections = []
    if (!a.workflow_type) a.workflow_type = 'process_flow'
    if (!a.flow_direction) a.flow_direction = 'left_to_right'
    if (!a.workflow_title) a.workflow_title = ''
    if (!a.workflow_insight) a.workflow_insight = ''
  }

  if (t === 'table') {
    if (!a.headers) a.headers = []
    if (!a.rows) a.rows = []
    if (!a.title) a.title = ''
  }

  return a
}

function normaliseZone(z) {
  if (!z) return null
  return {
    zone_id:          z.zone_id          || 'z1',
    zone_role:        z.zone_role        || 'primary_proof',
    message_objective:z.message_objective|| '',
    narrative_weight: z.narrative_weight || 'primary',
    artifacts:        (z.artifacts || []).map(normaliseArtifact).filter(Boolean),
    layout_hint:      { split: (z.layout_hint || {}).split || 'full' }
  }
}

function normaliseSlide(slide, plan) {
  const slideType = slide.slide_type || plan.slide_type || 'content'

  let zones = []

  if (slide.zones && Array.isArray(slide.zones) && slide.zones.length > 0) {
    zones = slide.zones.map(normaliseZone).filter(Boolean)
  } else {
    // Build default zones from archetype
    const archetype = slide.slide_archetype || inferArchetype(plan.section_type, plan.slide_index_in_section || 0)
    zones = defaultZonesForArchetype(archetype, plan.section_type).map(normaliseZone).filter(Boolean)
  }

  // Cap at 4 zones
  zones = zones.slice(0, 4)

  return {
    slide_number:                 slide.slide_number                 || plan.slide_number,
    section_name:                 slide.section_name                 || plan.section_name   || '',
    section_type:                 slide.section_type                 || plan.section_type   || '',
    slide_type:                   slideType,
    slide_archetype:              slide.slide_archetype              || inferArchetype(plan.section_type, 0),
    title:                        slide.title                        || plan.section_name   || ('Slide ' + plan.slide_number),
    subtitle:                     slide.subtitle                     || '',
    key_message:                  slide.key_message                  || plan.so_what        || '',
    visual_flow_hint:             slide.visual_flow_hint             || '',
    context_from_previous_slide:  slide.context_from_previous_slide  || '',
    zones:                        zones,
    speaker_note:                 slide.speaker_note                 || plan.purpose        || ''
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE PLAN BUILDER
// ═══════════════════════════════════════════════════════════════════════════════

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
        slide_number:             num,
        section_name:             section.section_name   || '',
        section_type:             section.section_type   || '',
        slide_type:               slideType,
        purpose:                  section.purpose        || '',
        key_content:              section.key_content    || [],
        so_what:                  section.so_what        || '',
        data_available:           section.data_available || false,
        slide_index_in_section:   i,
        suggested_archetype:      inferArchetype(section.section_type, i)
      })
      num++
    }

    if (num > slideCount) break
  }

  return plan
}


// ═══════════════════════════════════════════════════════════════════════════════
// BATCH WRITER
// ═══════════════════════════════════════════════════════════════════════════════

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
- For each slide, decide the archetype, write insight-led title, populate all zones with real artifacts
- Pull all numbers from the attached source document — no invented figures
- Title slides: zones = []
- Divider slides: zones = []
- Content slides: 1–4 zones, each with 1–2 artifacts
- Every chart: MUST have 3+ categories, matching values, no all-zeros
- clustered_bar: MUST have exactly 2 series
- Every insight_text: MUST have specific, data-driven points
- Workflows: fully populate nodes and connections
- Return ONLY a valid JSON array for these ${batchPlan.length} slides`

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
// REPAIR
// ═══════════════════════════════════════════════════════════════════════════════

async function repairSlide(slide, brief, contentB64) {
  console.log('Agent 4 — repairing slide', slide.slide_number, ':', slide.title)

  const prompt = `This slide has missing or invalid artifact content. Fix every zone with specific data from the source document.

CONTEXT:
Document type:  ${brief.document_type || '—'}
Key messages:   ${(brief.key_messages || []).join(' | ')}
Key data:       ${(brief.key_data_points || []).join(' | ')}

SLIDE TO FIX:
${JSON.stringify(slide, null, 2)}

Fix rules:
- Replace all placeholder or empty content with real, specific data
- Keep the same zones[] and artifact types — only fill in the content
- All numbers from the source document
- Charts: 3+ categories, matching values, no all-zeros
- insight_text: specific points with data
- workflows: fully populated nodes and connections
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


// ═══════════════════════════════════════════════════════════════════════════════
// MAIN RUNNER
// ═══════════════════════════════════════════════════════════════════════════════

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
  const archetypes = {}
  const artifactTypes = {}
  let totalZones = 0
  let totalArtifacts = 0

  allSlides.forEach(s => {
    archetypes[s.slide_archetype] = (archetypes[s.slide_archetype] || 0) + 1
    ;(s.zones || []).forEach(z => {
      totalZones++
      ;(z.artifacts || []).forEach(a => {
        totalArtifacts++
        artifactTypes[a.type] = (artifactTypes[a.type] || 0) + 1
      })
    })
  })

  console.log('Agent 4 complete')
  console.log('  Total slides:', allSlides.length)
  console.log('  Total zones:', totalZones)
  console.log('  Total artifacts:', totalArtifacts)
  console.log('  Archetypes:', JSON.stringify(archetypes))
  console.log('  Artifact types:', JSON.stringify(artifactTypes))
  console.log('  Placeholder remaining:', allSlides.filter(s => hasPlaceholderContent(s)).length)

  return allSlides
}
