// AGENT 3 - DOCUMENT ANALYST & DECK PLANNER
// Input:  state.contentB64 (PDF), state.slideCount
// Output: presentationBrief - deck context + full deck slide plan for Agent 4
//
// TWO-PHASE APPROACH:
//   Phase 1 — Read the entire document once; produce deck metadata + a lightweight
//              outline (every slot with title + role; structural slides fully populated;
//              content slides title/role only).
//   Phase 2 — Fill content slide detail in parallel batches of 2 slides each.
//              Each batch receives the document + the full outline and returns
//              strategic_objective, key_content, and layout signals for those 2 slides.

// ─── PHASE 1 SYSTEM PROMPT ────────────────────────────────────────────────────
const AGENT3_PHASE1_SYSTEM = `You are a senior management consultant with deep expertise in financial analysis,
strategy, and board-level communication. You have been asked to review a document and
plan a board-level presentation.

PART 1 — DOCUMENT ANALYSIS
Read the document and identify:
- What is the single most important insight - one punchy sentence a CEO finds immediately useful?
- What narrative flow best suits this content?
  - Financial: Situation -> Performance -> Drivers -> Outlook -> Actions
  - Market research: Context -> Market Dynamics -> Competitive Position -> Implications -> Recommendations
  - Strategy: Objective -> Current State -> Gap Analysis -> Options -> Recommended Path
  - Operational: Baseline -> Issues -> Root Cause -> Solutions -> Implementation
- Who is the audience and what tone is appropriate?

PART 2 — DECK OUTLINE
Plan the full deck structure holistically. You are responsible for:
- the title slide
- divider slides at genuine section boundaries only
- all content slides
- the thank-you slide

Plan exactly N content slides.
Always include exactly 1 title slide and exactly 1 thank-you slide.

Limits on Dividers
- Less than 10 content slides — maximum 2 dividers
- 10 to 15 content slides — maximum 3 dividers
- 15 to 25 content slides — maximum 4 dividers
- More than 25 content slides — maximum 6 dividers

DECK ASSEMBLY — STRUCTURAL SLIDE PLACEMENT

STEP 1 — GROUP SLIDES BY KEY MESSAGE CLUSTER
  Group slides that collectively address the same governing claim into one section.
  If a summary slide exists, use its claims as the section anchors.
  If no summary slide exists, group by thematic similarity.

STEP 2 — ASSIGN FINAL SLOT NUMBERS
  Build the ordered deck:
  1. Title slide — always slot 1
  2. For each section after the first: insert one divider, then the section's slides
  3. Thank-you slide — always last
  Assign slot_index sequentially.

STEP 3 — WRITE STRUCTURAL SLIDES (title / divider / thank_you)
  Include ALL fields for these slide types (see schema below).

For CONTENT slides — output only:
  slot_index, slide_type: "content", slide_title_draft, narrative_role
  (Do NOT write strategic_objective or key_content for content slides — those come in Phase 2)

NARRATIVE ROLE DEFINITIONS:
  summary                → Condenses key findings a board must retain.
  explainer_to_summary   → Unpacks the mechanism behind a summary claim.
  drill_down             → Decomposes an aggregate into its components.
  segmentation           → Splits a total across a meaningful dimension.
  trend_analysis         → Tracks a metric across time.
  waterfall_decomposition → Shows how a starting value reaches an ending value.
  benchmark_comparison   → Places a metric against a reference point.
  exception_highlight    → Draws attention to an anomaly or threshold breach.
  validation             → Tests a claim against independent evidence.
  context_setter         → Establishes the factual baseline before analytical slides.
  problem_statement      → Defines the problem the deck is responding to.
  risk_register          → Catalogues known risks with likelihood and impact.
  scenario_analysis      → Presents two or more plausible futures.
  decision_framework     → Structures a choice across options, criteria, trade-offs.
  recommendations        → States what the team proposes the board approve.
  methodology_note       → Documents definitions and data sources.
  additional_information → Supplementary detail supporting the deck's claims.

Return a single valid JSON object:
{
  "governing_thought": "string — the single most important insight in one punchy sentence",
  "audience": "string",
  "narrative_flow": "string",
  "data_heavy": true | false,
  "tone": "string",
  "key_messages": ["string"],
  "slides": [
    {
      "slot_index": 1,
      "slide_type": "title",
      "slide_title_draft": "string",
      "subtitle": "",
      "strategic_objective": "string",
      "narrative_role": "",
      "zone_count_signal": "unsure",
      "dominant_zone_signal": "unsure",
      "co_primary_signal": "no",
      "following_slide_claim": ""
    },
    {
      "slot_index": 2,
      "slide_type": "content",
      "slide_title_draft": "string",
      "narrative_role": "summary"
    },
    {
      "slot_index": 3,
      "slide_type": "divider",
      "slide_title_draft": "string",
      "subtitle": "",
      "strategic_objective": "string",
      "narrative_role": "",
      "zone_count_signal": "unsure",
      "dominant_zone_signal": "unsure",
      "co_primary_signal": "no",
      "following_slide_claim": ""
    }
  ]
}

CRITICAL OUTPUT RULES:
- Return ONLY the raw JSON object. No explanation. No preamble. No markdown fences.
- Your response must start with { and end with }.
- slides must contain ALL slots in final deck order.
- The deck must contain exactly 1 title, exactly N content slides, exactly 1 thank_you.
- Dividers are optional; follow the Limits on Dividers table.
- slot_index must be sequential starting at 1 with no gaps.
- Content slide slide_title_draft must be insight-led, not a topic label.
- governing_thought must be one punchy sentence a CEO finds immediately useful.
- key_messages must be specific and data-driven — no generic placeholders.`


// ─── PHASE 2 SYSTEM PROMPT ────────────────────────────────────────────────────
const AGENT3_PHASE2_SYSTEM = `You are a senior management consultant filling in the detail for specific slides
in a board-level presentation. You have already read the document and have the full deck outline.

Your task: for the content slides listed in the user message, provide the missing detail fields.

For each slide return:
{
  "slot_index": number,
  "strategic_objective": "one sentence — what this slide must achieve for the audience",
  "key_content": ["2-4 specific data points, facts, or claims from the document"],
  "zone_count_signal": "1" | "2" | "3" | "4" | "unsure",
  "dominant_zone_signal": "yes" | "no" | "unsure",
  "co_primary_signal": "yes" | "no",
  "following_slide_claim": "one-line statement of what the next slide establishes; empty string if last"
}

FIELD RULES:
- strategic_objective: one sentence, action-oriented — what must the audience believe after this slide?
- key_content: use actual numbers and facts from the document — no generic placeholders
- zone_count_signal: estimate of how many distinct content zones this slide needs
- dominant_zone_signal: "yes" if one zone carries the main insight, "no" if balanced
- co_primary_signal: "yes" only if two or more insights must receive exactly equal emphasis
- following_slide_claim: one line previewing the next slide's claim; "" if this is the last content slide

Return a JSON array of objects — one per slide in the batch, in the same order as requested.
No explanation. No preamble. No markdown fences. Start with [ and end with ].`


// ─── PHASE 1: document analysis + lightweight outline ─────────────────────────
async function _agent3Phase1(state, contentCount) {
  console.log('Agent 3 Phase 1 - analysing document, planning', contentCount, 'content slides')

  const messages = [{
    role: 'user',
    content: [
      {
        type: 'document',
        source: {
          type: 'base64',
          media_type: 'application/pdf',
          data: state.contentB64
        }
      },
      {
        type: 'text',
        text: `Analyse this document thoroughly and produce the deck outline.

Plan exactly ${contentCount} content slides, plus the full deck structure around them.
Include exactly 1 title slide and exactly 1 thank-you slide.
Insert divider slides only at genuine narrative section breaks.

For content slides, output ONLY: slot_index, slide_type, slide_title_draft, narrative_role.
For structural slides (title, divider, thank_you), output ALL fields.

Follow the top-down approach: lead with the governing thought, then prove it, then recommend action.
key_messages must be the exact sentences a consultant would put on an executive summary slide.
Return ONLY valid JSON.`
      }
    ]
  }]

  // Phase 1 only needs lightweight output: metadata + brief slide entries
  // ~60 tokens per content slot (title + role only) + ~200 per structural slide + 2000 overhead
  const maxTokens = Math.min(8000, Math.max(4000, contentCount * 60 + 2000))
  console.log('Agent 3 Phase 1 - max_tokens:', maxTokens)

  const raw = await callClaude(AGENT3_PHASE1_SYSTEM, messages, maxTokens)
  console.log('Agent 3 Phase 1 - response length:', raw.length)

  const trimmed = raw.trimEnd()
  if (trimmed && trimmed[trimmed.length - 1] !== '}') {
    console.warn('Agent 3 Phase 1 - response may be truncated. Last 100 chars:', trimmed.slice(-100))
  }

  return safeParseJSON(raw, null)
}


// ─── PHASE 2: fill detail for a batch of 2 content slides ─────────────────────
async function _agent3Phase2(state, outline, batch) {
  const batchIndexes = batch.map(s => s.slot_index).join(', ')
  console.log('Agent 3 Phase 2 - filling slots:', batchIndexes)

  const outlineSummary = (outline.slides || []).map(s =>
    s.slide_type === 'content'
      ? `  slot ${s.slot_index}: [content] "${s.slide_title_draft}" (${s.narrative_role})`
      : `  slot ${s.slot_index}: [${s.slide_type}] "${s.slide_title_draft}"`
  ).join('\n')

  const batchDescription = batch.map(s =>
    `  slot ${s.slot_index}: "${s.slide_title_draft}" — narrative role: ${s.narrative_role}`
  ).join('\n')

  const messages = [{
    role: 'user',
    content: [
      {
        type: 'document',
        source: {
          type: 'base64',
          media_type: 'application/pdf',
          data: state.contentB64
        }
      },
      {
        type: 'text',
        text: `DECK CONTEXT
Governing thought: ${outline.governing_thought || ''}
Audience: ${outline.audience || ''}
Narrative flow: ${outline.narrative_flow || ''}

FULL DECK OUTLINE
${outlineSummary}

YOUR TASK
Fill in the detail for these ${batch.length} content slide(s):
${batchDescription}

Return a JSON array with ${batch.length} object(s), one per slide above, in the same order.
Use actual numbers and facts from the document — no generic placeholders.
Return ONLY valid JSON starting with [ and ending with ].`
      }
    ]
  }]

  // ~600 tokens per slide (strategic_objective + 3-4 key_content items + signals + JSON)
  const maxTokens = Math.max(2000, batch.length * 600 + 400)

  const raw = await callClaude(AGENT3_PHASE2_SYSTEM, messages, maxTokens)
  console.log('Agent 3 Phase 2 slots', batchIndexes, '- response length:', raw.length)

  const trimmed = raw.trimEnd()
  if (trimmed && trimmed[trimmed.length - 1] !== ']') {
    console.warn('Agent 3 Phase 2 slots', batchIndexes, '- response may be truncated. Last 100 chars:', trimmed.slice(-100))
  }

  const parsed = safeParseJSON(raw, [])
  return Array.isArray(parsed) ? parsed : []
}


// ─── ASSEMBLY: merge Phase 1 outline + Phase 2 details ───────────────────────
function _agent3Assemble(outline, detailItems) {
  const detailMap = {}
  for (const d of detailItems) {
    if (d && d.slot_index != null) detailMap[d.slot_index] = d
  }

  const outlineSlides = (outline.slides || []).map(s => {
    if (s.slide_type !== 'content') {
      // Structural slides: already fully populated from Phase 1
      return {
        slide_number: s.slot_index,
        slide_type: s.slide_type,
        narrative_role: '',
        slide_title_draft: s.slide_title_draft || '',
        subtitle: s.subtitle || '',
        strategic_objective: s.strategic_objective || '',
        key_content: [],
        zone_count_signal: s.zone_count_signal || 'unsure',
        dominant_zone_signal: s.dominant_zone_signal || 'unsure',
        co_primary_signal: s.co_primary_signal || 'no',
        following_slide_claim: s.following_slide_claim || ''
      }
    }

    // Content slides: merge outline + phase 2 detail
    const d = detailMap[s.slot_index] || {}
    return {
      slide_number: s.slot_index,
      slide_type: 'content',
      narrative_role: s.narrative_role || 'explainer_to_summary',
      slide_title_draft: s.slide_title_draft || '',
      subtitle: '',
      strategic_objective: d.strategic_objective || '',
      key_content: Array.isArray(d.key_content) ? d.key_content : [],
      zone_count_signal: String(d.zone_count_signal || 'unsure'),
      dominant_zone_signal: String(d.dominant_zone_signal || 'unsure'),
      co_primary_signal: String(d.co_primary_signal || 'no'),
      following_slide_claim: d.following_slide_claim || ''
    }
  })

  // Sort by slot_index and renumber
  const sorted = outlineSlides
    .sort((a, b) => a.slide_number - b.slide_number)
    .map((s, i, arr) => ({
      ...s,
      slide_number: i + 1,
      slide_type: ['title', 'divider', 'content', 'thank_you'].includes(s.slide_type) ? s.slide_type : 'content',
      narrative_role: s.slide_type === 'content' ? (s.narrative_role || 'explainer_to_summary') : '',
      zone_count_signal: ['1', '2', '3', '4', 'unsure'].includes(String(s.zone_count_signal)) ? String(s.zone_count_signal) : 'unsure',
      dominant_zone_signal: ['yes', 'no', 'unsure'].includes(String(s.dominant_zone_signal)) ? String(s.dominant_zone_signal) : 'unsure',
      co_primary_signal: ['yes', 'no'].includes(String(s.co_primary_signal)) ? String(s.co_primary_signal) : 'no',
      following_slide_claim: s.following_slide_claim || (arr[i + 1] ? (arr[i + 1].slide_title_draft || '') : '')
    }))

  return {
    governing_thought: outline.governing_thought || '',
    audience: outline.audience || '',
    narrative_flow: outline.narrative_flow || '',
    data_heavy: outline.data_heavy !== false,
    tone: outline.tone || 'confident',
    key_messages: Array.isArray(outline.key_messages) ? outline.key_messages : [],
    slides: sorted
  }
}


// ─── VALIDATION ───────────────────────────────────────────────────────────────
function _agent3Validate(brief, contentCount) {
  const slides = brief.slides || []
  const contentSlides = slides.filter(s => s.slide_type === 'content')
  const structuralCounts = slides.reduce((acc, s) => {
    acc[s.slide_type] = (acc[s.slide_type] || 0) + 1
    return acc
  }, {})
  const maxDividers = contentCount < 10 ? 2 : contentCount <= 15 ? 3 : contentCount <= 25 ? 4 : 6

  if (contentSlides.length !== contentCount) {
    console.warn('Agent 3 - content slide count mismatch: got', contentSlides.length, 'expected', contentCount)
    return false
  }
  if ((structuralCounts.title || 0) !== 1) {
    console.warn('Agent 3 - expected 1 title slide, got', structuralCounts.title || 0)
    return false
  }
  if ((structuralCounts.thank_you || 0) !== 1) {
    console.warn('Agent 3 - expected 1 thank_you slide, got', structuralCounts.thank_you || 0)
    return false
  }
  if ((structuralCounts.divider || 0) > maxDividers) {
    console.warn('Agent 3 - too many dividers:', structuralCounts.divider, 'max:', maxDividers)
    return false
  }
  return true
}


// ─── MAIN ENTRY POINT ─────────────────────────────────────────────────────────
async function runAgent3(state) {
  const contentCount = state.slideCount
  console.log('Agent 3 starting (two-phase) - content slides:', contentCount)

  const fallback = buildFallbackBrief(contentCount)

  // ── Phase 1: read document, build outline ──────────────────────────────────
  const outline = await _agent3Phase1(state, contentCount)

  if (!outline || !Array.isArray(outline.slides) || outline.slides.length === 0) {
    console.warn('Agent 3 Phase 1 - failed to produce outline, using fallback')
    return fallback
  }

  const contentSlots = outline.slides.filter(s => s.slide_type === 'content')
  if (contentSlots.length !== contentCount) {
    console.warn('Agent 3 Phase 1 - content slot count mismatch: got', contentSlots.length, 'expected', contentCount, '— using fallback')
    return fallback
  }

  console.log('Agent 3 Phase 1 complete -', outline.slides.length, 'total slots,', contentSlots.length, 'content slots')

  // ── Phase 2: fill detail in parallel batches of 2 ─────────────────────────
  const batches = []
  for (let i = 0; i < contentSlots.length; i += 2) {
    batches.push(contentSlots.slice(i, i + 2))
  }

  console.log('Agent 3 Phase 2 - firing', batches.length, 'batch(es) in parallel')
  const batchResults = await Promise.all(
    batches.map(batch => _agent3Phase2(state, outline, batch))
  )

  const allDetails = batchResults.flat()
  console.log('Agent 3 Phase 2 complete -', allDetails.length, 'slides detailed')

  // ── Assembly ───────────────────────────────────────────────────────────────
  const brief = _agent3Assemble(outline, allDetails)

  if (!_agent3Validate(brief, contentCount)) {
    console.warn('Agent 3 - assembled brief failed validation, using fallback')
    return fallback
  }

  console.log('Agent 3 complete - brief summary:')
  console.log('  Governing thought:', brief.governing_thought)
  console.log('  Total slides:', brief.slides.length)
  console.log('  Content slides:', brief.slides.filter(s => s.slide_type === 'content').length)
  console.log('  Key messages:', (brief.key_messages || []).length)
  console.log('  Data heavy:', brief.data_heavy)

  return brief
}


function buildFallbackBrief(contentSlideCount) {
  const slides = [{
    slide_number: 1,
    slide_type: 'title',
    narrative_role: '',
    slide_title_draft: 'Executive Review',
    subtitle: '',
    strategic_objective: 'Frame the presentation for senior management.',
    key_content: [],
    zone_count_signal: 'unsure',
    dominant_zone_signal: 'unsure',
    co_primary_signal: 'no',
    following_slide_claim: ''
  }]

  for (let i = 1; i <= contentSlideCount; i++) {
    slides.push({
      slide_number: i + 1,
      slide_type: 'content',
      narrative_role: i === contentSlideCount ? 'recommendations' : 'explainer_to_summary',
      slide_title_draft: 'Content slide ' + i,
      subtitle: '',
      strategic_objective: 'Advance the audience toward the governing thought.',
      key_content: ['Content from uploaded document'],
      zone_count_signal: '2',
      dominant_zone_signal: 'yes',
      co_primary_signal: 'no',
      following_slide_claim: i < contentSlideCount ? ('Content slide ' + (i + 1)) : 'Close with clear next steps.'
    })
  }

  slides.push({
    slide_number: contentSlideCount + 2,
    slide_type: 'thank_you',
    narrative_role: '',
    slide_title_draft: 'Thank You',
    subtitle: '',
    strategic_objective: 'Close the presentation and reinforce the next step.',
    key_content: [],
    zone_count_signal: 'unsure',
    dominant_zone_signal: 'unsure',
    co_primary_signal: 'no',
    following_slide_claim: ''
  })

  return {
    governing_thought: 'Key insights from the document require further analysis.',
    audience: 'Senior Management',
    narrative_flow: 'Situation -> Analysis -> Recommendations',
    data_heavy: false,
    tone: 'confident',
    key_messages: [
      'Please review the source document for key messages',
      'Analysis could not be completed automatically',
      'Manual review recommended'
    ],
    slides
  }
}
