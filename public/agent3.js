// AGENT 3 - DOCUMENT ANALYST & DECK PLANNER
// Input:  state.contentB64 (PDF), state.slideCount
// Output: presentationBrief - deck context + full deck slide plan for Agent 4
//
// Agent 3 plans the FULL DECK:
// title, dividers, content slides, and thank-you.

const AGENT3_SYSTEM = `You are a senior management consultant with deep expertise in financial analysis,
strategy, and board-level communication. You have been asked to review a document and
plan a board-level presentation.

Your job has TWO parts:

PART 1 - DOCUMENT ANALYSIS
Read the document and identify:
- What is the single most important insight - one punchy sentence a CEO finds immediately useful?
- What narrative flow best suits this content?
  - Financial: Situation -> Performance -> Drivers -> Outlook -> Actions
  - Market research: Context -> Market Dynamics -> Competitive Position -> Implications -> Recommendations
  - Strategy: Objective -> Current State -> Gap Analysis -> Options -> Recommended Path
  - Operational: Baseline -> Issues -> Root Cause -> Solutions -> Implementation
- Who is the audience and what tone is appropriate?

PART 2 - FULL DECK PLAN
Plan the full deck holistically.
You are responsible for:
- the title slide
- divider slides at genuine section boundaries only
- all content slides
- the thank-you slide

Plan exactly N content slides.
Always include exactly 1 title slide and exactly 1 thank-you slide.
You may insert up to 3 divider slides if the narrative genuinely benefits from them.

For each slide define:
- slide_number: final sequential deck number starting at 1
- slide_type: one of "title", "divider", "content", "thank_you"
- narrative_role: required for content slides; use "" for title/divider/thank_you
- slide_title_draft: draft title for the slide
- strategic_objective: one sentence - what this slide must achieve for the audience
- key_content: 2-4 specific data points, facts, or claims from the document this slide will use
- zone_count_signal: 1 | 2 | 3 | 4 | unsure
- dominant_zone_signal: yes | no | unsure
- co_primary_signal: yes | no
- following_slide_claim: one-line statement of what the next slide will establish; use "" if not applicable

Each content slide plan carries purpose, high_level_narrative, and key_content from Agent 3.
Read these three fields and assign exactly one narrative_role from the definitions below.
Lock it before doing anything else — it controls slide_intent defaults and Phase 3 artifact gates.

  narrative_role tells you:
  - What analytical job this slide has in the overall proof chain
  - What slide_intent to use as default (see mapping below)
  - What artifact constraints apply in Phase 3

  NARRATIVE ROLE DEFINITIONS:
  summary                → A single slide that condenses the key findings of a section or deck into
                           the fewest possible claims the board needs to retain.
  explainer_to_summary   → Unpacks the mechanism or logic behind a summary claim — answers "how
                           does this number work" before the board asks.
  drill_down             → Decomposes an aggregate from summary or explainer_to_summary into its components to show where the result
                           comes from and which component drives it.
  segmentation           → Splits a total across a meaningful dimension (geography, product,
                           customer type) to reveal which segment is driving performance.
  trend_analysis         → Tracks a metric across time to identify direction, inflection points,
                           and whether the current position is improving or deteriorating.
  waterfall_decomposition → Shows how a starting value reaches an ending value through a sequence
                            of additive and subtractive components.
  benchmark_comparison   → Places a metric against an external or internal reference point to
                           establish whether performance is strong, weak, or on-par.
  exception_highlight    → Draws attention to an anomaly, outlier, or threshold breach that
                           requires the board's awareness or action.
  validation             → Tests a claim, assumption, or model output against independent
                           evidence to confirm or challenge its credibility.
  context_setter         → Establishes the factual baseline — market conditions, prior period,
                           or structural constraints — before analytical slides make claims.
  problem_statement      → Defines the problem the deck is responding to: its scale, cause,
                           and consequence if unaddressed.
  risk_register          → Catalogues known risks with likelihood and impact, so the board
                           can prioritise oversight and mitigation.
  scenario_analysis      → Presents two or more plausible futures under different assumptions
                           so the board can understand the range of outcomes.
  decision_framework     → Structures a choice by laying out options, criteria, and trade-offs
                           so the board can make a well-reasoned decision.
  recommendations        → States what the team proposes the board approve, fund, or direct —
                           with the rationale and the ask made explicit.
  methodology_note       → Documents definitions, data sources, or calculation logic that the
                           board needs to trust the numbers but does not need to analyse.
  additional_information → Supplementary detail that supports the deck's claims but does not
                           carry a standalone analytical argument.

NARRATIVE ROLE → slide_intent DEFAULT MAPPING:
  Use this as the starting position. Override only if content analysis gives strong reason.

  summary                    → prove
  explainer_to_summary       → prove
  drill_down                 → explain
  segmentation               → prove
  trend_analysis             → explain
  waterfall_decomposition    → explain
  benchmark_comparison       → prove
  exception_highlight        → alert
  validation                 → prove
  problem_statement          → alert
  context_setter             → update
  risk_register              → alert
  scenario_analysis          → decide
  decision_framework         → decide
  recommendations            → decide
  methodology_note           → explain
  additional_information     → explain

  
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DECK ASSEMBLY — STRUCTURAL SLIDE PLACEMENT
(run once after Phase 1 is complete for all content slides)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

STEP 1 — GROUP SLIDES BY KEY MESSAGE CLUSTER
  Read the key_message of every content slide locked in Phase 1.
  Group slides that collectively address the same governing claim or theme into one section.
  The grouping rule: slides whose key_messages all support the same top-level point
  belong in the same section — they should not be separated by a divider.

  If a summary slide exists, use its claims as the section anchors:
  each section = all slides whose key_messages prove one claim on the summary slide.

  If no summary slide exists, group by thematic similarity of key_messages.

Limits on Dividers
- Less than 10 content slides - maximum 2 dividers
- 10 to 15 content slides - maximum 3 dividers
- 15 to 25 content slides - maximum 4 dividers
- More than 25 content slides - maximum 6 dividers

STEP 2 — ASSIGN FINAL SLIDE NUMBERS
  Build the ordered deck:
  1. Title slide — always position 1 (slide_type: "title")
  2. For each section after the first: insert one divider, then the section's slides
  3. Thank-you slide — always last (slide_type: "thank_you")
  Assign slide_number sequentially across the full deck.

STEP 3 — WRITE STRUCTURAL SLIDES

  title (slide_type: "title"):
    - title: short presentation name, 4–8 words
    - subtitle: audience / date / context if relevant
    - key_message: governing thought of the full deck
    - zones: []

  divider (slide_type: "divider"):
    - title: name of the section that follows — derived from the shared theme of its slides, 3–5 words
    - subtitle: empty string
    - key_message: one sentence — what this group of slides will collectively prove
    - zones: []

  thank_you (slide_type: "thank_you"):
    - title: "Thank You" or equivalent closing phrase
    - subtitle: presenter name / contact if relevant
    - key_message: one sentence — what the audience must do next
    - zones: []

Return a single valid JSON object:
{
  "governing_thought": "string - the single most important insight in one punchy sentence",
  "audience": "string - who this presentation is for",
  "narrative_flow": "string - name of the flow pattern being used",
  "data_heavy": true,
  "tone": "string - recommended tone",
  "key_messages": ["string"],
  "slides": [
    {
      "slide_number": 1,
      "slide_type": "title",
      "narrative_role": "",
      "slide_title_draft": "string",
      "strategic_objective": "string",
      "key_content": ["string", "string"],
      "zone_count_signal": "unsure",
      "dominant_zone_signal": "unsure",
      "co_primary_signal": "no",
      "following_slide_claim": "string"
    }
  ]
}

CRITICAL OUTPUT RULES:
- Return ONLY the raw JSON object. No explanation. No preamble. No markdown fences.
- Your response must start with { and end with }.
- slides must contain the full ordered deck.
- The deck must contain exactly 1 title slide, exactly N content slides, and exactly 1 thank_you slide.
- Divider slides are optional, but use them only at genuine section boundaries and never more than 3.
- slide_number must be final deck numbering after structural slides are inserted.
- Every key_content item must reference actual content from the document - no generic placeholders.
- slide_title_draft must be specific and insight-led for content slides, not a topic label.
- governing_thought must be a single punchy sentence a CEO finds immediately useful.
- key_messages must be specific and data-driven - no generic placeholders.`


async function runAgent3(state) {
  const contentCount = state.slideCount

  console.log('Agent 3 starting - content slides:', contentCount)

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
        text: `Analyse this document thoroughly and produce a presentation brief.

Plan exactly ${contentCount} content slides, plus the full deck structure around them.
Include exactly 1 title slide and exactly 1 thank-you slide.
Insert divider slides only when they reflect real narrative section breaks.

Follow the top-down approach: lead with the governing thought, then prove it, then recommend action.
Be specific - use actual numbers and facts from the document, not generic placeholders.
key_messages must be the exact sentences a consultant would put on an executive summary slide.
Return ONLY valid JSON.`
      }
    ]
  }]

  const raw = await callClaude(AGENT3_SYSTEM, messages, 8000)

  console.log('Agent 3 - raw response length:', raw.length)
  if (raw.length < 200) console.warn('Agent 3 - suspiciously short response:', raw)

  const fallback = buildFallbackBrief(contentCount)
  const brief = safeParseJSON(raw, fallback)

  if (!brief.slides || !Array.isArray(brief.slides) || brief.slides.length === 0) {
    console.warn('Agent 3 - no slides in brief, using fallback')
    return fallback
  }

  let slides = brief.slides.map((s, i) => ({
    slide_number: Number(s.slide_number) || (i + 1),
    slide_type: String(s.slide_type || '').toLowerCase(),
    narrative_role: s.narrative_role || '',
    slide_title_draft: s.slide_title_draft || '',
    strategic_objective: s.strategic_objective || '',
    key_content: Array.isArray(s.key_content) ? s.key_content : [],
    zone_count_signal: String(s.zone_count_signal || 'unsure'),
    dominant_zone_signal: String(s.dominant_zone_signal || 'unsure'),
    co_primary_signal: String(s.co_primary_signal || 'no'),
    following_slide_claim: s.following_slide_claim || ''
  }))

  const contentSlides = slides.filter(s => s.slide_type === 'content')
  const structuralCounts = slides.reduce((acc, s) => {
    acc[s.slide_type] = (acc[s.slide_type] || 0) + 1
    return acc
  }, {})

  if (contentSlides.length !== contentCount) {
    console.warn('Agent 3 - content slide count is', contentSlides.length, 'not', contentCount, '- using fallback')
    return fallback
  }
  if ((structuralCounts.title || 0) !== 1 || (structuralCounts.thank_you || 0) !== 1 || (structuralCounts.divider || 0) > 3) {
    console.warn('Agent 3 - invalid structural slide counts, using fallback')
    return fallback
  }

  slides = slides
    .sort((a, b) => a.slide_number - b.slide_number)
    .map((s, i, arr) => ({
      ...s,
      slide_number: i + 1,
      slide_type: ['title', 'divider', 'content', 'thank_you'].includes(s.slide_type) ? s.slide_type : 'content',
      narrative_role: s.slide_type === 'content' ? (s.narrative_role || 'explainer_to_summary') : '',
      key_content: Array.isArray(s.key_content) ? s.key_content : [],
      zone_count_signal: ['1', '2', '3', '4', 'unsure'].includes(String(s.zone_count_signal)) ? String(s.zone_count_signal) : 'unsure',
      dominant_zone_signal: ['yes', 'no', 'unsure'].includes(String(s.dominant_zone_signal)) ? String(s.dominant_zone_signal) : 'unsure',
      co_primary_signal: ['yes', 'no'].includes(String(s.co_primary_signal)) ? String(s.co_primary_signal) : 'no',
      following_slide_claim: s.following_slide_claim || (arr[i + 1] ? (arr[i + 1].slide_title_draft || '') : '')
    }))

  brief.slides = slides

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
    strategic_objective: 'Frame the presentation for senior management.',
    key_content: ['Title slide for the generated presentation'],
    zone_count_signal: 'unsure',
    dominant_zone_signal: 'unsure',
    co_primary_signal: 'no',
    following_slide_claim: 'Key insights from the document require further analysis.'
  }]

  for (let i = 1; i <= contentSlideCount; i++) {
    slides.push({
      slide_number: i + 1,
      slide_type: 'content',
      narrative_role: i === contentSlideCount ? 'recommendations' : 'explainer_to_summary',
      slide_title_draft: 'Content slide ' + i,
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
    strategic_objective: 'Close the presentation and reinforce the next step.',
    key_content: ['Closing slide'],
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
