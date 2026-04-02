// ─── AGENT 3 — DOCUMENT ANALYST & PRESENTATION STRATEGIST ────────────────────
// Input:  state.contentB64 (PDF), state.slideCount
// Output: presentationBrief — deck context + content slide plan for Agent 4
//
// Agent 3 plans CONTENT SLIDES ONLY.
// Structural slides (title, dividers, thank-you) are added by Agent 4.

const AGENT3_SYSTEM = `You are a senior management consultant with deep expertise in financial analysis,
strategy, and board-level communication. You have been asked to review a document and
plan the content slides for a board-level presentation.

Your job has TWO parts:

PART 1 — DOCUMENT ANALYSIS
Read the document and identify:
- What is the single most important insight — one punchy sentence a CEO finds immediately useful?
- What narrative flow best suits this content?
  - Financial:        Situation → Performance → Drivers → Outlook → Actions
  - Market research:  Context → Market Dynamics → Competitive Position → Implications → Recommendations
  - Strategy:         Objective → Current State → Gap Analysis → Options → Recommended Path
  - Operational:      Baseline → Issues → Root Cause → Solutions → Implementation
- Who is the audience and what tone is appropriate?

PART 2 — CONTENT SLIDE PLAN
Plan exactly N content slides following the narrative flow above.
Structural slides (title, dividers, thank-you) are NOT your responsibility — do not include them.

For each content slide define:
  - content_slide_number: sequential integer starting at 1
  - purpose: one sentence — what this slide must achieve for the audience
  - high_level_narrative: the analytical story this slide tells and the specific insight it delivers
  - key_content: 2–4 specific data points, facts, or claims from the document this slide will use

Return a single valid JSON object:
{
  "governing_thought": "string — the single most important insight in one punchy sentence",
  "audience": "string — who this presentation is for",
  "narrative_flow": "string — name of the flow pattern being used",
  "data_heavy": true or false,
  "tone": "string — recommended tone (confident / cautious / urgent / balanced)",
  "key_messages": [
    "string — key message with specific data from the document"
  ],
  "content_slides": [
    {
      "content_slide_number": 1,
      "purpose": "string",
      "high_level_narrative": "string",
      "key_content": ["string", "string"]
    }
  ]
}

CRITICAL OUTPUT RULES:
- Return ONLY the raw JSON object. No explanation. No preamble. No markdown fences.
- Your response must start with { and end with }.
- content_slides must contain exactly N items — one per requested content slide.
- Every key_content item must reference actual content from the document — no generic placeholders.
- high_level_narrative must be a specific analytical claim, not a topic label.
- governing_thought must be a single punchy sentence a CEO finds immediately useful.
- key_messages must be specific and data-driven — no generic placeholders.`


async function runAgent3(state) {
  const contentCount = state.slideCount  // content slides only — structural added by Agent 4

  console.log('Agent 3 starting — content slides:', contentCount)

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
        text: `Analyse this document thoroughly and produce a presentation brief.

Plan exactly ${contentCount} content slides.
Do NOT include structural slides (title, dividers, thank-you) — Agent 4 handles those.

Follow the top-down approach: lead with the governing thought, then prove it, then recommend action.
Be specific — use actual numbers and facts from the document, not generic placeholders.
key_messages must be the exact sentences a consultant would put on an executive summary slide.
Return ONLY valid JSON.`
      }
    ]
  }]

  const raw = await callClaude(AGENT3_SYSTEM, messages, 8000)

  console.log('Agent 3 — raw response length:', raw.length)
  if (raw.length < 200) console.warn('Agent 3 — suspiciously short response:', raw)

  const fallback = buildFallbackBrief(contentCount)
  const brief    = safeParseJSON(raw, fallback)

  if (!brief.content_slides || !Array.isArray(brief.content_slides) || brief.content_slides.length === 0) {
    console.warn('Agent 3 — no content_slides in brief, using fallback')
    return fallback
  }

  // Ensure correct count
  if (brief.content_slides.length !== contentCount) {
    console.warn('Agent 3 — content_slides count is', brief.content_slides.length, 'not', contentCount, '— trimming/padding')
    while (brief.content_slides.length < contentCount) {
      const n = brief.content_slides.length + 1
      brief.content_slides.push({
        content_slide_number: n,
        purpose:              'Additional content slide',
        high_level_narrative: '',
        key_content:          []
      })
    }
    brief.content_slides = brief.content_slides.slice(0, contentCount)
  }

  // Normalize fields
  brief.content_slides = brief.content_slides.map((s, i) => ({
    content_slide_number: s.content_slide_number || (i + 1),
    purpose:              s.purpose              || '',
    high_level_narrative: s.high_level_narrative || '',
    key_content:          Array.isArray(s.key_content) ? s.key_content : []
  }))

  console.log('Agent 3 complete — brief summary:')
  console.log('  Governing thought:', brief.governing_thought)
  console.log('  Content slides:',    brief.content_slides.length)
  console.log('  Key messages:',      (brief.key_messages || []).length)
  console.log('  Data heavy:',        brief.data_heavy)

  return brief
}


// ─── FALLBACK BRIEF ──────────────────────────────────────────────────────────

function buildFallbackBrief(contentSlideCount) {
  const slides = []
  for (let i = 1; i <= contentSlideCount; i++) {
    slides.push({
      content_slide_number: i,
      purpose:              'Content slide ' + i,
      high_level_narrative: '',
      key_content:          ['Content from uploaded document']
    })
  }

  return {
    governing_thought: 'Key insights from the document require further analysis.',
    audience:          'Senior Management',
    narrative_flow:    'Situation → Analysis → Recommendations',
    data_heavy:        false,
    tone:              'confident',
    key_messages:      [
      'Please review the source document for key messages',
      'Analysis could not be completed automatically',
      'Manual review recommended'
    ],
    content_slides: slides
  }
}


// ─── NARRATIVE ROLE INFERENCE ─────────────────────────────────────────────────
// Fallback used by Agent 4 when LLM does not return a narrative_role for a slide.

function inferNarrativeRole(section, index) {
  const legacyMap = {
    title:              'title',
    executive_summary:  'summary',
    divider:            'divider',
    recommendations:    'recommendations',
    conclusion:         'recommendations',
    appendix:           'additional_information',
    financial_data:     'explainer_to_summary',
    market_analysis:    'explainer_to_summary',
    strategic_analysis: 'explainer_to_summary',
    operational_review: 'explainer_to_summary',
  }

  if (section.section_type && legacyMap[section.section_type]) {
    return legacyMap[section.section_type]
  }

  if (index === 0) return 'title'
  return 'explainer_to_summary'
}
