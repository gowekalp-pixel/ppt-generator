// ─── AGENT 3 — DOCUMENT ANALYST & PRESENTATION STRATEGIST ────────────────────
// Input:  state.contentB64 (PDF), state.slideCount
// Output: presentationBrief — a rich JSON object that gives Agent 4
//         everything it needs to write the actual slides
//
// Agent 3 does NOT create slides.
// Agent 3 deeply understands the document and defines the presentation strategy.
// Agent 4 uses this brief to build the slide content.

const AGENT3_SYSTEM = `You are a senior management consultant with deep expertise in financial analysis, 
strategy, and board-level communication. You have been asked to review a document and 
prepare a comprehensive presentation brief for your team who will build the actual slides.

Your job has TWO parts:

PART 1 — DOCUMENT ANALYSIS
Read the document thoroughly and extract:
- What type of document is this? (financial report, market research, strategy doc, operational review etc)
- What is the single most important insight or message in this document?
- What are the 3-5 key messages a senior leader must walk away with?
- What data points, metrics, or facts are most significant?
- What problems or risks are identified?
- What recommendations or next steps are suggested?
- Is this data-heavy (lots of numbers/charts) or narrative-heavy (qualitative analysis)?

PART 2 — PRESENTATION STRATEGY (Top-Down, McKinsey Style)
Structure the presentation using the top-down approach:
- Lead with the answer — what is the governing thought? (the single sentence that summarises everything)
- Then prove it — what sections of evidence support this governing thought?
- Then act — what should the audience do as a result?

Think carefully about the right narrative flow for THIS specific content:
- For financial content: Situation → Performance → Drivers → Outlook → Actions
- For market research: Context → Market Dynamics → Competitive Position → Implications → Recommendations  
- For strategy: Objective → Current State → Gap Analysis → Options → Recommended Path
- For operational: Baseline → Issues → Root Cause → Solutions → Implementation

Return a single valid JSON object with this exact structure:
{
  "document_type": "string — what kind of document this is",
  "document_summary": "string — 2-3 sentence summary of the entire document",
  "governing_thought": "string — the single most important insight in one punchy sentence",
  "audience": "string — who this presentation is for",
  "narrative_flow": "string — name of the flow pattern being used",
  "data_heavy": true or false,
  "key_messages": [
    "string — key message 1 with specific data if available",
    "string — key message 2 with specific data if available",
    "string — key message 3 with specific data if available"
  ],
  "key_data_points": [
    "string — important metric or fact from the document",
    "string — important metric or fact from the document"
  ],
  "risks_and_issues": [
    "string — identified risk or problem"
  ],
  "recommendations": [
    "string — recommended action or next step"
  ],
  "sections": [
    {
      "section_number": 1,
      "section_name": "string — name of this section",
      "section_type": "string — one of: title, executive_summary, divider, financial_data, market_analysis, strategic_analysis, operational_review, recommendations, appendix, conclusion",
      "purpose": "string — what this section must achieve for the audience",
      "key_content": ["string — specific content point to cover in this section"],
      "suggested_slide_count": number,
      "data_available": true or false,
      "so_what": "string — the insight or takeaway from this section, not just the data"
    }
  ],
  "total_slides": number,
  "opening_guidance": "string — how Agent 4 should open the presentation",
  "closing_guidance": "string — how Agent 4 should close the presentation",
  "tone": "string — recommended tone (e.g. confident, cautious, urgent, balanced)"
}

IMPORTANT RULES:
- total_slides must equal exactly the number requested
- The sum of suggested_slide_count across all sections must equal total_slides
- Always include a title slide (1 slide) and conclusion/next steps (1 slide)
- Include 2-3 section divider slides for a clean structure
- key_messages must be specific and data-driven where possible — not generic
- governing_thought must be a single sentence that a CEO would find immediately useful
- Return ONLY valid JSON. No explanation. No markdown fences.`


async function runAgent3(state) {
  console.log('Agent 3 starting — analysing document, slide count:', state.slideCount)

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

The final presentation must have exactly ${state.slideCount} slides total.
The sum of all suggested_slide_count values in sections must equal ${state.slideCount}.

Remember:
- Top-down approach: lead with the governing thought, then prove it, then recommend action
- Be specific — use actual numbers and facts from the document, not generic placeholders
- The key_messages should be the exact sentences a consultant would put on an executive summary slide
- Return ONLY valid JSON.`
      }
    ]
  }]

  const raw = await callClaude(AGENT3_SYSTEM, messages, 3000)

  console.log('Agent 3 — raw response length:', raw.length)

  const fallback = buildFallbackBrief(state.slideCount)
  const brief    = safeParseJSON(raw, fallback)

  // ── Validate and fix ──────────────────────────────────────────────────────

  // Ensure it has the required fields
  if (!brief.sections || !Array.isArray(brief.sections) || brief.sections.length === 0) {
    console.warn('Agent 3 — no sections in brief, using fallback')
    return fallback
  }

  // Enforce total_slides matches state.slideCount
  if (brief.total_slides !== state.slideCount) {
    console.warn('Agent 3 — total_slides mismatch, correcting:', brief.total_slides, '→', state.slideCount)
    brief.total_slides = state.slideCount
  }

  // Fix section slide counts if they don't add up
  const sectionTotal = brief.sections.reduce((sum, s) => sum + (s.suggested_slide_count || 1), 0)
  if (sectionTotal !== state.slideCount) {
    console.warn('Agent 3 — section slide counts sum to', sectionTotal, 'not', state.slideCount, '— redistributing')
    brief.sections = redistributeSlides(brief.sections, state.slideCount)
  }

  // Ensure every section has required fields
  brief.sections = brief.sections.map((s, i) => ({
    section_number:       s.section_number       || (i + 1),
    section_name:         s.section_name         || 'Section ' + (i + 1),
    section_type:         s.section_type         || 'content',
    purpose:              s.purpose              || '',
    key_content:          Array.isArray(s.key_content) ? s.key_content : [],
    suggested_slide_count:s.suggested_slide_count|| 1,
    data_available:       s.data_available       !== undefined ? s.data_available : false,
    so_what:              s.so_what              || ''
  }))

  console.log('Agent 3 complete — brief summary:')
  console.log('  Document type:', brief.document_type)
  console.log('  Governing thought:', brief.governing_thought)
  console.log('  Sections:', brief.sections.length)
  console.log('  Total slides:', brief.total_slides)
  console.log('  Key messages:', (brief.key_messages || []).length)
  console.log('  Data heavy:', brief.data_heavy)

  return brief
}


// ─── SLIDE REDISTRIBUTION ────────────────────────────────────────────────────
// If section slide counts don't add up to the target, redistribute fairly

function redistributeSlides(sections, target) {
  // Give each section at least 1 slide
  const base     = sections.map(s => ({ ...s, suggested_slide_count: 1 }))
  let remaining  = target - base.length

  if (remaining < 0) {
    // Too many sections — trim to target
    return base.slice(0, target)
  }

  // Distribute remaining slides to content sections (not title/divider)
  const contentIdxs = base
    .map((s, i) => ({ i, type: s.section_type }))
    .filter(s => !['title', 'divider'].includes(s.type))
    .map(s => s.i)

  let idx = 0
  while (remaining > 0) {
    const i = contentIdxs[idx % contentIdxs.length]
    base[i].suggested_slide_count++
    remaining--
    idx++
  }

  return base
}


// ─── FALLBACK BRIEF ──────────────────────────────────────────────────────────
// Used if Claude fails to return valid JSON

function buildFallbackBrief(slideCount) {
  const contentSlides = slideCount - 4 // minus title, exec summary, divider, conclusion

  return {
    document_type:    'Business Document',
    document_summary: 'Document analysis failed — Agent 4 will work from the raw content.',
    governing_thought:'Key insights from the document require further analysis.',
    audience:         'Senior Management',
    narrative_flow:   'Situation → Analysis → Recommendations',
    data_heavy:       false,
    key_messages:     [
      'Please review the source document for key messages',
      'Analysis could not be completed automatically',
      'Manual review recommended'
    ],
    key_data_points:  [],
    risks_and_issues: [],
    recommendations:  [],
    sections: [
      {
        section_number:        1,
        section_name:         'Title',
        section_type:         'title',
        purpose:              'Opening title slide',
        key_content:          [],
        suggested_slide_count: 1,
        data_available:        false,
        so_what:              ''
      },
      {
        section_number:        2,
        section_name:         'Executive Summary',
        section_type:         'executive_summary',
        purpose:              'Key takeaways for senior management',
        key_content:          ['Key message 1', 'Key message 2', 'Key message 3'],
        suggested_slide_count: 1,
        data_available:        false,
        so_what:              'The three things the audience must remember'
      },
      {
        section_number:        3,
        section_name:         'Main Content',
        section_type:         'strategic_analysis',
        purpose:              'Core content from the document',
        key_content:          ['Content from uploaded document'],
        suggested_slide_count: Math.max(1, contentSlides),
        data_available:        true,
        so_what:              'Key insight from the content'
      },
      {
        section_number:        4,
        section_name:         'Next Steps',
        section_type:         'conclusion',
        purpose:              'Recommended actions and owners',
        key_content:          ['Action items', 'Owners', 'Timeline'],
        suggested_slide_count: 1,
        data_available:        false,
        so_what:              'What the audience should do after this presentation'
      }
    ],
    total_slides:      slideCount,
    opening_guidance: 'Lead with the governing thought on the title slide',
    closing_guidance: 'End with clear, specific, owned next steps',
    tone:             'confident'
  }
}
