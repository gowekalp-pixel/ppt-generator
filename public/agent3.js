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

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
NARRATIVE ROLES — assign one to every section
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Each section must be assigned a narrative_role from the list below.
The role defines the analytical job of the slide and controls how Agent 4 renders it.

STRUCTURAL (no data, no proof chain — always 1 slide each):
  "title"                — Opening slide. Framing only, no analysis.
  "divider"              — Section break between major analytical blocks.
  "transition_narrative" — Connects two analytical sections: summarises what was just proven,
                           sets up what comes next. Narrative only, no new data.

OPENING / FRAMING:
  "context_setter"       — Establishes factual baseline before any claim is made: portfolio size,
                           time period, scope, definitions. Neutral — no governing thought yet.
  "problem_statement"    — States the complication or trigger: what changed, what risk emerged,
                           what decision must be made. The "Complication" in SCR structure.
  "summary"              — THE governing thought slide. States the conclusion first. All downstream
                           proof slides exist to validate this one slide. Maximum 1 per deck.
                           Agent 4 will register all cards/KPIs on this slide and deduplicate
                           them against subsequent slides.

                           WHEN TO INCLUDE a summary slide:
                           ✓ Board or C-suite audience who need the conclusion in 30 seconds
                           ✓ Multi-thread decks (6+ slides) where several independent analytical
                             sections need a single governing thought to unify them
                           ✓ Decision-seeking decks where the audience must approve or act
                           ✓ Situation is complex or politically sensitive — leading with the
                             answer frames everything that follows

                           WHEN TO SKIP the summary slide:
                           ✗ Short focused decks (≤5 slides) on a single topic — dive straight in
                           ✗ Sequential process or methodology presentations — logic unfolds step by step
                           ✗ Pure data/appendix decks — no governing thought, just reference material
                           ✗ Exploratory or hypothesis-generating decks where the conclusion is
                             not yet known — use problem_statement instead to open with the question
                           ✗ Operational updates or status reports where the audience just needs
                             the facts in order, not a pre-stated verdict

PROOF CHAIN (main analytical body — prove the summary claims):
  "explainer_to_summary" — Directly proves one specific claim on the summary slide. First level
                           of the proof hierarchy. Must set proves_claim to the summary bullet
                           it supports. May share data with summary but must go one analytical
                           level deeper — not just restate.
  "drill_down"           — Zooms into a specific sub-segment of an explainer slide: one geography,
                           one cohort, one industry. Narrow scope. Must set proves_claim to the
                           explainer section it supports.
  "segmentation"         — Compares multiple sub-groups side by side to reveal which are
                           outperforming, lagging, or anomalous. Different from drill_down which
                           goes deep into ONE segment — segmentation shows MULTIPLE segments.
  "trend_analysis"       — Specifically examines how a metric has evolved over time. Drives
                           line charts and period-over-period comparisons. Different from drill_down
                           which is a cross-section at one point in time.
  "waterfall_decomposition" — Explains how a top-line number breaks into its components or how it
                           moved from one period to another. Answers "what drives this number?"
  "benchmark_comparison" — Contextualises performance against an external reference: industry peers,
                           regulatory thresholds, historical norms, or internal targets. Answers
                           "is this number good or bad?"
  "exception_highlight"  — Calls out a specific anomaly, outlier, or deviation from the expected
                           pattern. Not a drill-down into normal behaviour — it surfaces something
                           that breaks the pattern and demands attention.
  "validation"           — Explicitly tests a stated assumption or strategic hypothesis with data.
                           "The thesis was X — here is what the data shows." Used to confirm or
                           challenge strategic bets or expansion assumptions.

DECISION SUPPORT:
  "risk_register"        — Structured inventory of identified risks with likelihood, impact, and
                           mitigation status. Used when the board must approve risk appetite
                           or sign off on mitigations collectively.
  "scenario_analysis"    — Shows how outcomes change under different assumptions: base / downside /
                           stress case. Answers "what if?" rather than "what is?"
  "decision_framework"   — Presents the structured logic for arriving at a decision: 2×2 matrix,
                           decision tree, criteria-weighted scorecard. Shows HOW the team
                           evaluated options, not just what they concluded.

SUPPLEMENTARY:
  "additional_information" — Enriching context that supports the narrative but is NOT in the direct
                           proof chain. If removed, the core argument still holds. No new KPIs
                           that contradict the summary.
  "methodology_note"     — Explains how data was collected, calculated, or classified. No analytical
                           claim. Typically 1 slide; appears after the slide that uses the method.

CLOSING:
  "recommendations"      — Actionable items with owners, timelines, and measurable outcomes, plus
                           forward-looking milestones and next review commitments. Must set
                           addresses_finding to the section(s) that identified the problem each
                           action resolves. Cards allowed only for target/milestone values — no
                           headline KPI cards restating actuals already shown elsewhere.
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PROOF CHAIN POINTERS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Use these two optional fields to create an explicit proof chain:

  "proves_claim": "string — the exact summary bullet or explainer claim this slide proves"
    → Required for: explainer_to_summary, drill_down, segmentation, trend_analysis,
      waterfall_decomposition, benchmark_comparison, exception_highlight, validation
    → Leave null for all other roles

  "addresses_finding": "string — the section name or finding this recommendation/action resolves"
    → Required for: recommendations
    → Leave null for all other roles

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Return a single valid JSON object with this exact structure:
{
  "document_type": "string — what kind of document this is",
  "governing_thought": "string — the single most important insight in one punchy sentence",
  "audience": "string — who this presentation is for (e.g. Board of Directors, Senior Management)",
  "narrative_flow": "string — name of the flow pattern being used",
  "data_heavy": true or false,
  "key_messages": [
    "string — key message 1 with specific data if available",
    "string — key message 2 with specific data if available",
    "string — key message 3 with specific data if available"
  ],
  "recommendations": [
    "string — recommended action or next step"
  ],
  "sections": [
    {
      "section_number": 1,
      "section_name": "string — name of this section",
      "narrative_role": "string — one of the roles defined above",
      "proves_claim": "string or null — the summary/explainer claim this slide proves",
      "addresses_finding": "string or null — the finding this recommendation resolves",
      "purpose": "string — what this section must achieve for the audience",
      "key_content": ["string — specific content point to cover in this section"],
      "suggested_slide_count": number,
      "data_available": true or false,
      "so_what": "string — the insight or takeaway from this section, not just the data"
    }
  ],
  "total_slides": number,
  "opening_guidance": "string — one sentence: how Agent 4 should frame the title slide and opening",
  "closing_guidance": "string — one sentence: how Agent 4 should close — what the last slide must achieve",
  "tone": "string — recommended tone (e.g. confident, cautious, urgent, balanced)"
}

SLIDE COUNT DEFINITION — READ CAREFULLY:
The slide count passed to you is the number of CONTENT slides the user wants.
Content slides = all slides EXCEPT title, divider, and thank-you slides.
You MUST add structural slides on top of the content count:
  - 1 title slide (always)
  - 1 thank-you / closing slide (always, narrative_role: "additional_information" or "recommendations")
  - 1–2 divider slides (as needed for section breaks)
  total_slides = content_slide_count + 1 (title) + 1 (closing) + divider count
Example: user requests 12 slides → total_slides = 12 + 1 + 1 + dividers (e.g. 2) = 16
The sum of suggested_slide_count across ALL sections (including title, dividers, closing) must equal total_slides.

CRITICAL OUTPUT RULES:
- Return ONLY the raw JSON object. No explanation. No preamble. No markdown fences.
- Your response must start with { and end with }. Nothing before or after.
- total_slides must equal content_slide_count + structural slides as defined above.
- The sum of suggested_slide_count across all sections must equal total_slides.
- Always include a title slide (narrative_role: "title").
- Include 1-2 divider slides for clean structure. Do not overuse dividers.
- Always include a closing/thank-you slide as the last section.
- Include a summary slide (narrative_role: "summary") ONLY when the presentation type, audience,
  and slide count warrant it — follow the WHEN TO INCLUDE / WHEN TO SKIP guidance above.
  If included, there must be exactly one. If skipped, open with context_setter or problem_statement instead.
- key_messages must be specific and data-driven — no generic placeholders.
- governing_thought must be a single punchy sentence a CEO finds immediately useful.
- opening_guidance and closing_guidance must be actionable one-sentence directives, not generic advice.
- Every key_content item must reference actual content from the document, not generic filler.
- proves_claim must be null for structural, opening, decision-support, and supplementary roles.`


async function runAgent3(state) {
  // state.slideCount = the number the USER requested = content slides only
  // We add structural slides (title + closing + dividers) on top.
  // Claude is told the content count and the total it must produce.
  const contentCount   = state.slideCount          // e.g. 12
  const structuralMin  = 3                          // title(1) + closing(1) + at least 1 divider
  const totalTarget    = contentCount + structuralMin  // e.g. 15 — Claude may add a second divider

  console.log('Agent 3 starting — content slides:', contentCount, '| total (incl. structural):', totalTarget)

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

The user has requested ${contentCount} CONTENT slides.
Content slides exclude: title slide, divider slides, and the closing/thank-you slide.
You must add structural slides on top:
  - 1 title slide
  - 1 closing/thank-you slide
  - 1–2 divider slides (use your judgment based on section count)
total_slides = ${contentCount} + 1 (title) + 1 (closing) + divider count
Minimum total_slides = ${totalTarget}. Adjust up by 1 if a second divider is warranted.

The sum of all suggested_slide_count values across ALL sections (including structural slides) must equal total_slides.
Content sections (everything except title, dividers, closing) must sum to exactly ${contentCount}.

Remember:
- Top-down approach: lead with the governing thought, then prove it, then recommend action
- Be specific — use actual numbers and facts from the document, not generic placeholders
- The key_messages should be the exact sentences a consultant would put on an executive summary slide
- Return ONLY valid JSON.`
      }
    ]
  }]

  const raw = await callClaude(AGENT3_SYSTEM, messages, 8000)

  console.log('Agent 3 — raw response length:', raw.length)
  if (raw.length < 200) console.warn('Agent 3 — suspiciously short response:', raw)

  const fallback = buildFallbackBrief(contentCount)
  const brief    = safeParseJSON(raw, fallback)

  // ── Validate and fix ──────────────────────────────────────────────────────

  // Ensure it has the required fields
  if (!brief.sections || !Array.isArray(brief.sections) || brief.sections.length === 0) {
    console.warn('Agent 3 — no sections in brief, using fallback')
    return fallback
  }

  // Compute the actual total from sections so we can validate correctly
  const actualTotal = brief.sections.reduce((sum, s) => sum + (s.suggested_slide_count || 1), 0)

  // Verify content slides = contentCount (structural roles excluded from count)
  const contentSectionTotal = brief.sections
    .filter(s => !FIXED_SINGLE_SLIDE_ROLES.has(s.narrative_role))
    .reduce((sum, s) => sum + (s.suggested_slide_count || 1), 0)
  if (contentSectionTotal !== contentCount) {
    console.warn('Agent 3 — content slide count is', contentSectionTotal, 'not', contentCount, '— redistributing')
    brief.sections = redistributeSlides(brief.sections, contentCount)
  }

  // Sync total_slides to actual section sum (after any redistribution)
  const correctedTotal = brief.sections.reduce((sum, s) => sum + (s.suggested_slide_count || 1), 0)
  if (brief.total_slides !== correctedTotal) {
    console.warn('Agent 3 — total_slides mismatch, correcting:', brief.total_slides, '→', correctedTotal)
    brief.total_slides = correctedTotal
  }

  // Legacy path kept for backwards compat (old code checked against state.slideCount directly)
  const sectionTotal = brief.sections.reduce((sum, s) => sum + (s.suggested_slide_count || 1), 0)
  if (sectionTotal !== brief.total_slides) {
    console.warn('Agent 3 — section slide counts sum to', sectionTotal, 'not', brief.total_slides, '— redistributing')
    brief.sections = redistributeSlides(brief.sections, contentCount)
  }

  // Ensure every section has required fields
  brief.sections = brief.sections.map((s, i) => ({
    section_number:        s.section_number        || (i + 1),
    section_name:          s.section_name          || 'Section ' + (i + 1),
    narrative_role:        s.narrative_role        || inferNarrativeRole(s, i),
    proves_claim:          s.proves_claim          || null,
    addresses_finding:     s.addresses_finding     || null,
    purpose:               s.purpose               || '',
    key_content:           Array.isArray(s.key_content) ? s.key_content : [],
    suggested_slide_count: s.suggested_slide_count || 1,
    data_available:        s.data_available        !== undefined ? s.data_available : false,
    so_what:               s.so_what               || ''
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
// target = the number of CONTENT slides requested by the user (excludes structural).
// Structural roles (title, divider, closing/thank-you, transition_narrative, methodology_note)
// are always exactly 1 slide and never receive extras.
// The function preserves structural sections at 1 each and distributes
// the content target across non-structural sections only.

const FIXED_SINGLE_SLIDE_ROLES = new Set([
  'title', 'divider', 'transition_narrative', 'methodology_note'
])

// Roles that are structural (1 slide each, never redistributed)
const STRUCTURAL_ROLES = new Set([
  'title', 'divider', 'transition_narrative', 'methodology_note', 'additional_information'
])

function redistributeSlides(sections, contentTarget) {
  // Lock structural sections at 1 slide each; reset content sections to 1 each
  const base = sections.map(s => ({ ...s, suggested_slide_count: 1 }))

  // Content sections = everything that isn't structural
  const contentIdxs = base
    .map((s, i) => ({ i, role: s.narrative_role }))
    .filter(({ role }) => !STRUCTURAL_ROLES.has(role))
    .map(({ i }) => i)

  // remaining = how many extra slides to distribute among content sections
  // (contentTarget - 1 per content section already assigned above)
  let remaining = contentTarget - contentIdxs.length

  if (remaining < 0) {
    // More content sections than target — trim content sections, keep all structural
    const structuralIdxs = new Set(
      base.map((s, i) => ({ s, i }))
        .filter(({ s }) => STRUCTURAL_ROLES.has(s.narrative_role))
        .map(({ i }) => i)
    )
    const trimmed = base.filter((s, i) => structuralIdxs.has(i))
    const contentSections = base.filter((s, i) => !structuralIdxs.has(i)).slice(0, contentTarget)
    return [...trimmed, ...contentSections]
  }

  // Distribute remaining extras round-robin across content sections
  let idx = 0
  while (remaining > 0 && contentIdxs.length > 0) {
    base[contentIdxs[idx % contentIdxs.length]].suggested_slide_count++
    remaining--
    idx++
  }

  return base
}


// ─── NARRATIVE ROLE INFERENCE ─────────────────────────────────────────────────
// Fallback for sections that came back from Claude without a narrative_role.
// Maps legacy section_type values and positional heuristics to a best-guess role.

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

  // Positional heuristics: first section → title, rest → explainer
  if (index === 0) return 'title'
  return 'explainer_to_summary'
}


// ─── FALLBACK BRIEF ──────────────────────────────────────────────────────────
// Used if Claude fails to return valid JSON

function buildFallbackBrief(contentSlideCount) {
  // contentSlideCount = what the user requested (excludes title, dividers, closing)
  // Structural slots added on top: title(1) + divider(1) + closing(1) = 3
  // Content slots: context_setter(1) + recommendations(1) + proof(remaining)
  const proofSlides = Math.max(1, contentSlideCount - 2)  // context + recommendations consume 2

  return {
    document_type:    'Business Document',
    governing_thought:'Key insights from the document require further analysis.',
    audience:         'Senior Management',
    narrative_flow:   'Situation → Analysis → Recommendations',
    data_heavy:       false,
    key_messages:     [
      'Please review the source document for key messages',
      'Analysis could not be completed automatically',
      'Manual review recommended'
    ],
    recommendations:  [],
    sections: [
      {
        section_number:        1,
        section_name:          'Title',
        narrative_role:        'title',
        proves_claim:          null,
        addresses_finding:     null,
        purpose:               'Opening title slide',
        key_content:           [],
        suggested_slide_count: 1,
        data_available:        false,
        so_what:               ''
      },
      {
        section_number:        2,
        section_name:          'Context',
        narrative_role:        'context_setter',
        proves_claim:          null,
        addresses_finding:     null,
        purpose:               'Establish scope, time period, and factual baseline',
        key_content:           ['Document scope', 'Time period covered', 'Key definitions'],
        suggested_slide_count: 1,
        data_available:        false,
        so_what:               'Audience understands what they are reviewing and why'
      },
      {
        section_number:        3,
        section_name:          'Analysis',
        narrative_role:        'divider',
        proves_claim:          null,
        addresses_finding:     null,
        purpose:               'Section break before analytical detail',
        key_content:           [],
        suggested_slide_count: 1,
        data_available:        false,
        so_what:               ''
      },
      {
        section_number:        4,
        section_name:          'Main Content',
        narrative_role:        'explainer_to_summary',
        proves_claim:          null,
        addresses_finding:     null,
        purpose:               'Core content from the document',
        key_content:           ['Content from uploaded document'],
        suggested_slide_count: Math.max(1, proofSlides),
        data_available:        true,
        so_what:               'Key insight from the content'
      },
      {
        section_number:        5,
        section_name:          'Recommended Actions',
        narrative_role:        'recommendations',
        proves_claim:          null,
        addresses_finding:     'Main Content',
        purpose:               'Actionable next steps with owners and timelines',
        key_content:           ['Action items', 'Owners', 'Timeline'],
        suggested_slide_count: 1,
        data_available:        false,
        so_what:               'What the audience must do after this presentation'
      },
      {
        section_number:        6,
        section_name:          'Thank You',
        narrative_role:        'additional_information',
        proves_claim:          null,
        addresses_finding:     null,
        purpose:               'Closing slide with contact details and next steps',
        key_content:           [],
        suggested_slide_count: 1,
        data_available:        false,
        so_what:               ''
      },
    ],
    total_slides: contentSlideCount + 3,  // content + title + divider + closing
    opening_guidance: 'Open by establishing context and scope — let the content speak before drawing conclusions',
    closing_guidance: 'Close with recommendations and forward-looking actions — ensure every key finding has an owner',
    tone:             'confident'
  }
}
