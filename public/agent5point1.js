// ─── AGENT 5.1 — BIG 4 SENIOR PARTNER REVIEWER ───────────────────────────────
// Input:  state.designedSpec   — from Agent 5 (render-ready layout spec)
//         state.slideManifest  — from Agent 4 (original content + zones)
//         state.outline        — from Agent 3 (brief, governing thought)
//         state.brandRulebook  — from Agent 2 (design tokens)
//
// Output: { reviewedSpec, reviewReport }
//
// Flow:
//   1. Build condensed reading copy for fast review
//   2. One Claude call — senior partner reviews against 5 criteria
//   3. Apply quick text fixes (titles, key messages) directly
//   4. Batch-redesign ALL slides with critical/visual/structural issues
//      by re-running Agent 5 with partner feedback embedded in the prompt
//   5. Single-slide fallback if batch redesign fails
//   6. Attach review metadata to every slide


// ═══════════════════════════════════════════════════════════════════════════════
// SYSTEM PROMPT — SENIOR PARTNER REVIEW
// ═══════════════════════════════════════════════════════════════════════════════

var AGENT51_REVIEW_SYSTEM = `You are a Senior Partner at a Big 4 management consulting firm.
You are conducting a final gate review before this deck goes to a board-level client.
You have reviewed hundreds of decks. You are direct, specific, and uncompromising on quality.

You will receive a condensed slide manifest and the deck's governing thought.

Review against FIVE criteria. Reference exact slide numbers and quote actual titles/messages.

═══════════════════════════════════════
CRITERION 1 — NARRATIVE FLOW
═══════════════════════════════════════
Does the sequence of key_messages form a tight, logical argument?
- Does each slide build on or respond to the previous?
- Is the governing thought reinforced and resolved by the final slide?
- Flag any abrupt jumps, circular logic, or missing bridges.
- Flag if the opening does not establish stakes and the close does not own next steps.

═══════════════════════════════════════
CRITERION 2 — TITLE QUALITY
═══════════════════════════════════════
Every content slide title MUST state the CONCLUSION, not the topic.
Standard: If a title could appear on any deck about this subject, it fails.

FAIL examples (topic titles): "Revenue Analysis" / "Market Overview" / "Team Structure"
PASS examples (insight titles): "Revenue grew 18% YoY driven solely by Product X" / "Market growing at 22% CAGR with €4B untapped headroom" / "Talent gap in ops will block scale without immediate action"

For every failing title, provide the EXACT replacement wrapped in double quotes in the fix field.
The replacement must be specific, quantified where data exists, and opinionated.

═══════════════════════════════════════
CRITERION 3 — VISUAL VARIETY & ARTIFACT FIT
═══════════════════════════════════════
- Trend data → must use line chart, not bullets
- Comparisons → bar or column chart, not a table of text
- Numbers quoted in bullets → should be a stat callout or chart
- Process / sequence → workflow diagram, not numbered bullets
- Are there 3+ consecutive insight_text-only slides? Flag for visual relief.
- Does each artifact type match the insight type of that zone?

═══════════════════════════════════════
CRITERION 4 — ZONE STRUCTURE & SO WHAT
═══════════════════════════════════════
- Every data-heavy slide MUST have an implication/So What zone
- No slide should contain only raw data with zero interpretation
- Does every content slide have one clear PRIMARY zone that carries the message?
- Are supporting zones genuinely supportive or just padding?

═══════════════════════════════════════
CRITERION 5 — BOARD-READINESS
═══════════════════════════════════════
- Can a CEO extract the slide's point in under 5 seconds?
- Is every key_message specific, quantified, and directional?
- Does the deck close with clear, owned, time-bound next steps?
- Would any slide embarrass the firm in front of a PE or board audience?

═══════════════════════════════════════
OUTPUT FORMAT
═══════════════════════════════════════
Return ONLY a valid JSON object:
{
  "overall_rating": "approved" | "minor_revisions" | "major_revisions",
  "approved": true | false,
  "governing_thought_alignment": "1-2 sentences — is the governing thought carried through?",
  "narrative_assessment": "2-3 sentences — is the argument tight? where does it break?",
  "issues": [
    {
      "slide_number": <number>,
      "criterion": "narrative" | "title" | "visual" | "zone_structure" | "board_readiness",
      "severity": "critical" | "moderate" | "minor",
      "issue": "specific description — quote the actual title or key_message that fails",
      "fix": "exact correction — for title: put the EXACT replacement in double quotes"
    }
  ],
  "strengths": ["strength referencing specific slide numbers"],
  "summary": "3-4 sentence overall verdict as you would deliver it in a verbal debrief"
}

Rules:
- Cap at 12 issues total. Ruthlessly prioritise critical ones.
- Every issue must reference a specific slide_number.
- For title issues the fix MUST contain the exact replacement title in double quotes.
- Return ONLY the JSON. No explanation. No markdown. No preamble.`


// ═══════════════════════════════════════════════════════════════════════════════
// SYSTEM PROMPT — FEEDBACK-DRIVEN REDESIGN
// ═══════════════════════════════════════════════════════════════════════════════

var AGENT51_REDESIGN_SYSTEM = `You are a senior presentation designer rebuilding slides based on a Big 4 partner's review.

The partner has flagged specific issues for each slide. Rebuild EVERY flagged slide to fully address the partner's comments.

Rules:
- Each slide object contains a _partner_feedback array — address EVERY item in it
- For criterion "title": use the EXACT replacement title from required_fix
- For criterion "visual": change the artifact type as directed (e.g. bullets → bar chart)
- For criterion "zone_structure": add, remove, or restructure zones as directed
- For criterion "board_readiness": sharpen the key message and title to be specific and quantified
- For criterion "narrative": adjust the slide's focus and message per the feedback
- Preserve all original content data (categories, series values, points, card text, rows, etc.)
- Apply brand design tokens exactly — colors, fonts, slide dimensions
- Compute all coordinates in decimal inches to 2 decimal places
- FULLY specify every artifact's style sub-objects:
  chart: chart_style{} + series_style[]
  workflow: workflow_style{} + nodes[] with x/y/w/h + connections[] with path[]
  table: table_style{} + column_widths[] + row_heights[]
  cards: card_style{} + card_frames[] with x/y/w/h
  insight_text standard: style{} + heading_style{} + body_style{}
  insight_text grouped:  heading_style{} + group_layout + group_header_style{} + group_bullet_box_style{} + bullet_style{} + group_gap_in + header_to_box_gap_in
- Return ONLY a valid JSON array of exactly the requested slides. No markdown. No explanation.`


// ═══════════════════════════════════════════════════════════════════════════════
// QUICK-FIX APPLICATOR
// Applies text-level fixes directly to the designed spec — no redesign needed.
// ═══════════════════════════════════════════════════════════════════════════════

function applyFix51(issue, fixedSpec) {
  const slideNum = issue.slide_number
  const dIdx     = fixedSpec.findIndex(s => s.slide_number === slideNum)
  if (dIdx < 0) return

  const slide  = fixedSpec[dIdx]
  const fixTxt = issue.fix || ''

  // ── TITLE FIX ──────────────────────────────────────────────────────────────
  if (issue.criterion === 'title') {
    const patterns = [
      /["\u201c\u201d]([^"\u201c\u201d]{10,200})["\u201c\u201d]/,
      /"([^"]{10,200})"/,
      /change title to[:\s]+(.{10,200})/i,
      /replace with[:\s]+(.{10,200})/i,
      /use[:\s]+"?(.{10,200})"?\s*$/i
    ]
    let newTitle = null
    for (const pat of patterns) {
      const m = fixTxt.match(pat)
      if (m && m[1] && m[1].trim().length >= 10) { newTitle = m[1].trim(); break }
    }
    if (newTitle) {
      fixedSpec[dIdx].title = newTitle
      if (slide.title_block) slide.title_block.text = newTitle
      console.log('Agent 5.1 title fix S' + slideNum + ':', newTitle.slice(0, 70))
    }
  }

  // ── KEY MESSAGE FIX ────────────────────────────────────────────────────────
  if (issue.criterion === 'board_readiness') {
    const patterns = [
      /["\u201c\u201d]([^"\u201c\u201d]{15,250})["\u201c\u201d]/,
      /"([^"]{15,250})"/,
      /change to[:\s]+(.{15,250})/i,
      /replace with[:\s]+(.{15,250})/i
    ]
    let newKM = null
    for (const pat of patterns) {
      const m = fixTxt.match(pat)
      if (m && m[1] && m[1].trim().length >= 15) { newKM = m[1].trim(); break }
    }
    if (newKM) {
      fixedSpec[dIdx].key_message = newKM
      console.log('Agent 5.1 key_message fix S' + slideNum)
    }
  }

  // ── NARRATIVE FIX — add partner note to speaker note ──────────────────────
  if (issue.criterion === 'narrative') {
    const existing = fixedSpec[dIdx].speaker_note || ''
    fixedSpec[dIdx].speaker_note = (existing ? existing + ' ' : '') +
      '[PARTNER NOTE: ' + fixTxt + ']'
    console.log('Agent 5.1 narrative note S' + slideNum)
  }

  // ── VISUAL FLAG — mark for manual attention if no redesign ────────────────
  if (issue.criterion === 'visual' && issue.severity === 'minor') {
    fixedSpec[dIdx].partner_flag = issue.issue
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// FEEDBACK-DRIVEN REDESIGN
// Re-runs Agent 5 design for a batch of slides, with partner feedback baked in.
// ═══════════════════════════════════════════════════════════════════════════════

async function redesignSlidesWithFeedback(manifestSlides, brand, brief, issuesBySlide, batchNum) {
  const slideNums = manifestSlides.map(s => s.slide_number)
  console.log('Agent 5.1 redesign batch', batchNum, ':', slideNums.join(', '))

  // Embed partner feedback directly into each slide's manifest data
  const annotated = manifestSlides.map(ms => {
    const issues = (issuesBySlide[ms.slide_number] || []).map(i => ({
      criterion:    i.criterion,
      severity:     i.severity,
      issue:        i.issue,
      required_fix: i.fix
    }))
    return { ...ms, _partner_feedback: issues }
  })

  const prompt =
    buildBrandBrief(brand, brief) +
    '\n\nSLIDES TO REDESIGN — WITH PARTNER FEEDBACK (batch ' + batchNum + '):\n' +
    JSON.stringify(annotated, null, 2) +
    '\n\nINSTRUCTIONS:' +
    '\n- Each slide has a _partner_feedback array — address EVERY item' +
    '\n- For criterion "title": apply the EXACT replacement from required_fix' +
    '\n- For criterion "visual": change the artifact type as directed' +
    '\n- For criterion "zone_structure": restructure zones as directed; add implication zone if missing' +
    '\n- For criterion "board_readiness": sharpen title + key_message to be specific and quantified' +
    '\n- For criterion "narrative": shift the slide\'s focus per the feedback' +
    '\n- Preserve all content data (categories, series, points, cards, rows, headers)' +
    '\n- Apply brand design tokens exactly' +
    '\n- Compute exact coordinates in decimal inches (2dp)' +
    '\n- FULLY specify all artifact style sub-objects (chart_style, series_style[], workflow_style, card_frames[], etc.)' +
    '\n- Return a valid JSON array of exactly ' + manifestSlides.length + ' slides'

  const raw    = await callClaude(AGENT51_REDESIGN_SYSTEM, [{ role: 'user', content: prompt }], 7000)
  const parsed = safeParseJSON(raw, null)

  if (!Array.isArray(parsed)) {
    console.warn('Agent 5.1 redesign batch', batchNum, '— parse failed. Raw length:', raw.length)
    console.warn('  First 300 chars:', raw.slice(0, 300))
    return null
  }

  console.log('Agent 5.1 redesign batch', batchNum, '— got', parsed.length, 'slides')
  return parsed
}


// ═══════════════════════════════════════════════════════════════════════════════
// HELPER: which issues trigger a full redesign vs a quick text fix?
// ═══════════════════════════════════════════════════════════════════════════════

function needsRedesign(issue) {
  if (issue.severity === 'critical') return true
  if (issue.severity === 'moderate' && (issue.criterion === 'visual' || issue.criterion === 'zone_structure')) return true
  return false
}


// ═══════════════════════════════════════════════════════════════════════════════
// MAIN RUNNER
// ═══════════════════════════════════════════════════════════════════════════════

async function runAgent51(state) {
  const designedSpec = state.designedSpec
  const manifest     = state.slideManifest
  const brief        = state.outline
  const brand        = state.brandRulebook

  if (!designedSpec || !designedSpec.length) {
    console.warn('Agent 5.1 — no designedSpec, skipping review')
    return { reviewedSpec: [], reviewReport: buildEmptyReport() }
  }

  console.log('Agent 5.1 starting — reviewing', designedSpec.length, 'slides')

  // ── Step 1: Build condensed reading copy for review ───────────────────────
  const condensed = designedSpec.map(s => ({
    slide_number:    s.slide_number,
    slide_type:      s.slide_type,
    slide_archetype: s.slide_archetype,
    title:           s.title,
    key_message:     s.key_message,
    zones_summary:   (s.zones_summary || (s.zones || []).map(z => ({
      role:      z.zone_role,
      weight:    z.narrative_weight,
      artifacts: (z.artifacts || []).map(a => a.type)
    }))).map(z => ({
      role:      z.role      || z.zone_role,
      weight:    z.weight    || z.narrative_weight,
      artifacts: z.artifacts || z.artifact_types || []
    }))
  }))

  const reviewInput =
    'PRESENTATION TYPE: ' + ((brief || {}).document_type    || 'Business document')  + '\n' +
    'GOVERNING THOUGHT: ' + ((brief || {}).governing_thought || 'Not specified')      + '\n' +
    'NARRATIVE FLOW: '    + ((brief || {}).narrative_flow    || 'Not specified')      + '\n' +
    'TONE: '              + ((brief || {}).tone              || 'professional')       + '\n\n' +
    'SLIDE MANIFEST:\n'   + JSON.stringify(condensed, null, 2)

  // ── Step 2: Partner review call ───────────────────────────────────────────
  let critique = null
  try {
    const raw = await callClaude(AGENT51_REVIEW_SYSTEM, [{ role: 'user', content: reviewInput }], 3000)
    critique  = safeParseJSON(raw, null)
    if (critique) {
      console.log('Agent 5.1 review done:', critique.overall_rating,
        '|', (critique.issues || []).length, 'issues',
        '| approved:', critique.approved)
    } else {
      console.warn('Agent 5.1 — review parse failed, proceeding without revisions')
    }
  } catch(e) {
    console.warn('Agent 5.1 — review call failed:', e.message)
  }

  if (!critique || !(critique.issues || []).length) {
    // No issues — return as-is with review metadata attached
    const reviewedSpec = designedSpec.map(s => ({
      ...s,
      partner_review: {
        overall_rating: critique ? critique.overall_rating : 'not_reviewed',
        issues: [],
        flagged: false
      }
    }))
    return { reviewedSpec, reviewReport: critique || buildEmptyReport() }
  }

  // ── Step 3: Triage issues ─────────────────────────────────────────────────
  const allIssues = critique.issues || []

  // Build per-slide index
  const issuesBySlide = {}
  allIssues.forEach(i => {
    if (!issuesBySlide[i.slide_number]) issuesBySlide[i.slide_number] = []
    issuesBySlide[i.slide_number].push(i)
  })

  // Slides that get a full Agent 5 redesign with feedback
  const redesignSlideNums = [...new Set(
    allIssues.filter(needsRedesign).map(i => i.slide_number)
  )]

  // Issues that can be fixed with text edits (no redesign)
  const quickFixIssues = allIssues.filter(i =>
    !needsRedesign(i) || !redesignSlideNums.includes(i.slide_number)
  )

  console.log('Agent 5.1 triage — redesign:', redesignSlideNums.length,
    'slides | quick-fix:', quickFixIssues.length, 'issues')

  // Deep copy of spec to mutate
  const fixedSpec = JSON.parse(JSON.stringify(designedSpec))

  // ── Step 4: Apply quick text fixes (titles, key messages, notes) ─────────
  quickFixIssues.forEach(issue => {
    try { applyFix51(issue, fixedSpec) }
    catch(e) { console.warn('Agent 5.1 quick-fix error S' + issue.slide_number + ':', e.message) }
  })

  // ── Step 5: Feedback-driven redesign via Agent 5 ─────────────────────────
  if (redesignSlideNums.length > 0 && manifest && brand) {
    // Get manifest entries for slides to redesign
    const slidesForRedesign = redesignSlideNums
      .map(num => (manifest || []).find(s => s.slide_number === num))
      .filter(Boolean)

    // Batch into groups of 3 (same as Agent 5 batching)
    const BATCH_SIZE = 3
    const batches = []
    for (let i = 0; i < slidesForRedesign.length; i += BATCH_SIZE) {
      batches.push(slidesForRedesign.slice(i, i + BATCH_SIZE))
    }

    for (let b = 0; b < batches.length; b++) {
      const batch = batches[b]
      let rebuilt = null

      try {
        rebuilt = await redesignSlidesWithFeedback(batch, brand, brief, issuesBySlide, b + 1)
      } catch(e) {
        console.warn('Agent 5.1 redesign batch', b + 1, 'failed:', e.message)
      }

      for (let i = 0; i < batch.length; i++) {
        const mSlide   = batch[i]
        const slideNum = mSlide.slide_number
        const dIdx     = fixedSpec.findIndex(s => s.slide_number === slideNum)
        if (dIdx < 0) continue

        // Try to match rebuilt result to this slide
        const match = rebuilt
          ? (rebuilt.find(r => r.slide_number === slideNum) || rebuilt[i] || null)
          : null

        if (match && match.canvas && (match.zones || []).length > 0) {
          const normalized = typeof normaliseDesignedSlide === 'function'
            ? normaliseDesignedSlide(match, mSlide, brand)
            : match
          // Merge: use rebuilt layout but preserve key content fields
          fixedSpec[dIdx] = {
            ...normalized,
            slide_number:    slideNum,
            slide_type:      mSlide.slide_type      || fixedSpec[dIdx].slide_type,
            slide_archetype: mSlide.slide_archetype || fixedSpec[dIdx].slide_archetype,
            title:           normalized.title       || mSlide.title,
            key_message:     normalized.key_message || mSlide.key_message,
            speaker_note:    mSlide.speaker_note    || fixedSpec[dIdx].speaker_note,
            _redesigned_by_51: true,
            _redesign_issues:  (issuesBySlide[slideNum] || []).map(i => i.criterion)
          }
          console.log('Agent 5.1 — S' + slideNum + ' redesigned (criteria:',
            (issuesBySlide[slideNum] || []).map(i => i.criterion).join(', ') + ')')
        } else {
          // Batch failed for this slide — fall back to single-slide redesign
          console.warn('Agent 5.1 — S' + slideNum + ' batch miss, trying per-slide fallback')
          if (typeof buildFallbackDesign === 'function') {
            try {
              const fb = await buildFallbackDesign(mSlide, brand, brief)
              if (fb && fb.canvas) {
                const normalizedFb = typeof normaliseDesignedSlide === 'function'
                  ? normaliseDesignedSlide(fb, mSlide, brand)
                  : fb
                fixedSpec[dIdx] = {
                  ...normalizedFb,
                  slide_number:    slideNum,
                  slide_type:      mSlide.slide_type,
                  slide_archetype: mSlide.slide_archetype,
                  title:           mSlide.title,
                  key_message:     mSlide.key_message,
                  speaker_note:    mSlide.speaker_note,
                  _redesigned_by_51: true,
                  _fallback_redesign: true
                }
                console.log('Agent 5.1 — S' + slideNum + ' fallback redesign OK')
              }
            } catch(fe) {
              console.warn('Agent 5.1 — S' + slideNum + ' fallback also failed:', fe.message)
            }
          }
        }
      }
    }
  }

  // ── Step 6: Attach review metadata to every slide ─────────────────────────
  const reviewedSpec = fixedSpec.map(slide => ({
    ...slide,
    partner_review: {
      overall_rating:  critique.overall_rating,
      issues:          issuesBySlide[slide.slide_number] || [],
      flagged:         (issuesBySlide[slide.slide_number] || []).some(i => i.severity === 'critical'),
      redesigned:      !!slide._redesigned_by_51
    }
  }))

  // ── Summary ───────────────────────────────────────────────────────────────
  const criticalCount = allIssues.filter(i => i.severity === 'critical').length
  const moderateCount = allIssues.filter(i => i.severity === 'moderate').length
  console.log('Agent 5.1 complete')
  console.log('  Rating:', critique.overall_rating, '| Approved:', critique.approved)
  console.log('  Issues — critical:', criticalCount, 'moderate:', moderateCount,
    'minor:', allIssues.filter(i => i.severity === 'minor').length)
  console.log('  Slides redesigned:', redesignSlideNums.length,
    '| Quick-fixed:', quickFixIssues.length)

  return {
    reviewedSpec,
    reviewReport: critique
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// EMPTY REPORT
// ═══════════════════════════════════════════════════════════════════════════════

function buildEmptyReport() {
  return {
    overall_rating:              'not_reviewed',
    approved:                    true,
    governing_thought_alignment: 'Review not available.',
    narrative_assessment:        'Review skipped.',
    issues:                      [],
    strengths:                   [],
    summary:                     'Partner review was skipped or unavailable. Manual review recommended before client delivery.'
  }
}
