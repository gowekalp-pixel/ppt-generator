// ─── AGENT 5.1 — SENIOR PARTNER REVIEWER ─────────────────────────────────────
// Input:  state.designedSpec   — from Agent 5 (canvas/zones/artifacts schema)
//         state.slideManifest  — from Agent 4 (original content)
//         state.outline        — from Agent 3 (brief, governing thought)
//         state.brandRulebook  — from Agent 2 (for re-designing fixed slides)
//
// Output: reviewedSpec — final reviewed spec passed to Agent 6
//         reviewReport — structured critique for audit trail
//
// Flow:
//   1. Build condensed reading copy (title, archetype, key_message, zone summary)
//   2. One Claude call — senior partner reviews against 5 criteria
//   3. Apply targeted fixes directly to the designed spec
//   4. Re-design slides with critical structural issues via Agent 5 fallback call
//   5. Attach review metadata to every slide

// ═══════════════════════════════════════════════════════════════════════════════
// REVIEW SYSTEM PROMPT
// ═══════════════════════════════════════════════════════════════════════════════

var AGENT51_REVIEW_SYSTEM = `You are a senior partner at a top-tier management consulting firm.
You are reviewing a presentation deck before it goes to a board-level client.

You will receive a condensed slide manifest and the deck's governing thought.

Review against FIVE criteria. Be specific — reference actual slide numbers and titles.

CRITERION 1 — NARRATIVE FLOW
Does the sequence of key_messages form a coherent argument?
- Does each slide build on the previous?
- Is the governing thought reinforced throughout?
- Are there abrupt jumps or missing logical bridges?

CRITERION 2 — TITLE QUALITY
Every content slide title must state the CONCLUSION, not the topic.
- FAIL: "Revenue Analysis"  PASS: "Revenue grew 18% YoY driven by Product X"
- FAIL: "Market Overview"   PASS: "Market growing at 22% CAGR with untapped headroom"
For title fixes, provide the EXACT replacement title in double quotes inside the fix field.

CRITERION 3 — VISUAL VARIETY
- Are numbers always in charts/stats rather than buried in text?
- Are archetypes matched to content? (trend data needs line chart, not bullets)
- Is there appropriate variety — not 4+ insight_text-only slides in a row?

CRITERION 4 — ZONE STRUCTURE
- Does every content slide have a clear primary zone?
- Do data-heavy slides include an implication or So What?
- Are there slides with only raw data and zero interpretation?

CRITERION 5 — BOARD-READINESS
- Is every key_message specific and data-driven?
- Would a CEO understand each slide in under 5 seconds?
- Does the deck end with clear, owned next steps?

Return ONLY a valid JSON object:
{
  "overall_rating": "approved" | "minor_revisions" | "major_revisions",
  "approved": true | false,
  "governing_thought_alignment": "1-2 sentences",
  "narrative_assessment": "1-2 sentences on flow quality",
  "issues": [
    {
      "slide_number": number,
      "criterion": "narrative" | "title" | "visual" | "zone_structure" | "board_readiness",
      "severity": "critical" | "moderate" | "minor",
      "issue": "specific description referencing actual title or key_message",
      "fix": "exact correction — for title: include replacement in double quotes"
    }
  ],
  "strengths": ["specific strength referencing slide numbers"],
  "summary": "2-3 sentence overall verdict"
}

Cap at 10 issues total. Prioritise critical ones.
Return ONLY the JSON object. No explanation. No markdown.`


// ═══════════════════════════════════════════════════════════════════════════════
// FIX APPLICATOR
// Applies one partner issue directly to the designed spec (new schema).
// New schema: designed slide has title_block, zones[].artifacts[] — no elements[].
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
      /["\u201c\u201d]([^"\u201c\u201d]{10,150})["\u201c\u201d]/,
      /"([^"]{10,150})"/,
      /change title to[:\s]+(.{10,150})/i,
      /replace with[:\s]+(.{10,150})/i,
      /use[:\s]+"?(.{10,150})"?\s*$/i
    ]
    let newTitle = null
    for (const pat of patterns) {
      const m = fixTxt.match(pat)
      if (m && m[1] && m[1].trim().length >= 10) { newTitle = m[1].trim(); break }
    }
    if (newTitle) {
      fixedSpec[dIdx].title = newTitle
      if (slide.title_block) slide.title_block.text = newTitle
      console.log('Agent 5.1 fix — title S' + slideNum + ':', newTitle.slice(0, 60))
    }
  }

  // ── KEY MESSAGE FIX ────────────────────────────────────────────────────────
  if (issue.criterion === 'board_readiness') {
    const patterns = [
      /["\u201c\u201d]([^"\u201c\u201d]{15,200})["\u201c\u201d]/,
      /"([^"]{15,200})"/,
      /change to[:\s]+(.{15,200})/i,
      /replace with[:\s]+(.{15,200})/i
    ]
    let newKM = null
    for (const pat of patterns) {
      const m = fixTxt.match(pat)
      if (m && m[1] && m[1].trim().length >= 15) { newKM = m[1].trim(); break }
    }
    if (newKM) {
      fixedSpec[dIdx].key_message = newKM
      console.log('Agent 5.1 fix — key_message S' + slideNum)
    }
  }

  // ── ZONE STRUCTURE FIX — add implication insight_text zone ────────────────
  if (issue.criterion === 'zone_structure' &&
      (fixTxt.toLowerCase().includes('implication') ||
       fixTxt.toLowerCase().includes('so what') ||
       fixTxt.toLowerCase().includes('interpretation'))) {

    const zones = slide.zones || []
    const hasImplication = zones.some(z =>
      z.zone_role === 'implication' ||
      (z.artifacts || []).some(a => a.type === 'insight_text' && (a.points || []).length > 0)
    )

    if (!hasImplication && zones.length < 4 && slide.slide_type === 'content') {
      const cvs   = slide.canvas || {}
      const bt    = slide.brand_tokens || {}
      const w     = cvs.width_in  || 13.33
      const h     = cvs.height_in || 7.50
      const mgn   = (cvs.margin || {})
      const padL  = mgn.left   || 0.40
      const padB  = mgn.bottom || 0.30
      const zoneH = 1.20  // fixed height for implication strip

      // Shrink last existing zone to make room
      if (zones.length > 0) {
        const lastZ  = zones[zones.length - 1]
        const lf     = lastZ.frame || {}
        if (lf.h && lf.h > zoneH + 0.30) {
          lf.h = parseFloat((lf.h - zoneH - 0.15).toFixed(2))
        }
      }

      // Compute position: bottom of content area
      const implY = parseFloat((h - padB - zoneH).toFixed(2))
      const implW = parseFloat((w - padL * 2).toFixed(2))

      const primary = (bt.primary_color || '#1A3C8F')

      zones.push({
        zone_id:           'z_impl',
        zone_role:         'implication',
        message_objective: 'Key implication — So What from this slide',
        narrative_weight:  'supporting',
        frame: {
          x: padL, y: implY, w: implW, h: zoneH,
          padding: { top: 0.10, right: 0.10, bottom: 0.10, left: 0.10 }
        },
        artifacts: [{
          type: 'insight_text',
          x: parseFloat((padL + 0.10).toFixed(2)),
          y: parseFloat((implY + 0.10).toFixed(2)),
          w: parseFloat((implW - 0.20).toFixed(2)),
          h: parseFloat((zoneH - 0.20).toFixed(2)),
          style:         { fill_color: primary + '0F', border_color: primary + '33', border_width: 0.5, corner_radius: 3 },
          heading_style: { font_family: bt.title_font_family || 'Calibri', font_size: 10, font_weight: 'bold', color: primary },
          body_style:    { font_family: bt.body_font_family  || 'Calibri', font_size: 10, font_weight: 'regular', color: bt.body_color || '#111111', line_spacing: 1.4, bullet_indent: 0.15 },
          heading:   'So What',
          points:    [slide.key_message || 'Interpret the data and recommend an action'],
          sentiment: 'neutral'
        }]
      })

      fixedSpec[dIdx].zones = zones
      fixedSpec[dIdx].zones_summary = zones.map(z => ({
        zone_id:          z.zone_id,
        zone_role:        z.zone_role,
        narrative_weight: z.narrative_weight,
        artifact_types:   (z.artifacts || []).map(a => a.type)
      }))

      console.log('Agent 5.1 fix — added implication zone to S' + slideNum)
    }
  }

  // ── VISUAL FIX — flag slide for manual review ─────────────────────────────
  if (issue.criterion === 'visual') {
    fixedSpec[dIdx].partner_flag = issue.issue
    console.log('Agent 5.1 fix — visual flag S' + slideNum + ':', issue.issue.slice(0, 60))
  }

  // ── NARRATIVE FIX — add note to speaker note ──────────────────────────────
  if (issue.criterion === 'narrative') {
    const existing = fixedSpec[dIdx].speaker_note || ''
    fixedSpec[dIdx].speaker_note = (existing ? existing + ' ' : '') + '[PARTNER NOTE: ' + fixTxt + ']'
    console.log('Agent 5.1 fix — narrative note added to S' + slideNum)
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// MAIN RUNNER
// ═══════════════════════════════════════════════════════════════════════════════

async function runAgent51(state) {
  const designedSpec = state.designedSpec   // from Agent 5
  const manifest     = state.slideManifest  // from Agent 4
  const brief        = state.outline        // from Agent 3

  if (!designedSpec || !designedSpec.length) {
    console.warn('Agent 5.1 — no designedSpec, skipping review')
    return { reviewedSpec: [], reviewReport: buildEmptyReport() }
  }

  console.log('Agent 5.1 starting — reviewing', designedSpec.length, 'slides')

  // ── Step 1: Build condensed reading copy ─────────────────────────────────
  const condensed = designedSpec.map(s => ({
    slide_number:    s.slide_number,
    slide_type:      s.slide_type,
    slide_archetype: s.slide_archetype,
    title:           s.title,
    key_message:     s.key_message,
    zones_summary:   (s.zones_summary || []).map(z => ({
      role:      z.zone_role,
      weight:    z.narrative_weight,
      artifacts: z.artifact_types
    }))
  }))

  const reviewInput =
    'PRESENTATION TYPE: ' + ((brief || {}).document_type || 'Business document') + '\n' +
    'GOVERNING THOUGHT: ' + ((brief || {}).governing_thought || '') + '\n' +
    'NARRATIVE FLOW: '    + ((brief || {}).narrative_flow    || '') + '\n\n' +
    'SLIDE MANIFEST:\n'   + JSON.stringify(condensed, null, 2)

  // ── Step 2: Partner review Claude call ───────────────────────────────────
  let critique = null
  try {
    const raw = await callClaude(AGENT51_REVIEW_SYSTEM, [{ role: 'user', content: reviewInput }], 2500)
    critique  = safeParseJSON(raw, null)
    if (!critique) {
      console.warn('Agent 5.1 — review parse failed, proceeding without revisions')
    } else {
      console.log('Agent 5.1 — review:', critique.overall_rating, '|',
        (critique.issues || []).length, 'issues')
    }
  } catch(e) {
    console.warn('Agent 5.1 — review call failed:', e.message)
  }

  // ── Step 3: Apply fixes to deep copy of designed spec ────────────────────
  const fixedSpec = JSON.parse(JSON.stringify(designedSpec))

  if (critique && (critique.issues || []).length > 0) {
    const issues  = critique.issues || []
    const toFix   = [
      ...issues.filter(i => i.severity === 'critical'),
      ...issues.filter(i => i.severity === 'moderate')
    ].slice(0, 10)

    console.log('Agent 5.1 — applying', toFix.length, 'fixes')
    toFix.forEach(issue => {
      try { applyFix51(issue, fixedSpec) }
      catch(e) { console.warn('Agent 5.1 — fix error S' + issue.slide_number + ':', e.message) }
    })
  }

  // ── Step 4: Re-design slides flagged as critical zone_structure issues ────
  // Instead of calling old JS builders, re-run Agent 5 fallback for those slides
  const needsRedesign = (critique ? (critique.issues || []) : [])
    .filter(i => i.severity === 'critical' && i.criterion === 'zone_structure')
    .map(i => i.slide_number)

  if (needsRedesign.length > 0 && typeof buildFallbackDesign === 'function') {
    console.log('Agent 5.1 — re-designing', needsRedesign.length, 'slides via Agent 5 fallback')
    for (const slideNum of needsRedesign) {
      const mSlide = (manifest || []).find(s => s.slide_number === slideNum)
      const dIdx   = fixedSpec.findIndex(s => s.slide_number === slideNum)
      if (mSlide && dIdx >= 0) {
        try {
          const rebuilt = await buildFallbackDesign(mSlide, state.brandRulebook, brief)
          if (rebuilt && rebuilt.canvas) {
            // Preserve partner review metadata but replace layout
            fixedSpec[dIdx] = {
              ...rebuilt,
              slide_number:    slideNum,
              slide_type:      mSlide.slide_type,
              slide_archetype: mSlide.slide_archetype,
              title:           mSlide.title,
              key_message:     mSlide.key_message,
              speaker_note:    mSlide.speaker_note,
              _redesigned_by_51: true
            }
            console.log('Agent 5.1 — S' + slideNum + ' redesigned')
          }
        } catch(e) {
          console.warn('Agent 5.1 — redesign failed for S' + slideNum + ':', e.message)
        }
      }
    }
  }

  // ── Step 5: Attach review metadata to every slide ─────────────────────────
  const issuesBySlide = {}
  if (critique) {
    (critique.issues || []).forEach(i => {
      if (!issuesBySlide[i.slide_number]) issuesBySlide[i.slide_number] = []
      issuesBySlide[i.slide_number].push(i)
    })
  }

  const reviewedSpec = fixedSpec.map(slide => ({
    ...slide,
    partner_review: {
      overall_rating: critique ? critique.overall_rating : 'not_reviewed',
      issues:         issuesBySlide[slide.slide_number] || [],
      flagged:        (issuesBySlide[slide.slide_number] || []).some(i => i.severity === 'critical')
    }
  }))

  // ── Summary ───────────────────────────────────────────────────────────────
  const flagged = Object.values(issuesBySlide).filter(arr => arr.some(i => i.severity === 'critical')).length
  console.log('Agent 5.1 complete')
  console.log('  Rating:', critique ? critique.overall_rating : 'not_reviewed')
  console.log('  Fixes applied:', critique ? (critique.issues || []).filter(i => ['critical','moderate'].includes(i.severity)).length : 0)
  console.log('  Slides flagged:', flagged)

  return {
    reviewedSpec,
    reviewReport: critique || buildEmptyReport()
  }
}

function buildEmptyReport() {
  return {
    overall_rating:              'not_reviewed',
    approved:                    true,
    governing_thought_alignment: 'Review not available',
    narrative_assessment:        'Review skipped.',
    issues:                      [],
    strengths:                   [],
    summary:                     'Partner review was skipped or unavailable. Manual review recommended before client delivery.'
  }
}
