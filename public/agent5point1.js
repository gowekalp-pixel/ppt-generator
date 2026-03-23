// ─── AGENT 5.1 — SENIOR PARTNER REVIEWER ─────────────────────────────────────
// Input:  state.designedSpec   — from Agent 5 (fully positioned elements[])
//         state.slideManifest  — from Agent 4 (original zone/artifact structure)
//         state.outline        — from Agent 3 (brief, governing thought)
//
// Output: reviewedSpec — Agent 5's designed spec after partner review and fixes
//         reviewReport — structured critique for audit trail
//
// Agent 5.1 plays the role of a senior partner reviewing the deck before it
// goes to the client. It critiques, then applies fixes, then signs off.
// It does NOT redesign slides from scratch — it makes targeted corrections.

// ═══════════════════════════════════════════════════════════════════════════════
// REVIEW SYSTEM PROMPT
// ═══════════════════════════════════════════════════════════════════════════════

const AGENT51_REVIEW_SYSTEM = `You are a senior partner at a top-tier management consulting firm.
You are reviewing a presentation before it goes to a board-level client.

You will receive:
1. A condensed slide manifest (titles, archetypes, zone roles, artifact types, key messages)
2. The presentation governing thought and document type

Review against FIVE criteria and return a structured JSON critique.

═══════════════════════════
CRITERION 1 — NARRATIVE FLOW
═══════════════════════════
Does the sequence of key_messages tell a coherent story?
- Does each slide build on the previous?
- Is the governing thought reinforced throughout?
- Are there abrupt transitions or logic gaps?

═══════════════════════════
CRITERION 2 — TITLE QUALITY
═══════════════════════════
Is every content slide title insight-led?
- Titles must state the CONCLUSION, not the topic
- FAIL: "Revenue Analysis" | PASS: "Revenue grew 18% driven by Product X"
- FAIL: "Market Overview"  | PASS: "Market growing at 22% CAGR with untapped headroom"
- Check every content slide title

═══════════════════════════
CRITERION 3 — VISUAL VARIETY AND APPROPRIATENESS
═══════════════════════════
- Are numbers always visualised (charts/stats) rather than buried in text?
- Is there appropriate visual variety — not 5+ bullet-only slides in a row?
- Are archetypes matched to content (e.g. trend data shown as line chart, not bullets)?
- Are there slides where a workflow would tell the story better than text?

═══════════════════════════
CRITERION 4 — ZONE STRUCTURE
═══════════════════════════
- Does each slide have a clear primary zone?
- Are implications and so-whats present on data slides?
- Are there slides with only data and no interpretation?
- Are there slides with too many zones causing clutter?

═══════════════════════════
CRITERION 5 — BOARD-READINESS
═══════════════════════════
- Is every key_message specific and data-driven?
- Are there any vague or generic statements?
- Does the deck communicate with the confidence and precision expected at board level?
- Would a senior executive understand each slide in under 5 seconds?

═══════════════════════════
OUTPUT FORMAT
═══════════════════════════

Return ONLY a valid JSON object:

{
  "overall_rating": "approved" | "minor_revisions" | "major_revisions",
  "approved": true | false,
  "governing_thought_alignment": "string — does the deck reinforce the governing thought?",
  "narrative_assessment": "string — 1-2 sentences on flow quality",
  "issues": [
    {
      "slide_number": number,
      "criterion": "narrative" | "title" | "visual" | "zone_structure" | "board_readiness",
      "severity": "critical" | "moderate" | "minor",
      "issue": "string — specific description of the problem",
      "fix": "string — exact correction to make"
    }
  ],
  "strengths": ["string — things done well"],
  "summary": "string — 2-3 sentence overall assessment for the team"
}

Be specific. Reference actual slide numbers, actual titles, actual key messages.
Do not give generic feedback.
Return ONLY the JSON object. No explanation. No markdown.`


// ═══════════════════════════════════════════════════════════════════════════════
// FIX APPLICATOR
// Applies partner critique to Agent 4 manifest + Agent 5 designed spec
// ═══════════════════════════════════════════════════════════════════════════════

function applyFix(issue, manifest, designedSpec) {
  const slideNum = issue.slide_number
  const mIdx = manifest.findIndex(s => s.slide_number === slideNum)
  const dIdx = designedSpec.findIndex(s => s.slide_number === slideNum)

  if (mIdx < 0 && dIdx < 0) return

  const fix = (issue.fix || '').toLowerCase()

  // ── Title fix ──────────────────────────────────────────────────────────────
  if (issue.criterion === 'title') {
    // Extract the proposed new title from fix text
    const titleMatch = issue.fix.match(/["']([^"']{10,100})["']/) ||
                       issue.fix.match(/change to:?\s*(.+)/i) ||
                       issue.fix.match(/use:?\s*(.+)/i)

    if (titleMatch) {
      const newTitle = titleMatch[1].trim()
      if (mIdx >= 0) manifest[mIdx].title = newTitle
      if (dIdx >= 0) {
        designedSpec[dIdx].title = newTitle
        // Update the title element in elements[]
        const titleEl = designedSpec[dIdx].elements.find(e => e.id === 'title' || e.id === 'sidebar_title')
        if (titleEl) titleEl.text = newTitle
      }
      console.log('Agent 5.1 — fixed title on slide', slideNum, '→', newTitle)
    }
  }

  // ── Key message fix ────────────────────────────────────────────────────────
  if (issue.criterion === 'board_readiness' && fix.includes('key message')) {
    const kmMatch = issue.fix.match(/["']([^"']{15,200})["']/)
    if (kmMatch) {
      const newKM = kmMatch[1].trim()
      if (mIdx >= 0) manifest[mIdx].key_message = newKM
      if (dIdx >= 0) designedSpec[dIdx].key_message = newKM
      console.log('Agent 5.1 — fixed key_message on slide', slideNum)
    }
  }

  // ── Zone structure fix — add implication zone ──────────────────────────────
  if (issue.criterion === 'zone_structure' && fix.includes('implication')) {
    if (mIdx >= 0) {
      const slide = manifest[mIdx]
      const hasImplication = (slide.zones||[]).some(z =>
        z.zone_role === 'implication' ||
        (z.artifacts||[]).some(a => a.type === 'insight_text')
      )

      if (!hasImplication && slide.key_message) {
        console.log('Agent 5.1 — adding implication zone to slide', slideNum)
        // Add a small implication insight_text zone — Agent 5 will re-render
        slide.zones = (slide.zones || []).slice(0, 3)  // max 3 existing zones
        slide.zones.push({
          zone_id:           'z_impl',
          zone_role:         'implication',
          message_objective: 'Key implication from this slide',
          narrative_weight:  'supporting',
          artifacts: [{
            type:      'insight_text',
            heading:   'So What',
            points:    [slide.key_message],
            sentiment: 'neutral'
          }],
          layout_hint: { split: inferImplicationSplit(slide.zones) }
        })
      }
    }
  }

  // ── Visual fix — suggest archetype change ──────────────────────────────────
  if (issue.criterion === 'visual') {
    if (fix.includes('chart') || fix.includes('visualise')) {
      console.log('Agent 5.1 — visual fix flagged for slide', slideNum, '(manual review recommended)')
      // Flag for manual review — don't auto-change archetypes as it requires new chart data
      if (dIdx >= 0) {
        designedSpec[dIdx].partner_flag = 'Visual: ' + issue.issue
      }
    }
  }
}

function inferImplicationSplit(existingZones) {
  const splits = existingZones.map(z => (z.layout_hint||{}).split || 'full')

  // If there's a full zone, change it to top_70 and add bottom_30 for implication
  if (splits.includes('full'))      return 'bottom_30'
  if (splits.includes('left_60'))   return 'bottom_full'  // would need restructuring
  if (splits.includes('top_30'))    return 'bottom_30'

  return 'bottom_50'
}


// ═══════════════════════════════════════════════════════════════════════════════
// MAIN RUNNER
// ═══════════════════════════════════════════════════════════════════════════════

async function runAgent51(state) {
  const designedSpec  = state.designedSpec   // from Agent 5
  const manifest      = state.slideManifest  // from Agent 4 (for fixing)
  const brief         = state.outline        // from Agent 3

  console.log('Agent 5.1 starting — reviewing', designedSpec.length, 'slides')

  // ── Step 1: Build condensed review input ──────────────────────────────────
  const condensed = designedSpec.map(s => ({
    slide_number:      s.slide_number,
    slide_type:        s.slide_type,
    slide_archetype:   s.slide_archetype,
    title:             s.title,
    key_message:       s.key_message,
    zones_summary:     s.zones_summary || []
  }))

  const reviewInput = `PRESENTATION TYPE: ${(brief||{}).document_type || '—'}
GOVERNING THOUGHT: ${(brief||{}).governing_thought || '—'}
NARRATIVE FLOW: ${(brief||{}).narrative_flow || '—'}

SLIDE MANIFEST TO REVIEW:
${JSON.stringify(condensed, null, 2)}`

  // ── Step 2: Partner review call ───────────────────────────────────────────
  let critique = null
  try {
    const raw     = await callClaude(AGENT51_REVIEW_SYSTEM, [{ role: 'user', content: reviewInput }], 2000)
    critique      = safeParseJSON(raw, null)

    if (!critique) {
      console.warn('Agent 5.1 — review parse failed, proceeding without revisions')
    } else {
      console.log('Agent 5.1 — review complete:',
        critique.overall_rating || '—',
        '|', (critique.issues||[]).length, 'issues found')
    }
  } catch(e) {
    console.warn('Agent 5.1 — review call failed:', e.message)
  }

  // ── Step 3: Apply fixes ───────────────────────────────────────────────────
  // Work on mutable copies
  const fixedManifest = JSON.parse(JSON.stringify(manifest))
  const fixedSpec     = JSON.parse(JSON.stringify(designedSpec))

  if (critique && (critique.issues||[]).length > 0) {
    const criticals = critique.issues.filter(i => i.severity === 'critical')
    const moderates = critique.issues.filter(i => i.severity === 'moderate')
    const toFix     = [...criticals, ...moderates].slice(0, 8)  // cap at 8 fixes

    console.log('Agent 5.1 — applying', toFix.length, 'fixes (', criticals.length, 'critical,', moderates.length, 'moderate)')

    toFix.forEach(issue => {
      try {
        applyFix(issue, fixedManifest, fixedSpec)
      } catch(e) {
        console.warn('Agent 5.1 — fix failed for slide', issue.slide_number, ':', e.message)
      }
    })
  }

  // ── Step 4: Re-render slides that had structural fixes ────────────────────
  // If zones were added (implication zones), those slides need re-rendering
  const slidesNeedingRerender = (critique ? (critique.issues||[]) : [])
    .filter(i => i.criterion === 'zone_structure')
    .map(i => i.slide_number)

  if (slidesNeedingRerender.length > 0) {
    console.log('Agent 5.1 — re-rendering', slidesNeedingRerender.length, 'slides after zone fixes')

    for (const slideNum of slidesNeedingRerender) {
      const mSlide = fixedManifest.find(s => s.slide_number === slideNum)
      const dIdx   = fixedSpec.findIndex(s => s.slide_number === slideNum)

      if (mSlide && dIdx >= 0 && mSlide.slide_type === 'content') {
        // Re-run Agent 5 element building for this slide
        // We need access to brand data — pass through state
        try {
          const brand      = state.brandRulebook
          const dim        = { w: brand.slide_width_inches||11.02, h: brand.slide_height_inches||8.29 }
          const colors     = getBrandColors(brand)
          const layoutMode = detectLayoutMode(brand)
          const footerLyt  = getSlideFooterLayout(brand, brand.slide_layouts||[])

          fixedSpec[dIdx].elements = buildContentSlide(mSlide, dim, colors, layoutMode, footerLyt)
          fixedSpec[dIdx].zones_summary = (mSlide.zones||[]).map(z => ({
            zone_id:          z.zone_id,
            zone_role:        z.zone_role,
            narrative_weight: z.narrative_weight,
            split:            (z.layout_hint||{}).split||'full',
            artifact_types:   (z.artifacts||[]).map(a=>a.type)
          }))
          console.log('  Re-rendered slide', slideNum)
        } catch(e) {
          console.warn('  Re-render failed for slide', slideNum, ':', e.message)
        }
      }
    }
  }

  // ── Step 5: Attach review metadata to each slide ─────────────────────────
  const issuesBySlide = {}
  if (critique) {
    (critique.issues||[]).forEach(i => {
      if (!issuesBySlide[i.slide_number]) issuesBySlide[i.slide_number] = []
      issuesBySlide[i.slide_number].push(i)
    })
  }

  const reviewedSpec = fixedSpec.map(slide => ({
    ...slide,
    partner_review: {
      overall_rating: critique ? critique.overall_rating : 'not_reviewed',
      issues:         issuesBySlide[slide.slide_number] || [],
      flagged:        (issuesBySlide[slide.slide_number]||[]).some(i => i.severity === 'critical')
    }
  }))

  // ── Summary log ───────────────────────────────────────────────────────────
  console.log('Agent 5.1 complete')
  console.log('  Overall rating:', critique ? critique.overall_rating : 'not_reviewed')
  console.log('  Fixes applied:', critique ? (critique.issues||[]).filter(i=>['critical','moderate'].includes(i.severity)).length : 0)
  console.log('  Slides flagged:', Object.values(issuesBySlide).filter(issues=>issues.some(i=>i.severity==='critical')).length)

  return {
    reviewedSpec,
    reviewReport: critique || {
      overall_rating:              'not_reviewed',
      approved:                    true,
      governing_thought_alignment: 'Review not available',
      narrative_assessment:        'Review not available',
      issues:                      [],
      strengths:                   [],
      summary:                     'Partner review was skipped or unavailable.'
    }
  }
}
