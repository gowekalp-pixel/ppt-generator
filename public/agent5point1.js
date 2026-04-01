// ─── AGENT 5.1 — SENIOR PARTNER REVIEW ──────────────────────────────────────
// Input:  state.designedSpec  — array of designed slides from Agent 5
//         state.brandTokens   — hoisted brand tokens
// Output: { reviewedSpec, reviewReport }
//
// Performs a lightweight quality review of the deck:
// - Checks slide sequence is correct (title first, closing last)
// - Flags any slides missing blocks
// - Passes through spec unchanged if no critical issues found
// Currently a passthrough with sequence guarantee — full Claude review TBD.

async function runAgent51(state) {
  const slides = state.designedSpec || []

  if (!slides.length) {
    console.warn('Agent 5.1 — no slides to review')
    return {
      reviewedSpec: [],
      reviewReport: { overall_rating: 'not_reviewed', issues: [] }
    }
  }

  // ── Guarantee sequence: sort by slide_number ascending ──────────────────────
  // Ensures title slide is first, closing last, regardless of batch completion order.
  const sorted = [...slides].sort((a, b) => (a?.slide_number || 0) - (b?.slide_number || 0))

  // ── Basic quality scan ───────────────────────────────────────────────────────
  const issues = []
  sorted.forEach(slide => {
    if (!slide) return
    if (!Array.isArray(slide.blocks) || slide.blocks.length === 0) {
      issues.push({
        slide_number: slide.slide_number,
        severity: 'critical',
        description: 'Slide ' + slide.slide_number + ' has no render blocks'
      })
    }
    if (!slide.canvas) {
      issues.push({
        slide_number: slide.slide_number,
        severity: 'critical',
        description: 'Slide ' + slide.slide_number + ' missing canvas definition'
      })
    }
  })

  const criticals = issues.filter(i => i.severity === 'critical').length
  const rating = criticals > 0 ? 'major_revisions' : issues.length > 0 ? 'minor_revisions' : 'approved'

  console.log('Agent 5.1 — review complete:', rating,
    '| slides:', sorted.length,
    '| issues:', issues.length,
    '(', criticals, 'critical )')
  console.log('Agent 5.1 — slide sequence:', sorted.map(s => s?.slide_number).join(', '))

  return {
    reviewedSpec: sorted,
    reviewReport: {
      overall_rating: rating,
      issues,
      slide_count: sorted.length
    }
  }
}
