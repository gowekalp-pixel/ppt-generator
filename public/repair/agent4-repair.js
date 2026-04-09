// ─── AGENT 4 — REPAIR / FALLBACK ─────────────────────────────────────────────
// This file is intentionally separate from agent4.js so repair and fallback
// logic for Agent 4 lives in one place.
//
// When adding a new artifact type:
//   1. Add a constraint entry to AGENT4_NARRATIVE_CONSTRAINTS if the new type
//      should be permitted or forbidden for any narrative_role.
//   2. Edit prompts/agent4/A4-SlideRepair.js if repair instruction language changes.
//
// Depends on (must load before this file):
//   prompts/agent4/A4-SlideRepair.js  → _A4_REPAIR_RULES
//   agent4.js                         → AGENT4_REPAIR_SYSTEM, buildBriefSummaryForAgent4

// ─── NARRATIVE CONSTRAINTS ────────────────────────────────────────────────────
// Artifact restrictions per narrative_role — used by buildRepairPrompt.
// Add an entry here whenever a new narrative_role has artifact restrictions.
const AGENT4_NARRATIVE_CONSTRAINTS = {
  recommendations:     { permitted: 'prioritization, initiative_map, comparison_table, cards (future targets/milestones only)', forbidden: 'chart showing actuals, workflow, matrix, driver_tree' },
  methodology_note:    { permitted: 'insight_text, table (definitions/rates only)', forbidden: 'cards, chart, prioritization, workflow, matrix, driver_tree' },
  context_setter:      { permitted: 'cards (baseline KPIs), stat_bar, comparison_table', forbidden: 'prioritization, workflow, matrix, driver_tree' },
  exception_highlight: { permitted: 'profile_card_set or cards (negative/warning sentiment), chart (supporting evidence)', forbidden: 'prioritization, workflow' },
}

// ─── REPAIR USER PROMPT BUILDER ───────────────────────────────────────────────
// Builds the per-slide user message for repairSlide.
function buildRepairPrompt(briefSummary, narrativeRole, slide, layoutNames) {
  const roleConstraint = AGENT4_NARRATIVE_CONSTRAINTS[narrativeRole]
  const constraintBlock = roleConstraint
    ? `\nNARRATIVE ROLE CONSTRAINTS (narrative_role = "${narrativeRole}"):
  PERMITTED artifact types: ${roleConstraint.permitted}
  FORBIDDEN artifact types: ${roleConstraint.forbidden}
  IMPORTANT: If any zone currently uses a FORBIDDEN artifact type, you MUST replace it with a PERMITTED type and populate it with real data.\n`
    : ''

  return `This slide has missing or invalid artifact content. Fix every zone with specific data from the source document.

CONTEXT:
Document type:  ${briefSummary.document_type || '—'}
Key messages:   ${briefSummary.key_messages || '—'}
Key data:       ${briefSummary.key_data_points || '—'}
Narrative role: ${narrativeRole || '—'}
${constraintBlock}
SLIDE TO FIX:
${JSON.stringify(slide, null, 2)}

${_A4_REPAIR_RULES
      .replace('__ZONE_SPLIT_RULE__', layoutNames && layoutNames.length >= 5
        ? 'set zone_split="full" for all zones; artifact arrangement only when a zone has 2 artifacts'
        : 'must be explicit for scratch composition; set artifact_coverage_hint on each artifact (dominant/compact/co-equal) when a zone has 2+ artifacts; single-artifact zones use "full"')
      .replace('__LAYOUT_HINT_RULE__', layoutNames && layoutNames.length >= 5
        ? 'set to "full" (Agent 5 uses selected_layout_name for positioning)'
        : 'mirror zone_split into layout_hint.split for compatibility')}`
}

// ─── REPAIR RUNNER ────────────────────────────────────────────────────────────
async function repairSlide(slide, brief, contentB64, layoutNames) {
  console.log('Agent 4 — repairing slide', slide.slide_number, ':', slide.title)

  const briefSummary = buildBriefSummaryForAgent4(brief)

  // Build narrative_role constraint block so the repair knows what is permitted/forbidden
  const narrativeRole = slide.narrative_role || ''
  const prompt = buildRepairPrompt(briefSummary, narrativeRole, slide, layoutNames)

  const messages = [{
    role: 'user',
    content: [
      { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: contentB64 } },
      { type: 'text', text: prompt }
    ]
  }]

  const raw     = await callClaude(AGENT4_REPAIR_SYSTEM, messages, 2000)
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
