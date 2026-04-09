// AGENT 4 — BATCH USER PROMPT: STATIC INSTRUCTIONS BLOCK
// Used by buildBatchPrompt() in agent4.js.
// __SLIDE_COUNT__ is replaced at runtime with batchPlan.length.
// Part of the user message (not the system prompt) — assembled by agent4.js.

const _A4_BATCH_INSTRUCTIONS = `SLIDE ORDER RULE (mandatory):
Output slides in EXACTLY the order listed above. Do NOT reorder slides for any reason.
The slide_number in your output must match the slide_number in the plan entry you are writing.
The downstream renderer assembles slides positionally — any reordering produces a broken deck.

INSTRUCTIONS:
- For each slide, start from the locked Agent 3 plan, then derive zones and artifacts from the message.
- Use narrative_role, zone_count_signal, dominant_zone_signal, co_primary_signal, and strategic_objective as the primary zone-planning inputs.
- Do NOT infer structure from any legacy field. Use only narrative_role, zone_count_signal, dominant_zone_signal, co_primary_signal, and strategic_objective for zone planning.
- Use this zone-count logic as the authoritative rule set (apply in order, first match wins):
  - narrative_role = methodology_note -> 1 zone, stop
  - narrative_role = transition_narrative -> 1 zone; insight_text only; no data artifacts permitted
  - narrative_role = summary -> 1-2 zones; prefer 1 if one dense synthesis artifact can carry the slide
  - co_primary_signal = yes -> 2 zones, CO-PRIMARY, side-by-side
  - narrative_role = explainer_to_summary -> 3-4 zones; dominant decomposition plus supporting proof
  - narrative_role = validation -> 2-3 zones; dominant proof plus supporting evidence
  - narrative_role = drill_down -> 2-3 zones; dominant decomposition plus support
  - narrative_role = benchmark_comparison -> 2 zones, equal weight unless dominant_zone_signal = yes
  - narrative_role = trend_analysis -> 2 zones, dominant proof plus implication support
  - narrative_role = segmentation -> 2-3 zones, comparison-led
  - narrative_role = waterfall_decomposition -> 2 zones, dominant proof plus explanation
  - narrative_role = scenario_analysis -> 3-4 zones, grid or structured comparison preferred
  - narrative_role = option_evaluation -> 3-4 zones, option comparison or criteria grid preferred
  - narrative_role = risk_assessment -> 2-3 zones, dominant register plus mitigation / implication support
  - narrative_role = recommendations -> 2-3 zones, recommendation plus rationale / ask support
  - narrative_role = exception_highlight -> 2 zones, dominant issue plus implication / action support
  - narrative_role = context_setter or problem_statement -> 2 zones, framing plus consequence / evidence support
  - zone_count_signal = 1, 2, 3, or 4 -> use that count as the baseline when no stronger rule above applies
  - strategic_objective implies comparing options, scenarios, or alternatives -> 3-4 zones
  - strategic_objective implies a single core proof with one takeaway -> 2 zones
  - strategic_objective implies a compact synthesis or note -> 1 zone
  - default -> 2 zones
- Apply these modifiers after selecting the baseline:
  - dominant_zone_signal = yes -> Zone 1 must be dominant
  - dominant_zone_signal = no and co_primary_signal = no -> prefer balanced weights across zones
  - scenario_analysis, option_evaluation, and recommendations should not collapse below 2 zones
- Write the title from slide_title_draft and sharpen it only if needed.
- Before finalizing artifacts for a content slide, choose ONE zone_structure that matches the zone count and narrative geometry:
  1-zone:  ZS01_single_full
  2-zone:  ZS02_stacked_equal | ZS03_side_by_side_equal
  3-zone:  ZS04_left_dominant_right_stack | ZS05_right_dominant_left_stack |
           ZS06_top_full_bottom_two | ZS07_top_two_bottom_dominant |
           ZS11_three_rows_equal | ZW01_three_columns_equal
  4-zone:  ZS08_quad_grid | ZS09_left_dominant_right_triptych |
           ZS10_top_full_bottom_three | ZW04_four_columns_equal |
           ZW02_three_columns_right_stack | ZW03_three_columns_left_stack
- After choosing zone_structure, decide which slot is dominant vs support, then pick allowed artifacts for each slot. For asymmetric structures, dominant slots may carry chart / workflow / table / reasoning artifacts, while support slots should prefer insight_text, grouped insight_text, compact cards, or compact charts.
- Pull all numbers from the attached source document — no invented figures
- Title slides: zones = []
- Divider slides: zones = []
- Content slides: 1–4 zones, each with 1–2 artifacts
- Structural pattern rules:
  - 1 zone / 1 artifact: only if the artifact is dense enough to carry the slide
  - 1 zone / 2 artifacts: only for tightly paired proof + interpretation structures
  - 2 zones / 2 artifacts: default clean structure, one artifact per zone
  - 2 zones / 3 artifacts: one paired zone + one solo zone
  - 2 zones / 4 artifacts: use sparingly, both zones must be dense and balanced
  - 3 zones / 3 artifacts: default dashboard / layered argument structure
  - 3 zones / 4 artifacts: only one zone may be paired
  - 3 zones / 5+ artifacts: avoid
  - 4 zones / 4 artifacts: simple compact dashboards only
  - 4 zones / 5+ artifacts: exceptional only, no reasoning artifacts
  - 2 zones / 2 artifacts examples: chart | insight_text, workflow | insight_text, table | insight_text, cards | workflow, cards | insight_text, chart | chart, chart | table
  - 2 zones / 3 artifacts examples: chart + insight_text | cards, workflow + insight_text | cards, cards | workflow + insight_text
  - 2 zones / 4 artifacts examples: chart + insight_text | workflow + insight_text, cards + insight_text | chart + insight_text, workflow + insight_text | table + insight_text
  - 3 zones / 3 artifacts examples: cards | workflow | insight_text, cards | chart | insight_text, chart | chart | insight_text
  - 3 zones / 4 artifacts examples: cards | workflow + insight_text | prioritization, chart | workflow + insight_text | insight_text, cards | chart + insight_text | insight_text
- In Scratch Mode, zone_split must be explicit for every zone.
- In Scratch Mode, if a zone has 2 artifacts, set artifact_arrangement and set artifact_coverage_hint on EACH artifact using semantic tokens: primary artifact → "dominant"; secondary artifact → "compact" (use "co-equal" if both are similar density and neither clearly dominates).
- In Scratch Mode, cards with 1–2 items are compact summary anchors only: keep their zone share at or below ~40% of the slide, prefer top strips or narrow side panes, and never let 2 sparse cards occupy a tall dominant zone.
- Card density rule: unless a single cards artifact contains 8+ cards, no individual card may imply more than ~15% of total slide area.
- Every chart: MUST have 3+ categories, matching values, no all-zeros; set artifact_header to the one-line insight the chart proves
- clustered_bar: MUST have exactly 2 series
- Every insight_text: MUST have specific, data-driven points; set artifact_header to a 2–4 word specific label naming the implication (e.g. "Risk Implication", "Growth Opportunity", "Action Required") — never use generic labels like "So What" or "Key Insight"
- Workflows: fully populate nodes and connections; set artifact_header to the one-line insight
- Workflow restrictions:
  - process_flow: left_to_right only, >=4 nodes, full-width zone, >=50% height
  - hierarchy: top_down_branching only, >=3 levels, >=50% width, full content height
  - decomposition: left_to_right or top_to_bottom / top_down_branching only; if >3 nodes it must own full width (left_to_right) or full height (vertical)
  - timeline: left_to_right only, >=4 nodes, full-width zone, >=50% height
  - information_flow: do not use
- Tables: set artifact_header to the one-line insight the table proves
- Matrix: fully populate axes, all 4 quadrants, and plotted points; set artifact_header
- Driver_tree: fully populate root and branches; set artifact_header
- Prioritization: fully populate ranked items sorted by importance; set artifact_header
- If matrix / driver_tree / prioritization is used, it must be in the PRIMARY zone and may be paired only with insight_text
- If a slide uses matrix / driver_tree / prioritization, do NOT add cards, chart, workflow, or table anywhere else on that slide
- ZS01_single_full is the only structure that should routinely host reasoning artifacts as the dominant full-slide construct
- 1 zone / 2 artifacts allowed pairs:
  - chart + insight_text
  - workflow + insight_text
  - table + insight_text
  - cards + insight_text only if cards >= 4
  - prioritization + insight_text
  - matrix + insight_text
  - driver_tree + insight_text
  - chart + table only when tightly linked
  - cards + workflow only when cards are a compact anchor and workflow is the main proof
  - never use matrix + chart, driver_tree + workflow, prioritization + cards, or two unrelated proof artifacts in one zone
- Return ONLY a valid JSON array for these __SLIDE_COUNT__ slides`
