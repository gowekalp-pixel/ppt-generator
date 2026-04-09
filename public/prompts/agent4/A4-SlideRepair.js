// AGENT 4 — REPAIR USER PROMPT: STATIC FIX RULES BLOCK
// Used by buildRepairPrompt() in agent4.js.
// Two placeholders are replaced at runtime by buildRepairPrompt():
//   __ZONE_SPLIT_RULE__   → layout-mode vs scratch-mode zone_split instruction
//   __LAYOUT_HINT_RULE__  → layout-mode vs scratch-mode layout_hint.split instruction
// Part of the user message (not the system prompt) — assembled by agent4.js.

const _A4_REPAIR_RULES = `Fix rules:
- Replace all placeholder or empty content with real, specific data from the source document
- Keep the same zones[] structure and zone count — do not add or remove zones
- Keep artifact types UNLESS they violate the NARRATIVE ROLE CONSTRAINTS above — if forbidden, replace with the best permitted type and populate fully
- Preserve the chosen zone_structure. For asymmetric zone structures, keep dense proof artifacts in the dominant slot and compact support artifacts in the smaller support slots.
- Keep structure compatible with these rules:
  - 1 zone / 2 artifacts only for tightly paired proof + interpretation
  - 2 zones / 2 artifacts is the default clean structure
  - 3 zones / 3 artifacts is the default dashboard structure
  - reasoning artifacts may pair only with insight_text
  - sparse cards must never be dominant alone
  - 2 zones / 3 artifacts should usually be one rich proof zone plus one compact supporting zone
  - 2 zones / 4 artifacts should be used sparingly and only when both zones are dense and balanced
- All numbers from the source document
- Charts: 3+ categories, matching values, no all-zeros; ensure artifact_header is set
- insight_text: specific points with data; ensure artifact_header is set
- Workflows: fully populated nodes and connections; ensure artifact_header is set
- Enforce workflow restrictions:
  - process_flow: left_to_right only, >=4 nodes, full-width zone, >=50% height
  - hierarchy: top_down_branching only, >=3 levels, >=50% width, full content height
  - decomposition: left_to_right or top_to_bottom / top_down_branching only; if >3 nodes it must own full width (left_to_right) or full height (vertical)
  - timeline: left_to_right only, >=4 nodes, full-width zone, >=50% height
  - information_flow: do not use
- Tables: ensure artifact_header is set
- Matrix: fully populate x_axis, y_axis, all 4 quadrants, and points; ensure artifact_header is set
- Driver_tree: fully populate root and branches; ensure artifact_header is set
- Prioritization: fully populate ranked action items; ensure artifact_header is set
- comparison_table: use the flat schema — columns[] as a string array, rows[].cells[{value, icon_type, subtext, tone}]; cells[0] is the option name with tone:"label"; data cells use tone:"positive"/"negative"/"neutral"; each data cell has EITHER value (metric/%, ₹) OR icon_type ("check"/"cross"/"partial") but never both; is_recommended:true on the best row; fully populate every cell with real values from the source document
- If matrix / driver_tree / prioritization is present, keep it only in the PRIMARY zone and pair it only with insight_text
- If a slide uses matrix / driver_tree / prioritization, do NOT add cards, chart, workflow, or table anywhere else on that slide
- selected_layout_name: choose from available layouts; set to "" if none available
- zone_split / artifact_arrangement / artifact_coverage_hint: __ZONE_SPLIT_RULE__
- layout_hint.split: __LAYOUT_HINT_RULE__
- In scratch composition, cards with 1–2 items must stay compact and must not occupy a dominant tall zone
- Unless a cards artifact has 8+ cards, no individual card may imply more than ~15% of total slide area
- Return ONLY a single JSON object for this one slide`
