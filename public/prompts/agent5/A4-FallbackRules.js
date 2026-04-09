// AGENT 5 — FALLBACK USER PROMPT: STATIC INSTRUCTIONS
// Used by buildFallbackDesign() in agent5.js.
// Part of the user message for the fallback (emergency) design call — assembled by agent5.js.

const _A5_FALLBACK_INSTRUCTIONS = `Build the best possible layout for this slide.
Preserve the title, key_message, zones structure and artifact content from the manifest.
CRITICAL: preserve the exact number of zones and the exact artifact types in each zone from the manifest.
Do NOT collapse charts, tables, workflows, cards, matrix, driver_tree, prioritization, comparison_table, initiative_map, profile_card_set, or risk_register into generic insight_text unless the manifest itself uses insight_text.
Choose the cleanest, most board-ready layout from the manifest structure itself: zone count, zone roles, artifact types, and selected_layout_name if present.
Return a single JSON object for this one slide.`
