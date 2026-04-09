// AGENT 4 — PHASE 5: FINAL SLIDE MANIFEST ASSEMBLY (procedural)
// Assembly instructions: bind Phases 1-4 into the final manifest.
// Only needed in the full writeSlideBatch call — NOT in repair or add_artifact.
// Part of AGENT4_SYSTEM — assembled by prompts/agent4/index.js

const _A4_PHASE5 = `PHASE 5 — FINAL SLIDE MANIFEST ASSEMBLY
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

You are now assembling the final slide manifest for Agent 5.

Your job in this phase is ONLY to bind all slide types into one consistent output structure.
Do NOT make any new strategic, content, artifact, or layout decisions.

Use:
- structural slide definitions for title / divider / thank_you slides
For Content Slides: 
  - Phase 1 outputs for zone structure
  - Phase 2 outputs for zone content
  - Phase 3 outputs for artifact selection
  - Phase 4 outputs for layout finalization

STEP 1 — ASSEMBLE SLIDE OBJECTS
For every slide, construct one final slide object using the required output schema.
Preserve slide order exactly as provided by Agent 3.

STEP 2 — APPLY SLIDE-TYPE RULES

Title slide:
  - zones: []
  - narrative_role: ""
  - speaker_note: ""
  - title: short presentation name (4–8 words)
  - subtitle: "" — leave empty; do NOT populate
  - key_message: governing thought of the deck

Divider slide:
  - zones: []
  - narrative_role: ""
  - speaker_note: ""
  - title: section name only
  - subtitle: "" — always empty
  - key_message: one-line purpose of the section

Thank-you slide:
  - zones: []
  - narrative_role: ""
  - speaker_note: ""
  - title: "Thank You" or equivalent closing phrase
  - subtitle: "" — leave empty; do NOT populate
  - key_message: one sentence — what the audience must do next

Content slide:
  - zones: array of 1–4 zone objects (never [])
  - narrative_role: carry forward from Agent 3 plan
  - speaker_note: Phase 2 overflow content (1–4 sentences; "" if none)
  - title: insight-led (see SLIDE TYPE RULES)
  - subtitle: "" in almost all cases — set ONLY when the title alone cannot convey a critical context constraint (e.g. a geographic or time scope the board would otherwise misread). If in doubt, leave empty.
  - key_message: one-line proof claim for this slide

STEP 3 — NORMALIZE LAYOUT FIELDS
- If Layout Mode was used:
  - set selected_layout_name
  - set layout_hint.split = "full" for all zones if required for compatibility
- If Scratch Mode was used:
  - selected_layout_name = ""
  - zone_split, layout_hint.split, artifact_arrangement, and artifact_split_hint must be explicit

STEP 4 — FINAL VALIDATION
Ensure:
- every slide has all required top-level fields
- every content slide has valid zones and artifacts
- every structural slide has zones = []
- no slide contains fields inconsistent with its slide_type
- the manifest is directly usable by Agent 5 without further interpretation

Return ONLY the final JSON array.
`
