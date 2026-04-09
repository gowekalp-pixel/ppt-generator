// AGENT 5 — QUALITY RULES
// Internal zone layout, layout decision rules, quality gates, render ownership, output rule.
// Assembled into AGENT5_SYSTEM by agent5.js.

const _A5_QUALITY_RULES = `*********************************************************************************
INTERNAL ZONE LAYOUT (2 artifacts)
*********************************************************************************

If 2 artifacts in a zone, split zone frame between them:
- chart + insight_text: chart 65%, insight 35%
- workflow + insight_text: workflow dominant
- chart + table: chart 65%, table 35% side by side
- cards + insight_text: cards dominant
- table + insight_text: table 65%, insight 35%
Artifacts must NOT overlap.

*********************************************************************************
LAYOUT DECISION RULES
*********************************************************************************

Translate Agent 4 layout_hint to zone geometry.
Content area = canvas minus margins minus title band.
Gutter between zones: 0.15 inches.

- full: single frame fills content area
- left_60 + right_40: split width 60/40
- left_50 + right_50: equal width split
- top_30 + bottom_70: split height 30/70
- top_left_50 + top_right_50 + bottom_full: two upper + full lower
- left_full_50 + top_right_50_h + bottom_right_50_h: left full + right stacked
- tl + tr + bl + br: 2Ã—2 grid, equal gutters

*********************************************************************************
QUALITY RULES
*********************************************************************************

- no missing fields
- no overlapping zones or artifacts
- body text â‰¥ 9pt; captions â‰¥ 8pt
- no placeholder values
- reflect narrative hierarchy in space allocation
- primary zones visually dominant

*********************************************************************************
RENDER OWNERSHIP:
- Agent 5 owns every render-critical detail for every inch of the slide canvas
- Agent 6 is only allowed to render what Agent 5 decided

OUTPUT RULE
*********************************************************************************

Return ONLY a valid JSON array.
No explanation. No markdown. No text outside JSON.`