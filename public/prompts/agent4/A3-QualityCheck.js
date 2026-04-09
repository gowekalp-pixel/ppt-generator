// AGENT 4 — PRE-OUTPUT QUALITY GATES
// Checklist Claude runs before emitting JSON.
// Needed in ALL calls: writeSlideBatch, repairSlide, and add_artifact.

const _A4_QUALITY_GATES = `PRE-OUTPUT QUALITY GATES
Run ALL gates before emitting JSON. Fix any failure before proceeding.
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

GATE 1 — CONTENT INTEGRITY
  [ ] No placeholder text anywhere
  [ ] No invented numbers — every figure sourced from the source document
  [ ] No vague wording — every point specific and actionable
  [ ] All content slides have insight-led titles
  [ ] Every slide title ≤ 10 words — count every word; rewrite if over
  [ ] Every insight_text standard: bullet count within density cap (compact zone → ≤3; standard zone → ≤6; dense zone → ≤8)
  [ ] Every insight_text standard: each bullet ≤ 12 words — rewrite if over; excess detail to speaker notes
  [ ] Every insight_text grouped: group × bullet count within density cap (compact → no groups allowed; standard → 2 sections × 3 pts; dense → 5 sections × 3 pts)
  [ ] Every insight_text grouped: each bullet ≤ 12 words — rewrite if over; excess detail to speaker notes
  [ ] Every prioritization: ≤ 5 items; title ≤ 8 words and NO numbers/% /currency; description ≤ 15 words
  [ ] Every prioritization qualifier: label = 1 word; value ≤ 4 words

GATE 2 — STRUCTURAL LIMITS
  [ ] Max 4 zones per slide
  [ ] Max 2 artifacts per zone
  [ ] At least 1 primary zone per content slide
  [ ] No more than 2 primary zones per slide

GATE 3 — ARTIFACT HARD CONSTRAINTS
  [ ] No slide contains more than one strategy artifact (matrix | driver_tree | prioritization | stat_bar | comparison_table | initiative_map | profile_card_set | risk_register) — count across ALL zones; if two are found, split onto separate slides
  [ ] Every bar/line/clustered_bar/horizontal_bar has ≥ 3 categories
  [ ] Every clustered_bar has exactly 2 series with matching units
  [ ] Every pie/donut has ≤ 5 segments
  [ ] dual_axis set to true wherever series have different units
  [ ] Every workflow has coherent, non-crossing nodes and connections
  [ ] Every PRIMARY artifact has artifact_header populated; SECONDARY artifacts use artifact_header only when needed
  [ ] No cards used for part-of-whole, portfolio mix, or mutually exclusive category data
  [ ] No cards used for status buckets, risk categories, or total-plus-components structures
  [ ] No cards-only slide — cards must be accompanied by chart/workflow/insight_text
  [ ] Every group_pie has 2–8 series (entities / pies)
  [ ] Every group_pie has ≤ 7 categories (slices) — HARD MAX 7
  [ ] Every group_pie series values[] length equals categories[] length
  [ ] Every group_pie is allocated ≥ 50% of the zone axis (width or height)
  [ ] group_pie with ≥ 5 pies is allocated ≥ 80% of the zone axis
  [ ] group_pie paired with insight_text standard only (in multi-zone slides)
  [ ] matrix / driver_tree / prioritization appear only in PRIMARY zones
  [ ] matrix / driver_tree / prioritization accompanied only by insight_text
  [ ] matrix occupies ≥ 70% of slide width AND ≥ 50% of slide height (hard minimum)
  [ ] stat_bar zone width: 1 bar col → ≥50%, 2 bar cols → ≥75%, 3 bar cols → 100% of slide width
  [ ] stat_bar zone height ≥ 30% slide height for 2 rows; each additional row adds ~11% (8 rows = 100%)
  [ ] Every stat_bar "bar" column is immediately followed by a "normal" column with empty header
  [ ] stat_bar has 2–8 rows and 1–3 "bar" columns (no two "bar" columns adjacent)
  [ ] Every left_to_right / timeline workflow spans FULL HORIZONTAL WIDTH
  [ ] No zone placed beside (left/right of) a left_to_right workflow
  [ ] Every top_to_bottom / top_down_branching workflow spans FULL VERTICAL HEIGHT
  [ ] No zone stacked above or below a top_to_bottom workflow
  [ ] Any workflow with > 3 nodes uses full-width (left_to_right) or full-height (top_to_bottom)

GATE 4 — ZONE SPATIAL COVERAGE
  [ ] All zone splits sum to 100% of content area — no gaps, no overlaps
  [ ] Every wide artifact is placed in a zone that gives it the required horizontal span
      from its family rules and active override rules
  [ ] Every tall artifact is placed in a zone that gives it the required vertical span
      from its family rules and active override rules
  [ ] PRIMARY zone occupies ≥ 60% of the split axis
  [ ] SECONDARY zone occupies ≤ 40% of the split axis
  [ ] 50/50 splits used only where both zones are CO-PRIMARY (zone_role = "co-primary") or both carry equal PRIMARY weight with no dominant zone

GATE 5 — LAYOUT CONSISTENCY (Layout Mode only)
  [ ] selected_layout_name is a valid name from the available layouts list
  [ ] All zones have layout_hint.split = "full"
  [ ] Left_to_right workflow → O1 override applied; widest single-column layout selected
  [ ] 4-card artifact → O4 override applied
  [ ] Tall horizontal_bar → O5 override applied
  [ ] Any workflow with > 3 nodes triggered the mandatory workflow override

GATE 6 — SLIDE COHERENCE
  [ ] Every zone's message_objective directly supports the slide's key_message
  [ ] No zone exists to fill space
  [ ] No text-only slide unless archetype is "summary" or "recommendation"
  [ ] All other archetypes contain at least one chart, cards, workflow,
      matrix, driver_tree, or prioritization
  [ ] No two zones show the same data in different formats
  [ ] No two zones communicate the same interpretation or implication
  [ ] Every zone adds genuinely new information not covered by any other zone on the slide

Return ONLY a valid JSON array. No explanation. No markdown. No text outside the JSON..`
