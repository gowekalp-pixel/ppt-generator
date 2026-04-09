// AGENT 4 — PHASE 1: ZONE STRUCTURE FINALIZATION
// Zone count rules, structure codes, consultant challenge round.
// Part of AGENT4_SYSTEM — assembled by prompts/agent4/index.js

const _A4_PHASE1 = `PHASE 1 — ZONE STRUCTURE FINALIZATION
(run for each content slide — slide numbers and narrative_role pre-assigned by Agent 3)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

BOARD REALITIES (govern all decisions in this phase)
  1. Boards read the title and Zone 1 first. Design Zone 1 as if it is the only zone they will read.
  2. Boards jump to implications. The zone sequence must anticipate that jump.
  3. If the slide only proves what they already believe, attention is lost.
  4. A technically correct but strategically inert slide fails.
  5. Every zone must earn its place.

STEP 0 — CHALLENGE THE TITLE
  Pressure-test the proposed title before doing anything else.
  A board-ready title must be:
  - declarative, not descriptive
  - specific enough to be falsifiable
  - conclusion-led, not topic-led
  - honest about what the slide will and will not prove
  If the title fails any test: rewrite it and note why.

STEP 1 — LOCK THE STRATEGIC OBJECTIVE
  State what the board must believe, decide, or feel differently after this slide.
  This is the cognitive destination, not the content summary.

STEP 2 — CONFIRM NARRATIVE_ROLE
  Read the narrative_role from Agent 3's slide plan. Do NOT re-derive or alter it.
  It is the primary artifact gate in Phase 2 STEP 0A.

STEP 3 — DERIVE ZONE CONFIGURATION
  Read these fields from Agent 3's slide plan and STEP 1–2 above.
  Commit to a zone configuration before assigning individual zones in STEP 4.

  INPUTS TO READ:
  - narrative_role          (from Agent 3 slide plan)
  - dominant_zone_signal    (from Agent 3 slide plan)
  - co_primary_signal       (from Agent 3 slide plan)

  ZONE COUNT RULES (apply in order, first match wins):
  - narrative_role is methodology_note -> 1 zone, stop
  - narrative_role is transition_narrative -> 1 zone; insight_text only; no data artifacts permitted
  - narrative_role is summary -> 1-2 zones; prefer 1 if one dense synthesis artifact can carry the slide
  - co_primary_signal is yes -> 2 zones, CO-PRIMARY, side-by-side
  - narrative_role is explainer_to_summary -> 3-4 zones; dominant decomposition plus supporting proof
  - narrative_role is validation -> 2-3 zones; dominant proof plus supporting evidence
  - narrative_role is drill_down -> 2-3 zones; dominant decomposition plus support
  - narrative_role is benchmark_comparison -> 2 zones, equal weight unless dominant_zone_signal is yes
  - narrative_role is trend_analysis -> 2 zones, dominant proof plus implication support
  - narrative_role is segmentation -> 2-3 zones, comparison-led
  - narrative_role is waterfall_decomposition -> 2 zones, dominant proof plus explanation
  - narrative_role is scenario_analysis -> 3-4 zones, grid or structured comparison preferred
  - narrative_role is option_evaluation -> 3-4 zones, option comparison or criteria grid preferred
  - narrative_role is risk_assessment -> 2-3 zones, dominant register plus mitigation / implication support
  - narrative_role is recommendations -> 2-3 zones, recommendation plus rationale / ask support
  - narrative_role is exception_highlight -> 2 zones, dominant issue plus implication / action support
  - narrative_role is context_setter or problem_statement -> 2 zones, framing plus consequence / evidence support
  - narrative_role is additional_information -> 1-2 zones; compact supporting evidence only
  - zone_count_signal is 1, 2, 3, or 4 -> use that count as the baseline when no stronger rule above applies
  - strategic_objective implies comparing options, scenarios, or alternatives -> 3-4 zones
  - strategic_objective implies a single core proof with one takeaway -> 2 zones
  - strategic_objective implies a compact synthesis or note -> 1 zone
  - default -> 2 zones

  MODIFIERS (adjust the baseline count after rules above):
  - dominant_zone_signal is yes -> Zone 1 must be DOMINANT regardless of count
  - dominant_zone_signal is no and co_primary_signal is no -> prefer balanced weights across zones
  - narrative_role is scenario_analysis, option_evaluation, or recommendations -> minimum 2 zones; options need space
  - if strategic_objective is primarily "align" or "understand" one core fact -> do not exceed 2 zones
    (choose the reading direction from narrative_role, dominant/co-primary signals, and strategic_objective)

STEP 4 — ASSIGN ZONES WITH STRATEGIC INTENT
  Assign 1 to 4 zones.
  Every zone must pass this test:
  "If I removed this zone, would the board reach a different and worse conclusion?"

  For each zone define:
  - role: PRIMARY | CO-PRIMARY | SECONDARY | SUPPORTING | OPTIONAL
  - strategic_purpose: 15-20 words on why this zone exists for this board on this slide — cannot be generic

STEP 5 — FINALIZE THE ZONE STRUCTURE 
  Choose exactly one code matching the zone count.

  1 zone:  ZS01_single_full

  2 zones: ZS02_stacked_equal | ZS03_side_by_side_equal

  3 zones: ZS04_left_dominant_right_stack | ZS05_right_dominant_left_stack |
           ZS06_top_full_bottom_two | ZS07_top_two_bottom_dominant |
           ZS11_three_rows_equal | ZW01_three_columns_equal

  4 zones: ZS08_quad_grid | ZS09_left_dominant_right_triptych |
           ZS10_top_full_bottom_three | ZW04_four_columns_equal |
           ZW02_three_columns_right_stack | ZW03_three_columns_left_stack

  Selection rules:
  - CO-PRIMARY → always side-by-side, never stacked
  - DOMINANT zone → must occupy the physically largest cell
  - waterfall_decomposition → prefer vertical codes: ZS02, ZS11, ZS04, ZS05
  - risk_assessment, option_evaluation → prefer grid/column codes: ZS08, ZW04, ZW01
  - scenario_analysis → prefer left-to-right codes: ZW01, ZS06
  - context_setter, explainer_to_summary → prefer stacked codes: ZS11, ZS02
  - benchmark_comparison → prefer side-by-side codes: ZS03, ZW01
  - ZW02 / ZW03 → wide canvas 4-zone layouts; dominant left or right column with a stacked pair on the opposite side; require wide canvas (width/height > 1.5)
  - OPTIONAL zone → choose a structure that degrades gracefully if the zone is skimmed
  - Wide canvas (width/height > 1.5): may use ZS.. and ZW.. structures
  - Standard canvas: prefer ZS.. structures

STEP 6 — CONSULTANT CHALLENGE ROUND
  Before finalizing, force these checks:
  - Does Zone 1 read alone with the title create the needed tension or conviction?
  - Is every zone answering a question the board is actually asking?
  - Could the objective be achieved with one fewer zone?
  - Is the sequence preventing misreading, or just adding detail?
  - If the board jumps to Zone 3 first, does the slide still work?
  - Is complexity truly required, or just performative thoroughness?

  ABSOLUTE CONSTRAINTS
  1. Maximum 4 zones; a fifth zone is a second slide.
  2. Zone labels are navigation aids only: 1–3 words.
  3. CO-PRIMARY zones are always side-by-side.
  4. In scenario_analysis and option_evaluation, Zone 1 establishes the current reality or decision criteria — never label it "Proof" or "Recommendation."
  5. In recommendations, the recommendation itself must land in the final zone, not Zone 1.
  6. In risk_assessment and option_evaluation, green/stable KPIs are SUBORDINATE weight.
  7. In alert slides (exception_highlight, problem_statement, risk_assessment), Zone 1 is DOMINANT.
  8. Never choose a structure with more cells than zones assigned.
  9. strategic_purpose cannot be generic.
`
