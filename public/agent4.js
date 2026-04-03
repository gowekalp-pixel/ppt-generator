// ─── AGENT 4 — SLIDE CONTENT ARCHITECT ────────────────────────────────────────
// Input:  state.outline    — presentation brief / slide plan from Agent 3
//         state.contentB64 — original PDF or source document
// Output: slideManifest    — flat JSON array, one object per slide
//
// Agent 4 decides WHAT each slide is trying to say, WHAT zones (messaging arcs)
// it needs, and WHAT artifacts sit inside each zone.
// Agent 5 decides final layout, coordinates, styling, and rendering.
//
// Key concepts:
//   Zone    = a messaging arc — one coherent argument unit
//   Artifact= the visual expression inside a zone (chart, table, workflow, etc.)
//   Each slide: max 4 zones, title/subtitle outside zones
//   Each zone:  max 2 artifacts

// ═══════════════════════════════════════════════════════════════════════════════
// SYSTEM PROMPT (pasted from consultant-authored spec)
// ═══════════════════════════════════════════════════════════════════════════════

const AGENT4_SYSTEM = `You are a senior management consultant acting as the slide content architect for a board-level presentation.
You will receive:
1. A structured presentation brief
2. A batch of slide plans
3. A source document for reference

Your role is to define the HIGH-LEVEL CONTENT STRUCTURE for each slide.
You do NOT design the final slide.
You do NOT decide coordinates, colors, fonts, or exact visual styling.
You DO decide:
- what the slide is trying to prove
- what messaging arcs (zones) it needs
- what artifacts belong inside each zone
- the spatial requirements those artifacts impose
- which layout satisfies those spatial requirements

Return ONLY a valid JSON array with one object per slide.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PHASE 1 — ZONE STRUCTURE FINALIZATION
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
  - narrative_role is summary -> 1-3 zones; prefer 1 if one dense synthesis artifact can carry the slide
  - co_primary_signal is yes -> 2 zones, CO-PRIMARY, side-by-side
  - narrative_role is explainer_to_summary, validation or drill_down -> 3-4 zones; dominant decomposition and supporting proof
  - narrative_role is benchmark_comparison -> 2 zones, equal weight unless dominant_zone_signal is yes
  - narrative_role is trend_analysis -> 2 zones, dominant proof plus implication support
  - narrative_role is segmentation -> 2-3 zones, comparison-led
  - narrative_role is waterfall_decomposition -> 2 zones, dominant proof plus explanation
  - narrative_role is scenario_analysis -> 3-4 zones, grid or structured comparison preferred
  - narrative_role is decision_framework -> 3-4 zones, option comparison or criteria grid preferred
  - narrative_role is risk_register -> 1-2 zones, dominant register plus mitigation / implication support
  - narrative_role is recommendations -> 1-2 zones, recommendation plus rationale / ask support
  - narrative_role is exception_highlight -> 2 zones, dominant issue plus implication / action support
  - narrative_role is context_setter or problem_statement -> 2 zones, framing plus consequence / evidence support
  - strategic_objective implies comparing options, scenarios, or alternatives -> 3-4 zones
  - strategic_objective implies a single core proof with one takeaway -> 2 zones
  - strategic_objective implies a compact synthesis or note -> 1 zone
  - default -> 2 zones

  MODIFIERS (adjust the baseline count after rules above):
  - dominant_zone_signal is yes -> Zone 1 must be DOMINANT regardless of count
  - dominant_zone_signal is no and co_primary_signal is no -> prefer balanced weights across zones
  - narrative_role is scenario_analysis, decision_framework, or recommendations -> minimum 2 zones; options need space
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
           ZS11_three_rows_equal | ZW01_three_columns_equal |
           ZW02_three_columns_right_stack | ZW03_three_columns_left_stack

  4 zones: ZS08_quad_grid | ZS09_left_dominant_right_triptych |
           ZS10_top_full_bottom_three | ZW04_four_columns_equal

  Selection rules:
  - CO-PRIMARY → always side-by-side, never stacked
  - DOMINANT zone → must occupy the physically largest cell
  - waterfall_decomposition → prefer vertical codes: ZS02, ZS11, ZS04, ZS05
  - risk_register, decision_framework → prefer grid/column codes: ZS08, ZW04, ZW01
  - scenario_analysis → prefer left-to-right codes: ZW01, ZS06
  - context_setter, explainer_to_summary → prefer stacked codes: ZS11, ZS02
  - benchmark_comparison → prefer side-by-side codes: ZS03, ZW01
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
  3. Zone question must be answerable by evidence alone.
  4. CO-PRIMARY zones are always side-by-side.
  5. In scenario_analysis and decision_framework, Zone 1 establishes the current reality or decision criteria — never label it "Proof" or "Recommendation."
  6. In recommendations, the recommendation itself must land in the final zone, not Zone 1.
  7. In risk_register and decision_framework, green/stable KPIs are SUBORDINATE weight.
  8. In alert slides (exception_highlight, problem_statement, risk_register), Zone 1 is DOMINANT.
  9. Never choose a structure with more cells than zones assigned.
  10. strategic_purpose cannot be generic.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PHASE 2 — ZONE CONTENT DERIVATION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

You are the Content Architect.
You have the locked zone structure from Phase 1, Agent 3's key content as
directional guidance, and the source document as the authoritative evidence base.

Your job: derive the best possible content for each zone.
Use Agent 3's key content to understand intent and direction.
Use the source document and your own analytical judgment to surface the
most compelling, accurate, and board-ready content for that zone.
You are NOT selecting artifacts. You are NOT designing layouts.

──────────────────────────────────────────────────────────
STEP 1 — DERIVE ZONE CONTENT
──────────────────────────────────────────────────────────
For each zone, working from its question and strategic_purpose:

  - Surface the strongest evidence, arguments, insights, or recommendations
    from the source document that answer the zone question.
  - Apply your analytical judgment: synthesise, prioritise, and sharpen.
    A board deserves the most incisive version of the content, not a transcript.
  - Ground every data claim in a specific figure from the source.
    Structure every argument or recommendation as explicit logic —
    not vague assertions.
  - Separate what belongs on the slide from what belongs in speaker notes.
    The slide carries the claim. The notes carry qualification and detail.

──────────────────────────────────────────────────────────
STEP 2 — PRODUCE ZONE CONTENT BRIEF
──────────────────────────────────────────────────────────
  {
    "slide_number":     1,
    "narrative_role": "explainer_to_summary",
    "slide_title_draft":  "string — declarative, specific, conclusion-led, honest",
    "zones:[
     {
      "strategic_purpose": {string"}
      "zone_role": "primary" | "co-primary" | "secondary" | "supporting" | "optional",
      "zone_id":          "z1",
      "zone_content":    ["<sharpest claim or fact with unit>", ...],
      "speaker_overflow": ["<supporting detail, qualifications, secondary data>", ...],
      "content_gap":      "string — note if source evidence is thin; null if sufficient"
      }
]

  CONSTRAINTS
  1. Every figure in slide_content[] must include its unit and source basis.
  2. No two zones on the same slide may carry overlapping slide_content[].
  3. content_gap is not a licence to invent — it signals Phase 2 to choose
     a lower-density artifact


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PHASE 3 — ARTIFACT FINALIZATION ACROSS ZONES
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

You are the Artifact Architect.
You receive the completed Zone and Zone Content Brief from Phase 2
Based on the output of phase 2, Your job is to assign the right artifact to each zone, enforce spatial and density constraints,
and produce a plan the visual designer can execute without ambiguity.
if there is a content_gap for a zone - recheck the main document to fill the information
You are NOT designing charts, colors, fonts, labels, or coordinates.
Each zone may have a MAXIMUM of 2 artifacts.

──────────────────────────────────────────────────────────
STEP 0A — ROLE-BASED ARTIFACT PRE-FILTER (execute before all other steps)
──────────────────────────────────────────────────────────
Read the slide_narrative_role locked in Phase 1.
Apply the constraints below before any data-shape or zone-role rules.
These are HARD constraints — no override permitted regardless of content.

  recommendations:
    PERMITTED:  prioritization (primary), initiative_map, comparison_table, insight_text|grouped,
                cards (target/milestone values only — not actuals), table (milestone summaries only — max 6 rows, max 4 columns; use comparison_table if options vs criteria)
    FORBIDDEN:  charts showing actuals, workflow, matrix, driver_tree
    Cards rule: card values must be future targets or milestones — do not restate actuals already shown on earlier slides

  methodology_note:
    PERMITTED:  insight_text, table (definitions/rates only — max 6 rows, max 4 columns)
    FORBIDDEN:  cards, chart, prioritization, workflow, matrix, driver_tree
    Reason: no analytical claim — data definitions only

  transition_narrative:
    PERMITTED:  insight_text only
    FORBIDDEN:  all data artifacts
    Reason: narrative connector — no new data introduced

  context_setter:
    PERMITTED:  cards (baseline KPIs), insight_text, stat_bar, comparison_table
    FORBIDDEN:  prioritization, workflow, matrix, driver_tree
    Reason: neutral baseline — no verdict, no action, no process

  exception_highlight:
    PERMITTED:  profile_card_set or cards (sentiment MUST be "negative" or "warning"), insight_text, chart (as supporting evidence)
    FORBIDDEN:  prioritization, workflow
    Cards rule: if profile_card_set or cards are used, every card sentiment field must be "negative" or "warning" — positive sentiment cards are blocked

──────────────────────────────────────────────────────────
STEP 0C — READ THE ZONE BRIEF STRATEGICALLY
──────────────────────────────────────────────────────────
  For each zone re-read the information 
  Ask: what is the MINIMUM artifact that answers this zone's question convincingly for a board?
  Start minimum. Add a second artifact only if the first leaves a material gap in the argument.

──────────────────────────────────────────────────────────
STEP 1 — ARTIFACT SELECTION
──────────────────────────────────────────────────────────

AVAILABLE ARTIFACT TYPES:
  chart:            bar | clustered_bar | horizontal_bar | line | area | pie | donut | combo | group_pie
  stat_bar
  insight_text:     standard | grouped
  cards
  profile_card_set
  workflow:         process_flow | hierarchy | decomposition | timeline
  table             ← LAST RESORT — only permitted for: (1) methodology_note rate/definition lookups; (2) recommendations milestone summaries. ALL other uses FORBIDDEN. Max 6 rows, max 4 columns.
  comparison_table
  initiative_map
  risk_register
  matrix
  driver_tree

─── ZONE ROLE → PERMITTED ARTIFACTS ───────────────────────
 Based on the inputs finalize the artifacts along with their content, with only below caveats. 
 As a management consultant, choose the best representation that should create the most impactful messaging for each slide.

  PRIMARY / DOMINANT zones:
  - Can have 2 Artifacts (Primary & Secondary)
  - 50%-100% of Slide's Horizontal or Vertical Area

  SECONDARY zones:
  - Can have 2 artifacts (Primary & Secondary)
  - Must answer a specific question, not just add detail

  CO-PRIMARY zones :
  - Can have 2 artifacts (Primary & Secondary)
  - Both zones must carry artifacts 
  - Both should be the same artifact class where data richness allows
  - Mismatched classes imply hierarchy — flag as CO-PRIMARY asymmetry if unavoidable

  FOR PRIMARY, SECONDARY AND CO-PRIMARY Zones, PRIMARY ARTIFACT SHOULD BE ONLY SELECTED FROM BELOW:
    - Charts
    - stat_bar
    - Comparison_table
    - initiative_map
    - risk_register
    - cards
    - profile_card_set
    - workflow
    - matrix
    - driver_tree
    - prioritization
  
  FOR PRIMARY, SECONDARY AND CO-PRIMARY Zones, PERMITTED SECOND ARTIFACT PAIRINGS and Size of the second artifact relative to the zone (e.g. max 30% of zone size):
  Primary artifact                          Permitted Second Artifact Size  
  Charts (Except pie, donut and Group Pie)  Insight_text (any subtype), max 30% of zone size
  group_pie                                 insight_text (any subtype), max 30% of zone size
  pie / donut                               cards (1–2) or insight_text, max 40% of zone size
  stat_bar                                  insight_text standard (callout only — 1–2 points), stacked above only, max 25% of zone size
  comparison_table                          insight_text standard, max 30% of zone size
  initiative_map                            insight_text standard (framing headline — 1–2 points), stacked above only, max 25% of zone size
  profile_card_set                          no second artifact permitted
  risk_register                             no second artifact permitted
  matrix                                    insight_text grouped, max 30% of zone size
  driver_tree                               insight_text grouped, max 30% of zone size
  prioritization                            no second artifact permitted 
  process_flow / timeline                   no second artifact permitted
  hierarchy / decomposition                 insight_text, max 30% of zone size
  cards (1–3)                               no second artifact permitted
  cards (4+)                                no second artifact permitted

SUPPORTING / SUBORDINATE zones:
  - Can have 1 artifact only
  - Compact only
  - Preferred: standard insight_text, cards (1–3), simple chart (≤6 categories), profile_card_set (≤3 profiles)
  - Avoid: workflow, driver_tree, matrix, large table, initiative_map, risk_register

OPTIONAL zones:
  - One artifact only, compact only: standard insight_text, cards (1–3), simple chart
  - If it cannot fit in compact form, the zone should not exist

─── CHART FAMILY SELECTION INDICATORS ─────────────────────────

  bar:            3+ categories, one series, no time axis (vertical columns)
  horizontal_bar: same as bar but rotated — prefer when labels are long OR categories > 6
  line:           trend over time (months/quarters/years as categories)
  pie:            composition — values sum to ~100%, HARD MAX 5 segments
                  If > 5 segments: HARD REJECT → convert to horizontal_bar automatically
  donut:          same as pie with centre callout; HARD MAX 5 segments
                  If > 5 segments: HARD REJECT → convert to horizontal_bar automatically
  waterfall:      bridge or variance — series items typed positive/negative/total
  clustered_bar:  EXACTLY 2 series compared across same 3+ categories
                  If only 1 series: use bar instead
                  HARD REJECT — units differ (e.g. count vs ₹, % vs volume):
                    → NEVER use clustered_bar when series units differ
                    → AUTOMATICALLY convert to combo (dual-axis) if trend is the insight
                    → AUTOMATICALLY convert to two separate bar artifacts if categories are the insight
                  Both series MUST share EXACTLY the same unit (both ₹, both %, both count, etc.)
                  SELF-CHECK before finalising any clustered_bar: "Do Series A and Series B share the same unit?" If NO → reject clustered_bar → replace with combo
  combo:          dual-axis overlay (bar + line); MANDATORY when two measures have different units
                  Use combo — NOT clustered_bar — whenever series units differ, regardless of
                  whether a time axis exists. Primary series → bar. Secondary series → line.
  group_pie:      multiple related pie charts — one pie per entity (e.g. industry, region, product);
                  all pies share the same slice categories (same distribution breakdown);
                  single shared legend above all pies; entity label below each pie centre-aligned.
                  Use when: comparing the SAME composition across 2–8 distinct entities is the insight.
                  HARD REJECT conditions:
                    entities (series) < 2     → convert to single pie
                    entities (series) > 8     → convert to clustered_bar or profile_cards
                    slices (categories) > 7   → convert to clustered_bar
                    slices (categories) < 2   → invalid
                  Zone allocation: requires ≥ 50% of zone horizontal OR vertical axis.
                  Layout direction:
                    If 100% horizontal available → arrange pies side-by-side in a single row
                    If 100% vertical available   → distribute pies across multiple rows
                    If 50–80% allocation         → MAX 4 pies
                    If > 80% allocation          → MAX 8 pies
 
 ─── Other artifacts SELECTION INDICATORS ───────────────────────── 
  stat_bar        
  A horizontal bar chart where each bar is accompanied by an inline qualitative
  annotation (label or descriptor) placed to the right of or alongside the bar.
  Use when: ranking is the primary message AND each entity needs a one-phrase
  qualifier that a separate insight_text zone would fail to connect visually.
  Examples: courier partners ranked by cost with "use case" label; SKUs ranked
  by revenue with "growth trajectory" label; cities ranked by orders with
  "priority tier" label.

  comparison_table 
  
  A structured grid where rows are options/candidates and columns are evaluation
  criteria. Each cell contains a judgment rating, not a raw number.
  Use when: the board needs to see WHICH option wins against WHICH criteria.
  The recommended option must be visually distinguished (recommended_option field).
  NEVER use plain table for option-vs-criteria data — comparison_table is mandatory.

  initiative_map
  A structured grid where each row is a parallel work stream or initiative.
  Unlike prioritization, rows have NO implied rank between them.
  Unlike table, each row is a self-contained action block, not a data lookup row.
  Use when: showing 3–6 parallel initiatives each described by 3–4 structured
  dimensions (e.g. owner, timeline, budget, KPI, status).
  NEVER use when rows have a rank order (use prioritization instead).
  NEVER use when rows are process steps (use workflow instead).

  profile_card_set
  An entity-first card layout where each card describes one entity using a set
  of key-value attribute pairs. Attributes may differ per entity (heterogeneous).
  Unlike cards (which are metric-first and structurally parallel), profile cards
  are read vertically within each card, not scanned horizontally across cards.
  Use when: rows describe entities with mixed attribute types (feature type +
  geography + revenue range + status all on the same entity row).
  NEVER use when metrics are structurally identical across entities (use cards).

  risk_register
  A severity-encoded list of risks or issues where row background color IS the
  primary data channel (not decoration). Multiple rows may have different severity
  levels simultaneously — this is what distinguishes it from highlight_rows on a
  plain table (which marks only one row at a time).
  Use when: risks or issues must be shown with per-row severity variation AND
  the board's eye must be directed to critical items without reading every cell.
  NEVER use plain table when severity-by-row is the primary signal.

─── CARDS SELECTION Indicators ────────────────────────────────

  Cards are ONLY valid when ALL four conditions are met:
  1. Metrics are independent (no shared denominator, no structural link between them)
  2. Metrics do NOT form a distribution or composition
  3. Metrics do NOT require comparison across categories
  4. Metrics are headline KPIs or executive summary stats

  If ANY of the following is true → use chart instead:
  - Categories form a whole (sum to 100% or a total)
  - Categories are mutually exclusive and comparable
  - The insight depends on relative size or ranking across categories
  - Values are status / risk / stage buckets of the same total
  - One value is an aggregate total and others are named components of that total

  AUTO-UPGRADE RULE (mandatory check before finalising):
  IF cards ≥ 3 AND values are numeric AND categories are comparable
  → Automatically upgrade to bar or pie chart + insight_text zone
  → Do not retain cards when data fits a chart

  Additional mandatory override:
  IF one card is an aggregate total AND remaining cards are named categories / statuses /
  stages of that same total
  → Automatically replace category cards with chart or decomposition workflow
  → At most ONE card may remain: the total / anchor KPI

  Card count limits:
  - Max 4 cards in full-width zones
  - Max 2 cards in side zones (left_X, right_X, tl, tr, bl, br)
  - Cards (4): zone MUST be full-width — no other zones on the same slide

  A slide CANNOT contain only cards. Cards must always be paired with at least one of:
  chart | workflow | table | insight_text | Stat_Bar

─── WORKFLOW SELECTION Indicators ─────────────────────────────

  process_flow:   linear sequence of steps; 4–6 nodes; left_to_right only;
                  requires full slide width and at least 50% slide height;
                  use when showing HOW something happens step by step

  timeline:       phased chronological progression; 4–6 nodes; left_to_right only;
                  must stretch across full slide width with at least 50% height;
                  use when showing WHEN things happen in sequence

  hierarchy:      parent-child reporting structure; minimum 3 levels;
                  top_down_branching only; requires at least 50% width and full content height;
                  use when showing WHO reports to whom or structural organisation

  decomposition:  top number or concept split into components; max 6 nodes;
                  direction: left_to_right or top_to_bottom or top_down_branching;
                  left_to_right requires full width; top_to_bottom requires full height;
                  use when showing WHAT something is made of

  Flow direction placement rules:
  - left_to_right / timeline:   MUST span full horizontal width; companion zones stacked
                                 above or below only; NEVER side-by-side
  - top_to_bottom / top_down_branching: MUST span full vertical height; companion zones
                                        left or right only; NEVER stacked above or below

  Node count rules:
  - process_flow: minimum 4 nodes
  - timeline: minimum 4 nodes
  - hierarchy: minimum 3 distinct levels
  - decomposition: max 6 nodes
  - Any workflow with ≥ 4 nodes triggers the full-width (left_to_right) or
    full-height (top_to_bottom) rule as a HARD OVERRIDE

─── INSIGHT TEXT SELECTION Indicators ─────────────────────────

  ZONE AREA = zone W% × zone H% as a fraction of total slide content area.
  Small zone  = zone area < 25%
  Medium zone = zone area 25%–50%
  Large zone  = zone area > 50%

  STANDARD MODE — point count cap by zone area (HARD MAX):
  ┌─────────────────┬───────────────┐
  │ Zone area       │ Max points[]  │
  ├─────────────────┼───────────────┤
  │ Small  (< 25%)  │ 3 points      │
  │ Medium (25–50%) │ 5 points      │
  │ Large  (> 50%)  │ 8 points      │
  └─────────────────┴───────────────┘
  - Each bullet ≤ 12 words. Count every word — articles, numbers, units each count as 1.
  - 1–2 data points per bullet maximum. Remaining detail goes to speaker notes.
  - Data-first phrasing: lead with the number or entity. No "This shows that…" preamble.
  - No compound sentences with dashes or semicolons — one idea per bullet.
  - If point count would exceed the zone cap: move lower-priority points to speaker notes.
  - If findings are thematically distinct and would exceed 4: switch to GROUPED mode.

  standard (1–2 points): compact callout or headline stat; annotation role only;
                          never the sole artifact in a PRIMARY zone
  standard (3 points+):  evidence list; use zone area cap above to determine max

  GROUPED MODE — groups × bullets cap by zone area (HARD MAX):
  ┌─────────────────┬──────────────────────────────┐
  │ Zone area       │ Max groups × bullets / group  │
  ├─────────────────┼──────────────────────────────┤
  │ Small  (< 25%)  │ 2 groups × 2 bullets          │
  │ Medium (25–50%) │ 3 groups × 3 bullets          │
  │ Large  (> 50%)  │ 5 groups × 3 bullets          │
  └─────────────────┴──────────────────────────────┘
  - Each bullet ≤ 12 words. 1–2 data points per bullet. Remainder to speaker notes.
  - Group headers: 2–4 words, no verbs, label the theme not the finding.
  - Board scans section headers first — headers must be self-explanatory at a glance.
  - grouped insight_text IS permitted in a PRIMARY zone if ≥ 4 substantive findings.

  INSIGHT TEXT VISUAL MODES:
  - STANDARD mode (points[]): flat list read top-to-bottom; use for sequential or independent findings
  - GROUPED mode (groups[]): board scans section headers first then reads bullets;
    use for 3+ thematically distinct finding clusters
  - Never use both points[] and groups[] in the same artifact

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PHASE 4 — LAYOUT FINALIZATION
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

You are the Layout Architect.
You receive the completed Slide Artifact Plan from Phase 3.
Translate zone structure + artifact assignments into precise spatial instructions.
Make NO new strategic decisions. Do NOT change artifacts or zone roles.
One judgment only: given these artifacts in these zones, what is the correct spatial
arrangement — and does a brand layout already exist, or must it be constructed from splits?

EXECUTION MODE
  LAYOUT MODE  — brand layouts provided (≥ 5 named layouts): use layout names; do not compute splits
  SCRATCH MODE — fewer than 5 brand layouts: compute zone splits from first principles

STEP 1 — CHARACTERISE THE CONTENT STRUCTURE

  1a. Classify artifacts:
      wide_artifacts     = artifacts with MIN_W ≥ 70% (process_flow, timeline, wide charts > 6 cat,
                           decomposition L→R > 3 nodes, cards ≥ 4, group_pie with ≥ 5 pies,
                           stat_bar [always], initiative_map [always], profile_card_set ≥ 4 profiles)
      tall_artifacts     = artifacts with MIN_H ≥ 55% (hierarchy, vertical decomposition,
                           horizontal_bar > 6 rows, driver_tree, matrix, prioritization > 5 rows,
                           risk_register [always], comparison_table ≥ 4 options)
      reasoning_artifacts = matrix | driver_tree | prioritization | comparison_table
      workflow_artifacts  = process_flow | timeline | hierarchy | decomposition
      structured_display  = stat_bar | initiative_map | profile_card_set | risk_register | comparison_table
      peer_zones          = zones where both are CO-PRIMARY (P2 or P5)

  1b. Canonical content structure (first match wins):
      wide_artifacts ≥ 1                             → WIDE_DOMINANT
      tall_artifacts ≥ 1 AND total_zones = 1         → TALL_DOMINANT
      total_zones = 1 AND total_artifacts = 1        → SINGLE
      total_zones = 1 AND total_artifacts = 2        → SINGLE_PAIR
      total_zones = 2 AND peer_zones = 1             → PEER_TWO
      total_zones = 2 AND peer_zones = 0             → PRIMARY_SUPPORT_TWO
      total_zones = 3 AND top zone spans full width  → TOP_PLUS_TWO
      total_zones = 3 AND left zone spans full height → LEFT_PLUS_TWO
      total_zones = 3 AND all peer                   → THREE_PEER
      total_zones = 4 AND all peer                   → QUAD_PEER
      total_zones = 4 AND one dominant               → DOMINANT_PLUS_THREE
      fallback                                       → PRIMARY_SUPPORT_TWO

STEP 2 — APPLY OVERRIDE RULES (check every artifact in order)

  O1. LEFT-TO-RIGHT WORKFLOW
      Trigger: process_flow, timeline, or left_to_right decomposition
      → Zone MUST span full horizontal width
      → Companion (if any) MUST be stacked above or below — NEVER side-by-side
      → content_structure forced to WIDE_DOMINANT
      → Mandatory for ≥ 4 nodes

      SELF-CHECK: If a left_to_right workflow shares horizontal space with any other zone
      → STOP. Discard layout. Redesign: workflow = full-width zone; companion = stacked zone.

  O2. TOP-DOWN WORKFLOW
      Trigger: hierarchy, or top_down / top_to_bottom decomposition
      → Zone MUST span full vertical height
      → Companion MUST be placed left or right — NEVER stacked
      → content_structure forced to TALL_DOMINANT
      → Mandatory for ≥ 3 levels or ≥ 4 nodes

  O3. WIDE CHART
      Trigger: any chart with > 6 categories
      → Zone MUST span full slide width
      → Companion insight_text embedded INSIDE same zone, not a separate zone
      → content_structure forced to WIDE_DOMINANT

  O4. FOUR-PLUS CARDS
      Trigger: cards ≥ 4
      → Zone MUST span full horizontal width
      → count = 4: no other zones on slide
      → count 5–10: companion stacked above or below only
      → content_structure forced to WIDE_DOMINANT

  O5. TALL HORIZONTAL BAR
      Trigger: horizontal_bar > 6 rows
      → Zone must occupy ≥ 65% slide height
      → Companion zone goes in remaining strip above
      → content_structure forced to TALL_DOMINANT

  O6. REASONING ARTIFACT
      Trigger: matrix, driver_tree, or prioritization
      → Zone must be DOMINANT — minimum 60% content area
      → Any companion zone must be subordinate
      → Side companion permitted only if zone ≥ 60% width

  O7. PORTRAIT / NON-WIDESCREEN
      Trigger: slide format is portrait or 4:3, OR width:height ratio < 1.5
      → ZW-series codes are PROHIBITED
      → Convert any ZW selection to the ZS equivalent

  O8. GROUP PIE
      Trigger: group_pie with ≥ 5 pies
      → Zone must occupy ≥ 80% of slide width (horizontal layout) OR ≥ 80% height (vertical layout)
      → content_structure forced to WIDE_DOMINANT (horizontal) or TALL_DOMINANT (vertical)
      → Companion insight_text MUST be in a separate stacked zone, not side-by-side with group_pie

      Trigger: group_pie with 2–4 pies
      → Zone must occupy ≥ 50% of slide width or height
      → Companion artifact may share the same zone or an adjacent zone

      Layout direction rule:
      → If slide width ≥ 1.5× height (widescreen): prefer horizontal arrangement (pies in one row)
      → If slide height ≥ slide width: prefer vertical arrangement (pies in multiple rows)

      SELF-CHECK: If a group_pie with ≥ 5 pies is allocated less than 80% of the slide axis
      → STOP. Increase zone allocation OR reduce pie count to ≤ 4.

  O9. STAT BAR
      Trigger: stat_bar (any row count)
      → Zone MUST span full horizontal slide width — label + bar track + value + annotation
        all require horizontal space in sequence; truncating width breaks the reading flow
      → content_structure forced to WIDE_DOMINANT
      → No zone may sit to the left or right of a stat_bar zone
      → Companion insight_text (if present) MUST be stacked above only (never below, never beside)
        — the stat_bar reading direction is top-to-bottom; a zone below steals eye flow
      → artifact_arrangement: single; artifact_coverage_hint: full; internal_alignment: fill

      SELF-CHECK: If stat_bar is placed in a side zone (left_X, right_X, tl, tr, bl, br)
      → STOP. Reassign stat_bar to a full-width zone and move any companion to a stacked zone above.

  O10. INITIATIVE MAP
      Trigger: initiative_map (any row/column count)
      → Zone MUST span full horizontal slide width — swim-lane columns require the full width
        to give each phase/dimension enough space for cell content
      → content_structure forced to WIDE_DOMINANT
      → No zone may sit to the left or right of an initiative_map zone
      → Companion insight_text (if present) MUST be stacked above only — max 25% slide height
      → artifact_arrangement: single; artifact_coverage_hint: full; internal_alignment: fill

      SELF-CHECK: If initiative_map is placed in a side zone
      → STOP. Reassign to full-width zone.

  O11. RISK REGISTER
      Trigger: risk_register (any item count)
      → Zone must occupy ≥ 65% of slide HEIGHT — severity-banded rows stack vertically;
        compressing height truncates the pip encoding and status pills
      → content_structure: SINGLE (no companion) or PRIMARY_SUPPORT_TWO (companion above only)
      → If companion zone exists: companion takes top 30%; risk_register takes bottom 70%
      → No zone may sit to the left or right of a risk_register zone
      → artifact_arrangement: single; artifact_coverage_hint: full; internal_alignment: fill

      SELF-CHECK: If risk_register zone height < 65%
      → STOP. Remove companion zone OR stack companion above at 25% height, risk_register below at 75%.

  O12. COMPARISON TABLE
      Trigger: comparison_table (any option/criteria count)
      → Zone must be DOMINANT — minimum 60% content area
      → Full width strongly preferred: criteria columns require horizontal space per column;
        narrow zones compress criteria headers below readability
      → If slide has two zones: comparison_table takes ≥ 60% width; companion takes ≤ 40%
        and should contain insight_text only — no second data artifact alongside a comparison_table
      → If comparison_table is the only artifact: assign full zone (SINGLE)
      → artifact_arrangement: single; artifact_coverage_hint: full; internal_alignment: fill

      SELF-CHECK: If comparison_table zone width < 55% of slide width
      → STOP. Either make it full width (SINGLE) or primary-left (left_60 or wider).

  O13. PROFILE CARD SET
      Trigger: profile_card_set ≥ 4 profiles
      → Zone MUST span full horizontal slide width — grid layout requires 3+ columns
      → content_structure forced to WIDE_DOMINANT
      → No companion zone permitted when profiles ≥ 4
      → artifact_arrangement: single; artifact_coverage_hint: full; internal_alignment: fill

      Trigger: profile_card_set 2–3 profiles
      → Zone must occupy ≥ 50% slide width
      → Companion insight_text permitted in adjacent zone (right side, ≤ 40% width)
      → artifact_arrangement: single; artifact_coverage_hint: full; internal_alignment: fill

      SELF-CHECK: If profile_card_set ≥ 4 profiles is not in a full-width zone
      → STOP. Reassign to full-width zone. Remove companion zone.

STEP 3 — SELECT LAYOUT

  LAYOUT MODE — map content_structure to brand layout:

    HARD CONSTRAINT FILTER: reject any candidate layout that violates any active override
    from Step 2 or gives any zone less space than the artifact's minimum geometry from Phase 3.

    10% SOFT SIZING GAP RULE: after hard constraints pass, a brand layout is acceptable only if
    each zone's sizing gap is within 10% of the target. If no layout satisfies both, fall back
    to SCRATCH MODE for this slide.

    content_structure             Artifact config             → Layout name pattern
    SINGLE                        1 wide artifact             → "Body" / "1 Across"
    SINGLE_PAIR                   2 artifacts, horizontal     → "2 Across" / "1 on 1"
    SINGLE_PAIR                   2 artifacts, stacked        → "Body" / "1 Across"
    PEER_TWO                      1 artifact per zone         → "2 Across" / "1 on 1"
    PEER_TWO                      2 artifacts per zone (4)    → "2 on 2" / "4 Block"
    PRIMARY_SUPPORT_TWO           chart primary               → "1 on 1" wide left
    PRIMARY_SUPPORT_TWO           workflow primary            → widest single-col
    THREE_PEER                    3 artifacts                 → "3 Across" / "3 Col"
    TOP_PLUS_TWO                  top wide + 2 below          → "1 on 2" / "1 on 3"
    LEFT_PLUS_TWO                 left tall + 2 right         → "2 on 1" stacked right
    QUAD_PEER                     4 peer artifacts            → "2 on 2" / "4 Block" / "4 Across"
    DOMINANT_PLUS_THREE           1 dominant + 3 support      → "1 on 3" / asymmetric
    WIDE_DOMINANT                 any wide artifact           → widest single-col layout
    TALL_DOMINANT                 any tall artifact           → tallest primary col layout

    Tiebreak: when two layouts match, select the one where the primary artifact zone is larger.
    If no brand layout survives hard constraints and 10% gap rule → switch to SCRATCH MODE.

  SCRATCH MODE — derive zone splits:

    ALLOWED SPLIT VALUES:
    Single zone:            full
    Two zones, side-by-side: left_50+right_50 | left_60+right_40 | left_40+right_60
    Two zones, stacked:     top_30+bottom_70 | top_40+bottom_60 | top_50+bottom_50
    Three zones (top+split): top_left_50 + top_right_50 + bottom_full
    Three zones (left+split): left_full_50 + top_right_50_h + bottom_right_50_h
    Four zones:             tl + tr + bl + br

    SPLIT SELECTION LOGIC:
    SINGLE / WIDE_DOMINANT / TALL_DOMINANT         → split = full
    SINGLE_PAIR (2 artifacts, 1 zone)              → split = full; resolve via artifact_arrangement
    PEER_TWO                                       → left_50 + right_50
                                                     (left_60 + right_40 if one is noticeably wider)
    PRIMARY_SUPPORT_TWO (chart/table/reasoning)    → left_60 + right_40
    PRIMARY_SUPPORT_TWO (wide artifact primary)    → top_40 + bottom_60 (primary in bottom)
    PRIMARY_SUPPORT_TWO (proof + annotation)       → top_30 + bottom_70
    TOP_PLUS_TWO                                   → top_left_50 + top_right_50 + bottom_full
    LEFT_PLUS_TWO                                  → left_full_50 + top_right_50_h + bottom_right_50_h
    THREE_PEER                                     → left_33 + mid_33 + right_33 (flag: three_peer_columns)
    QUAD_PEER / DOMINANT_PLUS_THREE                → tl + tr + bl + br

    STRUCTURED DISPLAY ARTIFACT SPLITS (override generic logic when triggered):
    stat_bar alone                                 → split = full
    stat_bar + companion insight_text              → top_30 + bottom_70 (insight_text top, stat_bar bottom)
    initiative_map alone                           → split = full
    initiative_map + companion insight_text        → top_25 + bottom_75 (insight_text top, initiative_map bottom)
    risk_register alone                            → split = full
    risk_register + companion                      → top_30 + bottom_70 (companion top, risk_register bottom)
    comparison_table alone                         → split = full
    comparison_table + companion insight_text      → left_65 + right_35
    profile_card_set ≥ 4 alone                     → split = full
    profile_card_set 2–3 + companion insight_text  → left_60 + right_40

    VISUAL WEIGHT CORRECTION (apply after assigning splits):
    PRIMARY zone must occupy ≥ 60% of content area
    SECONDARY zone must occupy ≤ 40%
    SUPPORTING zone must occupy ≤ 25%
    50/50 split only valid when both zones are CO-PRIMARY

    Correction table:
    primary + secondary side-by-side  → left_60 + right_40
    primary + secondary stacked       → top_40 + bottom_60
    primary + supporting stacked      → top_30 + bottom_70

STEP 4 — ARTIFACT PLACEMENT HINTS (all modes)

  For every artifact in every zone:

  artifact_arrangement (how multiple artifacts are arranged within the same zone)
    horizontal — side-by-side within zone
                 use for: chart+insight_text, cards+insight_text, table+insight_text
                 preferred internal split: 60/40 (proof left, annotation right)
    vertical   — stacked within zone
                 use for: workflow+insight_text (workflow top), chart+table
                 preferred internal split: 70/30 (proof top) or 60/40
    single     — one artifact only

  artifact_coverage_hint (approximate % of zone area this artifact should occupy)
    full      — single artifact in zone
    75pct     — workflow top in a vertical pair
    60pct     — primary of a horizontal pair
    40pct     — secondary of a horizontal pair
    compact   — annotation below a workflow; reasoning artifacts: never compact

  artifact_split_hint (% of zone width or height this artifact occupies when shared)
    100 — single artifact
     70 — workflow (vertical pair, top)
     60 — primary of horizontal pair
     50 — CO-PRIMARY pair (each)
     40 — secondary of horizontal pair
     30 — annotation (vertical pair, bottom)
     25 — supporting

  flow_direction (workflow artifacts only)
    left_to_right | top_to_bottom | top_down_branching

  internal_alignment
    fill       — chart, table, workflow, stat_bar, initiative_map, risk_register,
                 comparison_table, profile_card_set
    center     — cards (row)
    top_left   — insight_text (compact)
    top_center — insight_text (wide)

  STRUCTURED DISPLAY PLACEMENT HINTS (mandatory when artifact is stat_bar, initiative_map,
  risk_register, comparison_table, or profile_card_set):

    stat_bar
      artifact_arrangement:     single
      artifact_coverage_hint:   full
      artifact_split_hint:      100
      internal_alignment:       fill
      zone_constraint:          full_width  ← emit this field on the zone object

    initiative_map
      artifact_arrangement:     single
      artifact_coverage_hint:   full
      artifact_split_hint:      100
      internal_alignment:       fill
      zone_constraint:          full_width  ← emit this field on the zone object

    risk_register
      artifact_arrangement:     single
      artifact_coverage_hint:   full
      artifact_split_hint:      100
      internal_alignment:       fill
      zone_constraint:          full_width  ← emit this field on the zone object

    comparison_table
      artifact_arrangement:     single
      artifact_coverage_hint:   full
      artifact_split_hint:      100
      internal_alignment:       fill
      zone_constraint:          full_width_preferred  ← full width if alone; dominant-left if paired

    profile_card_set (≥ 4 profiles)
      artifact_arrangement:     single
      artifact_coverage_hint:   full
      artifact_split_hint:      100
      internal_alignment:       fill
      zone_constraint:          full_width

    profile_card_set (2–3 profiles)
      artifact_arrangement:     single
      artifact_coverage_hint:   full
      artifact_split_hint:      100
      internal_alignment:       fill
      zone_constraint:          dominant_left  ← min 50% width

STEP 5 — SELF-CHECK BEFORE OUTPUT

  [ ] O1 ACTIVE: no left_to_right workflow shares horizontal space with any zone
  [ ] O2 ACTIVE: no top-down workflow is stacked above or below another zone
  [ ] O9 ACTIVE: stat_bar is in a full-width zone; no sibling zone shares horizontal space
  [ ] O10 ACTIVE: initiative_map is in a full-width zone; no sibling zone shares horizontal space
  [ ] O11 ACTIVE: risk_register zone height ≥ 65% slide height; companion (if any) stacked above only
  [ ] O12 ACTIVE: comparison_table zone width ≥ 55% slide width; companion (if any) is insight_text only
  [ ] O13 ACTIVE: profile_card_set ≥ 4 profiles is in full-width zone with no companion
  [ ] PRIMARY zone occupies ≥ 60% content area (unless both zones are CO-PRIMARY)
  [ ] No SECONDARY zone is larger than PRIMARY zone
  [ ] No reasoning artifact (incl. comparison_table) is in a zone < 45% W or < 50% H
  [ ] No 4-card artifact shares slide with another zone
  [ ] LAYOUT MODE: selected layout has no more content cells than required
  [ ] SCRATCH MODE: all split values from the allowed list only
  [ ] Phase 3 Step 4 executed: every chart_header / table_header / workflow_header /
      comparison_header / initiative_header / risk_header / profile_header checked
      against slide title for near-repetition; differentiation applied where needed
  [ ] Every artifact has artifact_arrangement, artifact_coverage_hint,
      artifact_split_hint, and internal_alignment populated
  [ ] Every stat_bar / initiative_map / risk_register / comparison_table / profile_card_set
      has zone_constraint populated on its parent zone object
  [ ] ZW codes not used when slide_format width:height < 1.5

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
OUTPUT OBJECT — REQUIRED FIELDS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Each slide object must contain EXACTLY these top-level fields:

{
  "slide_number": number,
  "slide_type": "title" | "divider" | "content" | "thank_you",
  "narrative_role": "string — carry forward from Agent 3 plan; empty string for structural slides",
  "slide_archetype": "summary" | "trend" | "comparison" | "breakdown" | "driver_analysis" |
                     "process" | "recommendation" | "dashboard" | "proof" | "roadmap",
  "selected_layout_name": "string",
  "title": "string",
  "subtitle": "string",
  "key_message": "string",
  "zones": [ ... ],
  "speaker_note": "string"
}

SLIDE TYPE RULES

  Title slide:
  - title: short presentation name, 4–8 words
  - subtitle: audience / context / date if relevant
  - key_message: governing thought of the full deck
  - slide_archetype: "summary"
  - zones: []

  Divider slide:
  - title: section name only
  - subtitle: empty
  - key_message: one-line purpose of the section
  - slide_archetype: "summary"
  - zones: []

  Thank-you slide:
  - title: "Thank You" or equivalent closing phrase
  - subtitle: presenter name / contact if relevant
  - key_message: one sentence — what the audience must do next
  - slide_archetype: "summary"
  - zones: []

  Content slide:
  - title must be insight-led — never a generic topic label
    WRONG: "Revenue Analysis"   | RIGHT: "Premium mix drove most of the revenue uplift"
    WRONG: "Market Overview"    | RIGHT: "Market growing at 22% CAGR with untapped headroom"
    WRONG: "Geographic Risk"    | RIGHT: "North Zone concentration exceeds safe exposure threshold"
  - title HARD MAX: 10 words. Count every word — articles, prepositions, numbers each count as 1.
    If draft exceeds 10 words: cut modifiers, drop subsidiary clauses, keep the sharpest claim.
    Titles are scanned in 1–2 seconds on a board screen — one sharp assertion, not a paragraph.

ZONE OBJECT STRUCTURE

  Each zone must contain:
  {
    "zone_id": "z1",
    "artifacts": [ ... ],
    "zone_split": "<split value>",
    "artifact_arrangement": "horizontal" | "vertical" | null,
    "layout_hint": {
      "split": "<split value>",
      "artifact_arrangement": "horizontal" | "vertical" | null,
      "split_hint": [60, 40] | [50, 50] | null
    }
  }

ARTIFACT SCHEMAS

insight_text:
  {
    "type": "insight_text",
    "insight_header": "2–4 word specific label that names the implication — e.g. 'Risk Implication', 'Growth Opportunity', 'Action Required', 'Portfolio Risk' — never use generic labels like 'So What' or 'Key Insight'",
    "points": ["specific insight with data"],          ← STANDARD mode (flat list)
    "groups": [                                        ← GROUPED mode (thematic sections)
      { "header": "2–4 word label", "bullets": ["crisp point with data"] }
    ],
    "sentiment": "positive" | "warning" | "neutral"
  }
  Use either points[] OR groups[] — never both.

  STANDARD mode bullet rules:
  - Max points by slide area: < 25% → 4 pts; 25–50% → 6 pts; > 50% → 8 pts (HARD MAX)
  - Each bullet ≤ 12 words. 1–2 data points per bullet. Remainder goes to speaker notes.
  - Data-first phrasing: lead with the number or entity, not "This shows that…"
  - No compound sentences with dashes or semicolons — one idea per bullet.
  - If at the slide-area cap with findings remaining: move lower-priority findings to speaker notes.

  GROUPED mode bullet rules:
  - Grouped mode eligibility by slide area: < 30% → not allowed, revert to Standard; 30–50% → max 3 groups × 3 bullets each; > 50% → max 6 groups × 3 bullets each (HARD MAX)
  - Each bullet ≤ 12 words. 1–2 data points per bullet. Remainder to speaker notes.
  - Group headers: 2–4 words, no verbs, label the theme not the finding.

  Content integrity: preserve ALL facts, numbers, names, percentages exactly from source.
  Do NOT drop data-bearing facts. DO compress prose around the facts.

chart:
  {
    "type": "chart",
    "chart_type": "bar" | "line" | "pie" | "waterfall" | "clustered_bar" | "horizontal_bar" | "group_pie",
    "chart_decision": "one line: why this chart type was chosen",
    "chart_title": "",                                 ← leave empty when layout has header placeholder
    "chart_header": "the one-line insight the chart proves",
    "x_label": "string",
    "y_label": "string",
    "categories": ["string"],
    "series": [
      { "name": "string", "values": [number], "unit": "count|currency|percent|other",
        "types": ["positive"|"negative"|"total"] }
    ],
    "dual_axis": false,
    "secondary_series": [],
    "show_data_labels": true,
    "show_legend": true
  }
  chart_title is rendered INSIDE the plot area; chart_header is the zone heading.
  When a layout header placeholder exists, set chart_title: "" — never duplicate in both.

group_pie chart — use this schema instead when chart_type is "group_pie":
  {
    "type": "chart",
    "chart_type": "group_pie",
    "chart_decision": "one line: why group_pie was chosen over table or clustered_bar",
    "chart_title": "",
    "chart_header": "the one-line insight the group proves",
    "categories": ["Slice A", "Slice B", "Slice C"],  ← shared slice labels for ALL pies (max 7)
    "series": [                                        ← one entry per entity (pie); max 8 entries
      { "name": "Entity 1", "series_total": "₹39.7L", "values": [60, 25, 15], "unit": "percent" },
      { "name": "Entity 2", "series_total": "₹21.8L", "values": [70, 20, 10], "unit": "percent" },
      { "name": "Entity 3", "series_total": "₹2.0L",  "values": [50, 30, 20], "unit": "percent" }
    ],
    "show_legend": true,                               ← single shared legend above all pies
    "show_data_labels": true                           ← percentage labels on each slice
  }
  group_pie rules:
  - categories[] = the slice breakdown (same for all pies); max 7 entries
  - series[] = one entry per entity; series[i].name becomes the label BELOW pie i; max 8 entries
  - series[i].series_total (optional): pre-formatted absolute total displayed as a sub-label
    directly below the entity name under each pie — use when entities differ materially in
    absolute scale and the audience needs both composition and magnitude. Agent 4 must compute
    and format this value from source data (e.g. "₹39.7L", "23%", "$4.2M") — do not leave
    it blank or delegate calculation to Agent 5. Omit the field entirely if not meaningful.
  - Each series values[] must sum to ~100 (percentages) or represent a consistent unit
  - values[] length must equal categories[] length for every series
  - Do NOT use x_label, y_label, dual_axis, secondary_series for group_pie
  - Legend is always shared and rendered once above the group, below the chart_header

cards:
  {
    "type": "cards",
    "cards": [
      {
        "title": "string — metric/category label, max 4 words, no verbs",
        "subtitle": "string — primary number or %, max 8 characters",
        "body": "string — max 15 words, one crisp implication or supporting data point",
        "sentiment": "positive" | "negative" | "neutral"
      }
    ]
  }
  All cards in a zone must be parallel in structure.
  Card count vs zone: max 4 in full-width zones; max 2 in side zones.

workflow:
  {
    "type": "workflow",
    "workflow_type": "process_flow" | "hierarchy" | "decomposition" | "timeline",
    "flow_direction": "left_to_right" | "top_to_bottom" | "top_down_branching" | "bottom_up",
    "workflow_header": "string",
    "workflow_insight": "string",
    "nodes": [
      { "id": "n1", "label": "string", "value": "string", "description": "string", "level": 1 }
    ],
    "connections": [ { "from": "n1", "to": "n2", "type": "arrow" } ]
  }
  Node copy limits: label 2–5 words (hard max 18 chars); value 2–6 words; description 8–18 words.
  For left_to_right / timeline: value = short secondary above box; description = longer note below.
  For top_to_bottom / bottom_up: use only description; leave value empty.
  Max 6 nodes. Max 8 connections. No crossing connections.

table:
  {
    "type": "table",
    "table_header": "string — insight the table proves, not a topic label",
    "title": "string",
    "headers": ["string"],
    "rows": [["string"]],
    "highlight_rows": [0],
    "note": "string"
  }
  Include only columns that directly support message_objective.
  Preserve original data order from source unless sorting IS the insight.
  Max 6 rows. highlight_rows marks the single most important row.

matrix:
  {
    "type": "matrix",
    "matrix_type": "2x2",
    "matrix_header": "string",
    "x_axis": { "label": "string", "low_label": "string", "high_label": "string" },
    "y_axis": { "label": "string", "low_label": "string", "high_label": "string" },
    "quadrants": [ { "id": "q1", "name": "string", "insight": "string" } ],
    "points": [ { "label": "string", "x": "low|medium|high", "y": "low|medium|high" } ]
  }
  Max 6 points. Must define both axes and all 4 quadrants.

driver_tree:
  {
    "type": "driver_tree",
    "tree_header": "string",
    "root": { "label": "string", "value": "string" },
    "branches": [ { "label": "string", "value": "string", "children": [] } ]
  }
  Max 3 levels. Max 6–8 nodes. Root = outcome; children = drivers.
  NOT a process — do not use when showing sequence or steps.

prioritization:
  {
    "type": "prioritization",
    "priority_header": "string",
    "items": [
      {
        "rank": 1,
        "title": "string",
        "description": "string",
        "qualifiers": [ { "label": "string", "value": "string" } ]
      }
    ]
  }
  Max 5 items (HARD MAX). Must be sorted by importance, highest first.
  Field rules (HARD MAX):
  - title: strategic framing only — NO numbers, NO % values, NO currency amounts; ≤ 8 words.
    Verb-led or noun-phrase. Name the strategic action or theme, not the metric.
    WRONG: "Provision ₹17.66 L Punjab Dairy NPA"   RIGHT: "Provision Punjab Dairy Accounts"
    WRONG: "73% Delhi/NCR Outstanding Needs Review" RIGHT: "Review Delhi/NCR Concentration Risk"
  - description: data-driven backing for the title; ≤ 15 words. Include the key number(s).
    WRONG: "Implement monthly asset quality tracking for Delhi/NCR 1–3 year bucket (73% of outstanding)—establish early warning triggers for collection slowdown"
    RIGHT: "73% outstanding in 1–3yr bucket (₹18.48 Cr); establish monthly AQ triggers"
  - qualifiers: up to 2 per item; labels content-driven (not hardcoded).
    label: 1 word only (e.g. "Timeline", "Owner", "Impact", "Zone")
    value: ≤ 4 words (e.g. "Q2 2025", "Credit Team", "High")

stat_bar:
  {
    "type": "stat_bar",
    "stat_header": "string — the one-line insight the ranking proves",
    "annotation_style": "inline" | "trailing",
    "column_headers": {
      "label": "string",
      "metric": "string",
      "value": "string",
      "annotation": "string"
    },
    "rows": [
      {
        "id": "string — optional stable key",
        "label": "string — left-side row label",
        "value": number,
        "unit": "string — unit suffix for the displayed value",
        "display_value": "string — optional preformatted value label",
        "annotation": "string — right-side qualifier text",
        "annotation_representation": "text" | "pill",
        "bar_color": "string — optional override",
        "highlight": true | false
      }
    ]
  }
  annotation_style default: "trailing". Do NOT encode as raw table rows — emit semantic row
  objects a renderer can map to label / bar / value / annotation regions.
  highlight: set explicitly — never infer from rank.
  column_headers: optional; include only when column labels differ from defaults.

comparison_table:
  {
    "type": "comparison_table",
    "comparison_header": "string — the one-line insight the comparison proves",
    "criteria": [
      { "id": "string", "label": "string" }
    ],
    "options": [
      {
        "id": "string",
        "name": "string",
        "badge_text": "string — optional row badge e.g. 'recommended'",
        "cells": [
          {
            "criterion_id": "string — matches criteria[].id",
            "rating": "yes" | "partial" | "no" | "text",
            "display_value": "string — optional rendered symbol/text if not default",
            "note": "string — supporting note when rating is 'text'",
            "representation_type": "icon" | "text"
          }
        ]
      }
    ],
    "recommended_option_id": "string — id of the preferred option",
    "recommended_option": "string — name fallback only if id is unavailable"
  }
  recommended_option_id is preferred; recommended_option is a fallback only if id is unavailable.
  NEVER use plain table for option-vs-criteria data.

initiative_map:
  {
    "type": "initiative_map",
    "initiative_header": "string — the one-line framing of the initiative landscape",
    "dimension_labels": [
      { "id": "string", "label": "string" }
    ],
    "initiatives": [
      {
        "id": "string",
        "name": "string — track label shown at left",
        "subtitle": "string — optional track subtitle",
        "placements": [
          {
            "lane_id": "string — matches dimension_labels[].id",
            "title": "string",
            "subtitle": "string",
            "tags": ["string"],
            "footer": "string",
            "accent_tone": "primary" | "secondary" | "neutral"
          }
        ],
        "dimensions": [{ "label": "string", "value": "string" }]
      }
    ]
  }
  placements[] is preferred over dimensions[] (backward-compatible fallback only).
  NEVER use when rows have a rank order (use prioritization). NEVER use for process steps (use workflow).

profile_card_set:
  {
    "type": "profile_card_set",
    "profile_header": "string — the one-line framing of the entity set",
    "layout_direction": "horizontal" | "grid",
    "profiles": [
      {
        "id": "string",
        "entity_name": "string",
        "subtitle": "string — optional line directly below title",
        "badge_text": "string — optional KPI/tag pill at top-right",
        "secondary_items": [
          {
            "label": "string — left-side key",
            "value": "string or string[]",
            "representation_type": "text" | "chip_list" | "pill",
            "sentiment": "positive" | "negative" | "neutral" | "warning"
          }
        ],
        "attributes": [{ "key": "string", "value": "string", "sentiment": "string" }]
      }
    ]
  }
  layout_direction default: "horizontal".
  secondary_items[] is preferred over attributes[] (backward-compatible fallback only).
  NEVER use when metrics are structurally identical across entities (use cards instead).

risk_register:
  {
    "type": "risk_register",
    "risk_header": "string — the one-line framing of the risk landscape",
    "show_mitigation": false,
    "risks": [
      {
        "id": "string",
        "title": "string — short risk headline",
        "detail": "string — explanatory line under the title",
        "severity": "critical" | "high" | "medium" | "low",
        "owner": "string",
        "status": "string",
        "likelihood": 0 | 1 | 2 | 3,
        "impact": 0 | 1 | 2 | 3,
        "owner_tag": "string — optional display override for owner pill",
        "status_tag": "string — optional display override for status pill"
      }
    ]
  }
  show_mitigation: legacy fallback only — avoid if detail already includes mitigation.
  likelihood / impact: pip count 0–3; omit fields entirely if not relevant to the slide's message.
  NEVER use plain table when severity-by-row is the primary signal.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
PRE-OUTPUT QUALITY GATES
Run ALL gates before emitting JSON. Fix any failure before proceeding.
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

GATE 1 — CONTENT INTEGRITY
  [ ] No placeholder text anywhere
  [ ] No invented numbers — every figure sourced from the source document
  [ ] No vague wording — every point specific and actionable
  [ ] All content slides have insight-led titles
  [ ] Every slide title ≤ 10 words — count every word; rewrite if over
  [ ] Every insight_text standard: bullet count within zone-area cap (< 25% → ≤4; 25–50% → ≤6; > 50% → ≤8)
  [ ] Every insight_text standard: each bullet ≤ 12 words — rewrite if over; excess detail to speaker notes
  [ ] Every insight_text grouped: group × bullet count within zone-area cap (< 25% → 2×2; 25–50% → 3×3; > 50% → 5×3)
  [ ] Every insight_text grouped: each bullet ≤ 12 words — rewrite if over; excess detail to speaker notes
  [ ] Every prioritization: ≤ 5 items; title ≤ 8 words and NO numbers/% /currency; description ≤ 15 words
  [ ] Every prioritization qualifier: label = 1 word; value ≤ 4 words

GATE 2 — STRUCTURAL LIMITS
  [ ] Max 4 zones per slide
  [ ] Max 2 artifacts per zone
  [ ] At least 1 primary zone per content slide
  [ ] No more than 2 primary zones per slide

GATE 3 — ARTIFACT HARD CONSTRAINTS
  [ ] Every bar/line/clustered_bar/horizontal_bar has ≥ 3 categories
  [ ] Every clustered_bar has exactly 2 series with matching units
  [ ] Every pie/donut has ≤ 5 segments
  [ ] dual_axis set to true wherever series have different units
  [ ] Every workflow has coherent, non-crossing nodes and connections
  [ ] Every artifact has its header field populated (except cards)
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
  [ ] Every left_to_right / timeline workflow spans FULL HORIZONTAL WIDTH
  [ ] No zone placed beside (left/right of) a left_to_right workflow
  [ ] Every top_to_bottom / top_down_branching workflow spans FULL VERTICAL HEIGHT
  [ ] No zone stacked above or below a top_to_bottom workflow
  [ ] Any workflow with > 3 nodes uses full-width (left_to_right) or full-height (top_to_bottom)

GATE 4 — ZONE SPATIAL COVERAGE
  [ ] All zone splits sum to 100% of content area — no gaps, no overlaps
  [ ] Artifact with MIN_W ≥ 70% is in a zone covering ≥ 70% slide width
  [ ] Artifact with MIN_H ≥ 65% is in a zone covering ≥ 65% slide height
  [ ] PRIMARY zone occupies ≥ 60% of the split axis
  [ ] SECONDARY zone occupies ≤ 40% of the split axis
  [ ] 50/50 splits used only where both zones are narrative_weight = "primary"

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



// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE ARCHETYPE → DEFAULT ZONE STRUCTURE
// ═══════════════════════════════════════════════════════════════════════════════

function defaultZonesForArchetype(archetype) {
  switch (archetype) {

    case 'dashboard':
      return [
        { zone_id: 'z1', zone_role: 'summary', narrative_weight: 'primary',
          message_objective: 'Headline metrics at a glance',
          layout_hint: { split: 'top_30' },
          artifacts: [{ type: 'cards', cards: [] }] },
        { zone_id: 'z2', zone_role: 'supporting_evidence', narrative_weight: 'secondary',
          message_objective: 'Supporting detail',
          layout_hint: { split: 'bottom_70' },
          artifacts: [{ type: 'chart', chart_type: 'bar', categories: [], series: [] }] }
      ]

    case 'trend':
      return [
        { zone_id: 'z1', zone_role: 'primary_proof', narrative_weight: 'primary',
          message_objective: 'Show trend over time',
          layout_hint: { split: 'left_60' },
          artifacts: [{ type: 'chart', chart_type: 'line', categories: [], series: [] }] },
        { zone_id: 'z2', zone_role: 'implication', narrative_weight: 'secondary',
          message_objective: 'Interpret the trend',
          layout_hint: { split: 'right_40' },
          artifacts: [{ type: 'insight_text', insight_header: 'Trend Implication', points: [], sentiment: 'neutral' }] }
      ]

    case 'comparison':
      return [
        { zone_id: 'z1', zone_role: 'primary_proof', narrative_weight: 'primary',
          message_objective: 'Visual comparison across categories',
          layout_hint: { split: 'left_60' },
          artifacts: [{ type: 'chart', chart_type: 'bar', categories: [], series: [] }] },
        { zone_id: 'z2', zone_role: 'implication', narrative_weight: 'secondary',
          message_objective: 'What the comparison means',
          layout_hint: { split: 'right_40' },
          artifacts: [{ type: 'insight_text', insight_header: 'Key Insight', points: [], sentiment: 'neutral' }] }
      ]

    case 'breakdown':
      return [
        { zone_id: 'z1', zone_role: 'primary_proof', narrative_weight: 'primary',
          message_objective: 'Show composition or segmentation',
          layout_hint: { split: 'left_50' },
          artifacts: [{ type: 'chart', chart_type: 'pie', categories: [], series: [] }] },
        { zone_id: 'z2', zone_role: 'supporting_evidence', narrative_weight: 'secondary',
          message_objective: 'Detail behind the segments',
          layout_hint: { split: 'right_50' },
          artifacts: [{ type: 'insight_text', insight_header: 'Key Insight', points: [], sentiment: 'neutral' }] }
      ]

    case 'driver_analysis':
      return [
        { zone_id: 'z1', zone_role: 'primary_proof', narrative_weight: 'primary',
          message_objective: 'Show movement from baseline to result',
          layout_hint: { split: 'left_60' },
          artifacts: [{ type: 'chart', chart_type: 'waterfall', categories: [], series: [] }] },
        { zone_id: 'z2', zone_role: 'implication', narrative_weight: 'secondary',
          message_objective: 'Interpret the key drivers',
          layout_hint: { split: 'right_40' },
          artifacts: [{ type: 'insight_text', insight_header: 'Key Insight', points: [], sentiment: 'neutral' }] }
      ]

    case 'process':
    case 'roadmap':
      return [
        { zone_id: 'z1', zone_role: 'process', narrative_weight: 'primary',
          message_objective: 'Show process or roadmap structure',
          layout_hint: { split: 'top_60' },
          artifacts: [{ type: 'workflow', workflow_type: 'process_flow', flow_direction: 'left_to_right', nodes: [], connections: [] }] },
        { zone_id: 'z2', zone_role: 'implication', narrative_weight: 'secondary',
          message_objective: 'Key insight or action from the process',
          layout_hint: { split: 'bottom_40' },
          artifacts: [{ type: 'insight_text', insight_header: 'Process Implication', points: [], sentiment: 'neutral' }] }
      ]

    case 'recommendation':
      return [
        { zone_id: 'z1', zone_role: 'recommendation', narrative_weight: 'primary',
          message_objective: 'Recommended actions or priorities',
          layout_hint: { split: 'full' },
          artifacts: [{ type: 'cards', cards: [] }] }
      ]

    case 'proof':
      return [
        { zone_id: 'z1', zone_role: 'primary_proof', narrative_weight: 'primary',
          message_objective: 'Evidence supporting the claim',
          layout_hint: { split: 'left_60' },
          artifacts: [{ type: 'chart', chart_type: 'bar', categories: [], series: [] }] },
        { zone_id: 'z2', zone_role: 'implication', narrative_weight: 'secondary',
          message_objective: 'Interpretation and implication of the evidence',
          layout_hint: { split: 'right_40' },
          artifacts: [{ type: 'insight_text', insight_header: 'Risk Implication', points: [], sentiment: 'neutral' }] }
      ]

    default: // summary
      return [
        { zone_id: 'z1', zone_role: 'summary', narrative_weight: 'primary',
          message_objective: 'Key summary of the section',
          layout_hint: { split: 'full' },
          artifacts: [{ type: 'insight_text', insight_header: 'Key Insight', points: [], sentiment: 'neutral' }] }
      ]
  }
}

function inferArchetype(sectionType, slideIndex) {
  switch (sectionType) {
    case 'financial_data':     return slideIndex === 0 ? 'dashboard' : 'comparison'
    case 'executive_summary':  return 'dashboard'
    case 'strategic_analysis': return slideIndex % 2 === 0 ? 'comparison' : 'breakdown'
    case 'market_analysis':    return 'comparison'
    case 'recommendations':    return 'recommendation'
    case 'conclusion':         return 'roadmap'
    case 'operational_review': return 'dashboard'
    default:                   return 'summary'
  }
}

function inferArchetypeFromZones(slide) {
  const zones = slide?.zones || []
  const artifacts = zones.flatMap(z => z.artifacts || []).map(a => String(a.type || '').toLowerCase())
  const zoneCount = zones.length
  const has = t => artifacts.includes(t)
  const hasOnlyInsight = artifacts.length > 0 && artifacts.every(t => t === 'insight_text')
  const hasCards = has('cards')
  const hasWorkflow = has('workflow')
  const hasChart = has('chart')
  const hasTable = has('table')
  const hasReasoning = artifacts.some(t => ['matrix', 'driver_tree', 'prioritization'].includes(t))

  if (slide?.slide_type === 'title' || slide?.slide_type === 'divider') return 'summary'
  if (hasReasoning) {
    if (has('prioritization')) return 'recommendation'
    if (has('driver_tree')) return 'driver_analysis'
    if (has('matrix')) return 'comparison'
  }
  if (hasWorkflow) {
    const workflowTypes = zones.flatMap(z => z.artifacts || []).filter(a => String(a.type || '').toLowerCase() === 'workflow')
    const hasRoadmapLike = workflowTypes.some(a => /timeline|roadmap/i.test(String(a.workflow_type || '')) || /timeline/i.test(String(a.flow_direction || '')))
    return hasRoadmapLike ? 'roadmap' : 'process'
  }
  if (hasChart && hasCards && zoneCount >= 3) return 'dashboard'
  if (hasCards && !hasChart && !hasTable && !hasWorkflow && !hasReasoning) return 'summary'
  if (hasChart && zoneCount >= 3) return 'dashboard'
  if (hasChart && zoneCount === 2) return 'comparison'
  if (hasTable) return 'proof'
  if (hasOnlyInsight) return 'summary'
  return slide?.slide_archetype || 'proof'
}

function compactList(arr, limit = 6, maxChars = 280) {
  const items = (arr || []).filter(Boolean).slice(0, limit).map(v => String(v).trim())
  const joined = items.join(' | ')
  return joined.length > maxChars ? joined.slice(0, maxChars) + '…' : joined
}

function buildBriefSummaryForAgent4(brief) {
  const b = brief || {}
  return {
    governing_thought: b.governing_thought || '',
    audience:          b.audience          || '',
    narrative_flow:    b.narrative_flow    || '',
    data_heavy:        b.data_heavy        || false,
    tone:              b.tone              || 'professional',
    key_messages:      Array.isArray(b.key_messages) ? b.key_messages : []
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// VALIDATION
// ═══════════════════════════════════════════════════════════════════════════════

function validateArtifact(artifact) {
  const t = (artifact.type || '').toLowerCase()

  if (t === 'chart') {
    const cats = artifact.categories || []
    const series = artifact.series || []
    const chartType = (artifact.chart_type || '').toLowerCase()

    // group_pie has its own validation rules
    if (chartType === 'group_pie') {
      if (cats.length < 2) return { valid: false, reason: 'group_pie needs 2+ categories (slices), got ' + cats.length }
      if (cats.length > 7) return { valid: false, reason: 'group_pie exceeds HARD MAX 7 slices, got ' + cats.length + ' — convert to clustered_bar' }
      if (series.length < 2) return { valid: false, reason: 'group_pie needs 2+ series (entities/pies), got ' + series.length + ' — convert to single pie' }
      if (series.length > 8) return { valid: false, reason: 'group_pie exceeds HARD MAX 8 entities, got ' + series.length + ' — convert to table or clustered_bar' }
      for (const s of series) {
        if ((s.values || []).length !== cats.length) {
          return { valid: false, reason: 'group_pie series "' + s.name + '" values count mismatch: ' + (s.values||[]).length + ' vs ' + cats.length }
        }
        if ((s.values || []).every(v => v === 0)) {
          return { valid: false, reason: 'group_pie series "' + s.name + '" has all-zero values' }
        }
      }
      return { valid: true }
    }

    // Need 2+ categories (3+ ideally but 2 is minimum after normalisation)
    if (cats.length < 2) return { valid: false, reason: 'chart needs 2+ categories, got ' + cats.length }

    if (!series.length) return { valid: false, reason: 'chart has no series' }

    // Values must match categories
    for (const s of series) {
      if ((s.values || []).length !== cats.length) {
        return { valid: false, reason: 'series values count mismatch: ' + (s.values||[]).length + ' vs ' + cats.length }
      }
    }

    // All-zero series
    for (const s of series) {
      if ((s.values || []).every(v => v === 0)) {
        return { valid: false, reason: 'chart series has all-zero values' }
      }
    }

    // clustered_bar needs 2 series
    if (artifact.chart_type === 'clustered_bar' && series.length < 2) {
      artifact.chart_type = 'bar' // auto-fix
    }

    return { valid: true }
  }

  if (t === 'stat_bar') {
    const rows = artifact.rows || []
    if (rows.length < 2) return { valid: false, reason: 'stat_bar needs 2+ rows' }
    if (!artifact.stat_header || String(artifact.stat_header).trim().length < 3) return { valid: false, reason: 'stat_bar missing stat_header' }
    if (!artifact.annotation_style) return { valid: false, reason: 'stat_bar missing annotation_style' }
    for (const row of rows) {
      if (!String(row?.label || '').trim()) return { valid: false, reason: 'stat_bar row missing label' }
      if (!Number.isFinite(+row?.value)) return { valid: false, reason: 'stat_bar row missing numeric value' }
    }
    if (rows.every(r => (+r.value || 0) === 0)) return { valid: false, reason: 'stat_bar has all-zero row values' }
    return { valid: true }
  }

  if (t === 'insight_text') {
    const groups = artifact.groups || []
    const points = artifact.points || []
    if (groups.length > 0) {
      // Grouped mode: valid if at least one group has bullets
      const hasBullets = groups.some(g => (g.bullets || []).length > 0)
      if (!hasBullets) return { valid: false, reason: 'insight_text grouped has no bullets in any group' }
      return { valid: true }
    }
    if (!points.length) return { valid: false, reason: 'insight_text has no points' }
    if (points.every(p => !p || p.trim().length < 5)) return { valid: false, reason: 'insight_text has only placeholder points' }
    return { valid: true }
  }

  if (t === 'cards') {
    const cards = artifact.cards || []
    if (!cards.length) return { valid: false, reason: 'cards has no items' }
    return { valid: true }
  }

  if (t === 'workflow') {
    const nodes = artifact.nodes || []
    if (!nodes.length) return { valid: false, reason: 'workflow has no nodes' }
    if (String(artifact.workflow_type || '').toLowerCase() === 'information_flow') return { valid: false, reason: 'information_flow is not allowed' }
    return { valid: true }
  }

  if (t === 'table') {
    if (!(artifact.headers || []).length) return { valid: false, reason: 'table has no headers' }
    if (!(artifact.rows || []).length) return { valid: false, reason: 'table has no rows' }
    return { valid: true }
  }

  if (t === 'matrix') {
    if (!artifact.x_axis?.label || !artifact.y_axis?.label) return { valid: false, reason: 'matrix missing axis labels' }
    if ((artifact.quadrants || []).length !== 4) return { valid: false, reason: 'matrix must define 4 quadrants' }
    if (!(artifact.points || []).length) return { valid: false, reason: 'matrix has no points' }
    return { valid: true }
  }

  if (t === 'driver_tree') {
    if (!artifact.root?.label) return { valid: false, reason: 'driver_tree missing root label' }
    if (!(artifact.branches || []).length) return { valid: false, reason: 'driver_tree has no branches' }
    return { valid: true }
  }

  if (t === 'prioritization') {
    if (!(artifact.items || []).length) return { valid: false, reason: 'prioritization has no items' }
    if ((artifact.items || []).some(i => i.rank == null || !i.title)) return { valid: false, reason: 'prioritization items missing rank/title' }
    return { valid: true }
  }

  if (t === 'comparison_table') {
    if (!(artifact.criteria || []).length) return { valid: false, reason: 'comparison_table has no criteria' }
    if (!(artifact.options || []).length) return { valid: false, reason: 'comparison_table has no options' }
    if ((artifact.options || []).some(o => !(o.cells || []).length)) return { valid: false, reason: 'comparison_table option missing cells' }
    return { valid: true }
  }

  if (t === 'initiative_map') {
    if (!(artifact.dimension_labels || []).length) return { valid: false, reason: 'initiative_map has no dimension_labels' }
    if (!(artifact.initiatives || []).length) return { valid: false, reason: 'initiative_map has no initiatives' }
    return { valid: true }
  }

  if (t === 'risk_register') {
    if (!(artifact.risks || []).length) return { valid: false, reason: 'risk_register has no risks' }
    if ((artifact.risks || []).some(r => !r.title || !r.severity)) return { valid: false, reason: 'risk_register risk missing title or severity' }
    return { valid: true }
  }

  if (t === 'profile_card_set') {
    if (!(artifact.profiles || []).length) return { valid: false, reason: 'profile_card_set has no profiles' }
    if ((artifact.profiles || []).some(p => !p.entity_name)) return { valid: false, reason: 'profile_card_set profile missing entity_name' }
    return { valid: true }
  }

  return { valid: true }
}

function validateReasoningArtifactUsage(slide) {
  const reasoningTypes = new Set(['matrix', 'driver_tree', 'prioritization'])
  const zones = slide.zones || []
  for (const zone of zones) {
    const arts = zone.artifacts || []
    const reasoningArts = arts.filter(a => reasoningTypes.has((a.type || '').toLowerCase()))
    if (!reasoningArts.length) continue
    if ((zone.narrative_weight || 'primary') !== 'primary') return false
    if (reasoningArts.length > 1) return false
    if (arts.some(a => !reasoningTypes.has((a.type || '').toLowerCase()) && (a.type || '').toLowerCase() !== 'insight_text')) return false
  }
  return true
}

function validateWorkflowUsage(slide) {
  const zones = slide?.zones || []
  for (const zone of zones) {
    for (const artifact of (zone.artifacts || [])) {
      if (artifactType(artifact) === 'workflow' && !validateWorkflowArtifactAndZone(artifact, zone)) return false
    }
  }
  return true
}

function artifactType(artifact) {
  return String((artifact || {}).type || '').toLowerCase()
}

const ZONE_STRUCTURE_LIBRARY = {
  ZS01_single_full: {
    zoneCount: 1,
    slots: ['primary_full'],
    scratchSplits: ['full']
  },
  ZS02_stacked_equal: {
    zoneCount: 2,
    slots: ['top_primary', 'bottom_secondary'],
    scratchSplits: ['top_50', 'bottom_50']
  },
  ZS03_side_by_side_equal: {
    zoneCount: 2,
    slots: ['left_primary', 'right_secondary'],
    scratchSplits: ['left_50', 'right_50']
  },
  ZS04_left_dominant_right_stack: {
    zoneCount: 3,
    slots: ['left_dominant', 'right_top_support', 'right_bottom_support'],
    scratchSplits: ['left_60', 'top_right_50', 'bottom_full']
  },
  ZS05_right_dominant_left_stack: {
    zoneCount: 3,
    slots: ['left_top_support', 'left_bottom_support', 'right_dominant'],
    scratchSplits: ['top_left_50', 'bottom_full', 'right_60']
  },
  ZS06_top_full_bottom_two: {
    zoneCount: 3,
    slots: ['top_anchor', 'bottom_left_support', 'bottom_right_support'],
    scratchSplits: ['top_35', 'left_50', 'right_50']
  },
  ZS07_top_two_bottom_dominant: {
    zoneCount: 3,
    slots: ['top_left_support', 'top_right_support', 'bottom_dominant'],
    scratchSplits: ['top_left_50', 'top_right_50', 'bottom_60']
  },
  ZS08_quad_grid: {
    zoneCount: 4,
    slots: ['top_left', 'top_right', 'bottom_left', 'bottom_right'],
    scratchSplits: ['tl', 'tr', 'bl', 'br']
  },
  ZS09_left_dominant_right_triptych: {
    zoneCount: 4,
    slots: ['left_dominant', 'right_top_support', 'right_mid_support', 'right_bottom_support'],
    scratchSplits: ['left_60', 'top_right_50', 'tr', 'br']
  },
  ZS10_top_full_bottom_three: {
    zoneCount: 4,
    slots: ['top_anchor', 'bottom_left_support', 'bottom_mid_support', 'bottom_right_support'],
    scratchSplits: ['top_35', 'bl', 'bottom_full', 'br']
  },
  ZS11_three_rows_equal: {
    zoneCount: 3,
    slots: ['row_1', 'row_2', 'row_3'],
    scratchSplits: ['top_33', 'top_33', 'bottom_34']
  },
  ZW01_three_columns_equal: {
    zoneCount: 3,
    canvasFamily: 'wide',
    slots: ['col_1', 'col_2', 'col_3'],
    scratchSplits: ['left_33', 'left_33', 'right_34']
  },
  ZW02_three_columns_right_stack: {
    zoneCount: 4,
    canvasFamily: 'wide',
    slots: ['left_col', 'mid_col', 'right_top', 'right_bottom'],
    scratchSplits: ['left_33', 'left_33', 'top_right_50', 'br']
  },
  ZW03_three_columns_left_stack: {
    zoneCount: 4,
    canvasFamily: 'wide',
    slots: ['left_top', 'left_bottom', 'mid_col', 'right_col'],
    scratchSplits: ['top_left_50', 'bl', 'left_33', 'right_34']
  },
  ZW04_four_columns_equal: {
    zoneCount: 4,
    canvasFamily: 'wide',
    slots: ['col_1', 'col_2', 'col_3', 'col_4'],
    scratchSplits: ['tl', 'tr', 'bl', 'br']
  }
}

function zoneStructureDef(id) {
  return ZONE_STRUCTURE_LIBRARY[id] || null
}

function artifactCardCount(artifact) {
  return artifactType(artifact) === 'cards' ? ((artifact.cards || []).length || 0) : 0
}

function artifactWorkflowNodeCount(artifact) {
  return artifactType(artifact) === 'workflow' ? ((artifact.nodes || []).length || 0) : 0
}

function artifactWorkflowLevelCount(artifact) {
  if (artifactType(artifact) !== 'workflow') return 0
  const levels = new Set((artifact.nodes || []).map(n => Number.isFinite(+n?.level) ? +n.level : null).filter(v => v != null))
  return levels.size
}

function artifactDriverTreeNodeCount(artifact) {
  if (artifactType(artifact) !== 'driver_tree') return 0
  const branches = artifact.branches || []
  return 1 + branches.length + branches.reduce((sum, b) => sum + (((b && b.children) || []).length || 0), 0)
}

function artifactTableShape(artifact) {
  return artifactType(artifact) === 'table'
    ? { cols: ((artifact.headers || []).length || 0), rows: ((artifact.rows || []).length || 0) }
    : { cols: 0, rows: 0 }
}

function artifactPrioritizationCount(artifact) {
  return artifactType(artifact) === 'prioritization' ? ((artifact.items || []).length || 0) : 0
}

function artifactInsightGroupCount(artifact) {
  return artifactType(artifact) === 'insight_text' ? ((artifact.groups || []).length || 0) : 0
}

function isReasoningArtifact(artifact) {
  return ['matrix', 'driver_tree', 'prioritization'].includes(artifactType(artifact))
}

function isSparseCardsArtifact(artifact) {
  return artifactType(artifact) === 'cards' && artifactCardCount(artifact) > 0 && artifactCardCount(artifact) <= 2
}

function isDenseSoloArtifact(artifact) {
  const t = artifactType(artifact)
  if (t === 'cards') return artifactCardCount(artifact) >= 8
  if (t === 'matrix') return true
  if (t === 'driver_tree') return artifactDriverTreeNodeCount(artifact) >= 6
  if (t === 'prioritization') return artifactPrioritizationCount(artifact) >= 5
  if (t === 'insight_text') return artifactInsightGroupCount(artifact) >= 3
  if (t === 'table') {
    const shape = artifactTableShape(artifact)
    return shape.cols >= 5 && shape.rows >= 10
  }
  if (t === 'workflow') return artifactWorkflowNodeCount(artifact) >= 6
  return false
}

function isSubstantialArtifact(artifact) {
  const t = artifactType(artifact)
  if (t === 'insight_text') {
    return ((artifact.points || []).length || 0) >= 3 || artifactInsightGroupCount(artifact) >= 2
  }
  if (t === 'workflow') return artifactWorkflowNodeCount(artifact) >= 4
  if (t === 'table') {
    const shape = artifactTableShape(artifact)
    return shape.cols >= 4 || shape.rows >= 6
  }
  if (t === 'chart') return ((artifact.categories || []).length || 0) >= 4
  if (t === 'cards') return artifactCardCount(artifact) >= 4
  if (t === 'prioritization') return artifactPrioritizationCount(artifact) >= 4
  if (t === 'driver_tree') return artifactDriverTreeNodeCount(artifact) >= 5
  if (t === 'matrix') return true
  return false
}

function zoneStructureArtifactProfile(zone) {
  const arts = zone?.artifacts || []
  return {
    count: arts.length,
    types: arts.map(artifactType),
    hasInsight: arts.some(a => artifactType(a) === 'insight_text'),
    hasReasoning: arts.some(isReasoningArtifact),
    hasWorkflow: arts.some(a => artifactType(a) === 'workflow'),
    hasChart: arts.some(a => artifactType(a) === 'chart'),
    hasTable: arts.some(a => artifactType(a) === 'table'),
    hasCards: arts.some(a => artifactType(a) === 'cards'),
    denseSolo: arts.length === 1 && isDenseSoloArtifact(arts[0]),
    compactCards: arts.some(isSparseCardsArtifact),
    groupedInsightOnly: arts.length === 1 && artifactType(arts[0]) === 'insight_text' && artifactInsightGroupCount(arts[0]) >= 3
  }
}

function isDominantSlot(slotName) {
  return /dominant|primary|anchor|full/.test(String(slotName || '').toLowerCase())
}

function isSupportSlot(slotName) {
  return /support|secondary|top_|bottom_|left_|right_|row_|col_/.test(String(slotName || '').toLowerCase()) && !isDominantSlot(slotName)
}

function validateZoneForStructureSlot(zone, slotName, structureId) {
  const arts = zone?.artifacts || []
  const profile = zoneStructureArtifactProfile(zone)
  if (!arts.length) return false

  if (structureId === 'ZS01_single_full') {
    if (profile.count === 1) {
      return profile.denseSolo || profile.groupedInsightOnly
    }
    if (profile.count === 2) {
      const key = zoneArtifactPairKey(zone)
      return new Set([
        'chart+insight_text',
        'driver_tree+insight_text',
        'insight_text+matrix',
        'insight_text+prioritization',
        'insight_text+workflow',
        'insight_text+table'
      ]).has(key)
    }
    return false
  }

  if (structureId === 'ZS08_quad_grid' || structureId === 'ZW04_four_columns_equal') {
    if (profile.count !== 1) return false
    if (profile.hasReasoning || profile.hasTable) return false
    return true
  }

  if (profile.hasReasoning) {
    if (!isDominantSlot(slotName)) return false
    if (profile.count > 2) return false
    return arts.every(a => isReasoningArtifact(a) || artifactType(a) === 'insight_text')
  }

  if (profile.hasCards && profile.compactCards && isDominantSlot(slotName)) return false
  if (profile.hasTable && isSupportSlot(slotName) && profile.count > 1) return false
  if (profile.hasWorkflow && isSupportSlot(slotName)) {
    const wf = arts.find(a => artifactType(a) === 'workflow')
    const wfType = String(wf?.workflow_type || '').toLowerCase()
    if (['process_flow', 'timeline'].includes(wfType)) return false
  }

  return true
}

function inferZoneStructure(slide) {
  const zones = slide?.zones || []
  const count = zones.length
  const allArtifacts = zones.flatMap(z => z.artifacts || [])
  const hasReasoning = allArtifacts.some(isReasoningArtifact)
  const hasCompactCards = zones.some(isCompactCardsZone)
  const hasWideWorkflow = zones.some(z => (z.artifacts || []).some(a => {
    const t = artifactType(a)
    const wfType = String(a?.workflow_type || '').toLowerCase()
    return t === 'workflow' && ['process_flow', 'timeline'].includes(wfType)
  }))
  const allSolo = zones.every(z => (z.artifacts || []).length === 1)

  if (count <= 1) return 'ZS01_single_full'
  if (count === 2) {
    if (hasReasoning || hasWideWorkflow) return 'ZS02_stacked_equal'
    return 'ZS03_side_by_side_equal'
  }
  if (count === 3) {
    if (hasCompactCards) return 'ZS06_top_full_bottom_two'
    if (hasWideWorkflow) return 'ZS07_top_two_bottom_dominant'
    return 'ZS04_left_dominant_right_stack'
  }
  if (count === 4) {
    if (allSolo) return 'ZS08_quad_grid'
    return 'ZS09_left_dominant_right_triptych'
  }
  return 'ZS08_quad_grid'
}

function applyZoneStructureMetadata(slide) {
  if (!slide || slide.slide_type !== 'content') return slide
  const structureId = slide.zone_structure || inferZoneStructure(slide)
  const def = zoneStructureDef(structureId) || zoneStructureDef(inferZoneStructure(slide))
  const zones = (slide.zones || []).slice(0, def?.zoneCount || 4).map((zone, idx) => ({
    ...zone,
    zone_slot: zone.zone_slot || def?.slots?.[idx] || `slot_${idx + 1}`
  }))
  return {
    ...slide,
    zone_structure: def ? structureId : inferZoneStructure(slide),
    zones
  }
}

function validateZoneStructureRules(slide) {
  if (!slide || slide.slide_type !== 'content') return true
  const structureId = slide.zone_structure || inferZoneStructure(slide)
  const def = zoneStructureDef(structureId)
  const zones = slide.zones || []
  if (!def) return false
  if (zones.length !== def.zoneCount) return false
  return zones.every((zone, idx) => validateZoneForStructureSlot(zone, zone.zone_slot || def.slots[idx], structureId))
}

function parseZoneSplitForConstraints(zone) {
  const split = String(zone?.zone_split || zone?.layout_hint?.split || 'full').toLowerCase()
  if (split === 'full' || split === 'bottom_full') return { fullWidth: true, fullHeight: split === 'full', widthFrac: 1, heightFrac: split === 'full' ? 1 : 0.5 }
  if (split === 'top_full') return { fullWidth: true, fullHeight: false, widthFrac: 1, heightFrac: 0.5 }
  const m = split.match(/^(left|right|top|bottom)_(\d{1,3})$/)
  if (m) {
    const side = m[1]
    const pct = Math.max(1, Math.min(99, parseInt(m[2], 10) || 0)) / 100
    if (side === 'left' || side === 'right') return { fullWidth: false, fullHeight: true, widthFrac: pct, heightFrac: 1 }
    return { fullWidth: true, fullHeight: false, widthFrac: 1, heightFrac: pct }
  }
  if (['top_left_50', 'top_right_50', 'tl', 'tr', 'bl', 'br'].includes(split)) return { fullWidth: false, fullHeight: false, widthFrac: 0.5, heightFrac: 0.5 }
  return { fullWidth: false, fullHeight: false, widthFrac: 0.5, heightFrac: 0.5 }
}

function validateWorkflowArtifactAndZone(workflowArtifact, zone) {
  if (artifactType(workflowArtifact) !== 'workflow') return true
  const workflowType = String(workflowArtifact.workflow_type || '').toLowerCase()
  const flow = String(workflowArtifact.flow_direction || '').toLowerCase()
  const nodeCount = artifactWorkflowNodeCount(workflowArtifact)
  const levelCount = artifactWorkflowLevelCount(workflowArtifact)
  const zoneShape = parseZoneSplitForConstraints(zone)

  if (workflowType === 'information_flow') return false

  if (workflowType === 'process_flow') {
    if (flow !== 'left_to_right') return false
    if (nodeCount < 4) return false
    if (!zoneShape.fullWidth) return false
    if (zoneShape.heightFrac < 0.5) return false
  }

  if (workflowType === 'hierarchy') {
    if (flow !== 'top_down_branching') return false
    if (levelCount < 3) return false
    if (zoneShape.widthFrac < 0.5) return false
    if (!zoneShape.fullHeight) return false
  }

  if (workflowType === 'decomposition') {
    if (!['left_to_right', 'top_to_bottom', 'top_down_branching'].includes(flow)) return false
    if (nodeCount > 3 && flow === 'left_to_right' && !zoneShape.fullWidth) return false
    if (nodeCount > 3 && ['top_to_bottom', 'top_down_branching'].includes(flow) && !zoneShape.fullHeight) return false
  }

  if (workflowType === 'timeline') {
    if (flow !== 'left_to_right') return false
    if (nodeCount < 4) return false
    if (!zoneShape.fullWidth) return false
    if (zoneShape.heightFrac < 0.5) return false
  }

  return true
}

function zoneArtifactPairKey(zone) {
  return ((zone?.artifacts || []).map(artifactType).sort().join('+'))
}

function slideHasDominantReasoningArtifact(slide) {
  const zones = slide?.zones || []
  const reasoningZones = zones.filter(z => (z.artifacts || []).some(isReasoningArtifact))
  if (!reasoningZones.length) return false
  return reasoningZones.some(z => (z.narrative_weight || '').toLowerCase() === 'primary')
}

function validateStructuralPatternRules(slide) {
  if (!slide || slide.slide_type !== 'content') return true
  const zones = slide.zones || []
  const zoneCount = zones.length
  const counts = zones.map(z => (z.artifacts || []).length)
  const totalArtifacts = counts.reduce((sum, n) => sum + n, 0)
  const allArtifacts = zones.flatMap(z => z.artifacts || [])
  const reasoningCount = allArtifacts.filter(isReasoningArtifact).length

  if (zoneCount === 1 && totalArtifacts === 1) {
    return isDenseSoloArtifact(allArtifacts[0])
  }

  if (zoneCount === 1 && totalArtifacts === 2) {
    const arts = zones[0].artifacts || []
    const types = arts.map(artifactType)
    const hasInsight = types.includes('insight_text')
    const pair = types.slice().sort().join('+')
    if (pair === 'cards+insight_text') return artifactCardCount(arts.find(a => artifactType(a) === 'cards')) >= 4
    if (['chart+insight_text', 'workflow+insight_text', 'table+insight_text', 'matrix+insight_text', 'driver_tree+insight_text', 'insight_text+prioritization'].includes(pair)) {
      const nonInsight = arts.find(a => artifactType(a) !== 'insight_text')
      return hasInsight && isSubstantialArtifact(nonInsight)
    }
    if (pair === 'cards+workflow') {
      const cards = arts.find(a => artifactType(a) === 'cards')
      const workflow = arts.find(a => artifactType(a) === 'workflow')
      return artifactCardCount(cards) > 0 && artifactCardCount(cards) <= 2 && artifactWorkflowNodeCount(workflow) >= 4
    }
    if (pair === 'chart+table') return isSubstantialArtifact(arts[0]) || isSubstantialArtifact(arts[1])
    return false
  }

  if (zoneCount === 2 && totalArtifacts === 2) {
    if (!counts.every(n => n === 1)) return false
    if (reasoningCount > 1) return false
    const sparsePrimaryCards = zones.some(z => (z.narrative_weight || '').toLowerCase() === 'primary' && (z.artifacts || []).some(isSparseCardsArtifact))
    return !sparsePrimaryCards
  }

  if (zoneCount === 2 && totalArtifacts === 3) {
    if (counts.slice().sort((a, b) => a - b).join(',') !== '1,2') return false
    const pairedZone = zones.find(z => (z.artifacts || []).length === 2)
    const otherZone = zones.find(z => (z.artifacts || []).length === 1)
    const pairedArts = pairedZone?.artifacts || []
    const pairedReasoning = pairedArts.filter(isReasoningArtifact)
    const pairKey = zoneArtifactPairKey(pairedZone)
    const allowedPairs = new Set([
      'chart+insight_text',
      'insight_text+workflow',
      'insight_text+table',
      'insight_text+prioritization',
      'driver_tree+insight_text',
      'insight_text+matrix',
      'cards+insight_text'
    ])
    if (pairedReasoning.length > 0 && pairedArts.some(a => !isReasoningArtifact(a) && artifactType(a) !== 'insight_text')) return false
    if ((otherZone?.artifacts || []).some(isReasoningArtifact) && pairedReasoning.length > 0) return false
    if (pairKey === 'cards+insight_text' && artifactCardCount(pairedArts.find(a => artifactType(a) === 'cards')) < 4) return false
    if (pairKey !== 'chart+table' && pairKey !== 'cards+workflow' && !allowedPairs.has(pairKey)) return false
    if (pairKey === 'cards+workflow') {
      const cards = pairedArts.find(a => artifactType(a) === 'cards')
      const workflow = pairedArts.find(a => artifactType(a) === 'workflow')
      if (!(artifactCardCount(cards) > 0 && artifactCardCount(cards) <= 2 && artifactWorkflowNodeCount(workflow) >= 4)) return false
    }
    if (pairKey === 'chart+table') {
      if (!pairedArts.every(isSubstantialArtifact)) return false
    }
    if ((otherZone?.artifacts || []).some(isReasoningArtifact) && reasoningCount > 1) return false
    return true
  }

  if (zoneCount === 2 && totalArtifacts === 4) {
    if (!counts.every(n => n === 2)) return false
    if (reasoningCount > 0) return false
    return zones.every(z => {
      const types = (z.artifacts || []).map(artifactType)
      const pair = types.slice().sort().join('+')
      return ['chart+insight_text', 'workflow+insight_text', 'insight_text+table', 'cards+insight_text'].includes(pair)
    })
  }

  if (zoneCount === 3 && totalArtifacts === 3) {
    if (!counts.every(n => n === 1)) return false
    const sparsePrimaryCards = zones.some(z => (z.narrative_weight || '').toLowerCase() === 'primary' && (z.artifacts || []).some(isSparseCardsArtifact))
    const hasPrimary = zones.some(z => (z.narrative_weight || '').toLowerCase() === 'primary')
    const hasSecondaryOrSupporting = zones.some(z => ['secondary', 'supporting'].includes((z.narrative_weight || '').toLowerCase()))
    return !sparsePrimaryCards
      && hasPrimary
      && hasSecondaryOrSupporting
  }

  if (zoneCount === 3 && totalArtifacts === 4) {
    if (counts.slice().sort((a, b) => a - b).join(',') !== '1,1,2') return false
    const pairedZones = zones.filter(z => (z.artifacts || []).length === 2)
    if (pairedZones.length !== 1) return false
    const pairedArts = pairedZones[0].artifacts || []
    if (pairedArts.some(isReasoningArtifact)) {
      if ((pairedZones[0].narrative_weight || '').toLowerCase() !== 'primary') return false
      if (pairedArts.some(a => !isReasoningArtifact(a) && artifactType(a) !== 'insight_text')) return false
    }
    if (reasoningCount > 0 && !slideHasDominantReasoningArtifact(slide)) return false
    return true
  }

  if (zoneCount === 3 && totalArtifacts >= 5) return false

  if (zoneCount === 4 && totalArtifacts === 4) {
    return counts.every(n => n === 1)
  }

  if (zoneCount === 4 && totalArtifacts >= 5) {
    if (reasoningCount > 0) return false
    return counts.every(n => n <= 2)
  }

  return true
}

function enforceReasoningArtifactUsage(slide) {
  const reasoningTypes = new Set(['matrix', 'driver_tree', 'prioritization'])
  if (!slide || slide.slide_type !== 'content') return slide
  const zones = (slide.zones || []).map(zone => {
    const arts = zone.artifacts || []
    const reasoningArts = arts.filter(a => reasoningTypes.has(String(a.type || '').toLowerCase()))
    if (!reasoningArts.length) return zone
    const primaryReasoning = reasoningArts[0]
    const insightArts = arts.filter(a => String(a.type || '').toLowerCase() === 'insight_text')
    return {
      ...zone,
      narrative_weight: 'primary',
      artifacts: [primaryReasoning].concat(insightArts.slice(0, 1))
    }
  })
  return { ...slide, zones }
}

function enforceStructuralPatternRules(slide) {
  if (!slide || slide.slide_type !== 'content') return slide
  const zones = (slide.zones || []).map(zone => {
    const arts = zone.artifacts || []
    if ((zone.narrative_weight || '').toLowerCase() === 'primary' && arts.some(isSparseCardsArtifact) && arts.length === 1) {
      return {
        ...zone,
        narrative_weight: zone.zone_role === 'summary' ? 'secondary' : 'secondary'
      }
    }
    return zone
  })
  return { ...slide, zones }
}

function validateSlideArtifactMix(slide) {
  if (slide.slide_type !== 'content') return true
  const artifacts = []
  ;(slide.zones || []).forEach(zone => (zone.artifacts || []).forEach(art => artifacts.push((art.type || '').toLowerCase())))
  if (!artifacts.length) return false
  if (artifacts.every(t => t === 'cards')) return false
  if (!validateStructuralPatternRules(slide)) return false
  if (!validateZoneStructureRules(slide)) return false
  return true
}

function hasPlaceholderContent(slide) {
  if (slide.slide_type !== 'content') return false
  if (!slide.zones || !slide.zones.length) return true
  if (!slide.key_message || slide.key_message.trim().length < 10) return true
  if (!validateReasoningArtifactUsage(slide)) return true
  if (!validateWorkflowUsage(slide)) return true
  if (!validateSlideArtifactMix(slide)) return true

  for (const zone of slide.zones) {
    for (const artifact of (zone.artifacts || [])) {
      const check = validateArtifact(artifact)
      if (!check.valid) return true
    }
  }
  return false
}


// ═══════════════════════════════════════════════════════════════════════════════
// NORMALISE
// ═══════════════════════════════════════════════════════════════════════════════

function normaliseArtifact(a) {
  if (!a || !a.type) return null
  // Resolve type early: chart+chart_type:stat_bar → stat_bar
  if (String(a.type).toLowerCase() === 'chart' && String(a.chart_type || '').toLowerCase() === 'stat_bar') {
    a.type = 'stat_bar'
  }
  const t = a.type.toLowerCase()
  if (a.artifact_coverage_hint != null) {
    const n = parseFloat(a.artifact_coverage_hint)
    a.artifact_coverage_hint = Number.isFinite(n) ? Math.max(1, Math.min(100, n)) : undefined
  }

  if (t === 'stat_bar') {
    if (!a.stat_header) a.stat_header = a.chart_header || ''
    if (!a.stat_decision) a.stat_decision = a.chart_decision || ''
    if (!a.rows) a.rows = []
    if (!a.column_headers) a.column_headers = {}
    if (!a.annotation_style) a.annotation_style = 'trailing'
    a.rows = a.rows.map((row, idx) => ({
      id:                      row?.id || `row_${idx + 1}`,
      label:                   row?.label || '',
      value:                   typeof row?.value === 'number' ? row.value : (parseFloat(String(row?.value || '').replace(/[^0-9.-]/g, '')) || 0),
      unit:                    row?.unit || '',
      display_value:           row?.display_value || '',
      annotation:              row?.annotation || '',
      annotation_representation: row?.annotation_representation || 'text',
      bar_color:               row?.bar_color || '',
      highlight:               row?.highlight === true
    }))
  }

  if (t === 'chart') {
    if (!a.categories) a.categories = []
    if (!a.series) a.series = []
    if (!a.chart_type) a.chart_type = 'bar'
    if (!a.chart_title) a.chart_title = ''
    if (!a.chart_header) a.chart_header = ''
    if (!a.chart_insight) a.chart_insight = ''
    if (a.show_data_labels === undefined) a.show_data_labels = true

    // Normalise series values to numbers
    a.series = a.series.map(s => ({
      name:   s.name   || '',
      values: (s.values || []).map(v => typeof v === 'number' ? v : parseFloat(String(v).replace(/[^0-9.-]/g,'')) || 0),
      unit:   s.unit   || (a.chart_type === 'group_pie' ? 'percent' : undefined),
      types:  s.types  || null
    }))

    // group_pie auto-fixes
    if (a.chart_type === 'group_pie') {
      // Too few entities → downgrade to single pie
      if (a.series.length < 2) a.chart_type = 'pie'
      // Too many entities → fallback to bar (data integrity preserved; flag in conflict log)
      else if (a.series.length > 8) a.chart_type = 'bar'
      // Too many slices → convert to clustered_bar
      else if (a.categories.length > 7) a.chart_type = 'clustered_bar'
    }

    // Auto-fix clustered_bar with 1 series
    if (a.chart_type === 'clustered_bar' && a.series.length < 2) {
      a.chart_type = 'bar'
    }

    // Align values length to categories length
    a.series = a.series.map(s => {
      while (s.values.length < a.categories.length) s.values.push(0)
      s.values = s.values.slice(0, a.categories.length)
      return s
    })
  }

  if (t === 'insight_text') {
    if (!a.points) a.points = []
    // Map insight_header → heading for backward compatibility with Agent 5/6
    if (!a.heading) a.heading = a.insight_header || ''
    if (!a.insight_header) a.insight_header = a.heading
    if (!a.sentiment) a.sentiment = 'neutral'
    // Normalise points — flatten if they're objects
    a.points = a.points.map(p => typeof p === 'string' ? p : (p.text || p.point || JSON.stringify(p)))
  }

  if (t === 'cards') {
    if (!a.cards) a.cards = []
    a.cards = a.cards.map(c => ({
      title:     c.title     || c.header || '',
      subtitle:  c.subtitle  || '',
      body:      c.body      || '',
      sentiment: c.sentiment || 'neutral'
    }))
  }

  if (t === 'workflow') {
    if (!a.nodes) a.nodes = []
    if (!a.connections) a.connections = []
    if (!a.workflow_type) a.workflow_type = 'process_flow'
    if (!a.flow_direction) a.flow_direction = 'left_to_right'
    if (!a.workflow_header) a.workflow_header = ''
    if (!a.workflow_title) a.workflow_title = ''
    if (!a.workflow_insight) a.workflow_insight = ''
  }

  if (t === 'table') {
    if (!a.headers) a.headers = []
    if (!a.rows) a.rows = []
    if (!a.title) a.title = ''
    if (!a.table_header) a.table_header = ''
  }

  if (t === 'matrix') {
    if (!a.matrix_type) a.matrix_type = '2x2'
    if (!a.matrix_header) a.matrix_header = ''
    if (!a.x_axis) a.x_axis = { label: '', low_label: '', high_label: '' }
    if (!a.y_axis) a.y_axis = { label: '', low_label: '', high_label: '' }
    if (!a.quadrants) a.quadrants = []
    if (!a.points) a.points = []
  }

  if (t === 'driver_tree') {
    if (!a.tree_header) a.tree_header = ''
    if (!a.root) a.root = { label: '', value: '' }
    if (!a.branches) a.branches = []
  }

  if (t === 'prioritization') {
    if (!a.priority_header) a.priority_header = ''
    if (!a.items) a.items = []
    a.items = a.items.map((item, idx) => ({
      rank: item.rank != null ? item.rank : (idx + 1),
      title: item.title || '',
      description: item.description || '',
      qualifiers: Array.isArray(item.qualifiers)
        ? item.qualifiers.slice(0, 2).map(q => ({
            label: (q && q.label) || '',
            value: (q && q.value) || ''
          }))
        : [
            { label: '', value: '' },
            { label: '', value: '' }
          ]
    }))
  }

  return a
}

function normaliseZone(z) {
  if (!z) return null
  const zoneSplit = z.zone_split || (z.layout_hint || {}).split || 'full'
  const artifactArrangement = (z.layout_hint || {}).artifact_arrangement || z.artifact_arrangement || null
  const artifactSplitHint = Array.isArray(z.artifact_split_hint)
    ? z.artifact_split_hint
    : (Array.isArray((z.layout_hint || {}).split_hint) ? (z.layout_hint || {}).split_hint : (Array.isArray(z.split_hint) ? z.split_hint : null))
  return {
    zone_id:          z.zone_id          || 'z1',
    zone_slot:        z.zone_slot        || '',
    zone_role:        z.zone_role        || 'primary_proof',
    message_objective:z.message_objective|| '',
    narrative_weight: z.narrative_weight || 'primary',
    artifacts:        (z.artifacts || []).map(normaliseArtifact).filter(Boolean),
    zone_split:       zoneSplit,
    artifact_arrangement: artifactArrangement,
    layout_hint:      {
      split: zoneSplit,
      artifact_arrangement: artifactArrangement,
      split_hint: artifactSplitHint
    }
  }
}

function normaliseSlide(slide, plan) {
  const slideType = slide.slide_type || plan.slide_type || 'content'

  let zones = []

  // Title and Divider slides must never have zones — enforce unconditionally
  // regardless of what Claude returned.
  if (slideType !== 'title' && slideType !== 'divider') {
    if (slide.zones && Array.isArray(slide.zones) && slide.zones.length > 0) {
      zones = slide.zones.map(normaliseZone).filter(Boolean)
    } else {
      // Build default zones from archetype
      const archetype = slide.slide_archetype || 'proof'
      zones = defaultZonesForArchetype(archetype).map(normaliseZone).filter(Boolean)
    }

    // Cap at 4 zones
    zones = zones.slice(0, 4)
  }

  const normalized = {
    slide_number:                 slide.slide_number                 || plan.slide_number,
    slide_type:                   slideType,
    narrative_role:               slide.narrative_role               || plan.narrative_role  || '',
    slide_archetype:              slide.slide_archetype              || 'proof',
    zone_structure:               slide.zone_structure               || '',
    selected_layout_name:         slide.selected_layout_name         || '',
    title:                        slide.title                        || plan.slide_title_draft || ('Slide ' + plan.slide_number),
    subtitle:                     slide.subtitle                     || '',
    key_message:                  slide.key_message                  || '',
    visual_flow_hint:             slide.visual_flow_hint             || '',
    context_from_previous_slide:  slide.context_from_previous_slide  || '',
    zones:                        zones,
    speaker_note:                 slide.speaker_note                 || plan.strategic_objective || ''
  }
  const enforced = applyZoneStructureMetadata(enforceStructuralPatternRules(enforceReasoningArtifactUsage(normalized)))
  return {
    ...enforced,
    slide_archetype: inferArchetypeFromZones(enforced)
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE PLAN BUILDER
// ═══════════════════════════════════════════════════════════════════════════════

function buildSlidePlan(brief) {
  // Returns the full deck — structural slides (title, dividers, thank-you) pre-assigned
  // by Agent 3 along with all content slides with their narrative roles.
  const slides = brief.slides || []
  return slides.map((s, i) => ({
    slide_number:          s.slide_number          || (i + 1),
    slide_type:            s.slide_type            || 'content',
    narrative_role:        s.narrative_role        || '',
    slide_title_draft:     s.slide_title_draft     || '',
    strategic_objective:   s.strategic_objective   || '',
    key_content:           Array.isArray(s.key_content) ? s.key_content : [],
    zone_count_signal:     s.zone_count_signal     || 'unsure',
    dominant_zone_signal:  s.dominant_zone_signal  || 'unsure',
    co_primary_signal:     s.co_primary_signal     || 'no',
    following_slide_claim: s.following_slide_claim || ''
  }))
}


// ═══════════════════════════════════════════════════════════════════════════════
// BATCH WRITER
// ═══════════════════════════════════════════════════════════════════════════════

async function writeSlideBatch(batchPlan, brief, contentB64, batchNum, layoutNames, summaryCardRegistry = []) {
  console.log('Agent 4 — batch', batchNum, ': slides', batchPlan[0].slide_number, '–', batchPlan[batchPlan.length-1].slide_number)

  const hasLayouts = layoutNames.length >= 5
  const briefSummary = buildBriefSummaryForAgent4(brief)
  const compactBatchPlan = JSON.stringify((batchPlan || []).map(plan => ({
    slide_number:          plan.slide_number,
    slide_type:            plan.slide_type,
    narrative_role:        plan.narrative_role        || '',
    slide_title_draft:     plan.slide_title_draft     || '',
    strategic_objective:   plan.strategic_objective   || '',
    key_content:           plan.key_content           || [],
    zone_count_signal:     plan.zone_count_signal     || 'unsure',
    dominant_zone_signal:  plan.dominant_zone_signal  || 'unsure',
    co_primary_signal:     plan.co_primary_signal     || 'no',
    following_slide_claim: plan.following_slide_claim || ''
  })))
  const keyMsgLines = (briefSummary.key_messages || []).map((m, i) => `  ${i + 1}. ${m}`).join('\n') || '  —'
  const registryLine = summaryCardRegistry.length > 0
    ? `SUMMARY_CARD_REGISTRY (Phase 3 Step 0B — do not repeat these as cards on proof slides):\n${summaryCardRegistry.map(c => `  { title: "${c.title}", value: "${c.value}" }`).join('\n')}`
    : 'SUMMARY_CARD_REGISTRY: empty — no summary slide processed yet; skip deduplication'

  const prompt = `PRESENTATION BRIEF:
Governing thought: ${briefSummary.governing_thought || '—'}
Audience:          ${briefSummary.audience || '—'}
Narrative flow:    ${briefSummary.narrative_flow || '—'}
Data heavy:        ${briefSummary.data_heavy ? 'yes — prefer charts, tables, data-rich artifacts' : 'no — prefer insight_text, cards, workflow artifacts'}
Tone:              ${briefSummary.tone || 'professional'}
Key messages:
${keyMsgLines}

${registryLine}

AVAILABLE BRAND LAYOUTS (${layoutNames.length}): ${hasLayouts
  ? layoutNames.join(' | ')
  : layoutNames.length > 0 ? layoutNames.join(' | ') + ' — too few layouts; use zone_split / artifact_arrangement plus per-artifact artifact_coverage_hint for scratch geometry'
  : 'none — use zone_split / artifact_arrangement plus per-artifact artifact_coverage_hint for scratch geometry'}

${hasLayouts
  ? `*** LAYOUT MODE ACTIVE — ${layoutNames.length} content layouts provided ***
For EVERY content slide you write:
  1. Set selected_layout_name to the best-matching layout name from the list above.
  2. Set layout_hint.split = "full" for ALL zones on that slide.
  3. Do NOT use split values like left_50, right_50, etc. — those are only for scratch mode.`
  : '*** SCRATCH MODE — fewer than 5 content layouts; use zone_split / artifact_arrangement plus per-artifact artifact_coverage_hint for geometry; mirror legacy hints only for compatibility ***'}

SLIDES TO WRITE — batch ${batchNum} (${batchPlan.length} slides):
${compactBatchPlan}

SLIDE ORDER RULE (mandatory):
Output slides in EXACTLY the order listed above. Do NOT reorder slides for any reason.
The slide_number in your output must match the slide_number in the plan entry you are writing.
The downstream renderer assembles slides positionally — any reordering produces a broken deck.

INSTRUCTIONS:
- For each slide, start from the locked Agent 3 plan, then derive zones and artifacts from the message.
- Use narrative_role, zone_count_signal, dominant_zone_signal, co_primary_signal, and strategic_objective as the primary zone-planning inputs.
- Do NOT infer structure from any legacy field. Use only narrative_role, zone_count_signal, dominant_zone_signal, co_primary_signal, and strategic_objective for zone planning.
- Use this zone-count logic as the authoritative rule set:
  - narrative_role = methodology_note -> 1 zone, stop
  - narrative_role = summary -> 1-2 zones; prefer 1 if one dense synthesis artifact can carry the slide
  - co_primary_signal = yes -> 2 zones, co-primary, side-by-side
  - narrative_role = benchmark_comparison -> 2 zones, equal weight unless dominant_zone_signal = yes
  - narrative_role = trend_analysis -> 2 zones, dominant proof plus implication support
  - narrative_role = segmentation -> 2-3 zones, comparison-led
  - narrative_role = drill_down -> 2-3 zones, dominant decomposition plus support
  - narrative_role = waterfall_decomposition -> 2 zones, dominant proof plus explanation
  - narrative_role = scenario_analysis -> 3-4 zones, grid or structured comparison preferred
  - narrative_role = decision_framework -> 3-4 zones, option comparison or criteria grid preferred
  - narrative_role = risk_register -> 2-3 zones, dominant register plus mitigation or implication support
  - narrative_role = recommendations -> 2-3 zones, recommendation plus rationale or ask support
  - narrative_role = exception_highlight -> 2 zones, dominant issue plus implication or action support
  - narrative_role = context_setter or problem_statement -> 2 zones, framing plus consequence or evidence support
  - zone_count_signal = 1|2|3|4 -> use that count as the baseline when no stronger rule above applies
  - strategic_objective implies comparing options, scenarios, or alternatives -> 3-4 zones
  - strategic_objective implies a single core proof with one takeaway -> 2 zones
  - strategic_objective implies a compact synthesis or note -> 1 zone
  - default -> 2 zones
- Apply these modifiers after selecting the baseline:
  - dominant_zone_signal = yes -> Zone 1 must be dominant
  - dominant_zone_signal = no and co_primary_signal = no -> prefer balanced weights across zones
  - scenario_analysis, decision_framework, and recommendations should not collapse below 2 zones
- Write the title from slide_title_draft and sharpen it only if needed.
- Before finalizing artifacts for a content slide, choose ONE zone_structure that matches the zone count and narrative geometry:
  - ZS01_single_full
  - ZS02_stacked_equal
  - ZS03_side_by_side_equal
  - ZS04_left_dominant_right_stack
  - ZS05_right_dominant_left_stack
  - ZS06_top_full_bottom_two
  - ZS07_top_two_bottom_dominant
  - ZS08_quad_grid
  - ZS09_left_dominant_right_triptych
  - ZS10_top_full_bottom_three
  - ZS11_three_rows_equal
  - ZW01_three_columns_equal
  - ZW02_three_columns_right_stack
  - ZW03_three_columns_left_stack
  - ZW04_four_columns_equal
- After choosing zone_structure, decide which slot is dominant vs support, then pick allowed artifacts for each slot. For asymmetric structures, dominant slots may carry chart / workflow / table / reasoning artifacts, while support slots should prefer insight_text, grouped insight_text, compact cards, or compact charts.
- slide_archetype is a descriptive label only; it must summarize the final zone/artifact structure and must never drive artifact selection or layout choice
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
- In Scratch Mode, if a zone has 2 artifacts, set artifact_arrangement and set artifact_coverage_hint on EACH artifact so the hints sum to 100.
- In Scratch Mode, cards with 1–2 items are compact summary anchors only: keep their zone share at or below ~40% of the slide, prefer top strips or narrow side panes, and never let 2 sparse cards occupy a tall dominant zone.
- Card density rule: unless a single cards artifact contains 8+ cards, no individual card may imply more than ~15% of total slide area.
- Every chart: MUST have 3+ categories, matching values, no all-zeros; set chart_header to the one-line insight the chart proves
- clustered_bar: MUST have exactly 2 series
- Every insight_text: MUST have specific, data-driven points; set insight_header to a 2–4 word specific label naming the implication (e.g. "Risk Implication", "Growth Opportunity", "Action Required") — never use generic labels like "So What" or "Key Insight"
- Workflows: fully populate nodes and connections; set workflow_header to the one-line insight
- Workflow restrictions:
  - process_flow: left_to_right only, >=4 nodes, full-width zone, >=50% height
  - hierarchy: top_down_branching only, >=3 levels, >=50% width, full content height
  - decomposition: left_to_right or top_to_bottom / top_down_branching only; if >3 nodes it must own full width (left_to_right) or full height (vertical)
  - timeline: left_to_right only, >=4 nodes, full-width zone, >=50% height
  - information_flow: do not use
- Tables: set table_header to the one-line insight the table proves
- Matrix: fully populate axes, all 4 quadrants, and plotted points; set matrix_header
- Driver_tree: fully populate root and branches; set tree_header
- Prioritization: fully populate ranked items sorted by importance; set priority_header
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
- Return ONLY a valid JSON array for these ${batchPlan.length} slides`

  const messages = [{
    role: 'user',
    content: [
      { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: contentB64 } },
      { type: 'text', text: prompt }
    ]
  }]

  const raw    = await callClaude(AGENT4_SYSTEM, messages, 5000)
  const parsed = safeParseJSON(raw, null)

  if (!Array.isArray(parsed)) {
    console.warn('Agent 4 batch', batchNum, '— parse failed, raw length:', raw.length,
      '| first 300:', raw.slice(0, 300))
    return null
  }

  // Warn if Claude returned fewer slides than requested (token truncation sign)
  if (parsed.length < batchPlan.length) {
    console.warn('Agent 4 batch', batchNum, '— expected', batchPlan.length,
      'slides but got', parsed.length, '— some may be missing')
  }

  console.log('Agent 4 batch', batchNum, '— got', parsed.length, 'slides')
  return parsed
}


// ═══════════════════════════════════════════════════════════════════════════════
// LAYOUT PICKER — maps zone count to a best-effort layout name
// ═══════════════════════════════════════════════════════════════════════════════

function pickBestLayout(slide, layoutNames) {
  const zoneCount = (slide.zones || []).length || 1
  const zones = slide.zones || []
  const artifacts = zones.flatMap(z => z.artifacts || [])
  const artifactTypes = artifacts.map(a => (a.type || '').toLowerCase())
  const artifactCount = artifacts.length || 1
  const singleFullZone = zones.length === 1 && (((zones[0].layout_hint || {}).split || 'full') === 'full')
  const hasReasoningArtifact = artifactTypes.some(t => ['matrix', 'driver_tree', 'prioritization'].includes(t))
  const hasGroupedInsightOnly = artifactTypes.length > 0 && artifactTypes.every(t => t === 'insight_text')
  const hasWideWorkflow = artifacts.some(a => {
    const t = (a.type || '').toLowerCase()
    const dir = (a.flow_direction || '').toLowerCase()
    return t === 'workflow' && (dir === 'left_to_right' || dir === 'timeline')
  })
  const hasTallWorkflow = artifacts.some(a => {
    const t = (a.type || '').toLowerCase()
    const dir = (a.flow_direction || '').toLowerCase()
    return t === 'workflow' && (dir === 'top_to_bottom' || dir === 'top_down_branching')
  })
  const hasWideChart = artifacts.some(a => {
    const t = (a.type || '').toLowerCase()
    const cats = Array.isArray(a.categories) ? a.categories.length : 0
    return t === 'chart' && cats > 6
  })
  const hasTallHorizontalBar = artifacts.some(a => {
    const t = (a.type || '').toLowerCase()
    const chartType = (a.chart_type || '').toLowerCase()
    const cats = Array.isArray(a.categories) ? a.categories.length : 0
    return t === 'chart' && chartType === 'horizontal_bar' && cats > 6
  })
  const hasLargeGroupPie = artifacts.some(a => {
    const t = (a.type || '').toLowerCase()
    const chartType = (a.chart_type || '').toLowerCase()
    const pies = Array.isArray(a.series) ? a.series.length : 0
    return t === 'chart' && chartType === 'group_pie' && pies >= 5
  })
  const isOneZoneTwoArtifacts = zoneCount === 1 && artifactCount === 2
  const isTwoZoneFourArtifacts = zoneCount === 2 && artifactCount === 4

  const findByPatterns = (patterns) => {
    for (const pat of patterns) {
      const hit = layoutNames.find(n => pat.test(n))
      if (hit) return hit
    }
    return ''
  }

  if (hasWideWorkflow || hasReasoningArtifact || hasWideChart || hasLargeGroupPie) {
    const hit = findByPatterns([
      /body.?text|1\s*across|single|1\s*col/i,
      /title\s+and\s+content/i
    ])
    if (hit) return hit
  }

  if (hasTallWorkflow || hasTallHorizontalBar) {
    const hit = findByPatterns([
      /2\s*across|1.?on.?1|left.?right|2.?col/i,
      /body.?text|1\s*across|single|1\s*col/i
    ])
    if (hit) return hit
  }

  if (singleFullZone && hasGroupedInsightOnly) {
    const hit = findByPatterns([
      /body.?text|1\s*across|single|1\s*col/i,
      /2\s*column|2\s*col/i
    ])
    if (hit) return hit
  }

  if (isTwoZoneFourArtifacts) {
    const hit = findByPatterns([
      /2.?on.?2|4\s*block|four|grid/i,
      /4\s*across/i
    ])
    if (hit) return hit
  }

  if (isOneZoneTwoArtifacts) {
    const pair = artifactTypes.slice().sort().join('+')
    const horizontalPairs = new Set([
      'chart+insight_text',
      'cards+insight_text',
      'chart+cards',
      'table+insight_text'
    ])
    const verticalPairs = new Set([
      'workflow+insight_text',
      'chart+table'
    ])

    if (horizontalPairs.has(pair)) {
      const hit = findByPatterns([
        /2\s*across|1.?on.?1|left.?right|2.?col/i,
        /body.?text|1\s*across|single|1\s*col/i
      ])
      if (hit) return hit
    }

    if (verticalPairs.has(pair)) {
      const hit = findByPatterns([
        /body.?text|1\s*across|single|1\s*col/i,
        /1.?on.?2|1.?on.?3/i
      ])
      if (hit) return hit
    }
  }

  // Ordered preference patterns per zone count (matched against layout names)
  const byCount = {
    1: [/1\s*across|body.?text|single|1\s*col/i],
    2: [/2\s*across|1.?on.?1|left.?right|2.?col/i],
    3: [/3\s*across|1.?on.?2|2.?on.?1/i],
    4: [/2.?on.?2|3.?on.?3|four/i]
  }
  const patterns = byCount[Math.min(zoneCount, 4)] || byCount[1]
  for (const pat of patterns) {
    const hit = layoutNames.find(n => pat.test(n))
    if (hit) return hit
  }
  return ''
}

function layoutConflictsWithSlide(slide, layoutName) {
  if (!layoutName) return true
  const name = String(layoutName).toLowerCase()
  const zones = slide.zones || []
  const artifacts = zones.flatMap(z => z.artifacts || [])
  const artifactTypes = artifacts.map(a => (a.type || '').toLowerCase())
  const artifactCount = artifacts.length
  const singleFullZone = zones.length === 1 && (((zones[0]?.layout_hint || {}).split || 'full') === 'full')
  const isTwoZoneFourArtifacts = zones.length === 2 && artifactCount === 4

  if (singleFullZone && artifactTypes.some(t => ['matrix', 'driver_tree', 'prioritization'].includes(t))) {
    if (/3\s*column|3\s*col|4\s*block|2.?on.?2|3\s*across|4\s*across/i.test(name)) return true
  }

  if (singleFullZone && artifactCount === 2 && /3\s*column|3\s*col|4\s*block|2.?on.?2|3\s*across|4\s*across/i.test(name)) {
    return true
  }

  if (singleFullZone && artifactTypes.every(t => t === 'insight_text')) {
    if (/3\s*column|3\s*col|4\s*block|2.?on.?2|3\s*across|4\s*across/i.test(name)) return true
  }

  if (isTwoZoneFourArtifacts && /3\s*column|3\s*col|1.?on.?2|2.?on.?1/i.test(name)) {
    return true
  }

  // group_pie with ≥ 5 pies needs a wide single-column layout — narrow multi-col layouts conflict
  const hasLargeGroupPieConflict = artifacts.some(a => {
    return (a.type || '').toLowerCase() === 'chart' &&
           (a.chart_type || '').toLowerCase() === 'group_pie' &&
           (Array.isArray(a.series) ? a.series.length : 0) >= 5
  })
  if (hasLargeGroupPieConflict && /3\s*column|3\s*col|4\s*block|2.?on.?2|3\s*across|4\s*across|2\s*across|1.?on.?1/i.test(name)) {
    return true
  }

  return false
}

function hasCompatibleLayout(slide, layoutNames) {
  return !!pickBestLayout(slide, layoutNames)
}

function artifactDominanceScoreForScratch(artifact) {
  const type = String((artifact || {}).type || '').toLowerCase()
  const sentiment = String((artifact || {}).sentiment || '').toLowerCase()
  const pointCount = Array.isArray(artifact?.points) ? artifact.points.length : 0
  if (['matrix', 'driver_tree', 'prioritization'].includes(type)) return 100
  if (type === 'workflow') return 92
  if (type === 'chart') return 88
  if (type === 'table') return 76
  if (type === 'insight_text') return 72 + Math.min(pointCount, 5) * 2 + (sentiment === 'warning' ? 4 : 0)
  if (type === 'cards') {
    const cardCount = Array.isArray(artifact?.cards) ? artifact.cards.length : 0
    if (cardCount <= 2) return 34
    if (cardCount === 3) return 42
    return 52
  }
  return 60
}

function zoneDominanceScoreForScratch(zone) {
  const weight = String(zone?.narrative_weight || '').toLowerCase()
  const role = String(zone?.zone_role || '').toLowerCase()
  const artifacts = zone?.artifacts || []
  let score = weight === 'primary' ? 100 : weight === 'secondary' ? 72 : 60
  if (/primary|proof|recommendation|summary/.test(role)) score += 8
  if (/implication|supporting/.test(role)) score -= 4
  score += Math.max(...artifacts.map(artifactDominanceScoreForScratch), 0) * 0.2
  return score
}

function zoneHasOnlyCards(zone) {
  const arts = zone?.artifacts || []
  return arts.length > 0 && arts.every(a => String((a || {}).type || '').toLowerCase() === 'cards')
}

function zoneCardCount(zone) {
  const cardArtifacts = (zone?.artifacts || []).filter(a => String((a || {}).type || '').toLowerCase() === 'cards')
  return cardArtifacts.reduce((sum, art) => sum + ((art.cards || []).length || 0), 0)
}

function isCompactCardsZone(zone) {
  return zoneHasOnlyCards(zone) && zoneCardCount(zone) > 0 && zoneCardCount(zone) <= 2
}

function preferCompactCardsZone(zone, split) {
  return {
    ...zone,
    zone_split: split,
    layout_hint: { ...(zone.layout_hint || {}), split }
  }
}

function orderZonesForStructure(slide) {
  const zones = [...(slide?.zones || [])]
  const structureId = slide?.zone_structure || inferZoneStructure(slide)
  const def = zoneStructureDef(structureId)
  if (!def || zones.length <= 1) return zones

  const scored = zones.map((zone, idx) => ({
    zone,
    idx,
    score: zoneDominanceScoreForScratch(zone)
  })).sort((a, b) => b.score - a.score)

  if (['ZS04_left_dominant_right_stack', 'ZS09_left_dominant_right_triptych'].includes(structureId)) {
    const [dominant, ...rest] = scored
    return [dominant?.zone, ...rest.map(x => x.zone)].filter(Boolean)
  }
  if (structureId === 'ZS05_right_dominant_left_stack') {
    const [dominant, ...rest] = scored
    return [...rest.map(x => x.zone), dominant?.zone].filter(Boolean)
  }
  if (structureId === 'ZS06_top_full_bottom_two') {
    const compact = zones.find(isCompactCardsZone)
    const rest = zones.filter(z => z !== compact).sort((a, b) => zoneDominanceScoreForScratch(b) - zoneDominanceScoreForScratch(a))
    return compact ? [compact, ...rest] : scored.map(x => x.zone)
  }
  if (structureId === 'ZS07_top_two_bottom_dominant') {
    const [dominant, ...rest] = scored
    return [...rest.map(x => x.zone), dominant?.zone].filter(Boolean)
  }
  return scored.map(x => x.zone)
}

function applyZoneStructureScratchSplits(slide) {
  const structureId = slide.zone_structure || inferZoneStructure(slide)
  const def = zoneStructureDef(structureId)
  if (!def) return slide
  const orderedZones = orderZonesForStructure(slide)
  if (orderedZones.length !== def.zoneCount) return slide
  const zones = orderedZones.map((zone, idx) => {
    const split = def.scratchSplits[idx] || 'full'
    const slot = def.slots[idx] || `slot_${idx + 1}`
    return applyArtifactArrangementForScratch({
      ...zone,
      zone_slot: slot,
      zone_split: split,
      layout_hint: { ...(zone.layout_hint || {}), split }
    }, isDominantSlot(slot) ? 65 : 55)
  })
  return { ...slide, zone_structure: structureId, selected_layout_name: '', zones }
}

function applyArtifactArrangementForScratch(zone, dominantShare = 60) {
  const artifacts = zone?.artifacts || []
  if (artifacts.length < 2) return zone

  const scored = artifacts.map((art, idx) => ({
    idx,
    art,
    score: artifactDominanceScoreForScratch(art)
  }))
  scored.sort((a, b) => b.score - a.score)

  const dominantIdx = scored[0]?.idx ?? 0
  const secondaryShare = 100 - dominantShare
  let firstShare = dominantIdx === 0 ? dominantShare : secondaryShare
  let secondShare = 100 - firstShare

  const firstType = String((artifacts[0] || {}).type || '').toLowerCase()
  const secondType = String((artifacts[1] || {}).type || '').toLowerCase()

  if (firstType === 'cards') firstShare = Math.min(firstShare, 40)
  if (secondType === 'cards') secondShare = Math.min(secondShare, 40)

  if (firstType === 'cards' && secondType !== 'cards') secondShare = 100 - firstShare
  if (secondType === 'cards' && firstType !== 'cards') firstShare = 100 - secondShare
  if (firstType === 'cards' && secondType === 'cards') {
    firstShare = Math.min(firstShare, 40)
    secondShare = Math.min(secondShare, 40)
  }

  const coverage = artifacts.map((_, idx) => {
    if (artifacts.length === 2) return idx === 0 ? firstShare : secondShare
    if (idx === 0) return firstShare
    const rem = Math.max(0, 100 - firstShare)
    return rem / Math.max(artifacts.length - 1, 1)
  })
  const normalizedCoverage = coverage.map(v => Math.round(v * 100) / 100)
  const artifactsWithCoverage = artifacts.map((art, idx) => ({
    ...art,
    artifact_coverage_hint: normalizedCoverage[idx]
  }))

  return {
    ...zone,
    artifacts: artifactsWithCoverage,
    artifact_split_hint: artifacts.length === 2 ? [firstShare, secondShare] : normalizedCoverage,
    artifact_arrangement: 'vertical',
    split_hint: artifacts.length === 2 ? [firstShare, secondShare] : normalizedCoverage,
    layout_hint: {
      ...(zone.layout_hint || {}),
      artifact_arrangement: 'vertical',
      split_hint: artifacts.length === 2 ? [firstShare, secondShare] : normalizedCoverage
    }
  }
}

function assignScratchSplits(slide) {
  if (slide.slide_type !== 'content') return slide
  slide = applyZoneStructureMetadata(slide)
  if (validateZoneStructureRules(slide)) {
    return applyZoneStructureScratchSplits(slide)
  }
  const zones = (slide.zones || []).map(z => ({
    ...z,
    zone_split: z.zone_split || ((z.layout_hint || {}).split || 'full'),
    artifact_arrangement: (z.layout_hint || {}).artifact_arrangement || z.artifact_arrangement || null,
    artifact_split_hint: Array.isArray(z.artifact_split_hint) ? z.artifact_split_hint : (Array.isArray((z.layout_hint || {}).split_hint) ? (z.layout_hint || {}).split_hint : (Array.isArray(z.split_hint) ? z.split_hint : null)),
    layout_hint: {
      ...(z.layout_hint || {}),
      split: z.zone_split || ((z.layout_hint || {}).split || 'full'),
      artifact_arrangement: (z.layout_hint || {}).artifact_arrangement || z.artifact_arrangement || null,
      split_hint: Array.isArray(z.artifact_split_hint) ? z.artifact_split_hint : (Array.isArray((z.layout_hint || {}).split_hint) ? (z.layout_hint || {}).split_hint : (Array.isArray(z.split_hint) ? z.split_hint : null))
    }
  }))
  if (!zones.length) return slide

  const zoneArtifactTypes = zones.map(z => (z.artifacts || []).map(a => (a.type || '').toLowerCase()))
  const allArtifacts = zoneArtifactTypes.flat()
  const artifactCount = allArtifacts.length
  const hasReasoning = allArtifacts.some(t => ['matrix', 'driver_tree', 'prioritization'].includes(t))
  const compactCardZoneIndices = zones.map((z, idx) => ({ idx, compact: isCompactCardsZone(z) })).filter(x => x.compact).map(x => x.idx)
  const hasWideWorkflow = zones.some(z => (z.artifacts || []).some(a => {
    const t = (a.type || '').toLowerCase()
    const dir = (a.flow_direction || '').toLowerCase()
    return t === 'workflow' && (dir === 'left_to_right' || dir === 'timeline')
  }))
  const hasWideChart = zones.some(z => (z.artifacts || []).some(a => {
    const t = (a.type || '').toLowerCase()
    const cats = Array.isArray(a.categories) ? a.categories.length : 0
    return t === 'chart' && cats > 6
  }))

  if (zones.length === 1) {
    if (isCompactCardsZone(zones[0])) {
      zones[0] = applyArtifactArrangementForScratch(preferCompactCardsZone(zones[0], 'top_35'), 50)
    } else {
      zones[0] = applyArtifactArrangementForScratch({ ...zones[0], zone_split: 'full', layout_hint: { ...(zones[0].layout_hint || {}), split: 'full' } }, 60)
    }
  } else if (zones.length === 2 && artifactCount === 4) {
    const zoneScores = zones.map((z, idx) => ({ idx, score: zoneDominanceScoreForScratch(z) }))
    zoneScores.sort((a, b) => b.score - a.score)
    const dominantIdx = zoneScores[0]?.idx ?? 0
    const supportingIdx = dominantIdx === 0 ? 1 : 0

    if (compactCardZoneIndices.includes(dominantIdx) || compactCardZoneIndices.includes(supportingIdx)) {
      const cardIdx = compactCardZoneIndices.includes(dominantIdx) ? dominantIdx : supportingIdx
      const otherIdx = cardIdx === dominantIdx ? supportingIdx : dominantIdx
      zones[cardIdx] = applyArtifactArrangementForScratch(preferCompactCardsZone(zones[cardIdx], 'top_35'), 50)
      zones[otherIdx] = applyArtifactArrangementForScratch({
        ...zones[otherIdx],
        zone_split: 'bottom_65',
        layout_hint: { ...(zones[otherIdx].layout_hint || {}), split: 'bottom_65' }
      }, 65)
    } else {
      zones[dominantIdx] = applyArtifactArrangementForScratch({
        ...zones[dominantIdx],
        zone_split: 'left_60',
        layout_hint: { ...(zones[dominantIdx].layout_hint || {}), split: 'left_60' }
      }, 60)
      zones[supportingIdx] = applyArtifactArrangementForScratch({
        ...zones[supportingIdx],
        zone_split: 'right_40',
        layout_hint: { ...(zones[supportingIdx].layout_hint || {}), split: 'right_40' }
      }, 40)
    }
  } else if (zones.length === 2) {
    if (compactCardZoneIndices.length === 1) {
      const cardIdx = compactCardZoneIndices[0]
      const otherIdx = cardIdx === 0 ? 1 : 0
      zones[cardIdx] = applyArtifactArrangementForScratch(preferCompactCardsZone(zones[cardIdx], 'top_35'), 50)
      zones[otherIdx] = applyArtifactArrangementForScratch({
        ...zones[otherIdx],
        zone_split: 'bottom_65',
        layout_hint: { ...(zones[otherIdx].layout_hint || {}), split: 'bottom_65' }
      }, 65)
    } else if (hasReasoning || hasWideWorkflow || hasWideChart) {
      const zoneScores = zones.map((z, idx) => ({ idx, score: zoneDominanceScoreForScratch(z) }))
      zoneScores.sort((a, b) => b.score - a.score)
      const dominantIdx = zoneScores[0]?.idx ?? 0
      const secondaryIdx = dominantIdx === 0 ? 1 : 0
      zones[dominantIdx] = applyArtifactArrangementForScratch({
        ...zones[dominantIdx],
        zone_split: 'top_60',
        layout_hint: { ...(zones[dominantIdx].layout_hint || {}), split: 'top_60' }
      }, 60)
      zones[secondaryIdx] = applyArtifactArrangementForScratch({
        ...zones[secondaryIdx],
        zone_split: 'bottom_40',
        layout_hint: { ...(zones[secondaryIdx].layout_hint || {}), split: 'bottom_40' }
      }, 40)
    } else {
      const zoneScores = zones.map((z, idx) => ({ idx, score: zoneDominanceScoreForScratch(z) }))
      zoneScores.sort((a, b) => b.score - a.score)
      const dominantIdx = zoneScores[0]?.idx ?? 0
      const secondaryIdx = dominantIdx === 0 ? 1 : 0
      zones[dominantIdx] = applyArtifactArrangementForScratch({
        ...zones[dominantIdx],
        zone_split: 'left_60',
        layout_hint: { ...(zones[dominantIdx].layout_hint || {}), split: 'left_60' }
      }, 60)
      zones[secondaryIdx] = applyArtifactArrangementForScratch({
        ...zones[secondaryIdx],
        zone_split: 'right_40',
        layout_hint: { ...(zones[secondaryIdx].layout_hint || {}), split: 'right_40' }
      }, 40)
    }
  } else if (zones.length === 3) {
    if (compactCardZoneIndices.length === 1) {
      const cardIdx = compactCardZoneIndices[0]
      const otherIdxs = [0, 1, 2].filter(i => i !== cardIdx)
      zones[cardIdx] = applyArtifactArrangementForScratch(preferCompactCardsZone(zones[cardIdx], 'top_35'), 50)
      zones[otherIdxs[0]] = applyArtifactArrangementForScratch({ ...zones[otherIdxs[0]], zone_split: 'top_left_50', layout_hint: { ...(zones[otherIdxs[0]].layout_hint || {}), split: 'top_left_50' } }, 60)
      zones[otherIdxs[1]] = applyArtifactArrangementForScratch({ ...zones[otherIdxs[1]], zone_split: 'bottom_full', layout_hint: { ...(zones[otherIdxs[1]].layout_hint || {}), split: 'bottom_full' } }, 60)
    } else {
      zones[0] = applyArtifactArrangementForScratch({ ...zones[0], zone_split: 'top_left_50', layout_hint: { ...(zones[0].layout_hint || {}), split: 'top_left_50' } }, 60)
      zones[1] = applyArtifactArrangementForScratch({ ...zones[1], zone_split: 'top_right_50', layout_hint: { ...(zones[1].layout_hint || {}), split: 'top_right_50' } }, 60)
      zones[2] = applyArtifactArrangementForScratch({ ...zones[2], zone_split: 'bottom_full', layout_hint: { ...(zones[2].layout_hint || {}), split: 'bottom_full' } }, 60)
    }
  } else if (zones.length >= 4) {
    const splits = ['tl', 'tr', 'bl', 'br']
    zones.forEach((z, i) => {
      zones[i] = applyArtifactArrangementForScratch({ ...z, zone_split: splits[i] || 'full', layout_hint: { ...(z.layout_hint || {}), split: splits[i] || 'full' } }, 60)
    })
  }

  return { ...slide, selected_layout_name: '', zones }
}

function pruneAgent4SlideForOutput(slide) {
  if (!slide || typeof slide !== 'object') return slide
  const pruneArtifactForOutput = (artifact) => {
    if (!artifact || typeof artifact !== 'object') return artifact
    const rawType = String(artifact.type || '').toLowerCase()
    const chartType = String(artifact.chart_type || '').toLowerCase()
    const type = rawType === 'stat_bar' || (rawType === 'chart' && chartType === 'stat_bar')
      ? 'stat_bar'
      : rawType
    const coverage = artifact.artifact_coverage_hint != null
      ? { artifact_coverage_hint: artifact.artifact_coverage_hint }
      : {}
    if (type === 'insight_text') {
      return {
        type: 'insight_text',
        insight_header: artifact.insight_header || '',
        points: Array.isArray(artifact.points) ? artifact.points : [],
        groups: Array.isArray(artifact.groups) ? artifact.groups : [],
        sentiment: artifact.sentiment || 'neutral',
        ...coverage
      }
    }
    if (type === 'chart') {
      const chartType = (artifact.chart_type || 'bar').toLowerCase()
      if (chartType === 'group_pie') {
        return {
          type: 'chart',
          chart_type: 'group_pie',
          chart_title: artifact.chart_title || '',
          chart_header: artifact.chart_header || '',
          categories: Array.isArray(artifact.categories) ? artifact.categories : [],
          series: Array.isArray(artifact.series) ? artifact.series.map(s => ({
            name:   s.name   || '',
            values: Array.isArray(s.values) ? s.values : [],
            unit:   s.unit   || 'percent'
          })) : [],
          show_legend: artifact.show_legend !== false,
          show_data_labels: artifact.show_data_labels !== false,
          ...coverage
        }
      }
      return {
        type: 'chart',
        chart_type: artifact.chart_type || 'bar',
        chart_title: artifact.chart_title || '',
        chart_header: artifact.chart_header || '',
        x_label: artifact.x_label || '',
        y_label: artifact.y_label || '',
        categories: Array.isArray(artifact.categories) ? artifact.categories : [],
        series: Array.isArray(artifact.series) ? artifact.series : [],
        dual_axis: artifact.dual_axis === true,
        secondary_series: Array.isArray(artifact.secondary_series) ? artifact.secondary_series : [],
        show_data_labels: artifact.show_data_labels !== false,
        show_legend: artifact.show_legend === true,
        ...coverage
      }
    }
    if (type === 'stat_bar') {
      return {
        type: 'stat_bar',
        stat_header: artifact.stat_header || artifact.chart_header || '',
        stat_decision: artifact.stat_decision || artifact.chart_insight || '',
        column_headers: artifact.column_headers || {},
        rows: Array.isArray(artifact.rows) ? artifact.rows.map((row, idx) => ({
          id: row?.id || `row_${idx + 1}`,
          label: row?.label || '',
          value: row?.value,
          unit: row?.unit || '',
          display_value: row?.display_value || '',
          annotation: row?.annotation || '',
          annotation_representation: row?.annotation_representation || 'text',
          bar_color: row?.bar_color || '',
          highlight: row?.highlight === true
        })) : [],
        annotation_style: artifact.annotation_style || 'trailing',
        ...coverage
      }
    }
    if (type === 'cards') {
      return {
        type: 'cards',
        artifact_header_text: artifact.artifact_header_text || '',
        cards: Array.isArray(artifact.cards) ? artifact.cards.map(card => ({
          title: card?.title || '',
          subtitle: card?.subtitle || '',
          body: card?.body || '',
          sentiment: card?.sentiment || 'neutral'
        })) : [],
        ...coverage
      }
    }
    if (type === 'workflow') {
      return {
        type: 'workflow',
        workflow_type: artifact.workflow_type || 'process_flow',
        flow_direction: artifact.flow_direction || 'left_to_right',
        workflow_header: artifact.workflow_header || '',
        workflow_title: artifact.workflow_title || '',
        workflow_insight: artifact.workflow_insight || '',
        nodes: Array.isArray(artifact.nodes) ? artifact.nodes.map(node => ({
          id: node?.id || '',
          label: node?.label || '',
          value: node?.value || '',
          description: node?.description || '',
          level: node?.level != null ? node.level : 1
        })) : [],
        connections: Array.isArray(artifact.connections) ? artifact.connections.map(conn => ({
          from: conn?.from || '',
          to: conn?.to || '',
          type: conn?.type || 'arrow'
        })) : [],
        ...coverage
      }
    }
    if (type === 'table') {
      return {
        type: 'table',
        table_header: artifact.table_header || '',
        title: artifact.title || '',
        headers: Array.isArray(artifact.headers) ? artifact.headers : [],
        rows: Array.isArray(artifact.rows) ? artifact.rows : [],
        highlight_rows: Array.isArray(artifact.highlight_rows) ? artifact.highlight_rows : [],
        note: artifact.note || '',
        ...coverage
      }
    }
    if (type === 'matrix') {
      return {
        type: 'matrix',
        matrix_type: artifact.matrix_type || '2x2',
        matrix_header: artifact.matrix_header || '',
        x_axis: artifact.x_axis || { label: '', low_label: '', high_label: '' },
        y_axis: artifact.y_axis || { label: '', low_label: '', high_label: '' },
        quadrants: Array.isArray(artifact.quadrants) ? artifact.quadrants : [],
        points: Array.isArray(artifact.points) ? artifact.points : [],
        ...coverage
      }
    }
    if (type === 'driver_tree') {
      return {
        type: 'driver_tree',
        tree_header: artifact.tree_header || '',
        root: artifact.root || { label: '', value: '' },
        branches: Array.isArray(artifact.branches) ? artifact.branches : [],
        ...coverage
      }
    }
    if (type === 'prioritization') {
      return {
        type: 'prioritization',
        priority_header: artifact.priority_header || '',
        items: Array.isArray(artifact.items) ? artifact.items.map(item => ({
          rank: item?.rank,
          title: item?.title || '',
          description: item?.description || '',
          qualifiers: Array.isArray(item?.qualifiers) ? item.qualifiers.slice(0, 2).map(q => ({
            label: q?.label || '',
            value: q?.value || ''
          })) : []
        })) : [],
        ...coverage
      }
    }
    if (type === 'comparison_table') {
      return {
        type: 'comparison_table',
        comparison_header: artifact.comparison_header || artifact.table_header || '',
        criteria: Array.isArray(artifact.criteria) ? artifact.criteria.map(c => ({
          id: c?.id || '',
          label: c?.label || ''
        })) : [],
        options: Array.isArray(artifact.options) ? artifact.options.map(o => ({
          id: o?.id,
          name: o?.name || '',
          badge_text: o?.badge_text || undefined,
          cells: Array.isArray(o?.cells) ? o.cells.map(cell => ({
            criterion_id: cell?.criterion_id || '',
            rating: cell?.rating || 'text',
            display_value: cell?.display_value || undefined,
            note: cell?.note || undefined,
            representation_type: cell?.representation_type || undefined
          })) : []
        })) : [],
        recommended_option_id: artifact.recommended_option_id || undefined,
        recommended_option: artifact.recommended_option || undefined,
        ...coverage
      }
    }
    if (type === 'initiative_map') {
      return {
        type: 'initiative_map',
        initiative_header: artifact.initiative_header || artifact.table_header || '',
        dimension_labels: Array.isArray(artifact.dimension_labels) ? artifact.dimension_labels.map(d => ({
          id: d?.id || '',
          label: d?.label || ''
        })) : [],
        initiatives: Array.isArray(artifact.initiatives) ? artifact.initiatives.map(init => ({
          id: init?.id || undefined,
          name: init?.name || '',
          subtitle: init?.subtitle || undefined,
          placements: Array.isArray(init?.placements) ? init.placements.map(p => ({
            lane_id: p?.lane_id || '',
            title: p?.title || '',
            subtitle: p?.subtitle || undefined,
            tags: Array.isArray(p?.tags) ? p.tags : undefined,
            footer: p?.footer || undefined,
            accent_tone: p?.accent_tone || undefined
          })) : (Array.isArray(init?.dimensions) ? init.dimensions.map(d => ({
            lane_id: d?.label || '',
            title: d?.value || ''
          })) : [])
        })) : [],
        ...coverage
      }
    }
    if (type === 'risk_register') {
      return {
        type: 'risk_register',
        risk_header: artifact.risk_header || artifact.table_header || '',
        show_mitigation: artifact.show_mitigation !== false,
        risks: Array.isArray(artifact.risks) ? artifact.risks.map(r => ({
          id: r?.id || undefined,
          title: r?.title || '',
          detail: r?.detail || r?.description || '',
          severity: r?.severity || 'medium',
          owner: r?.owner || undefined,
          status: r?.status || undefined,
          likelihood: r?.likelihood != null ? r.likelihood : undefined,
          impact: r?.impact != null ? r.impact : undefined,
          owner_tag: r?.owner_tag || undefined,
          status_tag: r?.status_tag || undefined
        })) : [],
        ...coverage
      }
    }
    if (type === 'profile_card_set') {
      return {
        type: 'profile_card_set',
        profile_header: artifact.profile_header || artifact.artifact_header_text || '',
        layout_direction: artifact.layout_direction || 'horizontal',
        profiles: Array.isArray(artifact.profiles) ? artifact.profiles.map(p => ({
          id: p?.id || undefined,
          entity_name: p?.entity_name || '',
          subtitle: p?.subtitle || undefined,
          badge_text: p?.badge_text || undefined,
          secondary_items: Array.isArray(p?.secondary_items) ? p.secondary_items.map(item => ({
            label: item?.label || '',
            value: item?.value || '',
            representation_type: item?.representation_type || 'text',
            sentiment: item?.sentiment || undefined
          })) : (Array.isArray(p?.attributes) ? p.attributes.map(a => ({
            label: a?.key || '',
            value: a?.value || '',
            representation_type: 'text',
            sentiment: a?.sentiment || undefined
          })) : [])
        })) : [],
        ...coverage
      }
    }
    return { type: artifact.type || '' }
  }
  return {
    slide_number: slide.slide_number,
    slide_type: slide.slide_type,
    narrative_role: slide.narrative_role || '',
    slide_archetype: slide.slide_archetype,
    selected_layout_name: slide.selected_layout_name || '',
    title: slide.title || '',
    subtitle: slide.subtitle || '',
    key_message: slide.key_message || '',
    zones: (slide.zones || []).map(zone => {
      const splitHint = Array.isArray(zone.artifact_split_hint)
        ? zone.artifact_split_hint
        : (Array.isArray((zone.layout_hint || {}).split_hint) ? (zone.layout_hint || {}).split_hint : null)
      const arrangement = zone.artifact_arrangement || (zone.layout_hint || {}).artifact_arrangement || null
      const split = zone.zone_split || ((zone.layout_hint || {}).split || 'full')
      return {
        zone_id: zone.zone_id,
        zone_role: zone.zone_role,
        message_objective: zone.message_objective,
        narrative_weight: zone.narrative_weight,
        artifacts: (zone.artifacts || []).map(pruneArtifactForOutput),
        zone_split: split,
        artifact_arrangement: arrangement,
        artifact_split_hint: splitHint
      }
    }),
    speaker_note: slide.speaker_note || '',
    _was_repaired: slide._was_repaired || false
  }
}

// ═══════════════════════════════════════════════════════════════════════════════
// REPAIR
// ═══════════════════════════════════════════════════════════════════════════════

async function repairSlide(slide, brief, contentB64, layoutNames) {
  console.log('Agent 4 — repairing slide', slide.slide_number, ':', slide.title)

  const briefSummary = buildBriefSummaryForAgent4(brief)
  const prompt = `This slide has missing or invalid artifact content. Fix every zone with specific data from the source document.

CONTEXT:
Document type:  ${briefSummary.document_type || '—'}
Key messages:   ${briefSummary.key_messages || '—'}
Key data:       ${briefSummary.key_data_points || '—'}

SLIDE TO FIX:
${JSON.stringify(slide, null, 2)}

Fix rules:
- Replace all placeholder or empty content with real, specific data
- Keep the same zones[] and artifact types — only fill in the content
- Do NOT change the slide structure to fit slide_archetype; slide_archetype is metadata only
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
- Charts: 3+ categories, matching values, no all-zeros; ensure chart_header is set
- insight_text: specific points with data; ensure insight_header is set
- Workflows: fully populated nodes and connections; ensure workflow_header is set
- Enforce workflow restrictions:
  - process_flow: left_to_right only, >=4 nodes, full-width zone, >=50% height
  - hierarchy: top_down_branching only, >=3 levels, >=50% width, full content height
  - decomposition: left_to_right or top_to_bottom / top_down_branching only; if >3 nodes it must own full width (left_to_right) or full height (vertical)
  - timeline: left_to_right only, >=4 nodes, full-width zone, >=50% height
  - information_flow: do not use
- Tables: ensure table_header is set
- Matrix: fully populate x_axis, y_axis, all 4 quadrants, and points; ensure matrix_header is set
- Driver_tree: fully populate root and branches; ensure tree_header is set
- Prioritization: fully populate ranked action items; ensure priority_header is set
- If matrix / driver_tree / prioritization is present, keep it only in the PRIMARY zone and pair it only with insight_text
- If a slide uses matrix / driver_tree / prioritization, do NOT add cards, chart, workflow, or table anywhere else on that slide
- selected_layout_name: choose from available layouts; set to "" if none available
- zone_split / artifact_arrangement / artifact_coverage_hint: ${layoutNames && layoutNames.length >= 5 ? 'set zone_split="full" for all zones; artifact arrangement only when a zone has 2 artifacts' : 'must be explicit for scratch composition; use artifact_coverage_hint on each artifact when a zone has 2+ artifacts'}
- layout_hint.split: ${layoutNames && layoutNames.length >= 5 ? 'set to "full" (Agent 5 uses selected_layout_name for positioning)' : 'mirror zone_split into layout_hint.split for compatibility'}
- In scratch composition, cards with 1–2 items must stay compact and must not occupy a dominant tall zone
- Unless a cards artifact has 8+ cards, no individual card may imply more than ~15% of total slide area
- Return ONLY a single JSON object for this one slide`

  const messages = [{
    role: 'user',
    content: [
      { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: contentB64 } },
      { type: 'text', text: prompt }
    ]
  }]

  const raw     = await callClaude(AGENT4_SYSTEM, messages, 2000)
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


// ═══════════════════════════════════════════════════════════════════════════════
// MAIN RUNNER
// ═══════════════════════════════════════════════════════════════════════════════

async function runAgent4(state) {
  const brief      = state.outline
  const contentB64 = state.contentB64
  const brand      = state.brandRulebook || {}

  // Use pre-filtered content_layout_names from Agent 2 when available.
  // This excludes title, section-header, divider, blank, and thank-you layouts
  // so the "5+ layouts → use layout mode" threshold counts only usable content layouts.
  const _NON_CONTENT_TYPES = new Set(['title', 'sechead', 'blank'])
  const _isNonContent = (l) => {
    const t = (l.type || '').toLowerCase()
    const n = (l.name || l.layout_name || '').toLowerCase()
    return _NON_CONTENT_TYPES.has(t) ||
      /^blank$/i.test(n) ||
      /thank[\s_-]*you|end[\s_-]*slide|closing[\s_-]*slide|section[\s_-]*header|^section$|divider/i.test(n)
  }

  const layoutNames = brand.content_layout_names && brand.content_layout_names.length > 0
    ? brand.content_layout_names.filter(n => !_isNonContent({ name: n }))
    : (brand.layout_blueprints || brand.slide_layouts || [])
        .filter(l => !_isNonContent(l))
        .map(l => l.name || l.layout_name || '').filter(Boolean)

  const totalLayouts = (brand.layout_blueprints || brand.slide_layouts || []).length

  // buildSlidePlan returns the full deck from Agent 3 (structural + content slides)
  const slidePlan      = buildSlidePlan(brief)
  const contentPlan    = slidePlan.filter(s => s.slide_type === 'content')
  const structuralPlan = slidePlan.filter(s => s.slide_type !== 'content')
  const contentCount   = contentPlan.length

  console.log('Agent 4 starting — content slides:', contentCount, '| structural slides from Agent 3:', structuralPlan.length)
  console.log('  Brand layouts total:', totalLayouts, '| content layouts:', layoutNames.length,
    layoutNames.length >= 5 ? '→ layout mode (Agent 4 selects per slide)' : '→ zone-split mode')

  // Structural slides (title, divider, thank_you) come pre-defined from Agent 3.
  // Pass them through directly — no zone enrichment needed.
  let allSlides = structuralPlan.map(plan => normaliseSlide({
    slide_number: plan.slide_number,
    slide_type:   plan.slide_type,
    title:        plan.slide_title_draft || '',
    key_message:  plan.strategic_objective || ''
  }, plan))

  let summaryCardRegistry = []  // built from summary slide, threaded through all batches

  // Batch size capped at 3 content slides to reduce model overload.
  // Each batch re-sends the source PDF, so we pause 65 s between batches to reset the window.
  const BATCH_SIZE = 3
  const batches = []
  for (let i = 0; i < contentPlan.length; i += BATCH_SIZE) batches.push(contentPlan.slice(i, i + BATCH_SIZE))

  for (let b = 0; b < batches.length; b++) {
    if (b > 0) {
      console.log('Agent 4 — rate-limit pause 65 s before batch', b + 1)
      await new Promise(r => setTimeout(r, 65000))
    }
    const batch  = batches[b]
    const result = await writeSlideBatch(batch, brief, contentB64, b + 1, layoutNames, summaryCardRegistry)

    if (!result) {
      batch.forEach(plan => allSlides.push(normaliseSlide({}, plan)))
    } else {
      result.forEach(s => {
        // Content slide — match to plan entry by slide_number
        const plan = batch.find(p => p.slide_number === s.slide_number) || batch[0]
        const normalised = normaliseSlide(s, plan)
        allSlides.push(normalised)

        // If this is the summary slide, extract cards for deduplication in subsequent batches
        if (normalised.narrative_role === 'summary' || (normalised.zones || []).some(z => z.zone_role === 'summary')) {
          const summaryCards = (normalised.zones || [])
            .flatMap(z => z.artifacts || [])
            .filter(a => a.type === 'cards')
            .flatMap(a => a.cards || [])
            .filter(c => c.title && c.subtitle)
            .map(c => ({ title: String(c.title).trim(), value: String(c.subtitle).trim() }))
          if (summaryCards.length > 0) {
            summaryCardRegistry = summaryCards
            console.log('Agent 4 — summary card registry built:', summaryCardRegistry.map(c => `"${c.title}: ${c.value}"`).join(', '))
          }
        }
      })

      // Fallback: if any content plan entry produced no output, fill with blank
      batch.forEach(plan => {
        if (!allSlides.find(s => s.slide_number === plan.slide_number && s.slide_type === 'content')) {
          allSlides.push(normaliseSlide({}, plan))
        }
      })
    }

  }

  // Sort all slides by slide_number — structural + content batches may arrive out of order
  allSlides.sort((a, b) => (a.slide_number || 0) - (b.slide_number || 0))

  // Layout-mode enforcement: when 5+ content layouts exist every content slide must
  // have selected_layout_name set.  Claude sometimes misses this — fill gaps here.
  const hasLayouts = layoutNames.length >= 5
  if (hasLayouts) {
    allSlides = allSlides.map(s => {
      if (s.slide_type === 'content' && !hasCompatibleLayout(s, layoutNames)) {
        console.log('  No compatible layout for slide', s.slide_number, '→ using scratch splits')
        return assignScratchSplits(s)
      }
      if (s.slide_type === 'content' && (!s.selected_layout_name || _isNonContent({ name: s.selected_layout_name }) || layoutConflictsWithSlide(s, s.selected_layout_name))) {
        const assigned = pickBestLayout(s, layoutNames)
        console.log('  Auto-assigned layout for slide', s.slide_number, '→', assigned)
        return { ...s, selected_layout_name: assigned }
      }
      return s
    })
  } else {
    allSlides = allSlides.map(assignScratchSplits)
  }

  // Validate and repair
  const failed = allSlides.filter(s => hasPlaceholderContent(s))
  console.log('  Slides needing repair:', failed.length)

  // Repair in groups of 2 — each repair re-sends the PDF so we observe the same
  // 30k TPM rate limit as the main batches.  Pause 65 s between groups.
  const REPAIR_GROUP = 1
  for (let ri = 0; ri < failed.length; ri += REPAIR_GROUP) {
    if (ri > 0) {
      console.log('Agent 4 — rate-limit pause 65 s before repair group', Math.floor(ri / REPAIR_GROUP) + 1)
      await new Promise(r => setTimeout(r, 65000))
    }
    const group = failed.slice(ri, ri + REPAIR_GROUP)
    for (const slide of group) {
      const repaired = await repairSlide(slide, brief, contentB64, layoutNames)
      const idx = allSlides.findIndex(s => s.slide_number === slide.slide_number)
      if (repaired && idx >= 0) {
        const ns = normaliseSlide(repaired, slidePlan.find(p => p.slide_number === slide.slide_number) || slidePlan[0] || {})
        ns._was_repaired = true  // signals Agent 5 to process this slide solo to avoid token overflow
        allSlides[idx] = ns
        console.log('  Repaired slide', slide.slide_number)
      } else if (idx >= 0) {
        // Repair failed — mark anyway so Agent 5 treats it with extra care
        allSlides[idx] = { ...allSlides[idx], _was_repaired: true }
      }
    }
  }

  // Second enforcement pass: repaired slides may still be missing selected_layout_name
  if (hasLayouts) {
    allSlides = allSlides.map(s => {
      if (s.slide_type === 'content' && !hasCompatibleLayout(s, layoutNames)) {
        console.log('  Post-repair: no compatible layout for slide', s.slide_number, '→ using scratch splits')
        return assignScratchSplits(s)
      }
      if (s.slide_type === 'content' && (!s.selected_layout_name || _isNonContent({ name: s.selected_layout_name }) || layoutConflictsWithSlide(s, s.selected_layout_name))) {
        const assigned = pickBestLayout(s, layoutNames)
        console.log('  Post-repair layout assignment for slide', s.slide_number, '→', assigned)
        return { ...s, selected_layout_name: assigned }
      }
      return s
    })
  } else {
    allSlides = allSlides.map(assignScratchSplits)
  }

  // Summary log
  const archetypes = {}
  const artifactTypes = {}
  let totalZones = 0
  let totalArtifacts = 0

  allSlides.forEach(s => {
    archetypes[s.slide_archetype] = (archetypes[s.slide_archetype] || 0) + 1
    ;(s.zones || []).forEach(z => {
      totalZones++
      ;(z.artifacts || []).forEach(a => {
        totalArtifacts++
        artifactTypes[a.type] = (artifactTypes[a.type] || 0) + 1
      })
    })
  })

  console.log('Agent 4 complete')
  console.log('  Total slides:', allSlides.length)
  console.log('  Total zones:', totalZones)
  console.log('  Total artifacts:', totalArtifacts)
  console.log('  Archetypes:', JSON.stringify(archetypes))
  console.log('  Artifact types:', JSON.stringify(artifactTypes))
  console.log('  Placeholder remaining:', allSlides.filter(s => hasPlaceholderContent(s)).length)

  return allSlides.map(pruneAgent4SlideForOutput)
}
