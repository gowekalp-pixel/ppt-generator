// AGENT 4 — OUTPUT SCHEMA  [CHANGE-MANAGEMENT REPLICA]
// ─── Managed copy — production source: public/prompts/agent4/A1-ArtifactSchema.js
//     The /api/inject-artifact endpoint appends new artifact JSON schemas to THIS file.
//     It never touches the production source.
// Required slide/zone fields and all artifact schemas.

const _A4_OUTPUT_SCHEMA = `OUTPUT OBJECT — REQUIRED FIELDS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Each slide object must contain EXACTLY these top-level fields:

{
  "slide_number": number,
  "slide_type": "title" | "divider" | "content" | "thank_you",
  "narrative_role": "string — carry forward from Agent 3 plan; empty string for structural slides",
  "selected_layout_name": "string — empty string for structural slides",
  "title": "string",
  "subtitle": "string",
  "key_message": "string",
  "zones": [ ... ],   ← always [] for title / divider / thank_you
  "speaker_note": "string — content slides only: consolidate Phase 2 speaker overflow here
    (qualifications, secondary data, assumptions). 1–4 sentences. Empty string for
    structural slides and when there is no meaningful overflow."
}

SLIDE TYPE RULES — CONTENT SLIDE TITLE

  Content slide title must be insight-led — never a generic topic label
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
    "artifact_header": "2–4 word specific label that names the implication — e.g. 'Risk Implication', 'Growth Opportunity', 'Action Required', 'Portfolio Risk' — never use generic labels like 'So What' or 'Key Insight'",
    "points": ["specific insight with data"],          ← STANDARD mode (flat list)
    "groups": [                                        ← GROUPED mode (thematic sections)
      { "header": "2–4 word label", "bullets": ["crisp point with data"] }
    ],
    "sentiment": "positive" | "warning" | "neutral"
  }
  Use either points[] OR groups[] — never both.

  Content integrity: preserve ALL facts, numbers, names, percentages exactly from source.
  Do NOT drop data-bearing facts. DO compress prose around the facts.

chart:
  {
    "type": "chart",
    "chart_type": "bar" | "line" | "area" | "pie" | "donut" | "waterfall" | "clustered_bar" | "horizontal_bar" | "combo" | "group_pie",
    "chart_title": "",                                 ← leave empty when layout has header placeholder
    "artifact_header": "the one-line insight the chart proves",
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
  chart_title is rendered INSIDE the plot area; artifact_header is the zone heading.
  Use the base schema above for: bar, line, area, pie, waterfall, clustered_bar, horizontal_bar.
  Use the variant schemas below for donut, combo, and group_pie.

donut chart — use this schema instead when chart_type is "donut":
  {
    "type": "chart",
    "chart_type": "donut",
    "chart_title": "",
    "artifact_header": "the one-line insight the donut proves",
    "categories": ["Segment A", "Segment B"],          ← max 5 segments; HARD REJECT if > 5 → convert to horizontal_bar
    "series": [
      { "name": "string", "values": [number], "unit": "percent" }
    ],
    "center_label": "string — main callout in the donut centre; typically the total or anchor KPI",
    "show_data_labels": true,
    "show_legend": true
  }
  donut rules:
  - Use donut over pie when a centre callout (total, key metric, or anchor KPI) materially improves the message.
  - Do NOT use x_label, y_label, dual_axis, or secondary_series for donut.
  - center_label is mandatory for donut — if there is no meaningful centre value, use pie instead.
  - HARD MAX 5 segments; if > 5 → automatically convert to horizontal_bar.

combo chart — use this schema instead when chart_type is "combo":
  {
    "type": "chart",
    "chart_type": "combo",
    "chart_title": "",
    "artifact_header": "the one-line insight the dual-axis chart proves",
    "x_label": "string",
    "y_label": "string — left axis label (primary series unit)",
    "y2_label": "string — right axis label (secondary series unit)",
    "categories": ["string"],
    "series": [
      { "name": "string", "values": [number], "unit": "string" }  ← rendered as BARS on left axis
    ],
    "dual_axis": true,                                 ← always true for combo
    "secondary_series": [
      { "name": "string", "values": [number], "unit": "string" }  ← rendered as LINE on right axis
    ],
    "show_data_labels": true,
    "show_legend": true
  }
  combo rules:
  - dual_axis MUST be true — a combo chart without dual_axis is invalid.
  - series[] = primary data rendered as bars (left Y axis).
  - secondary_series[] = secondary data rendered as a line (right Y axis).
  - The two series MUST have different units (e.g. ₹ revenue on bars, % margin on line).
  - If both series share the same unit → reject combo → use clustered_bar instead.
  - y2_label is mandatory — always populate the right axis label.
  
group_pie chart — use this schema instead when chart_type is "group_pie":
  {
    "type": "chart",
    "chart_type": "group_pie",
    "chart_title": "",
    "artifact_header": "the one-line insight the group proves",
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
  - categories[] = the shared slice breakdown for all pies
  - series[] = one entry per entity; series[i].name becomes the label BELOW pie i
  - series[i].series_total (optional): pre-formatted absolute total displayed as a sub-label
    directly below the entity name under each pie — use when entities differ materially in
    absolute scale and the audience needs both composition and magnitude. Agent 4 must compute
    and format this value from source data (e.g. "₹39.7L", "23%", "$4.2M") — do not leave
    it blank or delegate calculation to Agent 5. Omit the field entirely if not meaningful.
  - values[] length must equal categories[] length for every series
  - Do NOT use x_label, y_label, dual_axis, secondary_series for group_pie
  - Legend is always shared and rendered once above the group, below the artifact_header

cards:
  {
    "type": "cards",
    "artifact_header": "string — one-line framing of the card set; required for primary zones",
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
    "artifact_header": "string — one-line insight the workflow proves",
    "nodes": [
      {
        "id": "n1",
        "node_label": "string",
        "primary_message": "string",
        "secondary_message": "string",
        "level": 1
      }
    ],
    "connections": [ { "from": "n1", "to": "n2", "type": "arrow" } ]
  }
  Node copy limits:
  - node_label: 2–5 words (hard max 18 chars)
  - primary_message: 4-10 words
  - secondary_message: 8–18 words
  Workflow node usage by type:
  - process_flow / timeline:
    - node_label = step / phase name
    - primary_message = short top message, KPI, or milestone cue
    - secondary_message = optional supporting note below
  - hierarchy / decomposition:
    - node_label is mandatory
    - primary_message and secondary_message are optional and should be used only if they improve clarity
  Max 6 nodes. Max 8 connections. No crossing connections.

table:
  {
    "type": "table",
    "artifact_header": "string — insight the table proves, not a topic label",
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
    "artifact_header": "string",
    "x_axis": { "label": "string — axis name (metadata only, not displayed)", "low_label": "string — self-contained: embed axis name + value (e.g. 'Low AOV  <₹2,000')", "high_label": "string — self-contained: embed axis name + value (e.g. 'High AOV  ₹3,500+')" },
    "y_axis": { "label": "string — axis name (metadata only, not displayed)", "low_label": "string — self-contained: embed axis name + value (e.g. 'Low Volume  <30k orders')", "high_label": "string — self-contained: embed axis name + value (e.g. 'High Volume  >50k orders')" },
    "quadrants": [
      {
        "id": "q1",
        "title": "string",
        "primary_message": "string — one-line axis descriptor (e.g. 'High ASP · low scale')",
        "tone": "positive" | "negative" | "neutral"
      }
    ],
    "points": [
      {
        "label": "string — full display name (e.g. 'Own Store')",
        "short_label": "string — 2-3 char abbreviation for the inner dot (e.g. 'OS')",
        "quadrant_id": "q1|q2|q3|q4 — which quadrant this point belongs to",
        "x": "number 0–100 — precise horizontal position as % of full grid width (e.g. 22, 68, 81)",
        "y": "number 0–100 — precise vertical position as % of full grid height (e.g. 15, 57, 83)",
        "emphasis": "high" | "medium" | "low"
      }
    ]
  }
  Max 12 points. Must define both axes and all 4 quadrants.
  Axis label rule:
  - low_label and high_label are the ONLY axis text rendered on the slide — make them self-contained.
  - Each must include the axis dimension name AND the threshold value (e.g. "Low AOV  <₹2k", "High Orders  >50k").
  - label is metadata only (not rendered); keep it concise (e.g. "Average Order Value").
  - Do NOT use generic labels like "Low" / "High" alone — the viewer must understand the axis from these alone.
  Quadrant usage:
  - id: q1=top-left, q2=top-right, q3=bottom-left, q4=bottom-right
  - title = quadrant strategic label (e.g. "Scale with margin")
  - primary_message = one-line axis descriptor showing WHERE in the matrix (e.g. "High ASP · high scale")
  - tone = "positive" (favourable quadrant), "negative" (unfavourable), "neutral" (monitor/mixed)
  Point placement — TWO-STEP mandatory process:
  STEP 1 — QUADRANT ASSIGNMENT (decide before any coordinates):
    - For each point, reason explicitly which quadrant (q1/q2/q3/q4) it belongs to based on the data.
    - Consider spread: if all points cluster in one quadrant the matrix is misleading — adjust axes or rethink.
    - Assign quadrant_id. The point MUST land inside that quadrant's half of the grid.
  STEP 2 — COORDINATE ASSIGNMENT (within the assigned quadrant):
    - q1 (top-left):    x ∈ 5–45,  y ∈ 55–95
    - q2 (top-right):   x ∈ 55–95, y ∈ 55–95
    - q3 (bottom-left): x ∈ 5–45,  y ∈ 5–45
    - q4 (bottom-right):x ∈ 55–95, y ∈ 5–45
    - Vary x and y within those ranges to reflect the point's relative strength on each axis.
    - Never place two points at the same (x, y) — offset by ≥8 units to avoid overlap.
    - Provide exact integer values (e.g. x=22, y=71) so Agent 5 can render without ambiguity.
  Label and short_label are the only display elements on the dot — do NOT add primary_message or secondary_message to points.
  Per-point insights belong in the paired insight_text artifact (see pairing rule below).
  - emphasis = "high" for the largest dot (typically the dominant or most critical item), "medium" default, "low" for minor items

driver_tree:
  {
    "type": "driver_tree",
    "artifact_header": "string",
    "root": {
      "node_label": "string",
      "primary_message": "string",
      "secondary_message": "string"
    },
    "branches": [
      {
        "node_label": "string",
        "primary_message": "string",
        "secondary_message": "string",
        "children": [
          {
            "node_label": "string",
            "primary_message": "string",
            "secondary_message": "string"
          }
        ]
      }
    ]
  }
  Max 3 levels. Max 6–8 nodes.
  Node usage:
  - node_label = driver name
  - primary_message = key value / main takeaway
  - secondary_message = optional supporting note
  Root = outcome; branches = main drivers; children = sub-drivers.
  NOT a process — do not use when showing sequence or steps.

prioritization:
  {
    "type": "prioritization",
    "artifact_header": "string",
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
    "artifact_header": "string — the one-line insight the ranking proves",
    "annotation_style": "inline" | "trailing",
    "scale_UL": number — upper limit of bar scale (e.g. 100 for percentages); omit to auto-scale to max row value,
    "column_headers": [
      {"id": "col1", "value": "string — column header label", "display_type": "text" | "bar" | "tag/pill" }
    ],
    "rows": [
      {
        "row_id": number — 1-based row index,
        "row_focus": "Y" | "N" — "Y" visually highlights this row; use sparingly (1–2 rows max),
        "cells": [
          {"col_id": "col1", "value": "string — cell value (numeric string for bar columns)"}
        ]
      }
    ]
  }
  column_headers rules:
    "text"   — plain text column (entity name or annotation). First text col = left label; others = trailing text.
    "bar"    — renders as a proportional horizontal bar. 1–3 bar columns allowed. Each cell value must be a numeric string.
              Per-column scale: add "scale_UL": number (upper limit) on the column header to fix the bar ceiling.
              Optionally add "scale_LL": number (lower limit) to override the auto lower limit (default: 50% of the column minimum).
              Bar fill fraction = (value − scale_LL) / (scale_UL − scale_LL) — bars are always normalised for visual differentiation.
              Example: on-time rate → scale_UL: 100 (scale_LL auto-computed as ~48.5 for a 97–100% dataset).
              Example: avg charge ₹0–₹15 → scale_UL: 15, scale_LL: 0.
    "normal" — secondary display value, right-aligned (e.g. a formatted metric like "₹8.2" or "16.5 Days").
  COLUMN PAIRING RULE: Every "bar" column MUST be immediately followed in column_headers by a "normal" column with an empty header ("value": "") — that normal column holds the bar's numeric value as readable text. Never place a "bar" column as the last column or adjacent to another "bar" column.
  scale_UL: set per bar column to fix the bar ceiling (e.g. 100 for percentages, realistic max for currency/count columns). Do NOT set a single global scale_UL at artifact level.
  scale_LL: optional override. If omitted, the renderer auto-computes it as 50% of the column minimum, ensuring bars always span a meaningful portion of the track width.
  row_focus "Y": highlighted row — never infer from rank, set explicitly. Do NOT mark all rows "Y".
  annotation_style default: "trailing".
  SIZE RULES (match zone allocation to these):
    Columns → minimum zone width: 
    1 column with "bar" representation → minimum 50% width;
    2 columns with "bar" representation → minimum 75% width;
    3 columns with "bar" representation → 100% width (full width);
    Rows → minimum zone height: 2 rows → ≥30% slide height; each extra row adds ~11% (8 rows = 100%).
    Minimum 2 rows, maximum 8 rows.

comparison_table:
  IMPORTANT: Use ONLY the flat schema below. Do NOT use column_headers[], criteria[], options[], or any other old field names — they are deprecated and will cause validation failure.
  {
    "type": "comparison_table",
    "artifact_header": "string — the one-line insight the comparison proves",
    "columns": ["Option Label", "Criterion 1", "Criterion 2"],
    "rows": [
      {
        "is_recommended": true,
        "badge": "Recommended",
        "cells": [
          { "value": "string — option name", "icon_type": null, "subtext": null, "tone": "label" },
          { "value": "8.1%", "icon_type": null, "subtext": "3× lower than COD", "tone": "positive" },
          { "value": null, "icon_type": "check", "subtext": "Best in class", "tone": "positive" }
        ]
      }
    ]
  }
  comparison_table cell rules — each data cell (cells[1..n]) has EITHER value OR icon_type, never both:
  - value:     the metric/percentage/currency text → rendered as a colored pill (e.g. "8.1%", "₹2,547")
               set to null when the cell conveys a verdict rather than a measurement
  - icon_type: named vector icon → rendered as an icon inside a colored circle
               set to null when the cell shows a metric value
               valid values: check | cross | partial | arrow_up | arrow_down | arrow_right |
                             star | warning | diamond | chevron
  - tone:      drives ALL coloring — agent 5 applies brand colors automatically
               "positive" = green  (use when the option wins on this criterion)
               "negative" = red    (use when the option loses on this criterion)
               "neutral"  = grey   (use for factual / neither-better context)
               "label"    = cells[0] only (the option name column)
  - subtext:   optional ≤6-word annotation below the value or icon. null if not needed.

  When to use value vs icon_type:
  - Measurable metric (%, ₹, count, ratio) → use value, set icon_type: null
  - Verdict with no specific metric (good/bad/partial) → use icon_type, set value: null
  - Never put text like "Yes" / "No" / "✓" in value — use icon_type: "check" / "cross" instead
  NEVER use plain table for option-vs-criteria data.

initiative_map:
  {
    "type": "initiative_map",
    "artifact_header": "string — the one-line framing of the initiative landscape",
    "column_headers": [
      { "id": "initiative", "label": "Initiative" },
      { "id": "c1", "label": "Phase 1 — Immediate" },
      { "id": "c2", "label": "Phase 2 — Next quarter" }
    ],
    "rows": [
      {
        "id": "string",
        "initiative_name": "string — bold row label, 2–5 words",
        "initiative_subtitle": "string — OPTIONAL muted sub-label (e.g. category); omit if initiative_name is already self-explanatory. Max 5 words. Must NOT repeat any word from initiative_name.",
        "cells": [
          {
            "column_id": "string — matches column_headers[].id",
            "primary_message": "string — headline for this cell; 2–5 words max",
            "secondary_message": "string — OPTIONAL detail line; include ONLY when primary_message is ≤4 words AND secondary adds non-redundant context; max 7 words; omit otherwise",
            "tags": [
              {
                "label": "string — short pill text (e.g. city name, priority band, owner); max 2 words",
                "tone": "primary" | "secondary" | "neutral"
              }
            ],
            "cell_tone": "primary" | "secondary" | "neutral"
          }
        ]
      }
    ]
  }
  initiative_map usage:
  - each row is one initiative / workstream; columns are structured dimensions (phase, owner, KPI, status).
  - DENSITY RULE: every cell must fit in ~0.7" of height. Be ruthless with brevity.
  - primary_message: 2–5 words — the single most important fact for that cell.
  - secondary_message: OPTIONAL — only add when primary_message is ≤4 words AND secondary adds a genuinely distinct data point. Max 7 words. Never restate primary_message in different words.
  - initiative_subtitle: OPTIONAL — only add when initiative_name alone is ambiguous. Max 5 words. Must not repeat words from initiative_name.
  - tags[] are optional pills — city names, owners, priority bands, confidence labels. Max 3 tags per cell.
  - each tag carries its own tone ("primary" | "secondary" | "neutral") for individual chip colour.
  - cell_tone drives the chip fill / border palette for chips with no individual tone override.
  - CRITICAL rendering rule: when tags[] is non-empty, primary_message is suppressed by the renderer — tags ARE the primary visual signal. Put any key metric into secondary_message (≤7 words), leave primary_message "".
  - When tags[] is empty: primary_message is the headline; secondary_message is optional supporting detail.
  - Typical city-tagged cell: tags=[{label:"Bangalore",tone:"primary"},{label:"Mumbai",tone:"primary"}], primary_message="", secondary_message="9,009 orders at ₹0.26Cr".
  - Typical tag-free cell: primary_message="~₹0.3–0.4Cr revenue", secondary_message="" (omit unless distinct).
  NEVER use when rows have a rank order (use prioritization). NEVER use for process steps (use workflow).

profile_card_set:
  {
    "type": "profile_card_set",
    "artifact_header": "string — the one-line framing of the entity set",
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
    "risk_header": "string — one-line framing of the risk landscape (used as section header, NOT as artifact_header — do NOT populate artifact_header for risk_register)",
    "severity_levels": [
      {
        "id": "string — e.g. 'level_1'",
        "label": "string — severity band heading shown in the colored band (e.g. 'Critical severity — immediate action required')",
        "tone": "critical" | "high" | "medium" | "low",
        "pip_levels": "number — total number of pip blocks in the scale (e.g. 5 means filled squares out of 5). All items in this severity_level share the same pip_levels value.",
        "item_details": [
          {
            "primary_message": "string — short risk headline (bold, ≤8 words)",
            "secondary_message": "string — supporting detail line (muted, ≤18 words)",
            "tags": [
              { "value": "string — short chip label (owner, team, category; ≤2 words)", "tone": "neutral" | "positive" | "negative" | "warning" }
            ],
            "pips": [
              { "label": "string — dimension name (e.g. 'Likelihood', 'Impact')", "intensity": "number — filled blocks out of pip_levels (e.g. 3 means 3 filled out of pip_levels total)" }
            ]
          }
        ]
      }
    ]
  }
  risk_register usage:
  - Do NOT populate artifact_header — risk_register is always a single-artifact zone and uses risk_header as its own internal section header.
  - severity_levels groups items by severity band. Order from worst to best: critical → high → medium → low. Omit unused levels.
  - tone on each severity_level drives band fill, dot color, and pip fill (critical=red, high=orange, medium=amber, low=gray).
  - pip_levels: set once per severity_level (typically 3–5). All pips across items in that level use this as the total scale. Choose consistently: if the data warrants finer granularity use 5, coarser use 3.
  - primary_message: bold title for the risk/issue (≤8 words).
  - secondary_message: supporting evidence line rendered smaller below the title (≤18 words).
  - tags[]: 1–3 short pill chips per item. Each has value (display text, ≤2 words) and tone (neutral=gray, positive=green, negative=red, warning=amber).
  - pips[]: 1–3 dimension assessments (e.g. Likelihood, Impact). intensity is numeric: number of filled blocks out of pip_levels (0 = none, pip_levels = fully filled). LLM decides intensity from data.
  NEVER use plain table when severity-by-row is the primary signal.
`
