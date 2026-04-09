// AGENT 5 — ALL ARTIFACT SCHEMAS
// Contains artifact contract intro + all 13 artifact type schemas:
//   1. INSIGHT TEXT (standard and grouped mode)
//   2. CHART           3. CARDS
//   4. WORKFLOW        5. TABLE
//   6. MATRIX          7. DRIVER_TREE      8. PRIORITIZATION
//   9. STAT_BAR       10. COMPARISON_TABLE 11. INITIATIVE_MAP
//  12. PROFILE_CARD_SET  13. RISK_REGISTER
// Assembled into AGENT5_SYSTEM by agent5.js.
// Edit here when adding, changing, or removing artifact types.

const _A5_ARTIFACTS = `*********************************************************************************
ARTIFACT CONTRACT
*********************************************************************************

Every artifact must be FULLY specified. No missing fields. No placeholders.

Allowed types: insight_text | chart | stat_bar | cards | workflow | table | matrix | driver_tree | prioritization | comparison_table | initiative_map | profile_card_set | risk_register

*********************************************************************************
1. INSIGHT TEXT
*********************************************************************************

********* STEP 1: SET insight_mode *********************************************************************************************************************************************************
Inspect the Agent 4 artifact:
- Agent 4 provides "groups[]"  â†’ set insight_mode: "grouped"
- Agent 4 provides "points[]"  â†’ set insight_mode: "standard"

********* STANDARD MODE SCHEMA *********************************************************************************************************************************************************************

{
  "type": "insight_text",
  "insight_mode": "standard",
  "x": number, "y": number, "w": number, "h": number,
  "style": {
    "fill_color": "hex or null",
    "border_color": "hex or null",
    "border_width": number,
    "corner_radius": number
  },
  "heading_style": {
    "font_family": "string",
    "font_size": number,
    "font_weight": "semibold" | "bold",
    "color": "hex"
  },
  "body_style": {
    "font_family": "string",
    "font_size": number,
    "font_weight": "regular",
    "color": "hex",
    "line_spacing": number,
    "indent_inches": number,
    "list_style": "tick_cross" | "numbered" | "bullet",
    "space_before_pt": number,
    "vertical_distribution": "spread"
  },
  "header_block": null or {
    "text": "string **” the artifact_header value from Agent 4",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string", "font_size": number, "font_weight": "semibold", "color": "hex",
    "style": "underline" | "brand_fill",
    "accent_color": "hex",
    "placeholder_ref": true or false
  }
}

Standard mode styling rules:
- style.fill_color: ALWAYS null — no background box or tint; insight_text renders directly on the slide background
- style.border_color: ALWAYS null; border_width: 0 — no box, no border of any kind
- heading_style.color: brand primary color; risk/alert accent only for "Risk Alert"
- list_style: "tick_cross" for positive/negative mix; "numbered" for sequential/ranked; "bullet" for parallel
- indent_inches: 0.12**“0.18; space_before_pt: 4**“8pt
- vertical_distribution "spread": distribute points evenly **” do NOT cluster at top
- body font_size proportional to artifact height:
    h < 2.0": 9**“11pt;  h 2.0**“3.5": 11**“14pt;  h > 3.5": 14**“18pt
- Do NOT pre-shrink font **” renderer auto-fits; use upper end of range
- heading_style.font_size = body_style.font_size + 2 to 4pt

********* GROUPED MODE SCHEMA ************************************************************************************************************************************************************************

{
  "type": "insight_text",
  "insight_mode": "grouped",
  "x": number, "y": number, "w": number, "h": number,
  "style": { "fill_color": null, "border_color": null, "border_width": 0, "corner_radius": 0 },
  "heading_style": {
    "font_family": "string",
    "font_size": number,
    "font_weight": "semibold" | "bold",
    "color": "hex"
  },
  "header_block": null or { ... same as standard mode ... },

  "group_layout": "columns" | "rows",

  "group_header_style": {
    "shape": "rounded_rect" | "circle_badge",
    "fill_color": "hex",
    "text_color": "hex",
    "font_family": "string",
    "font_size": number,
    "font_weight": "bold" | "semibold",
    "corner_radius": number,
    "w": number,
    "h": number
  },

  "group_bullet_box_style": {
    "fill_color": "hex or null",
    "border_color": "hex",
    "border_width": number,
    "corner_radius": number,
    "padding": { "top": number, "right": number, "bottom": number, "left": number }
  },

  "bullet_style": {
    "font_family": "string",
    "font_size": number,
    "font_weight": "regular",
    "color": "hex",
    "line_spacing": number,
    "indent_inches": number,
    "space_before_pt": number,
    "char": "string"
  },

  "group_gap_in": number,
  "header_to_box_gap_in": number
}

********* GROUPED MODE DESIGN RULES ******************************************************************************************************************************************************

GROUP LAYOUT **” choose based on zone dimensions and group count:
- "columns": use when zone w > zone h AND group count â‰¤ 3 (groups side-by-side)
- "rows":    use when zone h â‰¥ zone w OR group count â‰¥ 3 (groups stacked vertically, header left + box right)

GROUP HEADER SHAPE **” choose based on content semantics:
- "circle_badge": ONLY when groups represent numbered priority steps (1, 2, 3**¦) or sequential phases
  - Renders as a filled circle; priority number (1-based index) shown as bold text inside
- "rounded_rect": DEFAULT for all text-label group headers
  - columns layout: spans full column width as a header bar
  - rows layout: left-side label block beside the bullet box

GROUP HEADER COLORS:
- fill_color: brand primary color (EY yellow #FFE600, or dark brand color for contrast)
- text_color: #111111 on light/yellow fills; #FFFFFF on dark fills
- Do NOT use a subtle/light fill **” headers must be visually dominant

GROUP BULLET BOX:
- fill_color: null or very light near-white tint
- border_color: light brand border **” use brand secondary light or a neutral **” NEVER heavy/dark
- border_width: 0.5**“1.0pt
- corner_radius: match group_header_style.corner_radius for visual consistency

DIMENSION CALCULATION **” derive ALL sizes from content and available area:

  Let n       = number of groups
  Let f       = bullet_style.font_size (pt)
  Let line_h  = f / 72  (inches per line, 1pt = 1/72")
  Let header_block_h = height consumed by artifact-level header_block (0 if null)

  1. group_header_style.h   **” height of the group header shape:
       rounded_rect â†’ h = max(f Ã— 1.8 / 72,  artifact.h Ã— 0.06)
       circle_badge â†’ h = max(f Ã— 2.2 / 72,  artifact.h Ã— 0.08)   [w = h, always a circle]

  2. group_header_style.w   **” width of the group header shape:
       columns layout â†’ NOT specified here; renderer uses full column width
       rows layout, rounded_rect â†’ estimate from longest group header label:
           w = (max_header_chars Ã— f Ã— 0.55 / 72) + (f Ã— 2.0 / 72)
       rows layout, circle_badge â†’ w = h (square bounding box)

  3. group_gap_in **” gap between adjacent groups:
       columns â†’ max(artifact.w Ã— 0.015, 0.05)
       rows    â†’ max(artifact.h Ã— 0.015, 0.05)

  4. header_to_box_gap_in:
       = max(f Ã— 0.5 / 72, 0.03)

  5. group_bullet_box_style.padding:
       top = bottom = max(f Ã— 0.8 / 72, 0.05)
       left = right = max(f Ã— 1.0 / 72, 0.07)

  6. bullet_style.font_size **” derive from available box area:
       columns layout:
           col_w = (artifact.w - (n-1) Ã— group_gap_in) / n
           box_h = artifact.h - header_block_h - group_header_style.h - header_to_box_gap_in
           max_bullets_in_col = max bullets across all groups
           f = floor(box_h Ã— 72 / (max_bullets_in_col Ã— 1.5))  â†’ clamp to [8, 14]
       rows layout:
           total_row_h = artifact.h - header_block_h - (n-1) Ã— group_gap_in
           max_bullets_in_row = max bullets across all groups
           min_row_h = total_row_h Ã— (max_bullets_in_row / total_bullets)
           box_w = artifact.w - group_header_style.w - header_to_box_gap_in
           f = floor(min_row_h Ã— 72 / (max_bullets_in_row Ã— 1.5))  â†’ clamp to [8, 14]

  7. space_before_pt = max(f Ã— 0.4, 2)   (scales with font **” tighter than standard mode)
     indent_inches   = max(f Ã— 1.0 / 72, 0.08)

GEOMETRY **” columns layout (renderer uses these formulas, NOT hardcoded values):
- col_w[i]   = (artifact.w - (n-1) Ã— group_gap_in) / n          [equal width per column]
- header_h   = group_header_style.h                               [same for all columns]
- box_h[i]   = artifact.h - header_block_h - header_h - header_to_box_gap_in  [same for all columns]

GEOMETRY **” rows layout (renderer uses these formulas):
- total_row_h = artifact.h - header_block_h - (n-1) Ã— group_gap_in
- row_h[i]   = total_row_h Ã— (bullets[i].length / total_bullets)  [PROPORTIONAL to bullet count **” NOT equal]
               minimum row_h[i] = group_header_style.h + header_to_box_gap_in + (1 line of text)
- header_w   = group_header_style.w                               [same for all rows]
- box_w[i]   = artifact.w - header_w - header_to_box_gap_in      [same for all rows]


*********************************************************************************
2. CHART
*********************************************************************************

{
  "type": "chart",
  "x": number, "y": number, "w": number, "h": number,
  "dual_axis": false,
  "secondary_series": [],
  "chart_style": {
    "title_font_family": "string",
    "title_font_size": number,
    "axis_font_family": "string",
    "axis_font_size": number,
    "label_font_family": "string",
    "label_font_size": number,
    "title_color": "hex",
    "axis_color": "hex",
    "gridline_color": "hex",
    "legend_font_family": "string",
    "legend_font_size": number,
    "legend_color": "hex",
    "show_gridlines": false,
    "show_border": false,
    "border_color": null,
    "background_color": null
  },
  "series_style": [
    {
      "series_name": "string",
      "fill_color": "hex",
      "line_color": "hex",
      "line_width": number,
      "marker": "none" | "circle" | "square",
      "data_label_color": "hex"
    }
  ],
  "header_block": null or {
    "text": "string **” the artifact_header value from Agent 4",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string", "font_size": number, "font_weight": "semibold", "color": "hex",
    "style": "underline" | "brand_fill",
    "accent_color": "hex"
  }
}

Chart rules:
- use brand chart_palette in sequence for series
- primary series uses primary brand color
- data_label_color must use a brand text color token, preferably body_color; do not hardcode white
- assume Agent 6 renders chart data labels at Outside End, so choose label colors for readability on the slide background
- minimum axis font size: 8pt
- if chart + table in zone: chart takes 60**“75% of zone width
- show_gridlines must always be false; no chart should display gridlines

DUAL AXIS **” MANDATORY:
- If two or more series have DIFFERENT units (e.g. one is a count/number, another is â‚¹/currency/%, etc.)
  you MUST set dual_axis: true and list the secondary-axis series names in secondary_series[].
  NEVER plot metrics with different units on the same Y axis.
  The renderer will automatically display primary series as bars and secondary series as a line on the right Y axis.

LAYOUT SIZE HINTS based on category count:
These rules apply in BOTH scratch mode (free positioning) AND layout mode (template layouts).

SCRATCH MODE **” set zone dimensions directly:
- Column/vertical bar chart with > 6 categories: set zone w â‰¥ 7" (wide/horizontal stretch)
- Horizontal bar chart with > 6 categories: set zone h â‰¥ 5" (tall/vertical stretch)

LAYOUT MODE **” select the right layout from layout_map:
- layout_map[name].content_areas is an array of placeholders sorted leftâ†’right, topâ†’bottom.
  Each entry has w_in and h_in **” the actual rendered size of that placeholder in the template.
  content_areas[0] is always the primary (largest/first) content slot.
- Column/vertical bar chart with > 6 categories: scan layout_map for the layout whose
  content_areas[0].w_in is largest (â‰¥ 7" preferred). Set selected_layout_name to that layout.
- Horizontal bar chart with > 6 categories: scan layout_map for the layout whose
  content_areas[0].h_in is largest (â‰¥ 5" preferred). Set selected_layout_name to that layout.
- This overrides Agent 4's selected_layout_name when a better-fitting layout exists **”
  chart readability takes priority over the default layout choice.
- If no template layout meets the threshold, fall back to the widest/tallest available and
  note the constraint; do NOT invent coordinates that exceed the placeholder bounds.

HEADER DEDUPLICATION:
- If the zone has a header_ph_idx (layout mode), the artifact heading is written to the layout's
  header placeholder automatically. In that case set chart_title: "" **” do NOT repeat the same
  text as both an internal chart title and a header_block / layout header.
- Only use chart_title for a subtitle-style annotation inside the chart plot area when it adds
  information beyond what the zone header already says.

LEGEND POSITION (computed by renderer from chart-to-slide ratio **” do not specify in spec):
- chart.w > 60% of slide width                       â†’ legend RIGHT
- chart.h > 60% of slide height (and width â‰¤ 60%)   â†’ legend TOP
- All other cases: pie chart â†’ legend RIGHT; all other chart types â†’ legend TOP
- Legend font size is automatically capped at chart_header font size **” do not set legend_font_size above heading_style.font_size

PIE CHART CRITICAL: a pie has ONE series but MULTIPLE segments (one per category).
  The renderer colors each segment from series_style[i]. You MUST output one series_style
  entry PER DATA POINT, each with a UNIQUE fill_color drawn from chart_palette in order.
  Example for 3 categories: series_style = [
    { "series_name": "cat1", "fill_color": chart_palette[0], ... },
    { "series_name": "cat2", "fill_color": chart_palette[1], ... },
    { "series_name": "cat3", "fill_color": chart_palette[2], ... }
  ]
  NEVER repeat the same color for two segments. If chart_palette has fewer colors than
  segments, cycle through it (palette[i % palette.length]).

GROUP PIE CHART CRITICAL: chart_type = "group_pie" renders N independent pie charts in a group.
  Data model:
  - categories[] = the shared slice labels (same breakdown for EVERY pie) **” max 7
  - series[]     = one entry per entity/pie; series[i].name is the entity label shown BELOW pie i
  - series[i].values[] must have the same length as categories[]
  - series[i].unit should be "percent" (values sum to ~100)

  series_style: EXACTLY LIKE A SINGLE PIE **” one entry per SLICE (category), NOT per entity.
  The same colors are shared across every pie in the group. Output series_style.length === categories.length.
  Example for 3 slices, 3 entity pies:
    categories: ["Standard", "SMA-0", "Substandard"]
    series: [
      { "name": "Dairy",    "values": [73, 9, 11], "unit": "percent" },
      { "name": "Kirana",   "values": [43, 11, 36], "unit": "percent" },
      { "name": "Hardware", "values": [56, 9, 19], "unit": "percent" }
    ]
    series_style: [
      { "series_name": "Standard",    "fill_color": chart_palette[0], "data_label_color": body_color, "data_label_size": 9 },
      { "series_name": "SMA-0",       "fill_color": chart_palette[1], "data_label_color": body_color, "data_label_size": 9 },
      { "series_name": "Substandard", "fill_color": chart_palette[2], "data_label_color": body_color, "data_label_size": 9 }
    ]

  Legend: ALWAYS "top" **” one shared legend listing the SLICES (categories), rendered once above all pies.
    The legend entries come from categories[], colored by series_style[i].fill_color.
  Entity label: series[i].name is rendered BELOW each pie, center-aligned, in the brand accent color.
    Do NOT include entity names in the chart legend **” they appear as labels under each pie.
  series_total sub-label: if series[i].series_total is present and non-empty, render it as a
    second line directly below the entity name, center-aligned under that pie.
    Style: same horizontal alignment as the entity name; font size 1**“2pt smaller than the
    entity name; color: brand body_color or secondary text color (not accent).
    If series[i].series_total is absent or empty, render only the entity name **” no blank line.

  Layout size hints for group_pie:
  - group_pie with 5**“8 pies: set zone w â‰¥ 9" (needs near-full slide width)
  - group_pie with 2**“4 pies: zone w â‰¥ 5" (â‰¥ 50% of slide width)
  - show_legend must be true (single shared legend); the renderer places it at the top.

*********************************************************************************
CHART MICRO-LAYOUT OWNERSHIP:
- You must decide legend_position, data_label_size, and category_label_rotation in the spec
- You must decide all series colors, line widths, marker choices, and label colors
- Do NOT leave chart readability choices for the renderer

3. CARDS
*********************************************************************************

{
  "type": "cards",
  "cards_layout": "row" | "column" | "grid",
  "container": { "x": number, "y": number, "w": number, "h": number },
  "card_style": {
    "fill_color": "hex",
    "border_color": "hex or null",
    "border_width": number,
    "corner_radius": number,
    "shadow": false,
    "internal_padding": number
  },
  "title_style": {
    "font_family": "string",
    "font_size": number,
    "font_weight": "bold",
    "color": "hex"
  },
  "subtitle_style": {
    "font_family": "string",
    "font_size": number,
    "font_weight": "semibold",
    "color": "hex"
  },
  "body_style": {
    "font_family": "string",
    "font_size": number,
    "font_weight": "regular",
    "color": "hex",
    "line_spacing": number
  },
  "card_frames": [
    { "card_index": 0, "x": number, "y": number, "w": number, "h": number }
  ]
}

Card styling rules:
- card_frames: all cards must be EQUAL size **” same w and h; divide container evenly with 0.12" gutters
- card_style.fill_color: derive from card sentiment (from Agent 4 cards[i].sentiment):
    positive â†’ very light green tint (e.g., "#EEF7F0") or brand secondary light
    negative â†’ very light red/amber tint (e.g., "#FDF3F1") or brand warning light
    neutral  â†’ brand fill (light grey or brand secondary light, e.g., "#F5F5F5")
  If brand_tokens defines sentiment colors, use those instead
- card shape: square corners by default (corner_radius: 0). Do NOT use rounded pill cards.
- accent treatment: use a vertical accent strip on the LEFT side of the card, not a top strip
- when multiple cards exist in one artifact, vary accent-strip color using brand hierarchy:
    card 1 â†’ primary_color
    card 2 â†’ secondary_color
    card 3+ â†’ accent_colors in order
  Only fall back to sentiment color when a single card stands alone or no brand sequence exists
- title_style.font_size: SAME across all cards on the slide (pick one size, apply to all)
- subtitle_style.font_size: SAME across all cards (this is the headline metric **” make it the largest element, 18**“26pt)
- body_style.font_size: SAME across all cards (9**“11pt)
- card_style.internal_padding: 0.12**“0.18" **” consistent across all cards
- cards_layout: "row" when 3**“4 cards side by side; "column" when 2 cards stacked; "grid" for 4-card 2Ã—2

*********************************************************************************

4. WORKFLOW
*********************************************************************************

{
  "type": "workflow",
  "container": { "x": number, "y": number, "w": number, "h": number },
  "workflow_style": {
    "title_font_family": "string",
    "title_font_size": number,
    "title_color": "hex",
    "node_fill_color": "hex",
    "node_border_color": "hex",
    "node_border_width": number,
    "node_corner_radius": number,
    "node_title_font_family": "string",
    "node_title_font_size": number,
    "node_title_font_weight": "semibold" | "bold",
    "node_title_color": "hex",
    "node_value_font_family": "string",
    "node_value_font_size": number,
    "node_value_color": "hex",
    "connector_color": "hex",
    "connector_width": number,
    "arrowhead_style": "triangle" | "stealth",
    "connector_label_font_size": number,
    "connector_label_color": "hex",
    "node_inner_padding": number,
    "external_label_gap": number
  },
  "nodes": [
    { "id": "n1", "x": number, "y": number, "w": number, "h": number }
  ],
  "connections": [
    {
      "from": "n1",
      "to": "n2",
      "path": [ { "x": number, "y": number }, { "x": number, "y": number } ],
      "type": "arrow"
    }
  ],
  "header_block": null or {
    "text": "string **” the artifact_header value from Agent 4",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string", "font_size": number, "font_weight": "semibold", "color": "hex",
    "style": "underline" | "brand_fill",
    "accent_color": "hex"
  }
}

Workflow rules:
- max 6 nodes, max 8 connections, no crossing arrows
- follow flow_direction and workflow_type from Agent 4
- left_to_right: horizontal sequence, even spacing
- top_to_bottom: vertical stack, centered
- top_down_branching: root at top, children below, symmetric
- process_flow â†’ linear; hierarchy â†’ tree; decomposition â†’ branching; timeline â†’ equal phases
- all node x,y,w,h must be computed within container bounds
- connection path waypoints: straight or single-elbow only

Workflow message placement rules (MANDATORY):
- The PRIMARY message is the node label and must sit INSIDE the colored node box.
- For left_to_right / timeline:
  - value = short secondary line ABOVE the box, center-aligned
  - description = longer secondary line BELOW the box, center-aligned
  - if both exist, reserve space for both bands; keep the box visually dominant
  - if only one secondary message exists, prefer the below-box description band
- For top_to_bottom / bottom_up:
  - label stays inside the box
  - description is the ONLY external secondary message and must sit to the RIGHT of the box
  - do not place text above or below the box in these vertical flows
- For top_down_branching:
  - keep nodes compact and centered; external notes should be minimal and only used when necessary
- Never place long explanatory copy inside the node fill.

Workflow coordinate rules:
- Node sizing:
    process_flow / timeline (left_to_right): node w = (container.w âˆ’ (nâˆ’1)Ã—0.20) / n; node h must include any above/below message bands
    hierarchy / decomposition (top_down_branching): root node centered; children evenly spaced across width
    top_to_bottom / bottom_up: node box occupies roughly 35**“45% of container.w with the remaining width reserved for the right-side message band
- Node spacing: min 0.20" gap between adjacent nodes; evenly distribute remaining space
- Level assignment: all nodes at the same level must share the SAME y (horizontal) or x (vertical) coordinate
- Balance: for branching layouts, distribute child nodes symmetrically around parent x-center
- All node x,y must be within container.x/y and container.x+container.w / container.y+container.h
- Connection paths: start at center-right of "from" node, end at center-left of "to" node (left_to_right)
  For top_to_bottom: start at bottom-center, end at top-center

*********************************************************************************
WORKFLOW MICRO-LAYOUT OWNERSHIP:
- Reserve explicit whitespace for external value / description bands so labels do not collide with nodes or connectors
- You must decide the final node sizes, connector paths, node padding, and external text gaps

5. TABLE
*********************************************************************************

{
  "type": "table",
  "x": number, "y": number, "w": number, "h": number,
  "table_style": {
    "header_fill_color": "hex",
    "header_text_color": "hex",
    "header_font_family": "string",
    "header_font_size": number,
    "header_font_weight": "bold",
    "body_fill_color": "hex",
    "body_alt_fill_color": "hex or null",
    "body_text_color": "hex",
    "body_font_family": "string",
    "body_font_size": number,
    "body_font_weight": "regular",
    "grid_color": "hex",
    "grid_width": number,
    "highlight_fill_color": "hex or null",
    "highlight_text_color": "hex or null",
    "cell_padding": number
  },
  "column_widths": [number],
  "column_types": ["label" | "numeric" | "currency" | "percent" | "categorical"],
  "column_alignments": ["left" | "center" | "right"],
  "header_row_height": number,
  "row_heights": [number],
  "header_block": null or {
    "text": "string **” the artifact_header value from Agent 4",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string", "font_size": number, "font_weight": "semibold", "color": "hex",
    "style": "underline" | "brand_fill",
    "accent_color": "hex"
  }
}

Table styling rules:
- column_widths must sum exactly to table width (w field)
- body font min 9pt; header font min 10pt
- Column width heuristics (distribute table.w proportionally):
    Numeric columns (values, %, â‚¹): narrower **” typically 0.80**“1.20" each
    Label/name columns (first column, entity names): wider **” typically 1.50**“2.50"
    Short categorical columns: 0.80**“1.10"
- Column text alignment (enforce in table_style or column_align list if available):
    First / label column: left-align
    Numeric / currency / percent columns: right-align
    Header row: center-align all columns
- row_heights: all data rows equal height (0.30**“0.40"); header row slightly taller (0.35**“0.45")
- Zebra striping: set body_alt_fill_color to a very light tint of brand secondary (e.g., "#F7F8FA") for alternating rows
- highlight_rows: apply highlight_fill_color from brand accent to the highlight_rows indices from Agent 4
- table_style.cell_padding: 0.05**“0.08" (enforced by renderer; set as a hint here)

*********************************************************************************
TABLE MICRO-LAYOUT OWNERSHIP:
- You must output: column_widths, header_row_height, row_heights, column_types, and column_alignments
- Cell positions and frames (column_x_positions, row_y_positions, header_cell_frames, body_cell_frames) are computed automatically from these values **” do NOT output them
- The renderer must not infer table density, alignment, or spacing


6. MATRIX
*********************************************************************************

{
  "type": "matrix",
  "x": number, "y": number, "w": number, "h": number,
  "matrix_style": {
    "border_color": "hex",
    "border_width": number,
    "divider_color": "hex",
    "divider_width": number,
    "axis_label_font_family": "string",
    "axis_label_font_size": number,
    "axis_label_color": "hex",
    "quadrant_title_font_family": "string",
    "quadrant_title_font_size": number,
    "quadrant_body_font_family": "string",
    "quadrant_body_font_size": number,
    "positive_quadrant_fill": "hex **” light tint for favourable quadrants",
    "negative_quadrant_fill": "hex **” light tint for unfavourable quadrants",
    "neutral_quadrant_fill":  "hex **” light tint for neutral/monitor quadrants",
    "positive_title_color": "hex **” quadrant title text color for positive tone",
    "negative_title_color": "hex **” quadrant title text color for negative tone",
    "neutral_title_color":  "hex **” quadrant title text color for neutral tone",
    "positive_body_color":  "hex **” quadrant body text color for positive tone",
    "negative_body_color":  "hex **” quadrant body text color for negative tone",
    "neutral_body_color":   "hex **” quadrant body text color for neutral tone",
    "positive_point_fill":  "hex **” dot fill for points in positive quadrants",
    "negative_point_fill":  "hex **” dot fill for points in negative quadrants",
    "neutral_point_fill":   "hex **” dot fill for points in neutral quadrants",
    "point_label_font_family": "string",
    "point_label_font_size": number
  },
  "header_block": null or {
    "text": "string **” the artifact_header value from Agent 4",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string", "font_size": number, "font_weight": "semibold", "color": "hex",
    "style": "underline" | "brand_fill",
    "accent_color": "hex"
  }
}

Matrix rules:
- Quadrant fills are tone-driven (positive/negative/neutral set on each quadrant in Agent 4)
- Quadrant text colors (title, body) also driven by tone **” not a single global color
- Center dividers are dashed thin lines (not solid bars)
- Each quadrant renders: title (bold) + primary_message (italic axis descriptor). No secondary_message in quadrants.
- Each point renders as two shapes: filled abbreviation circle + outlined label bubble below
- Point dot color comes from quadrant_id â†’ that quadrant's tone (not a palette index)
- short_label (2-3 chars) goes inside the filled dot; full label goes in the bubble below
- Points carry NO primary_message or secondary_message **” those belong in the paired insight_text
- emphasis=high â†’ larger dot (0.26"), medium=0.20", low=0.16"
- Axis mid-labels (low_label/high_label) position at the center divider crosshair, NOT at outer edges **” they are the ONLY axis text rendered; they must be self-contained (axis name + threshold)
- xAxis.label and yAxis.label are metadata only **” do NOT render them; no rotated Y-axis label, no X-axis label below grid
- Point x/y are NUMERIC 0**“100 (percentage of full grid width/height). Y increases upward: y=100 plots at the TOP.
- Use quadrant_id to determine which quadrant a point belongs to for color; do NOT re-derive from x/y geometry

7. DRIVER_TREE
*********************************************************************************

{
  "type": "driver_tree",
  "x": number, "y": number, "w": number, "h": number,
  "tree_style": {
    "node_fill_color": "hex",
    "node_fill_color_secondary": "hex",
    "node_fill_color_leaf": "hex",
    "node_border_color": "hex",
    "node_border_width": number,
    "connector_color": "hex",
    "connector_width": "number **” connector line thickness in pts (1**“2pt typical; converted to inches by renderer)",
    "label_font_family": "string",
    "label_font_size": number,
    "label_color": "hex",
    "value_font_family": "string",
    "value_font_size": number,
    "value_color": "hex",
    "corner_radius": number
  },
  "header_block": null or {
    "text": "string **” the artifact_header value from Agent 4",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string", "font_size": number, "font_weight": "semibold", "color": "hex",
    "style": "underline" | "brand_fill",
    "accent_color": "hex"
  }
}

Driver tree rules:
- JS owns ALL geometry (node positions, sizes, connector paths, level distribution).
  Do NOT attempt to compute or output node x/y/w/h or connection waypoints.
- Your ONLY job is tree_style: choose node fill colors per level (root/branch/leaf),
  connector color and width, label fonts, corner_radius.
- Visual hierarchy is enforced by JS via node_fill_color (root) → node_fill_color_secondary
  (branch) → node_fill_color_leaf (leaf). Make these visually distinct.
- Max 3 levels; content (root label/value, branch labels/values, children) comes
  from the Agent 4 manifest — do NOT duplicate it in tree_style.

8. PRIORITIZATION
*********************************************************************************

{
  "type": "prioritization",
  "x": number, "y": number, "w": number, "h": number,
  "priority_style": {
    "row_fill_color": "hex",
    "row_border_color": "hex",
    "row_border_width": number,
    "row_corner_radius": number,
    "row_gap_in": number,
    "rank_palette": ["hex"],
    "rank_font_family": "string",
    "rank_font_size": number,
    "title_font_family": "string",
    "title_font_size": number,
    "title_color": "hex",
    "description_font_family": "string",
    "description_font_size": number,
    "description_color": "hex",
    "qualifier_text_color": "hex **” text color inside qualifier pills",
    "qualifier_value_palette": ["hex", "hex"],
    "qualifier_label_font_family": "string",
    "qualifier_label_font_size": number,
    "qualifier_value_font_family": "string",
    "qualifier_value_font_size": number
  },
  "header_block": null or {
    "text": "string **” the artifact_header value from Agent 4",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string", "font_size": number, "font_weight": "semibold", "color": "hex",
    "style": "underline" | "brand_fill",
    "accent_color": "hex"
  }
}

Prioritization rules:
- Use ONLY primitive geometry in the final blocks: rect, text_box
- Rows must be stacked vertically in rank order
- Each row contains:
  - left rank badge (two-layer design **” JS renders this automatically):
      Layer 1: rank-colored background rect (fully rounded) **” color = rank_palette[idx], so each rank has a distinct severity color
      Layer 2: white inner box at the same height as the background, shifted right so a thin colored strip shows on the left
      Text: "#N" in upper half and priority label (CRITICAL/HIGH/MEDIUM/LOW) in lower half **” both colored with rank_palette[idx]
  - rank_palette: set one color per rank (rank 1 = palette[0], rank 2 = palette[1], etc.) **” each rank must be visually distinct
  - action title
  - action description
  - up to 2 qualifier pills on the right
- rank_palette: set colors for each rank severity chip **” rank 1 gets palette[0], rank 2 gets palette[1], etc. Make each step visually distinct.
- qualifier_value_palette: MUST have exactly 2 entries **” qualifier 1 always uses palette[0], qualifier 2 always uses palette[1]. Use visually distinct colors so dual qualifiers are immediately distinguishable.
- Qualifier slots may be empty; do not render empty pills
- Rank 1 should be visually strongest; later ranks may step down subtly through the rank palette
- Title must dominate description; qualifiers must remain compact, secondary metadata

*********************************************************************************

9. STAT_BAR
*********************************************************************************

{
  "type": "stat_bar",
  "x": number, "y": number, "w": number, "h": number,
  "artifact_header": "string **” artifact section label",
  "column_headers": [
    {"id": "col1", "value": "string **” column header label", "display_type": "text" | "bar" | "normal"}
  ],
  "annotation_style": {
    "label_font_family": "string",
    "axis_color": "hex **” muted color for column header labels, dividers, and neutral bars",
    "annotation_color": "hex **” annotation text column color",
    "gridline_color": "hex **” header divider rule color",
    "border_color": "hex **” non-highlighted row border color",
    "background_color": "hex **” empty bar track fill and highlighted row background tint"
  },
  "scale_UL": number or null **” upper limit for bar scale; null = auto-scale to max row value,
  "rows": [
    {
      "row_id": number,
      "row_focus": "Y" | "N" **” "Y" = highlighted row,
      "cells": [
        {"col_id": "col1", "value": "string **” numeric string for bar columns, display text for others"}
      ]
    }
  ],
  "header_block": null or { "text": "string", "x": number, "y": number, "w": number, "h": number,
    "font_family": "string", "font_size": number, "font_weight": "semibold", "color": "hex",
    "style": "underline" | "brand_fill", "accent_color": "hex" }
}

Stat_bar rules:
- Layout (column widths, bar geometry, row heights) is computed by JS from x/y/w/h **” do NOT set internal positions
- column_headers: array of column descriptors. display_type "text" = label/annotation text; "bar" = proportional bar (1**“3 allowed, each with optional "scale_UL" on the column); "normal" = secondary display value
- COLUMN PAIRING: every "bar" column MUST be immediately followed by a "normal" column with "value": "" **” that column renders the bar's numeric value as readable text. No two "bar" columns may be adjacent.
- axis_color: use a muted caption gray (NOT a brand primary **” too vibrant for column headers)
- background_color: use a very light tint for the bar track background and highlighted row fill
- Highlighted bar fill comes automatically from brand chart_palette[0]; do not set it in annotation_style
- Max 8 rows; mark the most important row row_focus: "Y" (at most 1**“2 highlighted rows)

*********************************************************************************
10. COMPARISON_TABLE
*********************************************************************************

{
  "type": "comparison_table",
  "x": number, "y": number, "w": number, "h": number,
  "comparison_style": {
    "label_font_family": "string",
    "label_font_size": number,
    "body_font_family": "string",
    "body_font_size": number,
    "container_fill_color": "hex **” outer shell background (usually #FFFFFF)",
    "container_border_color": "hex",
    "container_border_width": number,
    "container_corner_radius": number,
    "grid_color": "hex **” horizontal row dividers",
    "recommended_fill_color": "hex **” highlight fill for the recommended option row",
    "yes_fill_color": "hex", "yes_text_color": "hex",
    "no_fill_color": "hex",  "no_text_color": "hex",
    "partial_fill_color": "hex", "partial_text_color": "hex",
    "neutral_fill_color": "hex"
  },
  "header_block": null or { ... }
}

Comparison table rules:
- Layout (column widths, row heights, cell positions) is computed by JS from x/y/w/h **” do NOT set internal positions
- Content (columns[], rows[]) comes from the Agent 4 manifest **” do NOT duplicate it here
- recommended_fill_color: light green tint (#EEF4E2 or brand equivalent)
- yes/no/partial: semantic signal colors **” green/red/amber respectively; do NOT use brand primary for these
- label_font_size: 9**“11pt; body_font_size: 8**“10pt
- container_corner_radius: 6**“10 for a card-like container

*********************************************************************************
11. INITIATIVE_MAP
*********************************************************************************

{
  "type": "initiative_map",
  "x": number, "y": number, "w": number, "h": number,
  "initiative_style": {
    "label_font_family": "string",
    "label_font_size": number,
    "body_font_family": "string",
    "body_font_size": number,
    "row_fill_color": "hex **” data row background (usually #FFFFFF)",
    "row_border_color": "hex **” row dividers and column separators",
    "row_border_width": number,
    "row_corner_radius": number,
    "primary_chip_fill": "hex **” tag chip fill for primary-tone cells (light tint of brand primary)",
    "secondary_chip_fill": "hex **” tag chip fill for secondary-tone cells (light tint of brand secondary)",
    "neutral_chip_fill": "hex **” tag chip fill for neutral-tone cells",
    "positive_color": "hex **” text color for positive / delta values"
  },
  "header_block": null or { ... }
}

Initiative map rules:
- Layout (track width, lane widths, row heights, cell text positions) is computed by JS from x/y/w/h **” do NOT set internal positions
- Content (column_headers, rows, cells, primary_message, secondary_message, tags) comes from the Agent 4 manifest **” do NOT duplicate it here
- row_border_color: use a light grid color (e.g. #D7DEE8)
- primary_chip_fill: light (~10%) tint of brand primary_color; secondary_chip_fill: light tint of brand secondary_color
- label_font_size: 9**“11pt; body_font_size: 8**“10pt

*********************************************************************************
12. PROFILE_CARD_SET
*********************************************************************************

{
  "type": "profile_card_set",
  "x": number, "y": number, "w": number, "h": number,
  "profile_style": {
    "label_font_family": "string",
    "label_font_size": number,
    "body_font_family": "string",
    "card_fill_color": "hex **” card background (usually #FFFFFF or very light brand tint)",
    "card_border_color": "hex",
    "card_border_width": number,
    "card_corner_radius": number,
    "muted_color": "hex **” subtitle and attribute-key label color (muted gray)",
    "divider_color": "hex **” horizontal line between card header and attribute body",
    "badge_fill_color": "hex **” KPI badge background (light green tint default)",
    "badge_border_color": "hex **” KPI badge border",
    "badge_text_color": "hex **” KPI badge text",
    "chip_fill_color": "hex **” attribute chip background (warm neutral default)",
    "chip_border_color": "hex **” attribute chip border",
    "chip_text_color": "hex **” attribute chip label text"
  },
  "header_block": null or { ... }
}

Profile card set rules:
- Grid layout (columns Ã— rows, card width/height, internal card proportions) is computed by JS from x/y/w/h and profile count **” do NOT set card positions
- Content (entity_name, subtitle, badge_text, attributes) comes from the Agent 4 manifest **” do NOT duplicate it here
- card_fill_color: #FFFFFF or a 5% tint of brand primary
- card_border_color: light gray or brand grid line color (#D7DEE8)
- label_font_size: 9**“13pt (controls the card entity name size)

*********************************************************************************
13. RISK_REGISTER
*********************************************************************************

{
  "type": "risk_register",
  "x": number, "y": number, "w": number, "h": number,
  "risk_style": {
    "label_font_family": "string",
    "label_font_size": number,
    "body_font_family": "string",
    "body_font_size": number,
    "primary_message_font_size": number,
    "secondary_message_font_size": number,
    "band_height_in": number,
    "row_height_in": number,
    "right_col_width_in": number,
    "row_border_color": "hex", "row_border_width": number, "row_corner_radius": number,
    "critical_fill_color": "hex", "high_fill_color": "hex", "medium_fill_color": "hex", "low_fill_color": "hex",
    "critical_badge_color": "hex", "high_badge_color": "hex", "medium_badge_color": "hex", "low_badge_color": "hex",
    "badge_text_color": "hex",
    "critical_text_color": "hex", "high_text_color": "hex", "medium_text_color": "hex", "low_text_color": "hex",
    "critical_pip_fill": "hex", "high_pip_fill": "hex", "medium_pip_fill": "hex", "low_pip_fill": "hex",
    "tag_positive_fill": "hex", "tag_positive_border": "hex", "tag_positive_text": "hex",
    "tag_negative_fill": "hex", "tag_negative_border": "hex", "tag_negative_text": "hex",
    "tag_warning_fill": "hex",  "tag_warning_border": "hex",  "tag_warning_text": "hex",
    "tag_neutral_fill": "hex",  "tag_neutral_border": "hex",  "tag_neutral_text": "hex"
  },
  "header_block": null or { ... }
}

Risk register rules:
- Layout (severity bands, row heights, pip positions, column widths) is computed by JS from x/y/w/h **” do NOT set internal positions
- Content (severity_levels[], each with label/tone/item_details[]; each item has primary_message, secondary_message, tags[], pips[]) comes from the Agent 4 manifest **” do NOT duplicate it here
- risk_register may have an artifact_header and header_block like any other artifact; the risk_header is an additional internal section header rendered by JS inside the body
- EXCEPTION: any slide that has exactly 1 zone with exactly 1 artifact must NOT have an artifact_header on that artifact (the slide title already serves as the header)

RISK_REGISTER STYLING **” decided by YOU (LLM) using brand tokens:
  Severity palette: use the brand's red/warm sequence at decreasing intensity **” critical=darkest, high=dark, medium=mid, low=light.
    If the brand has a primary red or error color, use it as the base. Otherwise use #DC2626 as semantic red base.
    critical_badge_color: brand error / darkest red (e.g. bt.primary_color if it is a red brand, else #DC2626)
    critical_fill_color:  very light tint of critical_badge_color (~8% opacity, e.g. #FDE8E8)
    critical_text_color:  dark shade of critical_badge_color for text legibility
    high_badge_color:     one step lighter/warmer than critical (e.g. #EA580C or brand secondary if warm)
    high_fill_color:      very light tint of high_badge_color
    high_text_color:      dark shade of high_badge_color
    medium_badge_color:   amber (e.g. #D97706 or brand accent if amber)
    medium_fill_color:    very light amber tint
    medium_text_color:    dark amber for text
    low_badge_color:      neutral gray (e.g. #6B7280)
    low_fill_color:       near-white gray tint
    low_text_color:       mid-gray
  pip_fill colors: match the badge_color for each severity level
  Tag chip colors: derive from brand tokens **”
    neutral tags: use bt.body_color tint for border, near-white fill
    negative tags: use critical_badge_color tint (same brand red family)
    positive tags: use brand success green if available, else #7AA243 family
    warning tags:  use medium_badge_color tint (amber family)
  Font sizes: YOU decide **” all primary_messages at one consistent size (recommend 11**”12pt), all secondary_messages at one consistent size (recommend 9**”10pt, 2pt smaller than primary)
  label_font_size: band header size (recommend 10pt)
  body_font_size: pip labels, item count (recommend 9pt)
  Layout density: YOU decide **” scale these to the artifact height and item count:
    band_height_in: height of each severity section-header band (recommend 0.28**”0.40”; default 0.34)
    row_height_in:  height of each risk item row (recommend 0.70**”1.10”; default 0.90)
                    use lower end when many items must fit, upper end for spacious decks
    right_col_width_in: width of the right column that holds pip grid + tags (recommend 1.60**”2.20”; default 1.90)

ARTIFACT HEADER
*********************************************************************************

Every artifact except cards gets a header_block **” a one-line label above the artifact
that names the insight it proves.

Source text from Agent 4:
  use artifact_header for the local artifact heading across all artifact types

If header text is empty or null: set header_block to null.

LAYOUT MODE **” detecting if the layout has a header area:
  Check layout_map[selected_layout_name].ph_count:
  - ph_count > 2: layout has extra placeholders beyond title+body
    â†’ set header_block.placeholder_ref: true
    â†’ Agent 6 will place header text into the dedicated placeholder
    â†’ x/y/w/h in header_block can be approximate (derived from body_placeholder top)
  - ph_count â‰¤ 2: layout has only title + body placeholders, no dedicated header
    â†’ set header_block.placeholder_ref: false
    â†’ compute header_block coordinates at the TOP of body_placeholder:
        x = body_placeholder.x_in
        y = body_placeholder.y_in
        w = body_placeholder.w_in
        h = 0.30
    â†’ adjust artifact y = body_placeholder.y_in + 0.30
    â†’ adjust artifact h = body_placeholder.h_in - 0.30
    â†’ choose style:
        "underline"  **” thin line under the header text using brand accent color
        "brand_fill" **” fill header area with brand primary/secondary, white text
      Default: "underline"

SCRATCH MODE:
  - Position header_block at the top of zone inner bounds
  - artifact y += header_block.h; artifact h -= header_block.h
  - Choose style same as above
`