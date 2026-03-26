// ─── AGENT 5 — SLIDE LAYOUT & VISUAL DESIGN ENGINE ───────────────────────────
// Input:  state.slideManifest   — output from Agent 4
//         state.brandRulebook   — brand guideline JSON from Agent 2
//         state.outline         — presentation brief from Agent 3
//
// Output: designedSpec — flat JSON array, one render-ready object per slide
//
// Architecture: Claude API call per batch of 4 slides.
// Claude receives the Agent 4 manifest + brand guideline + brief and returns
// a precise layout spec: canvas, brand_tokens, title_block, subtitle_block,
// zones (with fully positioned artifacts), and global_elements.
//
// Agent 5.1 then reviews this spec and applies targeted fixes.
// Agent 6 (python-pptx) consumes the final reviewed spec.

// ═══════════════════════════════════════════════════════════════════════════════
// SYSTEM PROMPT
// ═══════════════════════════════════════════════════════════════════════════════

var AGENT5_SYSTEM = `You are a senior presentation designer and layout system architect.

You will receive:
1. A slide content manifest created by Agent 4
2. A brand guideline JSON for the current deck
3. A presentation brief that explains the narrative flow and tone

═══════════════════════════
BATCH PROCESSING RULE
═══════════════════════════

You will receive slides in batches of 4–5.
Process ONLY the slides in this batch.
Return ONLY those slides in the JSON array.
Do not infer slides before or after this batch.
Do not summarize the entire deck.
Do not renumber slides.

═══════════════════════════
ROLE
═══════════════════════════

Your task is to convert each slide into an exact render-ready layout specification for PowerPoint generation.

You are NOT rewriting content.
You are NOT changing business meaning.

You ARE responsible for:
- spatial layout
- coordinates and dimensions
- typography
- color application
- chart styling
- workflow geometry
- table styling
- card styling
- alignment and spacing

Your output must be directly usable by a PPT rendering engine.

═══════════════════════════
COORDINATE SYSTEM
═══════════════════════════

- unit: inches
- origin: (0.00, 0.00) at top-left of slide
- x increases left → right
- y increases top → bottom
- all numeric values must be decimal inches
- round all numeric values to 2 decimal places

Applies to:
- canvas
- margins
- title/subtitle
- zones
- artifacts
- workflow nodes
- workflow connectors
- tables
- cards
- global elements

═══════════════════════════
BRAND GUIDELINE AUTHORITY RULE
═══════════════════════════

The brand guideline is the primary authority for design.

If the brand guideline defines:
- slide size
- fonts
- color palette
- typography hierarchy
- layout styles
- chart styles
- footer conventions

You MUST follow it.

DO NOT override brand-defined styles.

Only if the brand guideline is missing or incomplete:
- use neutral corporate defaults
- ensure readability and hierarchy
- avoid decorative styling

═══════════════════════════
TITLE & DIVIDER SLIDES (template mode)
═══════════════════════════

When slide_type is "title" or "divider" AND uses_template is true:
The master provides ALL visual elements — background, logo, decorations, footer.
Agent 6 places text directly into the master's title/subtitle placeholders.

Output ONLY:
- title_block: { text, font_family, font_size, font_weight, color } — NO x/y/w/h
- subtitle_block: same, or null
- zones: []
- global_elements: {}
- canvas.background: null
- layout_mode: true

═══════════════════════════
ALL CONTENT SLIDES — FIXED ELEMENTS (template mode)
═══════════════════════════

When slide_type is "content" AND uses_template is true, regardless of layout selection:
The master template positions slide title, subtitle, footer, and page number.

ALWAYS apply these rules for ALL content slides in template mode:
- title_block: text, font_family, font_size, font_weight, color ONLY — omit x/y/w/h
- subtitle_block: same — omit x/y/w/h
- global_elements: {} — master handles footer and page number automatically
- canvas.background: null

═══════════════════════════
CONTENT SLIDES — LAYOUT MODE (selected_layout_name is non-empty)
═══════════════════════════

When uses_template is true AND selected_layout_name is non-empty:
The pipeline automatically maps each zone to its content area slot in the named layout.
Your job is CONTENT QUALITY and VISUAL STYLE — not positioning.

Rules:
1. Set layout_mode: true
2. For each zone: set frame to null — the pipeline fills frame from the layout's content_areas
3. Do NOT set placeholder_idx on artifacts — the pipeline assigns the real PPTX placeholder idx
4. For each artifact (except cards): add header_block
   - layout_map[selected_layout_name].ph_count > 2 (multi-slot layout):
     → set header_block.placeholder_ref: true
     → x/y/w/h may be null (pipeline positions from placeholder)
   - ph_count ≤ 2: compute header_block at top of zone area; h = 0.30
5. Focus on brand-compliant styling: chart colors from chart_color_sequence, correct fonts,
   card_frames, insight_text body_style, table column_widths, workflow_style

═══════════════════════════
CONTENT SLIDES — SCRATCH MODE (selected_layout_name is empty)
═══════════════════════════

When uses_template is false OR selected_layout_name is empty:
Compute all coordinates from layout_hint splits.
- Set layout_mode: false
- zones: compute full frame coordinates from layout_hint splits
- Artifacts: compute x/y/w/h within zone bounds
- Artifact headers: compute header_block at top of zone inner bounds; shrink artifact area
- global_elements: include footer and page_number when uses_template is false
- canvas.background: set when uses_template is false

═══════════════════════════
OUTPUT STRUCTURE
═══════════════════════════

Each slide must return EXACTLY:

{
  "slide_number": number,
  "slide_type": "title" | "divider" | "content",
  "slide_archetype": "summary" | "trend" | "comparison" | "breakdown" | "driver_analysis" | "process" | "recommendation" | "dashboard" | "proof" | "roadmap",
  "layout_mode": true | false,
  "selected_layout_name": "string — brand layout name chosen by Agent 4, or empty string",
  "canvas": {
    "width_in": number,
    "height_in": number,
    "margin": { "left": number, "right": number, "top": number, "bottom": number },
    "background": { "color": "hex" }
  },
  "brand_tokens": {
    "title_font_family": "string",
    "body_font_family": "string",
    "caption_font_family": "string",
    "title_color": "hex",
    "body_color": "hex",
    "caption_color": "hex",
    "primary_color": "hex",
    "secondary_color": "hex",
    "accent_colors": ["hex"],
    "chart_palette": ["hex"]
  },
  "title_block": {
    "text": "string",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string",
    "font_size": number,
    "font_weight": "regular" | "semibold" | "bold",
    "color": "hex",
    "align": "left" | "center",
    "valign": "middle" | "top",
    "wrap": true
  },
  "subtitle_block": null or {
    "text": "string",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string",
    "font_size": number,
    "font_weight": "regular" | "semibold",
    "color": "hex",
    "align": "left" | "center",
    "valign": "middle" | "top",
    "wrap": true
  },
  "zones": [...],
  "global_elements": {
    "footer": {
      "show": true,
      "x": number, "y": number, "w": number, "h": number,
      "font_family": "string",
      "font_size": number,
      "color": "hex",
      "align": "left"
    },
    "page_number": {
      "show": true,
      "x": number, "y": number, "w": number, "h": number,
      "font_family": "string",
      "font_size": number,
      "color": "hex",
      "align": "right"
    }
  }
}

If subtitle_block is not needed set it to null.
global_elements is optional — include when appropriate.

═══════════════════════════
TITLE / SUBTITLE SIZING
═══════════════════════════

- title slides: 24–34 pt
- divider slides: 22–30 pt
- content slides: 16–22 pt
- subtitle is 9–16 pt smaller than title

═══════════════════════════
ZONES
═══════════════════════════

Each zone:

{
  "zone_id": "z1",
  "zone_role": "string",
  "message_objective": "string",
  "narrative_weight": "primary" | "secondary" | "supporting",
  "frame": {
    "x": number, "y": number, "w": number, "h": number,
    "padding": { "top": number, "right": number, "bottom": number, "left": number }
  },
  "artifacts": [...]
}

Rules:
- max 4 zones per slide
- zones must NOT overlap
- must respect layout_hint from Agent 4
- primary zones get more space than secondary
- each zone has 1–2 artifacts
- title and subtitle sit OUTSIDE zones

═══════════════════════════
ARTIFACT CONTRACT
═══════════════════════════

Every artifact must be FULLY specified. No missing fields. No placeholders.

Allowed types: insight_text | chart | cards | workflow | table

═══════════════════════════
1. INSIGHT TEXT
═══════════════════════════

─── STEP 1: SET insight_mode ───────────────────────────────────────────────────
Inspect the Agent 4 artifact:
- Agent 4 provides "groups[]"  → set insight_mode: "grouped"
- Agent 4 provides "points[]"  → set insight_mode: "standard"

─── STANDARD MODE SCHEMA ───────────────────────────────────────────────────────

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
    "text": "string — the insight_header value from Agent 4",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string", "font_size": number, "font_weight": "semibold", "color": "hex",
    "style": "underline" | "brand_fill",
    "accent_color": "hex",
    "placeholder_ref": true or false
  }
}

Standard mode styling rules:
- style.fill_color: null (transparent) or very light brand background tint
- style.border_color: subtle brand accent color or null — no heavy full-border box
- heading_style.color: brand primary color; risk/alert accent only for "Risk Alert"
- list_style: "tick_cross" for positive/negative mix; "numbered" for sequential/ranked; "bullet" for parallel
- indent_inches: 0.12–0.18; space_before_pt: 4–8pt
- vertical_distribution "spread": distribute points evenly — do NOT cluster at top
- body font_size proportional to artifact height:
    h < 2.0": 9–11pt;  h 2.0–3.5": 11–14pt;  h > 3.5": 14–18pt
- Do NOT pre-shrink font — renderer auto-fits; use upper end of range
- heading_style.font_size = body_style.font_size + 2 to 4pt

─── GROUPED MODE SCHEMA ────────────────────────────────────────────────────────

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

─── GROUPED MODE DESIGN RULES ──────────────────────────────────────────────────

GROUP LAYOUT — choose based on zone dimensions and group count:
- "columns": use when zone w > zone h AND group count ≤ 3 (groups side-by-side)
- "rows":    use when zone h ≥ zone w OR group count ≥ 3 (groups stacked vertically, header left + box right)

GROUP HEADER SHAPE — choose based on content semantics:
- "circle_badge": ONLY when groups represent numbered priority steps (1, 2, 3…) or sequential phases
  - Renders as a filled circle; priority number (1-based index) shown as bold text inside
- "rounded_rect": DEFAULT for all text-label group headers
  - columns layout: spans full column width as a header bar
  - rows layout: left-side label block beside the bullet box

GROUP HEADER COLORS:
- fill_color: brand primary color (EY yellow #FFE600, or dark brand color for contrast)
- text_color: #111111 on light/yellow fills; #FFFFFF on dark fills
- Do NOT use a subtle/light fill — headers must be visually dominant

GROUP BULLET BOX:
- fill_color: null or very light near-white tint
- border_color: light brand border — use brand secondary light or a neutral — NEVER heavy/dark
- border_width: 0.5–1.0pt
- corner_radius: match group_header_style.corner_radius for visual consistency

DIMENSION CALCULATION — derive ALL sizes from content and available area:

  Let n       = number of groups
  Let f       = bullet_style.font_size (pt)
  Let line_h  = f / 72  (inches per line, 1pt = 1/72")
  Let header_block_h = height consumed by artifact-level header_block (0 if null)

  1. group_header_style.h   — height of the group header shape:
       rounded_rect → h = max(f × 1.8 / 72,  artifact.h × 0.06)
       circle_badge → h = max(f × 2.2 / 72,  artifact.h × 0.08)   [w = h, always a circle]

  2. group_header_style.w   — width of the group header shape:
       columns layout → NOT specified here; renderer uses full column width
       rows layout, rounded_rect → estimate from longest group header label:
           w = (max_header_chars × f × 0.55 / 72) + (f × 2.0 / 72)
       rows layout, circle_badge → w = h (square bounding box)

  3. group_gap_in — gap between adjacent groups:
       columns → max(artifact.w × 0.015, 0.05)
       rows    → max(artifact.h × 0.015, 0.05)

  4. header_to_box_gap_in:
       = max(f × 0.5 / 72, 0.03)

  5. group_bullet_box_style.padding:
       top = bottom = max(f × 0.8 / 72, 0.05)
       left = right = max(f × 1.0 / 72, 0.07)

  6. bullet_style.font_size — derive from available box area:
       columns layout:
           col_w = (artifact.w - (n-1) × group_gap_in) / n
           box_h = artifact.h - header_block_h - group_header_style.h - header_to_box_gap_in
           max_bullets_in_col = max bullets across all groups
           f = floor(box_h × 72 / (max_bullets_in_col × 1.5))  → clamp to [8, 14]
       rows layout:
           total_row_h = artifact.h - header_block_h - (n-1) × group_gap_in
           max_bullets_in_row = max bullets across all groups
           min_row_h = total_row_h × (max_bullets_in_row / total_bullets)
           box_w = artifact.w - group_header_style.w - header_to_box_gap_in
           f = floor(min_row_h × 72 / (max_bullets_in_row × 1.5))  → clamp to [8, 14]

  7. space_before_pt = max(f × 0.4, 2)   (scales with font — tighter than standard mode)
     indent_inches   = max(f × 1.0 / 72, 0.08)

GEOMETRY — columns layout (renderer uses these formulas, NOT hardcoded values):
- col_w[i]   = (artifact.w - (n-1) × group_gap_in) / n          [equal width per column]
- header_h   = group_header_style.h                               [same for all columns]
- box_h[i]   = artifact.h - header_block_h - header_h - header_to_box_gap_in  [same for all columns]

GEOMETRY — rows layout (renderer uses these formulas):
- total_row_h = artifact.h - header_block_h - (n-1) × group_gap_in
- row_h[i]   = total_row_h × (bullets[i].length / total_bullets)  [PROPORTIONAL to bullet count — NOT equal]
               minimum row_h[i] = group_header_style.h + header_to_box_gap_in + (1 line of text)
- header_w   = group_header_style.w                               [same for all rows]
- box_w[i]   = artifact.w - header_w - header_to_box_gap_in      [same for all rows]

═══════════════════════════
2. CHART
═══════════════════════════

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
    "show_gridlines": true,
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
    "text": "string — the chart_header value from Agent 4",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string", "font_size": number, "font_weight": "semibold", "color": "hex",
    "style": "underline" | "brand_fill",
    "accent_color": "hex"
  }
}

Chart rules:
- use brand chart_palette in sequence for series
- primary series uses primary brand color
- minimum axis font size: 8pt
- if chart + table in zone: chart takes 60–75% of zone width

DUAL AXIS — MANDATORY:
- If two or more series have DIFFERENT units (e.g. one is a count/number, another is ₹/currency/%, etc.)
  you MUST set dual_axis: true and list the secondary-axis series names in secondary_series[].
  NEVER plot metrics with different units on the same Y axis.
  The renderer will automatically display primary series as bars and secondary series as a line on the right Y axis.

LAYOUT SIZE HINTS based on category count:
These rules apply in BOTH scratch mode (free positioning) AND layout mode (template layouts).

SCRATCH MODE — set zone dimensions directly:
- Column/vertical bar chart with > 6 categories: set zone w ≥ 7" (wide/horizontal stretch)
- Horizontal bar chart with > 6 categories: set zone h ≥ 5" (tall/vertical stretch)

LAYOUT MODE — select the right layout from layout_map:
- layout_map[name].content_areas is an array of placeholders sorted left→right, top→bottom.
  Each entry has w_in and h_in — the actual rendered size of that placeholder in the template.
  content_areas[0] is always the primary (largest/first) content slot.
- Column/vertical bar chart with > 6 categories: scan layout_map for the layout whose
  content_areas[0].w_in is largest (≥ 7" preferred). Set selected_layout_name to that layout.
- Horizontal bar chart with > 6 categories: scan layout_map for the layout whose
  content_areas[0].h_in is largest (≥ 5" preferred). Set selected_layout_name to that layout.
- This overrides Agent 4's selected_layout_name when a better-fitting layout exists —
  chart readability takes priority over the default layout choice.
- If no template layout meets the threshold, fall back to the widest/tallest available and
  note the constraint; do NOT invent coordinates that exceed the placeholder bounds.

HEADER DEDUPLICATION:
- If the zone has a header_ph_idx (layout mode), the artifact heading is written to the layout's
  header placeholder automatically. In that case set chart_title: "" — do NOT repeat the same
  text as both an internal chart title and a header_block / layout header.
- Only use chart_title for a subtitle-style annotation inside the chart plot area when it adds
  information beyond what the zone header already says.

LEGEND POSITION (auto-computed by renderer from frame size, but follow these intent rules):
- Chart w > 6" (horizontally stretched) → legend on right
- Chart h > 4.5" AND w ≤ 6" (vertically stretched) → legend on top
- All other cases → legend on right

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

═══════════════════════════
3. CARDS
═══════════════════════════

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
- card_frames: all cards must be EQUAL size — same w and h; divide container evenly with 0.12" gutters
- card_style.fill_color: derive from card sentiment (from Agent 4 cards[i].sentiment):
    positive → very light green tint (e.g., "#EEF7F0") or brand secondary light
    negative → very light red/amber tint (e.g., "#FDF3F1") or brand warning light
    neutral  → brand fill (light grey or brand secondary light, e.g., "#F5F5F5")
  If brand_tokens defines sentiment colors, use those instead
- title_style.font_size: SAME across all cards on the slide (pick one size, apply to all)
- subtitle_style.font_size: SAME across all cards (this is the headline metric — make it the largest element, 18–26pt)
- body_style.font_size: SAME across all cards (9–11pt)
- card_style.internal_padding: 0.12–0.18" — consistent across all cards
- cards_layout: "row" when 3–4 cards side by side; "column" when 2 cards stacked; "grid" for 4-card 2×2

═══════════════════════════
4. WORKFLOW
═══════════════════════════

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
    "connector_label_color": "hex"
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
    "text": "string — the workflow_header value from Agent 4",
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
- process_flow → linear; hierarchy → tree; decomposition → branching; timeline → equal phases
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
    process_flow / timeline (left_to_right): node w = (container.w − (n−1)×0.20) / n; node h must include any above/below message bands
    hierarchy / decomposition (top_down_branching): root node centered; children evenly spaced across width
    top_to_bottom / bottom_up: node box occupies roughly 35–45% of container.w with the remaining width reserved for the right-side message band
- Node spacing: min 0.20" gap between adjacent nodes; evenly distribute remaining space
- Level assignment: all nodes at the same level must share the SAME y (horizontal) or x (vertical) coordinate
- Balance: for branching layouts, distribute child nodes symmetrically around parent x-center
- All node x,y must be within container.x/y and container.x+container.w / container.y+container.h
- Connection paths: start at center-right of "from" node, end at center-left of "to" node (left_to_right)
  For top_to_bottom: start at bottom-center, end at top-center

═══════════════════════════
5. TABLE
═══════════════════════════

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
  "row_heights": [number],
  "header_block": null or {
    "text": "string — the table_header value from Agent 4",
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
    Numeric columns (values, %, ₹): narrower — typically 0.80–1.20" each
    Label/name columns (first column, entity names): wider — typically 1.50–2.50"
    Short categorical columns: 0.80–1.10"
- Column text alignment (enforce in table_style or column_align list if available):
    First / label column: left-align
    Numeric / currency / percent columns: right-align
    Header row: center-align all columns
- row_heights: all data rows equal height (0.30–0.40"); header row slightly taller (0.35–0.45")
- Zebra striping: set body_alt_fill_color to a very light tint of brand secondary (e.g., "#F7F8FA") for alternating rows
- highlight_rows: apply highlight_fill_color from brand accent to the highlight_rows indices from Agent 4
- table_style.cell_padding: 0.05–0.08" (enforced by renderer; set as a hint here)

═══════════════════════════
ARTIFACT HEADER
═══════════════════════════

Every artifact except cards gets a header_block — a one-line label above the artifact
that names the insight it proves.

Source text from Agent 4:
  insight_text → insight_header   chart → chart_header
  workflow     → workflow_header  table → table_header

If header text is empty or null: set header_block to null.

LAYOUT MODE — detecting if the layout has a header area:
  Check layout_map[selected_layout_name].ph_count:
  - ph_count > 2: layout has extra placeholders beyond title+body
    → set header_block.placeholder_ref: true
    → Agent 6 will place header text into the dedicated placeholder
    → x/y/w/h in header_block can be approximate (derived from body_placeholder top)
  - ph_count ≤ 2: layout has only title + body placeholders, no dedicated header
    → set header_block.placeholder_ref: false
    → compute header_block coordinates at the TOP of body_placeholder:
        x = body_placeholder.x_in
        y = body_placeholder.y_in
        w = body_placeholder.w_in
        h = 0.30
    → adjust artifact y = body_placeholder.y_in + 0.30
    → adjust artifact h = body_placeholder.h_in - 0.30
    → choose style:
        "underline"  — thin line under the header text using brand accent color
        "brand_fill" — fill header area with brand primary/secondary, white text
      Default: "underline"

SCRATCH MODE:
  - Position header_block at the top of zone inner bounds
  - artifact y += header_block.h; artifact h -= header_block.h
  - Choose style same as above

═══════════════════════════
INTERNAL ZONE LAYOUT (2 artifacts)
═══════════════════════════

If 2 artifacts in a zone, split zone frame between them:
- chart + insight_text: chart 65%, insight 35%
- workflow + insight_text: workflow dominant
- chart + table: chart 65%, table 35% side by side
- cards + insight_text: cards dominant
- table + insight_text: table 65%, insight 35%
Artifacts must NOT overlap.

═══════════════════════════
LAYOUT DECISION RULES
═══════════════════════════

Translate Agent 4 layout_hint to zone geometry.
Content area = canvas minus margins minus title band.
Gutter between zones: 0.15 inches.

- full: single frame fills content area
- left_60 + right_40: split width 60/40
- left_50 + right_50: equal width split
- top_30 + bottom_70: split height 30/70
- top_left_50 + top_right_50 + bottom_full: two upper + full lower
- left_full_50 + top_right_50_h + bottom_right_50_h: left full + right stacked
- tl + tr + bl + br: 2×2 grid, equal gutters

═══════════════════════════
QUALITY RULES
═══════════════════════════

- no missing fields
- no overlapping zones or artifacts
- body text ≥ 9pt; captions ≥ 8pt
- no placeholder values
- reflect narrative hierarchy in space allocation
- primary zones visually dominant

═══════════════════════════
OUTPUT RULE
═══════════════════════════

Return ONLY a valid JSON array.
No explanation. No markdown. No text outside JSON.`



// ═══════════════════════════════════════════════════════════════════════════════
// UTILITIES
// ═══════════════════════════════════════════════════════════════════════════════

// Round to 2 decimal places — kills JS float drift (10.219999... -> 10.22)
function r2(n) { return Math.round(n * 100) / 100 }


// ═══════════════════════════════════════════════════════════════════════════════
// BRAND TOKEN EXTRACTOR
// Strips slide_layouts and other bulky fields before sending to Claude.
// slide_layouts contains full placeholder XML — can be 10K+ tokens alone.
// Agent 5 only needs design tokens: colors, fonts, slide size.
// ═══════════════════════════════════════════════════════════════════════════════

function extractBrandTokens(brand) {
  const primaryLogo = brand.primary_logo || ((brand.logos || [])[0] || null)
  return {
    slide_width_inches:   r2(brand.slide_width_inches  || 13.33),
    slide_height_inches:  r2(brand.slide_height_inches || 7.50),
    primary_colors:       brand.primary_colors       || [],
    secondary_colors:     brand.secondary_colors     || [],
    background_colors:    brand.background_colors    || ['#FFFFFF'],
    text_colors:          brand.text_colors          || ['#111111'],
    accent_colors:        brand.accent_colors        || [],
    chart_colors:         brand.chart_colors         || [],
    chart_color_sequence: brand.chart_color_sequence || brand.chart_colors || [],
    all_colors:           brand.all_colors           || {},
    title_font:           brand.title_font           || {},
    body_font:            brand.body_font            || {},
    caption_font:         brand.caption_font         || {},
    typography_hierarchy: brand.typography_hierarchy || {},
    bullet_style:         brand.bullet_style         || { char: '•', indent_inches: 0.12, space_before_pt: 4 },
    insight_box_style:    brand.insight_box_style    || { fill_color: null, border_color: null, corner_radius: 4 },
    visual_style:         brand.visual_style         || 'corporate',
    color_scheme_name:    brand.color_scheme_name    || '',
    logo_asset:           primaryLogo ? {
      name:        primaryLogo.name || 'logo',
      mime_type:   primaryLogo.mime_type || 'image/png',
      width_px:    primaryLogo.width_px || 0,
      height_px:   primaryLogo.height_px || 0
      // base64 intentionally excluded — large; logo rendering uses brand.primary_logo directly
    } : null,
    logo_local_ref:       brand.primary_logo_local_ref || '',
    logo_position:        brand.logo_position        || 'top-right',
    spacing_notes:        brand.spacing_notes        || '',
    uses_template:        brand.uses_template        || false,
    // Compact layout map: name → { title_placeholder, body_placeholder, ph_count, content_areas, usage_guidance }
    // content_areas: large body placeholders (h > 0.5") ordered left→right, top→bottom.
    // Used in LAYOUT MODE — the pipeline maps zone[i] → content_areas[i] for frame + placeholder_idx.
    layout_map:           (brand.slide_layouts || []).reduce((acc, l) => {
      if (l.name) {
        const contentAreas = (l.placeholders || [])
          .filter(p => p.type === 'body' && (p.h_in || 0) > 0.5)
          .sort((a, b) => {
            const rowA = Math.round((a.y_in || 0) * 2)  // 0.5" row buckets
            const rowB = Math.round((b.y_in || 0) * 2)
            if (rowA !== rowB) return rowA - rowB
            return (a.x_in || 0) - (b.x_in || 0)       // left before right
          })
          .map(p => ({ idx: p.idx, x_in: p.x_in, y_in: p.y_in, w_in: p.w_in, h_in: p.h_in }))
        acc[l.name] = {
          ph_count:          l.ph_count          || 0,
          usage_guidance:    l.usage_guidance    || '',
          title_placeholder: l.title_placeholder || null,
          body_placeholder:  l.body_placeholder  || null,
          content_areas:     contentAreas
        }
      }
      return acc
    }, {})
    // slide_masters, layout_blueprints, master_blueprints intentionally excluded — too large for API
  }
}

function buildBrandBrief(brand, brief) {
  const b = brief || {}
  const tokens = extractBrandTokens(brand)
  return 'BRAND DESIGN TOKENS:\n' +
    JSON.stringify(tokens, null, 2) +
    '\n\nBRAND COMPLIANCE RULES (MUST follow exactly):' +
    '\n- Fonts: use title_font.family for all titles/headings, body_font.family for all body/bullet text' +
    '\n- Title font size: typography_hierarchy.title_size_pt (content slides), larger for title/divider slides' +
    '\n- Body font size: typography_hierarchy.body_size_pt — do NOT guess; use the extracted value' +
    '\n- Chart colors: use chart_color_sequence in order — each series/segment gets a DIFFERENT color' +
    '\n- Pie charts: series_style must have one entry PER DATA POINT (category), each with a unique fill_color' +
    '\n- Bullet char: bullet_style.char — use exactly this character, not substitutes' +
    '\n- Bullet spacing: bullet_style.space_before_pt — pass directly into body_style.space_before_pt' +
    '\n- Insight boxes: insight_box_style.fill_color and border_color — a left accent bar is always rendered; do NOT add a full perimeter border' +
    (tokens.uses_template
      ? '\n- TEMPLATE MODE ACTIVE: master provides background/logo/footer — set global_elements:{}, canvas.background:null' +
        '\n- Title/divider slides: text only on title_block/subtitle_block — omit x/y/w/h, set layout_mode:true' +
        '\n- Content slides with selected_layout_name: LAYOUT MODE — set layout_mode:true, zone.frame:null; do NOT set placeholder_idx (pipeline assigns from layout content_areas)' +
        '\n- Content slides without selected_layout_name: SCRATCH MODE — compute all coordinates from layout_hint'
      : '\n- SCRATCH MODE: compute all coordinates; specify background, footer, logo in global_elements') +
    '\n\nPRESENTATION BRIEF:' +
    '\nDocument type:     ' + (b.document_type     || 'Business document') +
    '\nGoverning thought: ' + (b.governing_thought || 'Key insights from the document') +
    '\nNarrative flow:    ' + (b.narrative_flow    || 'Situation to Recommendation') +
    '\nTone:              ' + (b.tone              || 'professional') +
    '\nData heavy:        ' + (b.data_heavy        ? 'yes' : 'no') +
    '\nLogo policy:       Use the provided logo asset when available and keep it inside safe margins'
}


// ═══════════════════════════════════════════════════════════════════════════════
// BATCH WRITER
// Sends one batch of slides to Claude and returns the array of layout specs
// ═══════════════════════════════════════════════════════════════════════════════

async function designSlideBatch(batchManifest, brand, brief, batchNum) {
  const slideNums = batchManifest.map(s => s.slide_number)
  console.log('Agent 5 batch', batchNum, ':', slideNums.join(', '))

  // Annotate each slide with its mode so Claude doesn't have to infer it
  const annotatedManifest = batchManifest.map(s => ({
    ...s,
    _mode: (brand.uses_template && s.selected_layout_name)
      ? 'layout_mode'
      : (brand.uses_template && (s.slide_type === 'title' || s.slide_type === 'divider'))
        ? 'template_title_divider'
        : 'scratch_mode'
  }))

  const prompt =
    buildBrandBrief(brand, brief) +
    '\n\nSLIDE BATCH ' + batchNum + ' (' + annotatedManifest.length + ' slides):\n' +
    JSON.stringify(annotatedManifest, null, 2) +
    '\n\nINSTRUCTIONS:' +
    '\n- Process ONLY these ' + batchManifest.length + ' slides' +
    '\n- Apply brand design tokens exactly' +
    '\n- Compute exact coordinates for every element (2 decimal places)' +
    '\n- FULLY specify all artifacts including all style sub-objects' +
    '\n- chart: must have chart_style and series_style[]' +
    '\n- workflow: must have workflow_style, nodes[] with x/y/w/h, connections[] with path[]' +
    '\n- table: must have table_style, column_widths[], row_heights[]' +
    '\n- cards: must have card_style, card_frames[] with x/y/w/h per card' +
    '\n- insight_text (standard mode): must have insight_mode:"standard", style, heading_style, body_style' +
    '\n- insight_text (grouped mode):  must have insight_mode:"grouped", heading_style, group_layout, group_header_style, group_bullet_box_style, bullet_style, group_gap_in, header_to_box_gap_in' +
    '\n- Return a valid JSON array of exactly ' + batchManifest.length + ' slide objects'

  const raw    = await callClaude(AGENT5_SYSTEM, [{ role: 'user', content: prompt }], 6000)
  const parsed = safeParseJSON(raw, null)

  if (!Array.isArray(parsed)) {
    console.warn('Agent 5 batch', batchNum, '-- parse failed. Raw length:', raw.length)
    console.warn('  First 400 chars:', raw.slice(0, 400))
    return null
  }

  console.log('Agent 5 batch', batchNum, '-- got', parsed.length, 'slides')
  return parsed
}


// ═══════════════════════════════════════════════════════════════════════════════
// SLIDE VALIDATOR
// Returns list of structural issues. Empty array = valid.
// ═══════════════════════════════════════════════════════════════════════════════

function validateDesignedSlide(slide) {
  const issues = []

  if (!slide.canvas)            issues.push('missing canvas')
  if (!slide.brand_tokens)      issues.push('missing brand_tokens')
  if (!slide.title_block)       issues.push('missing title_block')
  if (!slide.title_block?.text) issues.push('empty title')

  if (slide.slide_type === 'content') {
    if (!slide.zones || slide.zones.length === 0) issues.push('no zones')
  }

  ;(slide.zones || []).forEach((z, zi) => {
    // In layout mode, frames are filled post-process from content_areas — skip the check
    if (!z.frame && !slide.layout_mode) issues.push('z' + zi + ': missing frame')
    if (!z.artifacts?.length)           issues.push('z' + zi + ': no artifacts')
    ;(z.artifacts || []).forEach((a, ai) => {
      const p = 'z' + zi + '.a' + ai
      if (!a.type)                                     issues.push(p + ': missing type')
      if (a.type === 'chart'    && !a.chart_style)     issues.push(p + ': chart missing chart_style')
      if (a.type === 'chart'    && !a.series_style)    issues.push(p + ': chart missing series_style')
      if (a.type === 'workflow' && !a.nodes?.length)   issues.push(p + ': workflow missing nodes')
      if (a.type === 'workflow' && !a.workflow_style)  issues.push(p + ': workflow missing workflow_style')
      if (a.type === 'table'    && !a.table_style)     issues.push(p + ': table missing table_style')
      if (a.type === 'table'    && !a.column_widths)   issues.push(p + ': table missing column_widths')
      if (a.type === 'cards'    && !a.card_frames?.length) issues.push(p + ': cards missing card_frames')
      if (a.type === 'cards'    && !a.card_style)      issues.push(p + ': cards missing card_style')
      if (a.type === 'insight_text' && !a.heading_style) issues.push(p + ': insight_text missing heading_style')
      if (a.type === 'insight_text' && a.insight_mode === 'grouped' && !a.group_header_style) issues.push(p + ': grouped insight_text missing group_header_style')
      if (a.type === 'insight_text' && a.insight_mode === 'grouped' && !a.group_bullet_box_style) issues.push(p + ': grouped insight_text missing group_bullet_box_style')
      if (a.type === 'insight_text' && a.insight_mode === 'grouped' && !a.bullet_style) issues.push(p + ': grouped insight_text missing bullet_style')
      if (a.type === 'insight_text' && (!a.insight_mode || a.insight_mode === 'standard') && !a.body_style) issues.push(p + ': insight_text missing body_style')
    })
  })

  if (slide.global_elements?.logo?.show && !slide.global_elements.logo.image_base64) {
    issues.push('logo missing image_base64')
  }

  return issues
}

function clamp(n, min, max) {
  return Math.min(Math.max(n, min), max)
}

function rectWithin(rect, bounds) {
  const x = clamp(rect.x || 0, bounds.x, bounds.x + Math.max(0, bounds.w - 0.05))
  const y = clamp(rect.y || 0, bounds.y, bounds.y + Math.max(0, bounds.h - 0.05))
  const maxW = Math.max(0.05, bounds.x + bounds.w - x)
  const maxH = Math.max(0.05, bounds.y + bounds.h - y)
  return {
    ...rect,
    x: r2(x),
    y: r2(y),
    w: r2(clamp(rect.w || maxW, 0.05, maxW)),
    h: r2(clamp(rect.h || maxH, 0.05, maxH))
  }
}

function getZoneInnerBounds(zone) {
  const frame = zone.frame || {}
  const p = frame.padding || {}
  return {
    x: r2((frame.x || 0) + (p.left || 0)),
    y: r2((frame.y || 0) + (p.top || 0)),
    w: r2(Math.max(0.1, (frame.w || 0) - (p.left || 0) - (p.right || 0))),
    h: r2(Math.max(0.1, (frame.h || 0) - (p.top || 0) - (p.bottom || 0)))
  }
}

function normalizeTableSizing(artifact) {
  const cols = artifact.column_widths || []
  const rows = artifact.row_heights || []
  const totalW = cols.reduce((s, n) => s + (+n || 0), 0)
  const totalH = rows.reduce((s, n) => s + (+n || 0), 0)
  const colCount = Math.max(1, (artifact.headers || []).length || cols.length)
  const rowCount = Math.max(1, ((artifact.rows || []).length + 1) || rows.length)

  artifact.column_widths = totalW > 0
    ? cols.map(v => r2((+v || 0) * artifact.w / totalW))
    : Array.from({ length: colCount }, () => r2(artifact.w / colCount))

  artifact.row_heights = totalH > 0
    ? rows.map(v => r2((+v || 0) * artifact.h / totalH))
    : Array.from({ length: rowCount }, () => r2(artifact.h / rowCount))

  const ts = artifact.table_style || {}
  const density = rowCount * colCount
  const bodySize = density > 30 ? 8.5 : density > 20 ? 9 : (ts.body_font_size || 10)
  artifact.table_style = {
    ...ts,
    body_font_size: Math.max(8, bodySize),
    header_font_size: Math.max(9, Math.min(ts.header_font_size || 11, bodySize + 1))
  }
}

function enforceArtifactBounds(zone) {
  const inner = getZoneInnerBounds(zone)
  zone.artifacts = (zone.artifacts || []).map(artifact => {
    const a = { ...artifact }

    if (['insight_text', 'chart', 'table'].includes(a.type)) {
      Object.assign(a, rectWithin(a, inner))
    }

    if (a.type === 'insight_text') {
      if (a.insight_mode === 'grouped') {
        // For grouped mode: scale bullet font based on total bullet count across all groups
        const groups = a.groups || []
        const totalBullets = groups.reduce((s, g) => s + (g.bullets || []).length, 0)
        const maxBullets = groups.reduce((m, g) => Math.max(m, (g.bullets || []).length), 0)
        const baseSize = (a.bullet_style || {}).font_size || 10
        const fitted = totalBullets > 20 || maxBullets > 6 ? Math.max(8, baseSize - 2)
          : totalBullets > 12 || maxBullets > 4 ? Math.max(8.5, baseSize - 1)
          : baseSize
        a.bullet_style = { ...(a.bullet_style || {}), font_size: fitted }
        // Clamp spacing proportional to artifact dimensions — no hardcoded defaults
        const dimForGap = a.group_layout === 'rows' ? (a.h || 5) : (a.w || 10)
        const minGap = r2(Math.max(dimForGap * 0.01, 0.04))
        const maxGap = r2(Math.min(dimForGap * 0.03, 0.18))
        a.group_gap_in = r2(Math.min(maxGap, Math.max(minGap, a.group_gap_in || dimForGap * 0.015)))
        const f = (a.bullet_style || {}).font_size || 10
        a.header_to_box_gap_in = r2(Math.min(0.10, Math.max(f * 0.5 / 72, 0.02)))
      } else {
        const pointCount = (a.points || []).length
        const avgChars = pointCount ? Math.round((a.points || []).join(' ').length / pointCount) : 0
        const baseSize = (a.body_style || {}).font_size || 10
        const fitted = pointCount > 6 || avgChars > 120 ? Math.max(8, baseSize - 2)
          : pointCount > 4 || avgChars > 80 ? Math.max(8.5, baseSize - 1)
          : baseSize
        a.body_style = { ...(a.body_style || {}), font_size: fitted }
      }
    }

    if (a.type === 'cards') {
      const container = rectWithin(a.container || inner, inner)
      a.container = container
      a.card_frames = (a.card_frames || []).map(frame => rectWithin(frame, container))
      const longestBody = Math.max(0, ...(a.cards || []).map(c => String(c.body || '').length))
      const bodyBase = (a.body_style || {}).font_size || 10
      a.body_style = { ...(a.body_style || {}), font_size: longestBody > 180 ? Math.max(8, bodyBase - 1.5) : bodyBase }
    }

    if (a.type === 'workflow') {
      const container = rectWithin(a.container || inner, inner)
      a.container = container
      a.nodes = (a.nodes || []).map(node => rectWithin(node, container))
      a.connections = (a.connections || []).map(conn => ({
        ...conn,
        path: (conn.path || []).map(pt => ({
          x: r2(clamp(pt.x || 0, container.x, container.x + container.w)),
          y: r2(clamp(pt.y || 0, container.y, container.y + container.h))
        }))
      }))
      const nodeCount = (a.nodes || []).length
      if (nodeCount >= 5) {
        a.workflow_style = {
          ...(a.workflow_style || {}),
          node_title_font_size: Math.max(8, ((a.workflow_style || {}).node_title_font_size || 11) - 1),
          node_value_font_size: Math.max(7, ((a.workflow_style || {}).node_value_font_size || 10) - 1)
        }
      }
    }

    if (a.type === 'table') {
      normalizeTableSizing(a)
    }

    return a
  })
  return zone
}

function buildLogoElement(slide, brand) {
  const logo = brand.primary_logo || ((brand.logos || [])[0] || null)
  if (!logo || !logo.base64) return null

  const canvas = slide.canvas || {}
  const margin = canvas.margin || {}
  const slideW = canvas.width_in || brand.slide_width_inches || 13.33
  const slideH = canvas.height_in || brand.slide_height_inches || 7.5
  const maxW = slide.slide_type === 'title' ? 1.8 : 1.35
  const maxH = slide.slide_type === 'title' ? 0.7 : 0.45
  const ratio = (logo.width_px && logo.height_px) ? (logo.width_px / Math.max(logo.height_px, 1)) : 3
  let w = maxW
  let h = r2(w / Math.max(ratio, 0.5))
  if (h > maxH) {
    h = maxH
    w = r2(h * Math.max(ratio, 0.5))
  }

  const pos = (brand.logo_position || 'top-right').toLowerCase()
  const left = margin.left || 0.4
  const right = slideW - (margin.right || 0.4) - w
  const top = margin.top || 0.15
  const bottom = slideH - (margin.bottom || 0.3) - h

  let x = right
  let y = top
  if (pos.includes('left')) x = left
  if (pos.includes('bottom')) y = bottom
  if (pos.includes('center')) x = r2((slideW - w) / 2)

  return {
    show: true,
    x: r2(x),
    y: r2(y),
    w: r2(w),
    h: r2(h),
    image_base64: logo.base64,
    image_mime_type: logo.mime_type || 'image/png',
    preserve_aspect_ratio: true
  }
}

function applyBrandGuidelineOverrides(slide, manifestSlide, brand) {
  if (!slide || !brand) return slide

  const normalized = JSON.parse(JSON.stringify(slide))
  normalized.global_elements = normalized.global_elements || {}

  // In layout mode or template title/divider, coordinates are driven by the template.
  // Skip bounds enforcement — enforcing would corrupt placeholder-derived positions.
  const isLayoutMode = normalized.layout_mode === true ||
    (brand.uses_template && (normalized.slide_type === 'title' || normalized.slide_type === 'divider'))

  if (isLayoutMode) {
    // Just ensure text content is preserved from manifest
    if (manifestSlide?.title && normalized.title_block) {
      normalized.title_block.text = normalized.title_block.text || manifestSlide.title
    }
    if (manifestSlide?.subtitle && normalized.subtitle_block) {
      normalized.subtitle_block.text = normalized.subtitle_block.text || manifestSlide.subtitle
    }
    return normalized
  }

  // Scratch mode: apply full bounds enforcement
  const logo = buildLogoElement(normalized, brand)
  if (logo) {
    normalized.global_elements.logo = logo
    const tb = normalized.title_block
    if (tb && tb.y < logo.y + logo.h + 0.1 && tb.x < logo.x + logo.w) {
      tb.w = r2(Math.max(1.5, Math.min(tb.w || (logo.x - tb.x), logo.x - tb.x - 0.18)))
    }
  }

  const slideW = normalized.canvas?.width_in || brand.slide_width_inches || 13.33
  const slideH = normalized.canvas?.height_in || brand.slide_height_inches || 7.5
  const margin = normalized.canvas?.margin || { left: 0.4, right: 0.4, top: 0.15, bottom: 0.3 }
  const contentBounds = {
    x: margin.left || 0.4,
    y: margin.top || 0.15,
    w: slideW - (margin.left || 0.4) - (margin.right || 0.4),
    h: slideH - (margin.top || 0.15) - (margin.bottom || 0.3)
  }

  normalized.title_block = normalized.title_block ? rectWithin(normalized.title_block, contentBounds) : normalized.title_block
  normalized.subtitle_block = normalized.subtitle_block ? rectWithin(normalized.subtitle_block, contentBounds) : normalized.subtitle_block

  normalized.zones = (normalized.zones || []).map(zone => {
    const z = { ...zone, frame: rectWithin(zone.frame || contentBounds, contentBounds) }
    return enforceArtifactBounds(z)
  })

  if (manifestSlide?.title && normalized.title_block) normalized.title_block.text = normalized.title_block.text || manifestSlide.title
  if (manifestSlide?.subtitle && normalized.subtitle_block) normalized.subtitle_block.text = normalized.subtitle_block.text || manifestSlide.subtitle

  return normalized
}


// ═══════════════════════════════════════════════════════════════════════════════
// CLAUDE-BASED FALLBACK
// When a slide fails (bad parse, missing artifacts, truncation) — ask Claude to
// build the best possible layout for just that one slide.
// Uses a tight focused prompt: no batch overhead, brand tokens only, one slide.
// ═══════════════════════════════════════════════════════════════════════════════

var AGENT5_FALLBACK_SYSTEM = `You are a senior presentation designer.
A slide failed to render correctly and needs to be rebuilt from scratch.

Build the best possible board-ready corporate layout for this single slide.
Use the brand design tokens exactly — colors, fonts, slide size.
Return a single valid JSON object (not an array) matching the full slide schema.

CRITICAL — all artifacts must be FULLY specified:
- chart: include chart_style{} AND series_style[]
- workflow: include workflow_style{}, nodes[] with x/y/w/h, connections[] with path[]
- table: include table_style{}, column_widths[], row_heights[]
- cards: include card_style{}, card_frames[] with x/y/w/h per card
- insight_text standard: include insight_mode:"standard", style{}, heading_style{}, body_style{}
- insight_text grouped:  include insight_mode:"grouped", heading_style{}, group_layout, group_header_style{}, group_bullet_box_style{}, bullet_style{}, group_gap_in, header_to_box_gap_in

All coordinates in decimal inches, 2 decimal places.
Return ONLY a valid JSON object. No explanation. No markdown.`

async function buildFallbackDesign(manifestSlide, brand, brief) {
  console.log('Agent 5 -- fallback Claude call for S' + manifestSlide.slide_number)

  const tokens = extractBrandTokens(brand)
  const b      = brief || {}

  const prompt =
    'BRAND DESIGN TOKENS:\n' +
    JSON.stringify(tokens, null, 2) +
    '\n\nSLIDE TO REBUILD:\n' +
    JSON.stringify(manifestSlide, null, 2) +
    '\n\nContext:' +
    '\nDocument type: ' + (b.document_type || 'Business document') +
    '\nTone: ' + (b.tone || 'professional') +
    '\n\nBuild the best possible layout for this slide.' +
    '\nPreserve the title, key_message, zones structure and artifact content from the manifest.' +
    '\nChoose the cleanest, most board-ready layout given the archetype: ' + (manifestSlide.slide_archetype || 'summary') +
    '\nReturn a single JSON object for this one slide.'

  try {
    const raw    = await callClaude(AGENT5_FALLBACK_SYSTEM, [{ role: 'user', content: prompt }], 3000)
    const parsed = safeParseJSON(raw, null)

    // Claude may return an array with one item or a bare object
    const candidate = Array.isArray(parsed) ? parsed[0] : parsed

    if (candidate && typeof candidate === 'object' && candidate.canvas && candidate.zones) {
      const issues = validateDesignedSlide(candidate)
      if (issues.length === 0) {
        console.log('Agent 5 -- fallback Claude succeeded for S' + manifestSlide.slide_number)
        return candidate
      }
      console.warn('Agent 5 -- fallback Claude still has issues for S' + manifestSlide.slide_number + ':', issues.join('; '))
      // Return anyway -- better than nothing, zone repair will patch what it can
      return candidate
    }
  } catch (e) {
    console.warn('Agent 5 -- fallback Claude call failed for S' + manifestSlide.slide_number + ':', e.message)
  }

  // Last resort: minimal structurally valid object derived purely from brand tokens
  // No hardcoded content — pull everything from manifest and brand
  return buildMinimalSafeSlide(manifestSlide, tokens)
}


// ─── Minimal safe slide — only used if the fallback Claude call itself fails ──
// This is the true last resort. It still uses real brand values and real content.
function buildMinimalSafeSlide(manifestSlide, tokens) {
  const w         = r2(tokens.slide_width_inches  || 13.33)
  const h         = r2(tokens.slide_height_inches || 7.50)
  const primary   = (tokens.primary_colors    || ['#1A3C8F'])[0]
  const secondary = (tokens.secondary_colors  || ['#E8A020'])[0]
  const bg        = (tokens.background_colors || ['#FFFFFF'])[0]
  const titleFont = (tokens.title_font   || {}).family || 'Calibri'
  const bodyFont  = (tokens.body_font    || {}).family || 'Calibri'
  const isDark    = ['title', 'divider'].includes(manifestSlide.slide_type)

  const ct  = 1.05
  const cw  = r2(w - 0.80)
  const ch  = r2(h - ct - 0.40)

  // Pull real key_message points from the manifest zones if available
  const manifestPoints = []
  ;(manifestSlide.zones || []).forEach(z =>
    (z.artifacts || []).forEach(a => {
      if (a.type === 'insight_text') (a.points || []).forEach(p => manifestPoints.push(p))
    })
  )
  const bodyPoints = manifestPoints.length > 0
    ? manifestPoints.slice(0, 4)
    : [manifestSlide.key_message || 'See source document for details']

  return {
    slide_number:    manifestSlide.slide_number,
    slide_type:      manifestSlide.slide_type      || 'content',
    slide_archetype: manifestSlide.slide_archetype || 'summary',
    canvas: {
      width_in: w, height_in: h,
      margin:     { left: 0.40, right: 0.40, top: 0.15, bottom: 0.30 },
      background: { color: isDark ? primary : bg }
    },
    brand_tokens: {
      title_font_family:   titleFont,
      body_font_family:    bodyFont,
      caption_font_family: bodyFont,
      title_color:   isDark ? '#FFFFFF' : primary,
      body_color:    isDark ? '#CCDDFF' : '#111111',
      caption_color: '#888888',
      primary_color:   primary,
      secondary_color: secondary,
      accent_colors:   tokens.accent_colors || [],
      chart_palette:   tokens.chart_colors  || [primary, secondary, '#2E9E5B', '#C82333']
    },
    title_block: {
      text:        manifestSlide.title || '',
      x: 0.40, y: 0.15, w: cw,
      h:           isDark ? 2.00 : 0.75,
      font_family: titleFont,
      font_size:   isDark ? 30 : 18,
      font_weight: 'bold',
      color:       isDark ? '#FFFFFF' : primary,
      align: 'left', valign: 'middle', wrap: true
    },
    subtitle_block: manifestSlide.subtitle ? {
      text:        manifestSlide.subtitle,
      x: 0.40, y: isDark ? 2.60 : 0.95,
      w: cw, h: 0.45,
      font_family: bodyFont, font_size: 14, font_weight: 'regular',
      color:       isDark ? '#BBCCFF' : '#555555',
      align: 'left', valign: 'top', wrap: true
    } : null,
    zones: manifestSlide.slide_type === 'content' ? [{
      zone_id:           'z1',
      zone_role:         'primary_proof',
      message_objective: manifestSlide.key_message || '',
      narrative_weight:  'primary',
      frame: {
        x: 0.40, y: ct, w: cw, h: ch,
        padding: { top: 0.10, right: 0.10, bottom: 0.10, left: 0.10 }
      },
      artifacts: [{
        type: 'insight_text',
        x: 0.50, y: r2(ct + 0.10), w: r2(cw - 0.20), h: r2(ch - 0.20),
        style:         { fill_color: null, border_color: primary + '33', border_width: 0.5, corner_radius: 3 },
        heading_style: { font_family: titleFont, font_size: 12, font_weight: 'bold', color: primary },
        body_style:    { font_family: bodyFont, font_size: 11, font_weight: 'regular', color: '#111111', line_spacing: 1.4, bullet_indent: 0.15 },
        // Pass actual points through for Agent 6 to render
        heading: 'Key Insight',
        points:  bodyPoints,
        sentiment: 'neutral'
      }]
    }] : [],
    global_elements: {
      footer: {
        show: true, x: 0.40, y: r2(h - 0.26), w: 3.00, h: 0.20,
        font_family: bodyFont, font_size: 8, color: '#AAAAAA', align: 'left'
      },
      page_number: {
        show: true, x: r2(w - 0.88), y: r2(h - 0.26), w: 0.65, h: 0.20,
        font_family: bodyFont, font_size: 8, color: '#AAAAAA', align: 'right'
      }
    },
    _fallback: true
  }
}


// ═══════════════════════════════════════════════════════════════════════════════
// CONTENT MERGER
// Injects Agent 4 manifest artifact content into Agent 5 designed artifacts.
// Agent 5 produces layout + style. Agent 4 holds the actual data.
// This merge produces a single self-contained object Agent 6 can render directly.
//
// Matching strategy: zone index first, then zone_id string match as fallback.
// Within a zone: artifact matched by position index (Agent 4 and Agent 5 should
// produce the same number of artifacts per zone in the same order).
//
// Content fields injected per artifact type:
//   insight_text : heading, points[], sentiment
//   chart        : chart_type, chart_title, chart_insight, x_label, y_label,
//                  categories[], series[], show_data_labels, show_legend
//   cards        : cards[] (title, subtitle, body, sentiment per card)
//   workflow     : workflow_type, flow_direction, workflow_title, workflow_insight,
//                  node labels/values/descriptions/levels, connection from/to/type
//   table        : title, headers[], rows[][], highlight_rows[], note
// ═══════════════════════════════════════════════════════════════════════════════

function mergeContentIntoZones(designedZones, manifestZones) {
  if (!designedZones || !manifestZones) return designedZones || []

  return designedZones.map((dZone, zi) => {
    // Find the matching manifest zone — by index first, then by zone_id
    const mZone = manifestZones[zi]
               || manifestZones.find(z => z.zone_id === dZone.zone_id)
    if (!mZone) return dZone

    const mergedArtifacts = (dZone.artifacts || []).map((dArt, ai) => {
      // Match manifest artifact by position index
      const mArt = (mZone.artifacts || [])[ai]
      if (!mArt || mArt.type !== dArt.type) return dArt

      const t = dArt.type

      if (t === 'insight_text') {
        return {
          ...dArt,
          heading:        mArt.heading        || mArt.insight_header || dArt.heading   || 'Key Insight',
          insight_header: mArt.insight_header || mArt.heading        || dArt.insight_header || 'Key Insight',
          points:         mArt.points         || dArt.points         || [],
          sentiment:      mArt.sentiment      || dArt.sentiment      || 'neutral'
        }
      }

      if (t === 'chart') {
        return {
          ...dArt,
          chart_type:       mArt.chart_type       || dArt.chart_type       || 'bar',
          chart_header:     mArt.chart_header     || dArt.chart_header     || '',
          chart_title:      mArt.chart_title      || dArt.chart_title      || '',
          chart_insight:    mArt.chart_insight    || dArt.chart_insight    || '',
          x_label:          mArt.x_label          || dArt.x_label          || '',
          y_label:          mArt.y_label          || dArt.y_label          || '',
          categories:       mArt.categories       || dArt.categories       || [],
          series:           mArt.series           || dArt.series           || [],
          show_data_labels: mArt.show_data_labels !== undefined
                              ? mArt.show_data_labels : (dArt.show_data_labels !== false),
          show_legend:      mArt.show_legend      !== undefined
                              ? mArt.show_legend      : !!dArt.show_legend
        }
      }

      if (t === 'cards') {
        return {
          ...dArt,
          cards: mArt.cards || dArt.cards || []
        }
      }

      if (t === 'workflow') {
        // Merge node content (label, value, description, level) into designed nodes
        // Designed nodes have x/y/w/h; manifest nodes have the text content.
        const designedNodes   = dArt.nodes   || []
        const manifestNodes   = mArt.nodes   || []
        const mergedNodes = designedNodes.map((dn, ni) => {
          const mn = manifestNodes.find(n => n.id === dn.id) || manifestNodes[ni]
          if (!mn) return dn
          return {
            ...dn,                        // keep x, y, w, h from Agent 5
            label:       mn.label       || dn.label       || dn.id,
            value:       mn.value       || dn.value       || '',
            description: mn.description || dn.description || '',
            level:       mn.level       !== undefined ? mn.level : (dn.level || 1)
          }
        })

        // Merge connection from/to/type from manifest into designed connections
        // Designed connections have path[] waypoints; manifest has from/to/type.
        const designedConns  = dArt.connections || []
        const manifestConns  = mArt.connections || []
        const mergedConns = designedConns.map((dc, ci) => {
          const mc = manifestConns[ci]
                  || manifestConns.find(c => c.from === dc.from && c.to === dc.to)
          if (!mc) return dc
          return {
            ...dc,                        // keep path[] from Agent 5
            from: mc.from || dc.from,
            to:   mc.to   || dc.to,
            type: mc.type || dc.type || 'arrow'
          }
        })

        return {
          ...dArt,
          workflow_type:    mArt.workflow_type    || dArt.workflow_type    || 'process_flow',
          workflow_header:  mArt.workflow_header  || dArt.workflow_header  || '',
          flow_direction:   mArt.flow_direction   || dArt.flow_direction   || 'left_to_right',
          workflow_title:   mArt.workflow_title   || dArt.workflow_title   || '',
          workflow_insight: mArt.workflow_insight || dArt.workflow_insight || '',
          nodes:            mergedNodes,
          connections:      mergedConns
        }
      }

      if (t === 'table') {
        return {
          ...dArt,
          table_header:   mArt.table_header   || dArt.table_header   || '',
          title:          mArt.title          || dArt.title          || '',
          headers:        mArt.headers        || dArt.headers        || [],
          rows:           mArt.rows           || dArt.rows           || [],
          highlight_rows: mArt.highlight_rows || dArt.highlight_rows || [],
          note:           mArt.note           || dArt.note           || ''
        }
      }

      return dArt
    })

    return { ...dZone, artifacts: mergedArtifacts }
  })
}


// ═══════════════════════════════════════════════════════════════════════════════
// NORMALISER
// 1. Validates the designed slide structure
// 2. Merges Agent 4 artifact content into Agent 5 layout artifacts
// 3. Carries all manifest metadata through for Agent 5.1 and Agent 6
// ═══════════════════════════════════════════════════════════════════════════════

// Fills zone frames and artifact placeholder_idx from the layout's ordered content_areas.
// Called AFTER mergeContentIntoZones so artifact content comes from Agent 4, not Agent 5.
function applyLayoutZoneFrames(zones, layoutName, brand) {
  if (!layoutName) return zones
  const layouts = brand.slide_layouts || []
  const layout = layouts.find(l => (l.name || '').toLowerCase() === layoutName.toLowerCase())
    || layouts.find(l => (l.name || '').toLowerCase().includes(layoutName.toLowerCase()))
  if (!layout) return zones

  // Build ordered content areas on-the-fly (same logic as extractBrandTokens)
  const contentAreas = (layout.placeholders || [])
    .filter(p => p.type === 'body' && (p.h_in || 0) > 0.5)
    .sort((a, b) => {
      const rowA = Math.round((a.y_in || 0) * 2)
      const rowB = Math.round((b.y_in || 0) * 2)
      if (rowA !== rowB) return rowA - rowB
      return (a.x_in || 0) - (b.x_in || 0)
    })

  if (contentAreas.length === 0) return zones

  // Small body placeholders (h ≤ 0.5") are header labels, not content areas
  const headerPhs = (layout.placeholders || [])
    .filter(p => p.type === 'body' && (p.h_in || 0) <= 0.5)

  return zones.map((zone, zi) => {
    const ca = contentAreas[zi] || contentAreas[contentAreas.length - 1]
    const frame = {
      x: ca.x_in, y: ca.y_in, w: ca.w_in, h: ca.h_in,
      padding: { top: 0.08, right: 0.08, bottom: 0.08, left: 0.08 }
    }
    // Find paired header placeholder: same x-column as content area, positioned just above it
    const headerPh = headerPhs.find(p =>
      Math.abs((p.x_in || 0) - (ca.x_in || 0)) < 0.15 && (p.y_in || 0) < (ca.y_in || 0)
    )
    const artifacts = (zone.artifacts || []).map(a => ({ ...a, placeholder_idx: ca.idx }))
    return {
      ...zone,
      frame,
      header_ph_idx: headerPh ? headerPh.idx : null,
      artifacts
    }
  })
}

function normaliseDesignedSlide(designed, manifestSlide, brand) {
  if (!designed || typeof designed !== 'object') return null  // caller handles null -> fallback

  const branded = applyBrandGuidelineOverrides(designed, manifestSlide, brand)
  const issues = validateDesignedSlide(branded)
  if (issues.length > 0) {
    console.warn('Agent 5 -- S' + (branded.slide_number || '?') + ' issues:', issues.join('; '))
  }

  // Merge Agent 4 content into Agent 5 layout zones
  const mergedZones = mergeContentIntoZones(
    branded.zones || [],
    manifestSlide.zones || []
  )

  // Layout mode: fill zone frames + artifact placeholder_idx from the layout's content_areas.
  // This runs after merge so Agent 4's artifact content is already in place.
  const layoutName = manifestSlide.selected_layout_name || designed.selected_layout_name || ''
  const isLayoutMode = !!(designed.layout_mode || layoutName)
  const finalZones = isLayoutMode && layoutName
    ? applyLayoutZoneFrames(mergedZones, layoutName, brand)
    : mergedZones

  // Log merge summary
  const contentCounts = { insight_text: 0, chart: 0, cards: 0, workflow: 0, table: 0 }
  finalZones.forEach(z => (z.artifacts || []).forEach(a => {
    if (contentCounts[a.type] !== undefined) contentCounts[a.type]++
  }))
  console.log('  S' + manifestSlide.slide_number + ' merged content:',
    Object.entries(contentCounts).filter(([,n]) => n > 0).map(([t,n]) => t + ':' + n).join(' ') || 'none')

  return {
    ...branded,
    zones: finalZones,
    // Always override with manifest ground truth (Claude may drift on slide_number etc.)
    slide_number:          manifestSlide.slide_number,
    slide_type:            manifestSlide.slide_type            || designed.slide_type,
    slide_archetype:       manifestSlide.slide_archetype       || designed.slide_archetype || 'summary',
    // Layout mode fields — ground truth from Agent 4 manifest
    layout_mode:           designed.layout_mode                || false,
    selected_layout_name:  manifestSlide.selected_layout_name  || designed.selected_layout_name || '',
    section_name:          manifestSlide.section_name          || '',
    section_type:          manifestSlide.section_type          || '',
    // Slide-level content metadata
    title:            (designed.title_block || {}).text || manifestSlide.title    || '',
    subtitle:         (designed.subtitle_block || {}).text || manifestSlide.subtitle || '',
    key_message:      manifestSlide.key_message      || '',
    visual_flow_hint: manifestSlide.visual_flow_hint || '',
    speaker_note:     manifestSlide.speaker_note     || '',
    layout_name:      inferLayoutName(manifestSlide, brand),
    // Condensed zone summary for Agent 5.1 review
    zones_summary:    mergedZones.map(z => ({
      zone_id:          z.zone_id,
      zone_role:        z.zone_role,
      narrative_weight: z.narrative_weight,
      artifact_types:   (z.artifacts || []).map(a => a.type)
    })),
    _validation_issues: issues.length > 0 ? issues : undefined
  }
}

function inferLayoutName(manifestSlide, brand) {
  const st   = manifestSlide.slide_type      || 'content'
  const arch = manifestSlide.slide_archetype || ''
  const _NON_CONTENT_TYPES = new Set(['title', 'sechead', 'blank'])
  const isNonContent = (l) => {
    const t = (l.type || '').toLowerCase()
    const n = (l.name || '').toLowerCase()
    return _NON_CONTENT_TYPES.has(t) ||
      /^blank$/i.test(n) ||
      /thank[\s_-]*you|end[\s_-]*slide|closing[\s_-]*slide|section[\s_-]*header|^section$|divider|title slide/i.test(n)
  }
  const allLayouts = brand.slide_layouts || []
  const avail = st === 'content' ? allLayouts.filter(l => !isNonContent(l)) : allLayouts
  const find = (kws) => avail.find(l => kws.some(k => (l.name || '').toLowerCase().includes(k.toLowerCase())))

  if (st === 'title')   return (find(['Title Slide', 'title'])                     || {}).name || 'Title Slide'
  if (st === 'divider') return (find(['Section', 'Divider', 'section header'])     || {}).name || 'Section Divider'
  if (['recommendation','process','roadmap'].includes(arch))
    return (find(['3 Across','3 across','body text'])                              || {}).name || 'Body Text'
  if (['dashboard','summary'].includes(arch))
    return (find(['2 Across','1 Across','2 across','1 across'])                   || {}).name || '1 Across'
  return (find(['1 Across','Body Text','1 across','body text','2 Column','2 column']) || {}).name || (avail[0] || {}).name || 'Body Text'
}


// ═══════════════════════════════════════════════════════════════════════════════
// MAIN RUNNER
// ═══════════════════════════════════════════════════════════════════════════════

async function runAgent5(state) {
  const manifest = state.slideManifest
  const brand    = state.brandRulebook
  const brief    = state.outline || {}

  if (!manifest || !manifest.length) {
    console.error('Agent 5 -- slideManifest is empty')
    return []
  }

  const tokens = extractBrandTokens(brand)
  console.log('Agent 5 starting -- slides:', manifest.length)
  console.log('  Primary color:', (tokens.primary_colors || [])[0] || 'none')
  console.log('  Slide size:', tokens.slide_width_inches + '" x ' + tokens.slide_height_inches + '"')
  console.log('  Title font:', (tokens.title_font || {}).family || 'default')
  console.log('  Deck type:', brief.document_type || 'unknown')

  // Batch into groups of 3 — smaller batches mean more token headroom per slide,
  // which is the primary cause of incomplete artifact specs
  const BATCH_SIZE = 3
  const batches    = []
  for (let i = 0; i < manifest.length; i += BATCH_SIZE) {
    batches.push(manifest.slice(i, i + BATCH_SIZE))
  }
  console.log('  Batches:', batches.length, '(max', BATCH_SIZE, 'slides each)')

  const allDesigned = []

  for (let b = 0; b < batches.length; b++) {
    if (b > 0) {
      console.log('Agent 5 -- rate limit pause: waiting 65s before batch', b + 1, '...')
      await new Promise(r => setTimeout(r, 65000))
    }
    const batch  = batches[b]
    const result = await designSlideBatch(batch, brand, brief, b + 1)

    if (!result) {
      // Entire batch failed to parse — fall back per slide via Claude
      console.warn('Agent 5 -- batch', b + 1, 'failed entirely, running per-slide fallbacks')
      for (const ms of batch) {
        const fb = await buildFallbackDesign(ms, brand, brief)
        allDesigned.push(normaliseDesignedSlide(fb, ms, brand) || buildMinimalSafeSlide(ms, tokens))
      }
      continue
    }

    // Match each manifest slide to the returned result
    for (let i = 0; i < batch.length; i++) {
      const mSlide = batch[i]
      const match  = result.find(r => r.slide_number === mSlide.slide_number)
                  || result[i]
                  || null

      if (!match) {
        console.warn('Agent 5 -- no match for S' + mSlide.slide_number + ', running fallback')
        const fb = await buildFallbackDesign(mSlide, brand, brief)
        allDesigned.push(normaliseDesignedSlide(fb, mSlide, brand) || buildMinimalSafeSlide(mSlide, tokens))
        continue
      }

      const issues = validateDesignedSlide(match)

      if (issues.length === 0) {
        // Fully valid — normalise and accept
        allDesigned.push(normaliseDesignedSlide(match, mSlide, brand))
      } else {
        // Has issues — check if they're critical (missing zones/artifacts) or cosmetic
        const critical = issues.filter(i =>
          i.includes('missing canvas') ||
          i.includes('missing brand_tokens') ||
          i.includes('no zones') ||
          i.includes('no artifacts') ||
          i.includes('missing chart_style') ||
          i.includes('missing series_style') ||
          i.includes('missing workflow_style') ||
          i.includes('missing nodes') ||
          i.includes('missing table_style') ||
          i.includes('missing card_frames') ||
          i.includes('missing heading_style') ||
          i.includes('missing body_style')
        )

        if (critical.length > 0) {
          // Critical structural gaps — fallback Claude call for this slide
          console.warn('Agent 5 -- S' + mSlide.slide_number + ' has critical issues, running fallback:', critical.join('; '))
          const fb = await buildFallbackDesign(mSlide, brand, brief)
          allDesigned.push(normaliseDesignedSlide(fb, mSlide, brand) || buildMinimalSafeSlide(mSlide, tokens))
        } else {
          // Minor issues (e.g. empty title) — accept Claude's work with warnings
          console.warn('Agent 5 -- S' + mSlide.slide_number + ' minor issues (accepting):', issues.join('; '))
          allDesigned.push(normaliseDesignedSlide(match, mSlide, brand))
        }
      }
    }
  }

  // Summary
  const withIssues    = allDesigned.filter(s => s._validation_issues?.length > 0)
  const withFallback  = allDesigned.filter(s => s._fallback)
  const typeCounts    = {}
  allDesigned.forEach(s =>
    (s.zones || []).forEach(z =>
      (z.artifacts || []).forEach(a => { typeCounts[a.type] = (typeCounts[a.type] || 0) + 1 })
    )
  )

  console.log('Agent 5 complete')
  console.log('  Slides:', allDesigned.length,
    '| fallback:', withFallback.length,
    '| with issues:', withIssues.length)
  console.log('  Artifact types:', JSON.stringify(typeCounts))

  return allDesigned
}
