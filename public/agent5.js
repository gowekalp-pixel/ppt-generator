// ─── AGENT 5 — SLIDE LAYOUT & VISUAL DESIGN ENGINE ───────────────────────────
// Input:  state.slideManifest   — output from Agent 4
//         state.brandRulebook   — brand guideline JSON from Agent 2
//
// Output: designedSpec — flat JSON array, one render-ready object per slide
//
// Architecture: Claude API call per batch of 3 slides.
// Claude receives the Agent 4 manifest + brand guideline and returns
// a precise layout spec: canvas, brand_tokens, title_block, subtitle_block,
// zones (with fully positioned artifacts), and global_elements.
//
// Agent 5.1 then reviews this spec and applies targeted fixes.
// Agent 6 (python-pptx) consumes the final reviewed spec.

// Defensive global fallback: some older browser-served bundles or indirect
// preview/eval paths may still reference `bt` without a local binding.
// Keeping a harmless global object prevents a hard ReferenceError while the
// local per-function brand token bindings continue to be the primary source.
if (typeof globalThis.bt === 'undefined') globalThis.bt = {}
var bt = globalThis.bt

// ═══════════════════════════════════════════════════════════════════════════════
// SYSTEM PROMPT
// ═══════════════════════════════════════════════════════════════════════════════

var AGENT5_SYSTEM = `You are a senior presentation designer and layout system architect.

You will receive:
1. A slide content manifest created by Agent 4
2. A brand guideline JSON for the current deck

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
The renderer is NOT a designer.
You must decide every render-critical detail yourself.
Do not leave spacing, fitting, alignment, or artifact internals for Agent 6.

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

Font sizes:
- title slides: 24–34 pt
- divider slides: 22–30 pt
- content slides: 16–22 pt
- subtitle is 9–16 pt smaller than title

TITLE BLOCK HEIGHT (scratch mode only — in template/layout mode omit x/y/w/h):
Compute h from actual line count — do NOT use a generous fixed height:

  chars_per_line ≈ (title_block.w × 72) / (font_size_pt × 0.52)
  n_lines = ceil(title_char_count / chars_per_line)
  title_block.h = n_lines × (font_size_pt × 1.35 / 72) + 0.12

  Example: 90-char title, 20pt, w=12.5"
    chars_per_line = (12.5 × 72) / (20 × 0.52) = 900 / 10.4 ≈ 86
    n_lines = ceil(90 / 86) = 2
    h = 2 × (20 × 1.35 / 72) + 0.12 = 2 × 0.375 + 0.12 = 0.87"

  CRITICAL: zone frame y must start at title_block.y + title_block.h + 0.08" (minimum gap)
  Never add extra buffer to title_block.h — the renderer compacts it automatically.

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
- show_gridlines must always be false; no chart should display gridlines

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

LEGEND POSITION (computed by renderer from chart-to-slide ratio — do not specify in spec):
- chart.w > 60% of slide width                       → legend RIGHT
- chart.h > 60% of slide height (and width ≤ 60%)   → legend TOP
- All other cases: pie chart → legend RIGHT; all other chart types → legend TOP
- Legend font size is automatically capped at chart_header font size — do not set legend_font_size above heading_style.font_size

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
CHART MICRO-LAYOUT OWNERSHIP:
- You must decide legend_position, data_label_size, and category_label_rotation in the spec
- You must decide all series colors, line widths, marker choices, and label colors
- Do NOT leave chart readability choices for the renderer

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
- card shape: square corners by default (corner_radius: 0). Do NOT use rounded pill cards.
- accent treatment: use a vertical accent strip on the LEFT side of the card, not a top strip
- when multiple cards exist in one artifact, vary accent-strip color using brand hierarchy:
    card 1 → primary_color
    card 2 → secondary_color
    card 3+ → accent_colors in order
  Only fall back to sentiment color when a single card stands alone or no brand sequence exists
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
WORKFLOW MICRO-LAYOUT OWNERSHIP:
- Reserve explicit whitespace for external value / description bands so labels do not collide with nodes or connectors
- You must decide the final node sizes, connector paths, node padding, and external text gaps

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
  "column_x_positions": [number],
  "header_row_height": number,
  "row_heights": [number],
  "row_y_positions": [number],
  "header_cell_frames": [
    { "col_index": number, "x": number, "y": number, "w": number, "h": number }
  ],
  "body_cell_frames": [
    [
      { "row_index": number, "col_index": number, "x": number, "y": number, "w": number, "h": number }
    ]
  ],
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
- column_x_positions must identify the exact x-start of each column within the table frame
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
- row_y_positions must identify the exact y-start of header row and each data row within the table frame
- Zebra striping: set body_alt_fill_color to a very light tint of brand secondary (e.g., "#F7F8FA") for alternating rows
- highlight_rows: apply highlight_fill_color from brand accent to the highlight_rows indices from Agent 4
- table_style.cell_padding: 0.05–0.08" (enforced by renderer; set as a hint here)

═══════════════════════════
TABLE MICRO-LAYOUT OWNERSHIP:
- You must output final column_widths, column_x_positions, header_row_height, row_heights, row_y_positions, column_types, column_alignments, header_cell_frames, and body_cell_frames
- The renderer must not infer table density, alignment, or spacing

6. MATRIX
═══════════════════════════

{
  "type": "matrix",
  "x": number, "y": number, "w": number, "h": number,
  "matrix_style": {
    "border_color": "hex",
    "border_width": number,
    "divider_color": "hex",
    "divider_width": number,
    "axis_color": "hex",
    "axis_width": number,
    "axis_label_font_family": "string",
    "axis_label_font_size": number,
    "axis_label_color": "hex",
    "quadrant_title_font_family": "string",
    "quadrant_title_font_size": number,
    "quadrant_title_color": "hex",
    "quadrant_body_font_family": "string",
    "quadrant_body_font_size": number,
    "quadrant_body_color": "hex",
    "point_label_font_family": "string",
    "point_label_font_size": number,
    "point_label_color": "hex",
    "point_palette": ["hex"],
    "quadrant_fills": ["hex", "hex", "hex", "hex"]
  },
  "header_block": null or {
    "text": "string — the matrix_header value from Agent 4",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string", "font_size": number, "font_weight": "semibold", "color": "hex",
    "style": "underline" | "brand_fill",
    "accent_color": "hex"
  }
}

Matrix rules:
- Use ONLY primitive geometry in the final blocks: rect, text_box, rule, circle
- Reserve space for:
  - x axis title + low/high labels
  - y axis title + low/high labels
  - quadrant titles + quadrant insight lines
- Quadrant fills must be subtle light tints with clear contrast against labels and points
- Plot point positions semantically:
  - low = 25% of axis span
  - medium = 50%
  - high = 75%
- Y increases upward conceptually: y=high must plot nearer the TOP of the matrix
- Point labels must not collide with quadrant headings; offset labels around the marker if needed
- Use brand colors for points in sequence; max 6 points

7. DRIVER_TREE
═══════════════════════════

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
    "connector_width": number,
    "label_font_family": "string",
    "label_font_size": number,
    "label_color": "hex",
    "value_font_family": "string",
    "value_font_size": number,
    "value_color": "hex",
    "corner_radius": number
  },
  "header_block": null or {
    "text": "string — the tree_header value from Agent 4",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string", "font_size": number, "font_weight": "semibold", "color": "hex",
    "style": "underline" | "brand_fill",
    "accent_color": "hex"
  }
}

Driver tree rules:
- Use ONLY primitive geometry in the final blocks: rect, text_box, rule
- Root node centered at top; branch nodes on next row; leaf nodes on final row
- Max 3 levels
- Use orthogonal connectors only: vertical + horizontal segments
- Node labels inside the node box; values as second line or lower text block inside the same box
- Root must be visually dominant, level 2 medium, leaves smallest
- Keep branch distribution symmetric across the container

8. PRIORITIZATION
═══════════════════════════

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
    "rank_text_color": "hex",
    "title_font_family": "string",
    "title_font_size": number,
    "title_color": "hex",
    "description_font_family": "string",
    "description_font_size": number,
    "description_color": "hex",
    "qualifier_fill_color": "hex",
    "qualifier_text_color": "hex",
    "qualifier_value_palette": ["hex"],
    "qualifier_label_font_family": "string",
    "qualifier_label_font_size": number,
    "qualifier_value_font_family": "string",
    "qualifier_value_font_size": number
  },
  "header_block": null or {
    "text": "string — the priority_header value from Agent 4",
    "x": number, "y": number, "w": number, "h": number,
    "font_family": "string", "font_size": number, "font_weight": "semibold", "color": "hex",
    "style": "underline" | "brand_fill",
    "accent_color": "hex"
  }
}

Prioritization rules:
- Use ONLY primitive geometry in the final blocks: rect, text_box, circle
- Rows must be stacked vertically in rank order
- Each row contains:
  - left rank badge
  - action title
  - action description
  - up to 2 qualifier pills on the right
- Qualifier slots may be empty; do not render empty pills
- Rank 1 should be visually strongest; later ranks may step down subtly through the rank palette
- Title must dominate description; qualifiers must remain compact, secondary metadata

ARTIFACT HEADER
═══════════════════════════

Every artifact except cards gets a header_block — a one-line label above the artifact
that names the insight it proves.

Source text from Agent 4:
  insight_text → insight_header   chart → chart_header
  workflow     → workflow_header  table → table_header
  matrix       → matrix_header    driver_tree → tree_header
  prioritization → priority_header

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
RENDER OWNERSHIP:
- Agent 5 owns every render-critical detail for every inch of the slide canvas
- Agent 6 is only allowed to render what Agent 5 decided

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
          title_placeholder: l.title_placeholder || (l.master_summary || {}).title_placeholder || null,
          body_placeholder:  l.body_placeholder  || null,
          content_areas:     contentAreas
        }
      }
      return acc
    }, {})
    // slide_masters, layout_blueprints, master_blueprints intentionally excluded — too large for API
  }
}

function buildBrandBrief(brand) {
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
    '\n\nLogo policy: Use the provided logo asset when available and keep it inside safe margins'
}


// ═══════════════════════════════════════════════════════════════════════════════
// BATCH WRITER
// Sends one batch of slides to Claude and returns the array of layout specs
// ═══════════════════════════════════════════════════════════════════════════════

async function designSlideBatch(batchManifest, brand, batchNum) {
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
    buildBrandBrief(brand) +
    '\n\nSLIDE BATCH ' + batchNum + ' (' + annotatedManifest.length + ' slides):\n' +
    JSON.stringify(annotatedManifest, null, 2) +
    '\n\nINSTRUCTIONS:' +
    '\n- Process ONLY these ' + batchManifest.length + ' slides' +
    '\n- Apply brand design tokens exactly' +
    '\n- Compute exact coordinates for every element (2 decimal places)' +
    '\n- FULLY specify all artifacts including all style sub-objects' +
    '\n- chart: must have chart_style and series_style[]' +
    '\n- workflow: must have workflow_style, nodes[] with x/y/w/h, connections[] with path[]' +
    '\n- table: must have table_style, column_widths[], column_x_positions[], header_row_height, row_heights[], row_y_positions[], header_cell_frames[], body_cell_frames[]' +
    '\n- cards: must have card_style, card_frames[] with x/y/w/h per card' +
    '\n- matrix: must have matrix_style plus semantic fields from Agent 4 (x_axis, y_axis, quadrants, points)' +
    '\n- driver_tree: must have tree_style plus semantic fields from Agent 4 (root, branches)' +
    '\n- prioritization: must have priority_style plus semantic fields from Agent 4 (items[], qualifiers[])' +
    '\n- insight_text (standard mode): must have insight_mode:"standard", style, heading_style, body_style' +
    '\n- insight_text (grouped mode):  must have insight_mode:"grouped", heading_style, group_layout, group_header_style, group_bullet_box_style, bullet_style, group_gap_in, header_to_box_gap_in' +
    '\n- charts: include final legend_position, data_label_size, category_label_rotation, and series styling' +
    '\n- workflows: include final node geometry, connection paths, node_inner_padding, and external_label_gap' +
    '\n- tables: include final column_widths, column_x_positions, column_types, column_alignments, header_row_height, row_heights, row_y_positions, header_cell_frames, body_cell_frames, and cell_padding' +
    '\n- matrix: include final matrix_style and preserve semantic matrix content for block flattening' +
    '\n- driver_tree: include final tree_style and preserve root/branches for block flattening' +
    '\n- prioritization: include final priority_style and preserve ranked items/qualifiers for block flattening' +
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
      if (a.type === 'chart'    && a.chart_style && a.chart_style.legend_position == null) issues.push(p + ': chart missing legend_position')
      if (a.type === 'chart'    && a.chart_style && a.chart_style.data_label_size == null) issues.push(p + ': chart missing data_label_size')
      if (a.type === 'chart'    && a.chart_style && a.chart_style.category_label_rotation == null) issues.push(p + ': chart missing category_label_rotation')
      if (a.type === 'workflow' && !a.nodes?.length)   issues.push(p + ': workflow missing nodes')
      if (a.type === 'workflow' && !a.workflow_style)  issues.push(p + ': workflow missing workflow_style')
      if (a.type === 'workflow' && a.workflow_style && a.workflow_style.node_inner_padding == null) issues.push(p + ': workflow missing node_inner_padding')
      if (a.type === 'workflow' && a.workflow_style && a.workflow_style.external_label_gap == null) issues.push(p + ': workflow missing external_label_gap')
      if (a.type === 'workflow' && (a.connections || []).some(c => !Array.isArray(c.path) || c.path.length < 2)) issues.push(p + ': workflow connection missing path')
      if (a.type === 'table'    && !a.table_style)     issues.push(p + ': table missing table_style')
      if (a.type === 'table'    && !a.column_widths)   issues.push(p + ': table missing column_widths')
      if (a.type === 'table'    && !a.column_x_positions) issues.push(p + ': table missing column_x_positions')
      if (a.type === 'table'    && !a.row_heights)     issues.push(p + ': table missing row_heights')
      if (a.type === 'table'    && a.header_row_height == null) issues.push(p + ': table missing header_row_height')
      if (a.type === 'table'    && !a.row_y_positions) issues.push(p + ': table missing row_y_positions')
      if (a.type === 'table'    && !a.column_types)    issues.push(p + ': table missing column_types')
      if (a.type === 'table'    && !a.column_alignments) issues.push(p + ': table missing column_alignments')
      if (a.type === 'table'    && !a.header_cell_frames) issues.push(p + ': table missing header_cell_frames')
      if (a.type === 'table'    && !a.body_cell_frames) issues.push(p + ': table missing body_cell_frames')
      if (a.type === 'table'    && a.table_style && a.table_style.cell_padding == null) issues.push(p + ': table missing cell_padding')
      if (a.type === 'cards'    && !a.card_frames?.length) issues.push(p + ': cards missing card_frames')
      if (a.type === 'cards'    && !a.card_style)      issues.push(p + ': cards missing card_style')
      if (a.type === 'cards'    && !a.cards_layout)    issues.push(p + ': cards missing cards_layout')
      if (a.type === 'cards'    && !a.container)       issues.push(p + ': cards missing container')
      if (a.type === 'matrix'   && !a.matrix_style)    issues.push(p + ': matrix missing matrix_style')
      if (a.type === 'matrix'   && !a.x_axis?.label)   issues.push(p + ': matrix missing x_axis.label')
      if (a.type === 'matrix'   && !a.y_axis?.label)   issues.push(p + ': matrix missing y_axis.label')
      if (a.type === 'matrix'   && (a.quadrants || []).length !== 4) issues.push(p + ': matrix must define 4 quadrants')
      if (a.type === 'matrix'   && !(a.points || []).length) issues.push(p + ': matrix missing points')
      if (a.type === 'driver_tree' && !a.tree_style)   issues.push(p + ': driver_tree missing tree_style')
      if (a.type === 'driver_tree' && !a.root?.label)  issues.push(p + ': driver_tree missing root.label')
      if (a.type === 'driver_tree' && !(a.branches || []).length) issues.push(p + ': driver_tree missing branches')
      if (a.type === 'prioritization' && !a.priority_style) issues.push(p + ': prioritization missing priority_style')
      if (a.type === 'prioritization' && !(a.items || []).length) issues.push(p + ': prioritization missing items')
      if (a.type === 'prioritization' && (a.items || []).some(it => it.rank == null || !String(it.title || '').trim())) issues.push(p + ': prioritization items require rank and title')
      if (a.type === 'insight_text' && !a.heading_style) issues.push(p + ': insight_text missing heading_style')
      if (a.type === 'insight_text' && a.insight_mode === 'grouped' && !a.group_header_style) issues.push(p + ': grouped insight_text missing group_header_style')
      if (a.type === 'insight_text' && a.insight_mode === 'grouped' && !a.group_bullet_box_style) issues.push(p + ': grouped insight_text missing group_bullet_box_style')
      if (a.type === 'insight_text' && a.insight_mode === 'grouped' && !a.bullet_style) issues.push(p + ': grouped insight_text missing bullet_style')
      if (a.type === 'insight_text' && a.insight_mode === 'grouped' && a.group_gap_in == null) issues.push(p + ': grouped insight_text missing group_gap_in')
      if (a.type === 'insight_text' && a.insight_mode === 'grouped' && a.header_to_box_gap_in == null) issues.push(p + ': grouped insight_text missing header_to_box_gap_in')
      if (a.type === 'insight_text' && (!a.insight_mode || a.insight_mode === 'standard') && !a.body_style) issues.push(p + ': insight_text missing body_style')
      if (a.type === 'insight_text' && (!a.insight_mode || a.insight_mode === 'standard') && a.body_style && a.body_style.list_style == null) issues.push(p + ': insight_text missing list_style')
      if (a.type === 'insight_text' && (!a.insight_mode || a.insight_mode === 'standard') && a.body_style && a.body_style.line_spacing == null) issues.push(p + ': insight_text missing line_spacing')
      if (a.type === 'insight_text' && (!a.insight_mode || a.insight_mode === 'standard') && a.body_style && a.body_style.indent_inches == null) issues.push(p + ': insight_text missing indent_inches')
      if (a.type === 'insight_text' && (!a.insight_mode || a.insight_mode === 'standard') && a.body_style && a.body_style.space_before_pt == null) issues.push(p + ': insight_text missing space_before_pt')
    })
  })

  if (slide.global_elements?.logo?.show && !slide.global_elements.logo.image_base64) {
    issues.push('logo missing image_base64')
  }

  return issues
}

function rectsOverlap(a, b, gap = 0) {
  if (!a || !b) return false
  const ax1 = +a.x || 0
  const ay1 = +a.y || 0
  const ax2 = ax1 + (+a.w || 0)
  const ay2 = ay1 + (+a.h || 0)
  const bx1 = +b.x || 0
  const by1 = +b.y || 0
  const bx2 = bx1 + (+b.w || 0)
  const by2 = by1 + (+b.h || 0)
  return ax1 < bx2 - gap && ax2 > bx1 + gap && ay1 < by2 - gap && ay2 > by1 + gap
}

function rectArea(r) {
  return Math.max(0, +r.w || 0) * Math.max(0, +r.h || 0)
}

function validateRenderCompleteness(slide) {
  const issues = []
  const canvas = slide.canvas || {}
  const slideBounds = {
    x: 0,
    y: 0,
    w: +canvas.width_in || 0,
    h: +canvas.height_in || 0
  }
  const margin = canvas.margin || {}
  const contentBounds = {
    x: +margin.left || 0,
    y: +margin.top || 0,
    w: Math.max(0, (+canvas.width_in || 0) - (+margin.left || 0) - (+margin.right || 0)),
    h: Math.max(0, (+canvas.height_in || 0) - (+margin.top || 0) - (+margin.bottom || 0))
  }

  if (slide.slide_type === 'content') {
    if (!Array.isArray(slide.blocks) || slide.blocks.length === 0) issues.push('no blocks')
  }

  ;(slide.blocks || []).forEach((b, bi) => {
    const p = 'block' + bi
    if (!b.block_type) issues.push(p + ': missing block_type')
    if (!['image'].includes(b.block_type)) {
      for (const key of ['x', 'y', 'w', 'h']) {
        if (b[key] == null) issues.push(p + ': missing ' + key)
      }
    }
    if (b.x != null && b.y != null && b.w != null && b.h != null) {
      if (b.x < -0.01 || b.y < -0.01 ||
          b.x + b.w > slideBounds.w + 0.01 ||
          b.y + b.h > slideBounds.h + 0.01) {
        issues.push(p + ': outside canvas')
      }
    }
    if (['title', 'subtitle', 'footer', 'page_number', 'image', 'chart', 'table', 'workflow', 'bullet_list', 'rect', 'text_box', 'rule', 'circle'].includes(b.block_type)) {
      if (!b.artifact_type) issues.push(p + ': missing artifact_type')
      if (!b.artifact_subtype) issues.push(p + ': missing artifact_subtype')
      if (!b.fallback_policy) issues.push(p + ': missing fallback_policy')
      if (!b.block_role) issues.push(p + ': missing block_role')
    }
    if (b.block_type === 'chart') {
      if (!b.chart_style) issues.push(p + ': chart block missing chart_style')
      if (!b.series_style?.length) issues.push(p + ': chart block missing series_style')
      if (b.legend_position == null) issues.push(p + ': chart block missing legend_position')
      if (b.data_label_size == null) issues.push(p + ': chart block missing data_label_size')
      if (b.category_label_rotation == null) issues.push(p + ': chart block missing category_label_rotation')
    }
    if (b.block_type === 'table') {
      if (!b.column_widths?.length) issues.push(p + ': table block missing column_widths')
      if (!b.column_x_positions?.length) issues.push(p + ': table block missing column_x_positions')
      if (!b.row_heights?.length) issues.push(p + ': table block missing row_heights')
      if (!b.row_y_positions?.length) issues.push(p + ': table block missing row_y_positions')
      if (!b.column_types?.length) issues.push(p + ': table block missing column_types')
      if (!b.column_alignments?.length) issues.push(p + ': table block missing column_alignments')
      if (b.header_row_height == null) issues.push(p + ': table block missing header_row_height')
      if (!b.header_cell_frames?.length) issues.push(p + ': table block missing header_cell_frames')
      if (!b.body_cell_frames?.length) issues.push(p + ': table block missing body_cell_frames')
      if (b.headers?.length && b.column_widths?.length && b.headers.length !== b.column_widths.length) issues.push(p + ': headers/column_widths length mismatch')
      if (b.headers?.length && b.column_x_positions?.length && b.headers.length !== b.column_x_positions.length) issues.push(p + ': headers/column_x_positions length mismatch')
      if (b.rows?.length && b.row_heights?.length && b.rows.length !== b.row_heights.length) issues.push(p + ': rows/row_heights length mismatch')
      if (b.rows?.length && b.body_cell_frames?.length && b.rows.length !== b.body_cell_frames.length) issues.push(p + ': rows/body_cell_frames length mismatch')
    }
    if (b.block_type === 'workflow') {
      if (!b.nodes?.length) issues.push(p + ': workflow block missing nodes')
      if ((b.connections || []).some(c => !Array.isArray(c.path) || c.path.length < 2)) issues.push(p + ': workflow block missing connection path')
      if (!b.workflow_style) issues.push(p + ': workflow block missing workflow_style')
    }
  })

  const zoneFrames = (slide.zones || []).map((z, zi) => ({ ...z.frame, _idx: zi })).filter(z => z.x != null && z.y != null && z.w != null && z.h != null)
  for (let i = 0; i < zoneFrames.length; i++) {
    for (let j = i + 1; j < zoneFrames.length; j++) {
      if (rectsOverlap(zoneFrames[i], zoneFrames[j], 0.02)) {
        issues.push('zones overlap: z' + zoneFrames[i]._idx + ' & z' + zoneFrames[j]._idx)
      }
    }
  }

  let occupiedArea = 0
  ;(slide.zones || []).forEach((zone, zi) => {
    const inner = getZoneInnerBounds(zone)
    const arts = zone.artifacts || []
    const artRects = []
    ;(arts || []).forEach((a, ai) => {
      const p = 'z' + zi + '.a' + ai
      if (a.x != null && a.y != null && a.w != null && a.h != null) {
        const ar = { x: a.x, y: a.y, w: a.w, h: a.h, _id: p }
        artRects.push(ar)
        occupiedArea += rectArea(ar)
        if (!slide.layout_mode) {
          const fits =
            ar.x >= inner.x - 0.01 &&
            ar.y >= inner.y - 0.01 &&
            ar.x + ar.w <= inner.x + inner.w + 0.01 &&
            ar.y + ar.h <= inner.y + inner.h + 0.01
          if (!fits) issues.push(p + ': outside zone bounds')
        }
      }

      if (a.type === 'cards') {
        const frames = a.card_frames || []
        const cards = a.cards || []
        if (frames.length !== cards.length) issues.push(p + ': card_frames/cards length mismatch')
        if (frames.length > 1) {
          const ref = frames[0] || {}
          frames.forEach((fr, fi) => {
            if (Math.abs((fr.w || 0) - (ref.w || 0)) > 0.03 || Math.abs((fr.h || 0) - (ref.h || 0)) > 0.03) {
              issues.push(p + ': unequal card frame sizes')
            }
            if (fr.x == null || fr.y == null || fr.w == null || fr.h == null) {
              issues.push(p + '.card' + fi + ': incomplete frame')
            }
          })
        }
        for (let i = 0; i < frames.length; i++) {
          for (let j = i + 1; j < frames.length; j++) {
            if (rectsOverlap(frames[i], frames[j], 0.02)) {
              issues.push(p + ': overlapping card frames')
              break
            }
          }
        }
      }

      if (a.type === 'workflow') {
        const nodes = a.nodes || []
        for (let i = 0; i < nodes.length; i++) {
          for (let j = i + 1; j < nodes.length; j++) {
            if (rectsOverlap(nodes[i], nodes[j], 0.02)) {
              issues.push(p + ': overlapping workflow nodes')
              break
            }
          }
        }
      }
    })

    for (let i = 0; i < artRects.length; i++) {
      for (let j = i + 1; j < artRects.length; j++) {
        if (rectsOverlap(artRects[i], artRects[j], 0.02)) {
          issues.push('artifact overlap: ' + artRects[i]._id + ' & ' + artRects[j]._id)
        }
      }
    }
  })

  if (slide.slide_type === 'content' && contentBounds.w > 0 && contentBounds.h > 0) {
    const contentArea = contentBounds.w * contentBounds.h
    const fillRatio = occupiedArea / Math.max(contentArea, 0.01)
    if (fillRatio < 0.28) issues.push('content under-utilised: ' + fillRatio.toFixed(2))
    if (fillRatio > 0.92) issues.push('content over-packed: ' + fillRatio.toFixed(2))
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

function applyLayoutTitleFrames(slide, layoutName, brand) {
  if (!slide || !layoutName || !brand) return slide
  const layouts = brand.slide_layouts || []
  const layout = layouts.find(l => (l.name || '').toLowerCase() === layoutName.toLowerCase())
    || layouts.find(l => (l.name || '').toLowerCase().includes(layoutName.toLowerCase()))
  if (!layout) return slide

  const placeholders = layout.placeholders || []
  const isTitleType = p => {
    const t = String(p?.type || '').toLowerCase()
    return t === 'title' || t === 'center_title' || t === 'centertitle' || t === 'ctrtitle'
  }
  const titlePh = layout.title_placeholder
    || (layout.master_summary || {}).title_placeholder
    || placeholders.find(isTitleType)
    || placeholders
      .filter(p => String(p?.type || '').toLowerCase() !== 'body')
      .sort((a, b) => {
        const ay = a?.y_in ?? 99
        const by = b?.y_in ?? 99
        if (ay !== by) return ay - by
        return (b?.w_in ?? 0) - (a?.w_in ?? 0)
      })[0]
  if (!titlePh) return slide

  const out = JSON.parse(JSON.stringify(slide))
  const r2 = x => Math.round(x * 100) / 100
  const contentAreas = (layout.placeholders || [])
    .filter(p => p.type === 'body' && (p.h_in || 0) > 0.5)
    .sort((a, b) => {
      const rowA = Math.round((a.y_in || 0) * 2)
      const rowB = Math.round((b.y_in || 0) * 2)
      if (rowA !== rowB) return rowA - rowB
      return (a.x_in || 0) - (b.x_in || 0)
    })
  const topContentY = contentAreas.length ? Math.min(...contentAreas.map(p => p.y_in || 99)) : null

  if (out.title_block) {
    out.title_block = {
      ...out.title_block,
      x: r2(titlePh.x_in != null ? titlePh.x_in : out.title_block.x || 0.4),
      y: r2(titlePh.y_in != null ? titlePh.y_in : out.title_block.y || 0.15),
      w: r2(titlePh.w_in != null ? titlePh.w_in : out.title_block.w || 9.2),
      h: r2(titlePh.h_in != null ? titlePh.h_in : out.title_block.h || 0.7),
      align: out.title_block.align || 'left',
      valign: out.title_block.valign || 'middle',
      wrap: out.title_block.wrap !== false
    }
  }

  if (out.subtitle_block) {
    const subtitleX = titlePh.x_in != null ? titlePh.x_in : (out.subtitle_block.x != null ? out.subtitle_block.x : 0.4)
    const subtitleW = titlePh.w_in != null ? titlePh.w_in : (out.subtitle_block.w != null ? out.subtitle_block.w : 9.2)
    const defaultSubtitleY = (titlePh.y_in || 0.15) + (titlePh.h_in || 0.7) + 0.08
    let subtitleY = out.subtitle_block.y != null ? out.subtitle_block.y : defaultSubtitleY
    let subtitleH = out.subtitle_block.h != null ? out.subtitle_block.h : 0.35
    if (topContentY != null) {
      const maxBottom = topContentY - 0.08
      subtitleY = Math.min(subtitleY, Math.max(defaultSubtitleY, maxBottom - subtitleH))
      subtitleH = Math.max(0.18, Math.min(subtitleH, maxBottom - subtitleY))
    }
    out.subtitle_block = {
      ...out.subtitle_block,
      x: r2(subtitleX),
      y: r2(subtitleY),
      w: r2(subtitleW),
      h: r2(subtitleH),
      align: out.subtitle_block.align || 'left',
      valign: out.subtitle_block.valign || 'top',
      wrap: out.subtitle_block.wrap !== false
    }
  }

  return out
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
- table: include table_style{}, column_widths[], column_x_positions[], header_row_height, row_heights[], row_y_positions[], header_cell_frames[], body_cell_frames[]
- cards: include card_style{}, card_frames[] with x/y/w/h per card
- insight_text standard: include insight_mode:"standard", style{}, heading_style{}, body_style{}
- insight_text grouped:  include insight_mode:"grouped", heading_style{}, group_layout, group_header_style{}, group_bullet_box_style{}, bullet_style{}, group_gap_in, header_to_box_gap_in
- charts: include final legend_position, data_label_size, category_label_rotation, and series styling
- workflows: include final node geometry, connection paths, node_inner_padding, and external_label_gap
- tables: include final column_widths, column_x_positions, column_types, column_alignments, header_row_height, row_heights, row_y_positions, header_cell_frames, body_cell_frames, and cell_padding

All coordinates in decimal inches, 2 decimal places.
Return ONLY a valid JSON object. No explanation. No markdown.`

async function buildFallbackDesign(manifestSlide, brand) {
  console.log('Agent 5 -- fallback Claude call for S' + manifestSlide.slide_number)

  const tokens = extractBrandTokens(brand)

  const prompt =
    'BRAND DESIGN TOKENS:\n' +
    JSON.stringify(tokens, null, 2) +
    '\n\nSLIDE TO REBUILD:\n' +
    JSON.stringify(manifestSlide, null, 2) +
    '\n\nBuild the best possible layout for this slide.' +
    '\nPreserve the title, key_message, zones structure and artifact content from the manifest.' +
    '\nCRITICAL: preserve the exact number of zones and the exact artifact types in each zone from the manifest.' +
    '\nDo NOT collapse charts, tables, workflows, cards, matrix, driver_tree, or prioritization into generic insight_text unless the manifest itself uses insight_text.' +
    '\nChoose the cleanest, most board-ready layout given the archetype: ' + (manifestSlide.slide_archetype || 'summary') +
    '\nReturn a single JSON object for this one slide.'

  try {
    const raw    = await callClaude(AGENT5_FALLBACK_SYSTEM, [{ role: 'user', content: prompt }], 3000)
    const parsed = safeParseJSON(raw, null)

    // Claude may return an array with one item or a bare object
    const candidate = Array.isArray(parsed) ? parsed[0] : parsed

    if (candidate && typeof candidate === 'object' && candidate.canvas && candidate.zones) {
      const issues = validateDesignedSlide(candidate)
      const signatureIssues = validateFallbackStructure(candidate, manifestSlide)
      if (issues.length === 0 && signatureIssues.length === 0) {
        console.log('Agent 5 -- fallback Claude succeeded for S' + manifestSlide.slide_number)
        return candidate
      }
      console.warn('Agent 5 -- fallback Claude still has issues for S' + manifestSlide.slide_number + ':', issues.concat(signatureIssues).join('; '))
    }
  } catch (e) {
    console.warn('Agent 5 -- fallback Claude call failed for S' + manifestSlide.slide_number + ':', e.message)
  }

  // Last resort: minimal structurally valid object derived purely from brand tokens
  // No hardcoded content — pull everything from manifest and brand
  return buildMinimalSafeSlide(manifestSlide, tokens)
}

function manifestZoneArtifactSignature(slideLike) {
  return (slideLike?.zones || []).map(z => (z.artifacts || []).map(a => a.type || 'unknown'))
}

function validateFallbackStructure(candidate, manifestSlide) {
  const issues = []
  const candSig = manifestZoneArtifactSignature(candidate)
  const manifestSig = manifestZoneArtifactSignature(manifestSlide)
  if (candSig.length !== manifestSig.length) {
    issues.push('zone count mismatch vs manifest')
    return issues
  }
  for (let zi = 0; zi < manifestSig.length; zi++) {
    const mArts = manifestSig[zi]
    const cArts = candSig[zi] || []
    if (cArts.length !== mArts.length) {
      issues.push('z' + zi + ': artifact count mismatch vs manifest')
      continue
    }
    for (let ai = 0; ai < mArts.length; ai++) {
      if (String(cArts[ai] || '') !== String(mArts[ai] || '')) {
        issues.push('z' + zi + '.a' + ai + ': artifact type mismatch vs manifest (' + cArts[ai] + ' != ' + mArts[ai] + ')')
      }
    }
  }
  return issues
}

function makeHeaderBlockFromManifestArtifact(artifact, bt) {
  const text = (
    artifact?.insight_header ||
    artifact?.chart_header ||
    artifact?.table_header ||
    artifact?.workflow_header ||
    artifact?.matrix_header ||
    artifact?.tree_header ||
    artifact?.priority_header ||
    artifact?.heading ||
    ''
  )
  if (!text) return null
  return {
    text,
    x: 0, y: 0, w: 1, h: 0.3,
    font_family: bt.title_font_family || 'Arial',
    font_size: 11,
    font_weight: 'semibold',
    color: bt.primary_color || '#0078AE',
    style: 'underline',
    accent_color: bt.primary_color || '#0078AE'
  }
}

function buildSafeArtifactShell(manifestArt, bt) {
  const t = manifestArt?.type || 'insight_text'
  const header_block = makeHeaderBlockFromManifestArtifact(manifestArt, bt)
  if (t === 'chart') {
    return {
      type: 'chart',
      x: null, y: null, w: null, h: null,
      chart_style: {
        title_font_family: bt.title_font_family || 'Arial',
        title_font_size: 12,
        axis_font_family: bt.body_font_family || 'Arial',
        axis_font_size: 9,
        label_font_family: bt.body_font_family || 'Arial',
        label_font_size: 9,
        title_color: bt.primary_color || '#0078AE',
        axis_color: bt.body_color || '#111111',
        gridline_color: '#DDDDDD',
        legend_font_family: bt.body_font_family || 'Arial',
        legend_font_size: 9,
        legend_color: bt.body_color || '#111111',
        show_gridlines: false,
        show_border: false,
        border_color: null,
        background_color: null,
        legend_position: 'top',
        data_label_size: 9,
        category_label_rotation: 0
      },
      series_style: [],
      header_block
    }
  }
  if (t === 'table') {
    return {
      type: 'table',
      x: null, y: null, w: null, h: null,
      table_style: {
        header_fill_color: bt.primary_color || '#0078AE',
        header_text_color: '#FFFFFF',
        header_font_family: bt.title_font_family || 'Arial',
        header_font_size: 10,
        body_fill_color: '#FFFFFF',
        body_alt_fill_color: '#F7F8FA',
        body_text_color: bt.body_color || '#111111',
        body_font_family: bt.body_font_family || 'Arial',
        body_font_size: 9,
        grid_color: '#D7DEE8',
        grid_width: 0.5,
        highlight_fill_color: '#FFF4BF',
        cell_padding: 0.06
      },
      column_widths: [],
      column_x_positions: [],
      row_heights: [],
      header_row_height: null,
      row_y_positions: [],
      column_types: [],
      column_alignments: [],
      header_cell_frames: [],
      body_cell_frames: [],
      header_block
    }
  }
  if (t === 'workflow') {
    return {
      type: 'workflow',
      x: null, y: null, w: null, h: null,
      workflow_style: {
        node_fill_color: bt.primary_color || '#0078AE',
        node_border_color: '#FFFFFF',
        node_border_width: 1,
        node_title_font_family: bt.title_font_family || 'Arial',
        node_title_font_size: 10,
        node_title_color: '#FFFFFF',
        node_value_font_family: bt.body_font_family || 'Arial',
        node_value_font_size: 9,
        node_value_color: bt.body_color || '#111111',
        node_inner_padding: 0.08,
        external_label_gap: 0.08,
        connector_color: bt.primary_color || '#0078AE',
        connector_width: 0.5,
        node_corner_radius: 4
      },
      nodes: [],
      connections: [],
      container: null,
      header_block
    }
  }
  if (t === 'cards') {
    return {
      type: 'cards',
      x: null, y: null, w: null, h: null,
      cards_layout: manifestArt.cards_layout || 'column',
      container: null,
      card_frames: [],
      card_style: {
        fill_color: '#F5F5F5',
        border_color: '#DDDDDD',
        border_width: 0.75,
        corner_radius: 0,
        shadow: false,
        internal_padding: 0.12
      },
      title_style: {
        font_family: bt.title_font_family || 'Arial',
        font_size: 12,
        font_weight: 'bold',
        color: bt.primary_color || '#0078AE'
      },
      subtitle_style: {
        font_family: bt.body_font_family || 'Arial',
        font_size: 22,
        font_weight: 'bold',
        color: bt.body_color || '#111111'
      },
      body_style: {
        font_family: bt.body_font_family || 'Arial',
        font_size: 9,
        font_weight: 'regular',
        color: bt.body_color || '#111111',
        line_spacing: 1.2
      }
    }
  }
  if (t === 'matrix') {
    return {
      type: 'matrix',
      x: null, y: null, w: null, h: null,
      matrix_style: {},
      header_block
    }
  }
  if (t === 'driver_tree') {
    return {
      type: 'driver_tree',
      x: null, y: null, w: null, h: null,
      tree_style: {},
      header_block
    }
  }
  if (t === 'prioritization') {
    return {
      type: 'prioritization',
      x: null, y: null, w: null, h: null,
      priority_style: {},
      header_block
    }
  }

  const grouped = !!(manifestArt?.groups && manifestArt.groups.length)
  return grouped ? {
    type: 'insight_text',
    insight_mode: 'grouped',
    x: null, y: null, w: null, h: null,
    style: { fill_color: null, border_color: null, border_width: 0, corner_radius: 0 },
    heading_style: { font_family: bt.title_font_family || 'Arial', font_size: 12, font_weight: 'bold', color: bt.primary_color || '#0078AE' },
    group_layout: 'rows',
    group_header_style: { shape: 'rounded_rect', fill_color: bt.primary_color || '#0078AE', text_color: '#FFFFFF', font_family: bt.title_font_family || 'Arial', font_size: 10, font_weight: 'bold', corner_radius: 0.04, w: 1.4, h: 0.28 },
    group_bullet_box_style: { fill_color: null, border_color: '#CCCCCC', border_width: 0.75, corner_radius: 0.04, padding: { top: 0.08, right: 0.1, bottom: 0.08, left: 0.1 } },
    bullet_style: { font_family: bt.body_font_family || 'Arial', font_size: 10, font_weight: 'regular', color: bt.body_color || '#111111', line_spacing: 1.35, indent_inches: 0.1, space_before_pt: 3, char: '▶' },
    group_gap_in: 0.08,
    header_to_box_gap_in: 0.04,
    header_block
  } : {
    type: 'insight_text',
    insight_mode: 'standard',
    x: null, y: null, w: null, h: null,
    style: { fill_color: null, border_color: (bt.primary_color || '#0078AE') + '33', border_width: 0.5, corner_radius: 3 },
    heading_style: { font_family: bt.title_font_family || 'Arial', font_size: 12, font_weight: 'bold', color: bt.primary_color || '#0078AE' },
    body_style: { font_family: bt.body_font_family || 'Arial', font_size: 11, font_weight: 'regular', color: bt.body_color || '#111111', line_spacing: 1.4, indent_inches: 0.15, list_style: 'bullet', space_before_pt: 5, vertical_distribution: 'spread' },
    heading: manifestArt?.heading || manifestArt?.insight_header || 'Key Insight',
    header_block
  }
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
  const cw        = r2(w - 0.80)

  const fallbackBrandTokens = {
    title_font_family: titleFont,
    body_font_family: bodyFont,
    caption_font_family: bodyFont,
    title_color: isDark ? '#FFFFFF' : primary,
    body_color: isDark ? '#CCDDFF' : '#111111',
    caption_color: '#888888',
    primary_color: primary,
    secondary_color: secondary,
    accent_colors: tokens.accent_colors || [],
    chart_palette: tokens.chart_colors || [primary, secondary, '#2E9E5B', '#C82333']
  }

  const fallbackZones = manifestSlide.slide_type === 'content'
    ? (manifestSlide.zones || []).map((zone, zoneIdx) => ({
        zone_id: zone.zone_id || `z${zoneIdx + 1}` ,
        zone_role: zone.zone_role || (zoneIdx === 0 ? 'primary_proof' : 'supporting_evidence'),
        message_objective: zone.message_objective || manifestSlide.key_message || '',
        narrative_weight: zone.narrative_weight || (zoneIdx === 0 ? 'primary' : 'secondary'),
        frame: null,
        padding: zone.padding || null,
        layout_hint: zone.layout_hint || null,
        artifacts: (zone.artifacts || []).map(art => buildSafeArtifactShell(art, fallbackBrandTokens))
      }))
    : []

  return {
    slide_number: manifestSlide.slide_number,
    slide_type: manifestSlide.slide_type || 'content',
    slide_archetype: manifestSlide.slide_archetype || 'summary',
    canvas: {
      width_in: w, height_in: h,
      margin: { left: 0.40, right: 0.40, top: 0.15, bottom: 0.30 },
      background: { color: isDark ? primary : bg }
    },
    brand_tokens: {
      title_font_family: fallbackBrandTokens.title_font_family,
      body_font_family: fallbackBrandTokens.body_font_family,
      caption_font_family: fallbackBrandTokens.caption_font_family,
      title_color: fallbackBrandTokens.title_color,
      body_color: fallbackBrandTokens.body_color,
      caption_color: fallbackBrandTokens.caption_color,
      primary_color: fallbackBrandTokens.primary_color,
      secondary_color: fallbackBrandTokens.secondary_color,
      accent_colors: fallbackBrandTokens.accent_colors,
      chart_palette: fallbackBrandTokens.chart_palette
    },
    title_block: {
      text: manifestSlide.title || '',
      x: 0.40, y: 0.15, w: cw,
      h: isDark ? 2.00 : 0.75,
      font_family: titleFont,
      font_size: isDark ? 30 : 18,
      font_weight: 'bold',
      color: isDark ? '#FFFFFF' : primary,
      align: 'left', valign: 'middle', wrap: true
    },
    subtitle_block: manifestSlide.subtitle ? {
      text: manifestSlide.subtitle,
      x: 0.40, y: isDark ? 2.60 : 0.95,
      w: cw, h: 0.45,
      font_family: bodyFont, font_size: 14, font_weight: 'regular',
      color: isDark ? '#BBCCFF' : '#555555',
      align: 'left', valign: 'top', wrap: true
    } : null,
    zones: fallbackZones,
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
    layout_mode: manifestSlide.layout_mode !== false,
    selected_layout_name: manifestSlide.selected_layout_name || manifestSlide.layout_name || '',
    section_name: manifestSlide.section_name || '',
    section_type: manifestSlide.section_type || '',
    title: manifestSlide.title || '',
    subtitle: manifestSlide.subtitle || '',
    key_message: manifestSlide.key_message || '',
    visual_flow_hint: manifestSlide.visual_flow_hint || '',
    speaker_note: manifestSlide.speaker_note || '',
    _fallback: true
  }
}

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
//   insight_text : heading, insight_header, insight_mode, points[] (standard), groups[] (grouped), sentiment
//   chart        : chart_type, chart_title, chart_insight, x_label, y_label,
//                  categories[], series[], show_data_labels, show_legend
//   cards        : cards[] (title, subtitle, body, sentiment per card)
//   workflow     : workflow_type, flow_direction, workflow_title, workflow_insight,
//                  node labels/values/descriptions/levels, connection from/to/type
//   table        : title, headers[], rows[][], highlight_rows[], note
//   matrix       : matrix_type, matrix_header, x_axis, y_axis, quadrants[], points[]
//   driver_tree  : tree_header, root, branches[]
//   prioritization: priority_header, items[]
// ═══════════════════════════════════════════════════════════════════════════════

// ─── computeArtifactInternals ────────────────────────────────────────────────
// Post-processes merged zones and fills computed layout/sizing fields on each
// artifact IN PLACE, so generate_pptx.py can act as a pure renderer.
// Called after mergeContentIntoZones (and applyLayoutZoneFrames if used).
// ─────────────────────────────────────────────────────────────────────────────
function computeArtifactInternals(zones, canvas, brandTokens) {
  const round2 = x => Math.round(x * 100) / 100
  const bt = brandTokens || {}

  for (const zone of (zones || [])) {
    const artifacts = zone.artifacts || []
    const frame = zone.frame || {}

    // ── 1. Multi-artifact zone stacking ──────────────────────────────────────
    if (artifacts.length >= 2) {
      const needsCompute = artifacts.some(a => a.h == null || a.w == null || a.x == null || a.y == null)
      if (needsCompute) {
        const zx = frame.x || 0
        const zy = frame.y || 0
        const zw = frame.w || 0
        const zh = frame.h || 0
        const gap = 0.12
        const splitHint =
          (Array.isArray(zone.split_hint) ? zone.split_hint : null) ||
          (Array.isArray((zone.layout_hint || {}).split_hint) ? (zone.layout_hint || {}).split_hint : null)
        const arrangement =
          zone.artifact_arrangement ||
          (zone.layout_hint || {}).artifact_arrangement ||
          'vertical'

        let primaryFrac, secondaryFrac
        if (splitHint && Array.isArray(splitHint) && splitHint.length >= 2) {
          const total = splitHint[0] + splitHint[1]
          primaryFrac   = splitHint[0] / total
          secondaryFrac = splitHint[1] / total
        } else {
          primaryFrac   = 0.60
          secondaryFrac = 0.40
        }

        if (artifacts.length >= 2) {
          const firstType = String((artifacts[0]?.type || '')).toLowerCase()
          const secondType = String((artifacts[1]?.type || '')).toLowerCase()
          if (firstType === 'cards' && secondType !== 'cards') {
            primaryFrac = Math.min(primaryFrac, 0.40)
            secondaryFrac = 1 - primaryFrac
          } else if (secondType === 'cards' && firstType !== 'cards') {
            secondaryFrac = Math.min(secondaryFrac, 0.40)
            primaryFrac = 1 - secondaryFrac
          } else if (firstType === 'cards' && secondType === 'cards') {
            primaryFrac = Math.min(primaryFrac, 0.40)
            secondaryFrac = Math.min(secondaryFrac, 0.40)
          }
        }

        if (arrangement === 'horizontal') {
          const availW = zw - gap
          const primaryW = round2(availW * primaryFrac)
          const secondaryW = round2(availW * secondaryFrac)

          for (let i = 0; i < artifacts.length; i++) {
            const art = artifacts[i]
            if (i === 0) {
              art.x = round2(zx)
              art.y = round2(zy)
              art.w = primaryW
              art.h = round2(zh)
            } else {
              const remaining = artifacts.length - 1
              const eachW = round2((secondaryW - Math.max(0, remaining - 1) * gap) / Math.max(remaining, 1))
              art.x = round2(zx + primaryW + gap + (i - 1) * (eachW + gap))
              art.y = round2(zy)
              art.w = eachW
              art.h = round2(zh)
            }
          }
        } else {
          const availH = zh - gap
          const primaryH   = round2(availH * primaryFrac)
          const secondaryH = round2(availH * secondaryFrac)

          for (let i = 0; i < artifacts.length; i++) {
            const art = artifacts[i]
            if (i === 0) {
              art.x = round2(zx)
              art.y = round2(zy)
              art.w = round2(zw)
              art.h = primaryH
            } else {
              const remaining = artifacts.length - 1
              const eachH = round2((secondaryH - Math.max(0, remaining - 1) * gap) / Math.max(remaining, 1))
              art.x = round2(zx)
              art.y = round2(zy + primaryH + gap + (i - 1) * (eachH + gap))
              art.w = round2(zw)
              art.h = eachH
            }
          }
        }
      }
    }

    // ── Per-artifact computed fields ──────────────────────────────────────────
    for (const art of artifacts) {
      const artType = art.type

      // ── 2. Chart: _computed sub-object ─────────────────────────────────────
      if (artType === 'chart') {
        if (!art._computed) art._computed = {}
        const computed = art._computed
        const canvasW = (canvas && canvas.width_in) ? canvas.width_in : 10
        const canvasH = (canvas && canvas.height_in) ? canvas.height_in : 7.5

        // legend_position
        if (art.show_legend) {
          const widthRatio = (art.w || 0) / Math.max(canvasW, 0.1)
          const heightRatio = (art.h || 0) / Math.max(canvasH, 0.1)
          if (heightRatio > 0.60) computed.legend_position = 'top'
          else if (widthRatio > 0.60) computed.legend_position = 'right'
          else computed.legend_position = (art.chart_type === 'pie') ? 'right' : 'top'
        } else {
          computed.legend_position = 'none'
        }

        const cs = art.chart_style || {}
        const headerFs = ((art.header_block || {}).font_size) || cs.title_font_size || 11
        const maxLegendFs = Math.max(8, Math.min(headerFs - 1, 9))
        art.chart_style = {
          ...cs,
          legend_font_size: Math.min(cs.legend_font_size || maxLegendFs, maxLegendFs)
        }

        // data_label_size
        if (computed.data_label_size == null) {
          const base_size = cs.data_label_size || 9
          const n_cats = (art.categories || []).length || 1
          const density = Math.min(art.w || 0, art.h || 0) / n_cats
          const scale = Math.max(0.55, Math.min(1.0, density / 0.6))
          const computedSize = Math.round(base_size * scale)
          computed.data_label_size = Math.max(computedSize, Math.round(base_size * 0.55))
        }

        // category_label_rotation
        if (computed.category_label_rotation == null) {
          computed.category_label_rotation = (art.categories || []).length > 6 ? -45 : 0
        }
      }

      // ── 3. Table: column and row specs ─────────────────────────────────────
      if (artType === 'table') {
        const headers = art.headers || []
        const rows    = art.rows    || []
        const nCols   = headers.length

        if (nCols > 0) {
          const artW = art.w || 6

          // column_widths
          if (!art.column_widths || art.column_widths.length === 0) {
            const weights = []
            for (let c = 0; c < nCols; c++) {
              let maxLen = (headers[c] || '').length
              for (const row of rows) {
                if (c < row.length) maxLen = Math.max(maxLen, String(row[c] || '').length)
              }
              weights.push(Math.max(maxLen, 1))
            }
            const totalWeight = weights.reduce((s, w) => s + w, 0) || 1
            const colWidths = weights.map(w => round2(artW * w / totalWeight))
            // Fix rounding remainder on last column
            const widthSum = colWidths.reduce((s, w) => s + w, 0)
            colWidths[nCols - 1] = round2(colWidths[nCols - 1] + (artW - widthSum))
            art.column_widths = colWidths
          }

          // column_types
          if (!art.column_types) {
            const numPat = /^[\d,\.\%₹\$\-\+]+$/
            const types = []
            for (let c = 0; c < nCols; c++) {
              if (c === 0) {
                types.push('text')
              } else {
                const hits = rows.filter(row => c < row.length && numPat.test(String(row[c] || '').trim())).length
                types.push((hits / Math.max(rows.length, 1)) > 0.5 ? 'numeric' : 'text')
              }
            }
            art.column_types = types
          }

          // column_alignments
          if (!art.column_alignments) {
            art.column_alignments = (art.column_types || []).map(t => t === 'numeric' ? 'right' : 'left')
          }

          // row_heights + header_row_height
          if (!art.row_heights) {
            const artH = art.h || 2
            const available_h = artH - 0.35
            const n_data_rows = rows.length
            const row_h = Math.max(0.32, available_h / Math.max(n_data_rows, 1))
            art.row_heights = Array(n_data_rows).fill(round2(row_h))
            art.header_row_height = 0.35
          }

          // Explicit table grid geometry for Agent 6.
          const tableX = art.x || 0
          const tableY = art.y || 0
          const colWs = art.column_widths || []
          const dataRowHs = art.row_heights || []
          const headerH = art.header_row_height != null ? art.header_row_height : 0.35

          let curX = tableX
          art.column_x_positions = colWs.map(cw => {
            const x = round2(curX)
            curX += (+cw || 0)
            return x
          })

          let curY = tableY
          art.row_y_positions = [round2(curY)]
          curY += (+headerH || 0)
          for (const rh of dataRowHs) {
            art.row_y_positions.push(round2(curY))
            curY += (+rh || 0)
          }

          art.header_cell_frames = colWs.map((cw, ci) => ({
            col_index: ci,
            x: round2(art.column_x_positions[ci] || tableX),
            y: round2(tableY),
            w: round2(+cw || 0),
            h: round2(+headerH || 0)
          }))

          art.body_cell_frames = dataRowHs.map((rh, ri) =>
            colWs.map((cw, ci) => ({
              row_index: ri,
              col_index: ci,
              x: round2(art.column_x_positions[ci] || tableX),
              y: round2(art.row_y_positions[ri + 1] || tableY),
              w: round2(+cw || 0),
              h: round2(+rh || 0)
            }))
          )
        }
      }

      // ── 4. Cards: pre-compute card_frames ──────────────────────────────────
      if (artType === 'cards') {
        if (!art.card_frames || art.card_frames.length === 0) {
          const cards  = art.cards  || []
          const requestedLayout = String(art.cards_layout || art.layout || '').toLowerCase()
          const cs     = art.card_style || {}
          const gap    = cs.gap              || 0.12
          const count  = cards.length
          const ax     = art.x || 0
          const ay     = art.y || 0
          const aw     = art.w || 0
          const ah     = art.h || 0
          const aspect = ah > 0 ? aw / ah : 1
          const minReadableCardWidth = 1.45
          const minReadableCardHeight = 1.10
          const rowCardWidth = count > 0 ? (aw - gap * (count - 1)) / Math.max(count, 1) : aw
          const columnCardHeight = count > 0 ? (ah - gap * (count - 1)) / Math.max(count, 1) : ah
          const gridCols = count > 1 ? 2 : 1
          const gridRows = Math.ceil(count / Math.max(gridCols, 1))
          const gridCardWidth = (aw - gap * (gridCols - 1)) / Math.max(gridCols, 1)
          const gridCardHeight = (ah - gap * (gridRows - 1)) / Math.max(gridRows, 1)

          let layout = requestedLayout
          if (!['row', 'column', 'grid'].includes(layout)) {
            if (count <= 1) layout = 'row'
            else if (count === 2) layout = aspect >= 1 ? 'row' : 'column'
            else if (count === 3) layout = aspect >= 1 ? 'row' : 'column'
            else if (count === 4) layout = 'grid'
            else layout = aspect >= 1.15 ? 'row' : 'grid'
          }

          // Readability override: never keep a horizontal row when it makes KPI cards too narrow.
          if (layout === 'row' && rowCardWidth < minReadableCardWidth) {
            layout = count <= 3 ? 'column' : 'grid'
          }
          if (layout === 'grid' && (gridCardWidth < minReadableCardWidth || gridCardHeight < minReadableCardHeight)) {
            if (count <= 3) layout = 'column'
            else if (rowCardWidth >= minReadableCardWidth && aspect >= 1.15) layout = 'row'
          }
          if (layout === 'column' && columnCardHeight < minReadableCardHeight && rowCardWidth >= minReadableCardWidth) {
            layout = 'row'
          }

          const frames = []
          if (layout === 'row') {
            const card_w = round2((aw - gap * (count - 1)) / Math.max(count, 1))
            for (let i = 0; i < count; i++) {
              frames.push({ x: round2(ax + i * (card_w + gap)), y: round2(ay), w: card_w, h: round2(ah) })
            }
          } else if (layout === 'column') {
            const card_h = round2((ah - gap * (count - 1)) / Math.max(count, 1))
            for (let i = 0; i < count; i++) {
              frames.push({ x: round2(ax), y: round2(ay + i * (card_h + gap)), w: round2(aw), h: card_h })
            }
          } else {
            // grid or default
            const cols        = count > 1 ? 2 : 1
            const rows_count  = Math.ceil(count / cols)
            const card_w      = round2((aw - gap * (cols - 1)) / Math.max(cols, 1))
            const card_h      = round2((ah - gap * (rows_count - 1)) / Math.max(rows_count, 1))
            for (let i = 0; i < count; i++) {
              frames.push({
                x: round2(ax + (i % cols) * (card_w + gap)),
                y: round2(ay + Math.floor(i / cols) * (card_h + gap)),
                w: card_w,
                h: card_h
              })
            }
          }
          art.cards_layout = layout
          art.card_frames = frames
        }
      }

      // ── 5. insight_text (standard): font scaling ───────────────────────────
      if (artType === 'matrix') {
        const ms = art.matrix_style || {}
        const pointPalette = [
          bt.primary_color,
          bt.secondary_color,
          ...(bt.accent_colors || []),
          ...(bt.chart_palette || [])
        ].filter(Boolean)
        art.matrix_style = {
          border_color: ms.border_color || '#D7DEE8',
          border_width: ms.border_width != null ? ms.border_width : 0.8,
          divider_color: ms.divider_color || '#7B8794',
          divider_width: ms.divider_width != null ? ms.divider_width : 0.6,
          axis_color: ms.axis_color || '#4B5563',
          axis_width: ms.axis_width != null ? ms.axis_width : 0.7,
          axis_label_font_family: ms.axis_label_font_family || bt.body_font_family || 'Arial',
          axis_label_font_size: ms.axis_label_font_size || 10,
          axis_label_color: ms.axis_label_color || bt.body_color || '#374151',
          quadrant_title_font_family: ms.quadrant_title_font_family || bt.title_font_family || 'Arial',
          quadrant_title_font_size: ms.quadrant_title_font_size || 12,
          quadrant_title_color: ms.quadrant_title_color || '#2D3748',
          quadrant_body_font_family: ms.quadrant_body_font_family || bt.body_font_family || 'Arial',
          quadrant_body_font_size: ms.quadrant_body_font_size || 9,
          quadrant_body_color: ms.quadrant_body_color || '#374151',
          point_label_font_family: ms.point_label_font_family || bt.body_font_family || 'Arial',
          point_label_font_size: ms.point_label_font_size || 10,
          point_label_color: ms.point_label_color || '#111111',
          point_palette: (ms.point_palette && ms.point_palette.length ? ms.point_palette : pointPalette),
          quadrant_fills: (ms.quadrant_fills && ms.quadrant_fills.length === 4)
            ? ms.quadrant_fills
            : ['#FFF4BF', '#E4F2DE', '#F4F5F7', '#DDEFF5']
        }
      }

      if (artType === 'driver_tree') {
        const ts = art.tree_style || {}
        art.tree_style = {
          node_fill_color: ts.node_fill_color || '#EAF2FB',
          node_fill_color_secondary: ts.node_fill_color_secondary || '#EDF7F3',
          node_fill_color_leaf: ts.node_fill_color_leaf || '#F4F7FA',
          node_border_color: ts.node_border_color || '#D7DEE8',
          node_border_width: ts.node_border_width != null ? ts.node_border_width : 0.6,
          connector_color: ts.connector_color || '#7A8FA8',
          connector_width: ts.connector_width != null ? ts.connector_width : 0.5,
          label_font_family: ts.label_font_family || bt.title_font_family || 'Arial',
          label_font_size: ts.label_font_size || 11,
          label_color: ts.label_color || '#111111',
          value_font_family: ts.value_font_family || bt.body_font_family || 'Arial',
          value_font_size: ts.value_font_size || 10,
          value_color: ts.value_color || bt.primary_color || '#0078AE',
          corner_radius: ts.corner_radius != null ? ts.corner_radius : 6
        }
      }

      if (artType === 'prioritization') {
        const ps = art.priority_style || {}
        const rankPalette = [
          bt.secondary_color,
          bt.primary_color,
          ...(bt.accent_colors || []),
          ...(bt.chart_palette || [])
        ].filter(Boolean)
        const qualifierPalette = [
          bt.primary_color,
          bt.secondary_color,
          ...(bt.accent_colors || []),
          ...(bt.chart_palette || [])
        ].filter(Boolean)
        art.priority_style = {
          row_fill_color: ps.row_fill_color || '#FFFFFF',
          row_border_color: ps.row_border_color || '#D7DEE8',
          row_border_width: ps.row_border_width != null ? ps.row_border_width : 0.6,
          row_corner_radius: ps.row_corner_radius != null ? ps.row_corner_radius : 6,
          row_gap_in: ps.row_gap_in != null ? ps.row_gap_in : 0.16,
          rank_palette: (ps.rank_palette && ps.rank_palette.length ? ps.rank_palette : rankPalette),
          rank_font_family: ps.rank_font_family || bt.title_font_family || 'Arial',
          rank_font_size: ps.rank_font_size || 17,
          rank_text_color: ps.rank_text_color || '#FFFFFF',
          title_font_family: ps.title_font_family || bt.title_font_family || 'Arial',
          title_font_size: ps.title_font_size || 14,
          title_color: ps.title_color || '#1F2937',
          description_font_family: ps.description_font_family || bt.body_font_family || 'Arial',
          description_font_size: ps.description_font_size || 11,
          description_color: ps.description_color || '#374151',
          qualifier_fill_color: ps.qualifier_fill_color || '#EEF4E2',
          qualifier_text_color: ps.qualifier_text_color || '#1F2937',
          qualifier_value_palette: (ps.qualifier_value_palette && ps.qualifier_value_palette.length ? ps.qualifier_value_palette : qualifierPalette),
          qualifier_label_font_family: ps.qualifier_label_font_family || bt.body_font_family || 'Arial',
          qualifier_label_font_size: ps.qualifier_label_font_size || 10,
          qualifier_value_font_family: ps.qualifier_value_font_family || bt.title_font_family || 'Arial',
          qualifier_value_font_size: ps.qualifier_value_font_size || 10
        }
      }

      if (artType === 'insight_text' && art.insight_mode !== 'grouped') {
        const bs = art.body_style
        if (bs && bs.font_size != null) {
          const points      = art.points || []
          const n_points    = points.length
          if (n_points > 0) {
            const artH        = art.h || 0
            const body_h      = artH - 0.40 - 0.10 - 0.08
            const spec_fs     = bs.font_size
            const line_spacing  = bs.line_spacing    || 1.3
            const space_before  = bs.space_before_pt || 6
            const line_h_in   = spec_fs * line_spacing / 72
            const space_in    = space_before / 72
            const total_h     = n_points * line_h_in + (n_points - 1) * space_in

            if (total_h > body_h * 1.05) {
              const scaled_fs = Math.max(7, Math.floor(spec_fs * body_h / total_h))
              bs.font_size = scaled_fs
              if (bs.space_before_pt != null) {
                bs.space_before_pt = Math.round(bs.space_before_pt * scaled_fs / spec_fs)
              }
            }
          }
        }
      }
    } // end per-artifact loop
  } // end zones loop

  return zones
}


// ═══════════════════════════════════════════════════════════════════════════════
// BLOCK FLATTENER
// Converts the final processed slide spec into a flat, ordered blocks[] array.
// Each block is self-contained: block_type + x/y/w/h + type-specific fields.
// Called after computeArtifactInternals in normaliseDesignedSlide.
// generate_pptx.py reads blocks[] and dispatches each to a typed renderer.
// ═══════════════════════════════════════════════════════════════════════════════

function resolveArtifactSubtype(art) {
  if (!art || typeof art !== 'object') return 'generic'
  switch (art.type) {
    case 'chart':        return art.chart_type || 'generic'
    case 'insight_text': return art.insight_mode || 'standard'
    case 'workflow':     return art.workflow_type || art.flow_direction || 'workflow'
    case 'cards':        return art.cards_layout || 'cards'
    case 'table':        return art.table_subtype || 'standard'
    case 'matrix':       return art.matrix_type || '2x2'
    case 'driver_tree':  return 'driver_tree'
    case 'prioritization': return 'ranked_list'
    default:             return art.type || 'generic'
  }
}

function resolveArtifactHeaderText(art) {
  if (!art || typeof art !== 'object') return ''
  return ((art.header_block || {}).text ||
    art.insight_header ||
    art.chart_header ||
    art.table_header ||
    art.matrix_header ||
    art.tree_header ||
    art.priority_header ||
    art.workflow_header ||
    art.heading ||
    '')
}

function buildBlockFallbackPolicy(art, blockRole) {
  const artifactType = art?.type || 'generic'
  const artifactSubtype = resolveArtifactSubtype(art)
  return {
    allow_renderer_fallback: true,
    fallback_mode: 'subtype_default',
    trigger: 'missing_or_invalid_spec',
    artifact_type: artifactType,
    artifact_subtype: artifactSubtype,
    block_role: blockRole || 'artifact_body',
    fallback_key: artifactType + ':' + artifactSubtype
  }
}

function estimateHeaderBlockHeight(text, widthIn, fontSizePt) {
  const textStr = String(text || '').trim()
  if (!textStr) return 0.3
  const usableWidth = Math.max(0.6, Number(widthIn) || 0.6)
  const fontSize = Math.max(8, Number(fontSizePt) || 11)
  const charsPerLine = Math.max(10, Math.floor((usableWidth * 72) / (fontSize * 0.52)))
  const words = textStr.split(/\s+/).filter(Boolean)
  let lines = 1
  let lineLen = 0
  for (const word of words) {
    const nextLen = lineLen === 0 ? word.length : lineLen + 1 + word.length
    if (nextLen <= charsPerLine) lineLen = nextLen
    else {
      lines += 1
      lineLen = word.length
    }
  }
  const textHeight = lines * (fontSize * 1.28 / 72)
  return Math.max(0.3, Math.round((textHeight + 0.06) * 100) / 100)
}

function normalizeArtifactHeaderBands(zones) {
  const items = []
  for (const zone of (zones || [])) {
    for (const art of (zone.artifacts || [])) {
      const hb = art && art.header_block
      if (!hb || !hb.text) continue
      const style = hb.style || 'underline'
      if (style !== 'underline') continue
      const hy = Number(hb.y != null ? hb.y : art.y)
      const hw = Number(hb.w != null ? hb.w : art.w)
      const hfs = Number(hb.font_size || 11)
      if (!isFinite(hy) || !isFinite(hw)) continue
      const estimatedH = estimateHeaderBlockHeight(hb.text, hw, hfs)
      items.push({
        art,
        hb,
        y: hy,
        bottom: hy + Math.max(Number(hb.h || 0), estimatedH)
      })
    }
  }
  const groups = new Map()
  for (const item of items) {
    const key = String(Math.round(item.y * 8) / 8)
    if (!groups.has(key)) groups.set(key, [])
    groups.get(key).push(item)
  }
  for (const group of groups.values()) {
    if (group.length < 2) continue
    const alignedBottom = Math.max(...group.map(g => g.bottom))
    for (const g of group) {
      g.hb.h = Math.round(Math.max(Number(g.hb.h || 0), alignedBottom - g.y) * 100) / 100
    }
  }
}

function decorateArtifactBlocks(blocks, startIdx, endIdx, art, blockRole) {
  if (!art || startIdx >= endIdx) return
  const artifactType = art.type || 'generic'
  const artifactSubtype = resolveArtifactSubtype(art)
  const artifactHeaderText = resolveArtifactHeaderText(art)
  const fallbackPolicy = buildBlockFallbackPolicy(art, blockRole)
  for (let i = startIdx; i < endIdx; i++) {
    blocks[i] = {
      ...blocks[i],
      artifact_type: blocks[i].artifact_type || artifactType,
      artifact_subtype: blocks[i].artifact_subtype || artifactSubtype,
      artifact_header_text: blocks[i].artifact_header_text != null ? blocks[i].artifact_header_text : artifactHeaderText,
      block_role: blocks[i].block_role || blockRole,
      fallback_policy: blocks[i].fallback_policy || fallbackPolicy
    }
  }
}

function _matrixToBlocks(art, content_y, blocks, bt, r2) {
  const ms = art.matrix_style || {}
  const xAxis = art.x_axis || {}
  const yAxis = art.y_axis || {}
  const quadrants = art.quadrants || []
  const points = art.points || []
  const palette = ms.point_palette || [bt.primary_color, bt.secondary_color, ...(bt.accent_colors || []), ...(bt.chart_palette || [])].filter(Boolean)

  const ax = art.x || 0
  const ay = content_y
  const aw = art.w || 0
  const ah = r2((art.y || 0) + (art.h || 0) - content_y)

  const leftBand = r2(Math.min(Math.max(0.72, aw * 0.18), 1.05))
  const bottomBand = r2(Math.min(Math.max(0.62, ah * 0.16), 0.95))
  const topPad = 0.04
  const rightPad = 0.05

  const gridX = r2(ax + leftBand)
  const gridY = r2(ay + topPad)
  const gridW = r2(Math.max(1.6, aw - leftBand - rightPad))
  const gridH = r2(Math.max(1.4, ah - bottomBand - topPad))
  const midX = r2(gridX + gridW / 2)
  const midY = r2(gridY + gridH / 2)
  const quadW = r2(gridW / 2)
  const quadH = r2(gridH / 2)

  const fills = ms.quadrant_fills || ['#FFF4BF', '#E4F2DE', '#F4F5F7', '#DDEFF5']
  const quadRects = [
    { id: 'q1', x: gridX, y: gridY, w: quadW, h: quadH, fill: fills[0] },
    { id: 'q2', x: midX, y: gridY, w: quadW, h: quadH, fill: fills[1] },
    { id: 'q3', x: gridX, y: midY, w: quadW, h: quadH, fill: fills[2] },
    { id: 'q4', x: midX, y: midY, w: quadW, h: quadH, fill: fills[3] }
  ]

  blocks.push({
    block_type: 'rect',
    x: gridX, y: gridY, w: gridW, h: gridH,
    fill_color: '#FFFFFF',
    border_color: ms.border_color || '#D7DEE8',
    border_width: ms.border_width != null ? ms.border_width : 0.8,
    corner_radius: 0
  })

  quadRects.forEach(q => {
    blocks.push({
      block_type: 'rect',
      x: q.x, y: q.y, w: q.w, h: q.h,
      fill_color: q.fill,
      border_color: null,
      border_width: 0,
      corner_radius: 0
    })
  })

  blocks.push({
    block_type: 'rect',
    x: midX, y: gridY, w: 0.02, h: gridH,
    fill_color: ms.divider_color || '#7B8794',
    border_color: null,
    border_width: 0,
    corner_radius: 0
  })
  blocks.push({
    block_type: 'rect',
    x: gridX, y: midY, w: gridW, h: 0.02,
    fill_color: ms.divider_color || '#7B8794',
    border_color: null,
    border_width: 0,
    corner_radius: 0
  })

  const axisColor = ms.axis_color || '#4B5563'
  blocks.push({
    block_type: 'rect',
    x: gridX, y: r2(gridY + gridH), w: gridW, h: 0.02,
    fill_color: axisColor,
    border_color: null,
    border_width: 0,
    corner_radius: 0
  })
  blocks.push({
    block_type: 'rect',
    x: gridX, y: gridY, w: 0.02, h: gridH,
    fill_color: axisColor,
    border_color: null,
    border_width: 0,
    corner_radius: 0
  })

  const axisFont = ms.axis_label_font_family || bt.body_font_family || 'Arial'
  const axisFs = ms.axis_label_font_size || 10
  const axisTextColor = ms.axis_label_color || bt.body_color || '#374151'

  blocks.push({
    block_type: 'text_box',
    x: gridX,
    y: r2(gridY + gridH + 0.06),
    w: gridW,
    h: 0.22,
    text: xAxis.label || '',
    font_family: axisFont,
    font_size: axisFs + 1,
    bold: true,
    color: axisTextColor,
    align: 'center',
    valign: 'middle'
  })
  blocks.push({
    block_type: 'text_box',
    x: gridX,
    y: r2(gridY + gridH + 0.28),
    w: 0.6,
    h: 0.2,
    text: xAxis.low_label || 'Low',
    font_family: axisFont,
    font_size: axisFs,
    bold: false,
    color: axisTextColor,
    align: 'left',
    valign: 'middle'
  })
  blocks.push({
    block_type: 'text_box',
    x: r2(gridX + gridW - 0.6),
    y: r2(gridY + gridH + 0.28),
    w: 0.6,
    h: 0.2,
    text: xAxis.high_label || 'High',
    font_family: axisFont,
    font_size: axisFs,
    bold: false,
    color: axisTextColor,
    align: 'right',
    valign: 'middle'
  })

  blocks.push({
    block_type: 'text_box',
    x: ax,
    y: r2(gridY + gridH / 2 - 0.32),
    w: r2(Math.max(0.55, leftBand - 0.1)),
    h: 0.64,
    text: yAxis.label || '',
    font_family: axisFont,
    font_size: axisFs + 1,
    bold: true,
    color: axisTextColor,
    align: 'center',
    valign: 'middle'
  })
  blocks.push({
    block_type: 'text_box',
    x: r2(gridX - 0.52),
    y: gridY,
    w: 0.42,
    h: 0.2,
    text: yAxis.high_label || 'High',
    font_family: axisFont,
    font_size: axisFs,
    bold: false,
    color: axisTextColor,
    align: 'right',
    valign: 'middle'
  })
  blocks.push({
    block_type: 'text_box',
    x: r2(gridX - 0.52),
    y: r2(gridY + gridH - 0.2),
    w: 0.42,
    h: 0.2,
    text: yAxis.low_label || 'Low',
    font_family: axisFont,
    font_size: axisFs,
    bold: false,
    color: axisTextColor,
    align: 'right',
    valign: 'middle'
  })

  const quadMap = Object.fromEntries(quadrants.map(q => [String(q.id || '').toLowerCase(), q]))
  quadRects.forEach((rect, idx) => {
    const q = quadMap[rect.id] || quadrants[idx] || {}
    blocks.push({
      block_type: 'text_box',
      x: r2(rect.x + 0.14),
      y: r2(rect.y + 0.12),
      w: r2(rect.w - 0.28),
      h: 0.24,
      text: q.name || '',
      font_family: ms.quadrant_title_font_family || bt.title_font_family || 'Arial',
      font_size: ms.quadrant_title_font_size || 12,
      bold: true,
      color: ms.quadrant_title_color || '#2D3748',
      align: 'left',
      valign: 'top'
    })
    blocks.push({
      block_type: 'text_box',
      x: r2(rect.x + 0.14),
      y: r2(rect.y + 0.4),
      w: r2(rect.w - 0.28),
      h: 0.46,
      text: q.insight || '',
      font_family: ms.quadrant_body_font_family || bt.body_font_family || 'Arial',
      font_size: ms.quadrant_body_font_size || 9,
      bold: false,
      color: ms.quadrant_body_color || '#374151',
      align: 'left',
      valign: 'top'
    })
  })

  const semanticToRatio = { low: 0.25, medium: 0.50, high: 0.75 }
  const markerSize = 0.14
  points.slice(0, 6).forEach((pt, i) => {
    const px = r2(gridX + gridW * (semanticToRatio[String(pt.x || 'medium').toLowerCase()] || 0.5))
    const py = r2(gridY + gridH * (1 - (semanticToRatio[String(pt.y || 'medium').toLowerCase()] || 0.5)))
    const fill = palette[i % Math.max(palette.length, 1)] || bt.primary_color || '#0078AE'
    blocks.push({
      block_type: 'circle',
      x: r2(px - markerSize / 2),
      y: r2(py - markerSize / 2),
      w: markerSize,
      h: markerSize,
      fill_color: fill,
      font_color: '#FFFFFF',
      text: ''
    })

    const rightSpace = gridX + gridW - px
    const labelW = Math.min(1.1, Math.max(0.6, (String(pt.label || '').length || 8) * 0.08))
    const labelX = rightSpace > labelW + 0.18 ? r2(px + 0.12) : r2(px - labelW - 0.12)
    blocks.push({
      block_type: 'text_box',
      x: labelX,
      y: r2(py - 0.11),
      w: labelW,
      h: 0.24,
      text: pt.label || '',
      font_family: ms.point_label_font_family || bt.body_font_family || 'Arial',
      font_size: ms.point_label_font_size || 10,
      bold: false,
      color: ms.point_label_color || '#111111',
      align: 'left',
      valign: 'middle'
    })
  })
}

function _driverTreeToBlocks(art, content_y, blocks, bt, r2) {
  const ts = art.tree_style || {}
  const root = art.root || {}
  const branches = art.branches || []
  const ax = art.x || 0
  const ay = content_y
  const aw = art.w || 0
  const ah = r2((art.y || 0) + (art.h || 0) - content_y)

  const leafCount = branches.reduce((sum, b) => sum + Math.max((b.children || []).length, 1), 0) || Math.max(branches.length, 1)
  const hasThirdLevel = branches.some(b => (b.children || []).length > 0)
  const rowY = [
    r2(ay + 0.06),
    r2(ay + ah * 0.40),
    r2(ay + ah * 0.73)
  ]
  const rootW = r2(Math.min(Math.max(2.8, aw * 0.42), 3.9))
  const rootH = r2(Math.min(Math.max(0.9, ah * 0.20), 1.2))
  const branchW = r2(Math.min(Math.max(2.2, aw * 0.30), 3.0))
  const branchH = r2(Math.min(Math.max(0.8, ah * 0.18), 1.05))
  const leafW = r2(Math.min(Math.max(1.6, aw * 0.20), 2.2))
  const leafH = r2(Math.min(Math.max(0.7, ah * 0.16), 0.95))

  const rootX = r2(ax + (aw - rootW) / 2)
  const rootY = rowY[0]

  const leafCenters = []
  if (leafCount === 1) {
    leafCenters.push(r2(ax + aw / 2))
  } else {
    const left = ax + leafW / 2
    const usable = Math.max(0.5, aw - leafW)
    const step = usable / (leafCount - 1)
    for (let i = 0; i < leafCount; i++) leafCenters.push(r2(left + i * step))
  }

  let cursor = 0
  const branchLayout = branches.map((branch, i) => {
    const childCount = Math.max((branch.children || []).length, 1)
    const branchLeafCenters = leafCenters.slice(cursor, cursor + childCount)
    cursor += childCount
    const centerX = branchLeafCenters.length
      ? r2(branchLeafCenters.reduce((s, x) => s + x, 0) / branchLeafCenters.length)
      : r2(ax + aw / 2)
    return {
      branch,
      centerX,
      children: (branch.children || []).length
        ? branch.children.map((child, ci) => ({ child, centerX: branchLeafCenters[ci] }))
        : [{ child: null, centerX }]
    }
  })

  const pushNode = (x, y, w, h, fill, label, value, isRoot) => {
    blocks.push({
      block_type: 'rect',
      x, y, w, h,
      fill_color: fill,
      border_color: ts.node_border_color || '#D7DEE8',
      border_width: ts.node_border_width != null ? ts.node_border_width : 0.6,
      corner_radius: ts.corner_radius != null ? ts.corner_radius : 6
    })
    const labelH = value ? r2(h * 0.52) : r2(h * 0.72)
    blocks.push({
      block_type: 'text_box',
      x: r2(x + 0.1), y: r2(y + 0.08), w: r2(w - 0.2), h: labelH,
      text: label || '',
      font_family: ts.label_font_family || bt.title_font_family || 'Arial',
      font_size: isRoot ? (ts.label_font_size || 11) + 1 : (ts.label_font_size || 11),
      bold: false,
      color: ts.label_color || '#111111',
      align: 'center',
      valign: 'top'
    })
    if (value) {
      blocks.push({
        block_type: 'text_box',
        x: r2(x + 0.1), y: r2(y + h * 0.58), w: r2(w - 0.2), h: r2(h * 0.22),
        text: value,
        font_family: ts.value_font_family || bt.body_font_family || 'Arial',
        font_size: isRoot ? (ts.value_font_size || 10) + 2 : (ts.value_font_size || 10) + 1,
        bold: true,
        color: ts.value_color || bt.primary_color || '#0078AE',
        align: 'center',
        valign: 'middle'
      })
    }
  }

  const pushConnector = (x, y, w, h) => {
    blocks.push({
      block_type: 'rect',
      x: r2(x), y: r2(y), w: r2(w), h: r2(h),
      fill_color: ts.connector_color || '#7A8FA8',
      border_color: null,
      border_width: 0,
      corner_radius: 0
    })
  }

  pushNode(rootX, rootY, rootW, rootH, ts.node_fill_color || '#EAF2FB', root.label, root.value, true)

  const rootBottomX = r2(rootX + rootW / 2)
  const branchY = rowY[1]
  const branchBottomY = r2(branchY + branchH)
  const branchCenters = branchLayout.map(b => b.centerX)
  if (branchCenters.length) {
    const trunkBottomY = r2(branchY - 0.14)
    pushConnector(r2(rootBottomX - 0.01), r2(rootY + rootH), 0.02, r2(trunkBottomY - (rootY + rootH)))
    pushConnector(Math.min(...branchCenters), r2(trunkBottomY - 0.01), Math.max(0.02, Math.max(...branchCenters) - Math.min(...branchCenters)), 0.02)
  }

  branchLayout.forEach((entry, i) => {
    const bx = r2(entry.centerX - branchW / 2)
    pushNode(bx, branchY, branchW, branchH, ts.node_fill_color_secondary || '#EDF7F3', entry.branch.label, entry.branch.value, false)
    pushConnector(r2(entry.centerX - 0.01), r2(branchY - 0.14), 0.02, 0.14)

    if (!hasThirdLevel) return
    const childCenters = entry.children.map(c => c.centerX)
    const childY = rowY[2]
    const childTopY = childY
    pushConnector(r2(entry.centerX - 0.01), branchBottomY, 0.02, r2((childY - 0.14) - branchBottomY))
    if (childCenters.length > 1) {
      pushConnector(Math.min(...childCenters), r2(childY - 0.15), Math.max(0.02, Math.max(...childCenters) - Math.min(...childCenters)), 0.02)
    }

    entry.children.forEach(({ child, centerX }) => {
      const lx = r2(centerX - leafW / 2)
      pushConnector(r2(centerX - 0.01), r2(childY - 0.15), 0.02, 0.15)
      const label = child ? child.label : entry.branch.label
      const value = child ? child.value : entry.branch.value
      pushNode(lx, childTopY, leafW, leafH, ts.node_fill_color_leaf || '#F4F7FA', label, value, false)
    })
  })
}

function _prioritizationToBlocks(art, content_y, blocks, bt, r2) {
  const ps = art.priority_style || {}
  const items = (art.items || []).slice().sort((a, b) => (+a.rank || 999) - (+b.rank || 999)).slice(0, 5)
  const ax = art.x || 0
  const ay = content_y
  const aw = art.w || 0
  const ah = r2((art.y || 0) + (art.h || 0) - content_y)
  if (!items.length || aw <= 0 || ah <= 0) return

  const gap = ps.row_gap_in != null ? ps.row_gap_in : 0.16
  const rowH = r2((ah - gap * Math.max(0, items.length - 1)) / Math.max(items.length, 1))
  const rankSize = r2(Math.min(0.62, Math.max(0.42, rowH * 0.40)))
  const leftPad = 0.12
  const rightPad = 0.14
  const rankPalette = ps.rank_palette || [bt.secondary_color || '#E0B324', bt.primary_color || '#0078AE']
  const qualifierPalette = ps.qualifier_value_palette || [bt.primary_color || '#0078AE']

  items.forEach((item, idx) => {
    const rowY = r2(ay + idx * (rowH + gap))
    const rowX = ax
    const rowW = aw
    const qualifiers = Array.isArray(item.qualifiers) ? item.qualifiers.slice(0, 2) : []
    const nonEmptyQualifiers = qualifiers.filter(q => String(q?.label || '').trim() || String(q?.value || '').trim())
    const qualifierAreaW = nonEmptyQualifiers.length ? r2(Math.min(1.9, Math.max(1.45, aw * 0.26))) : 0
    const rankX = r2(rowX + leftPad)
    const rankY = r2(rowY + (rowH - rankSize) / 2)
    const textX = r2(rankX + rankSize + 0.14)
    const textW = r2(Math.max(1.1, rowW - (textX - rowX) - qualifierAreaW - rightPad - (qualifierAreaW ? 0.14 : 0)))
    const qualifierX = qualifierAreaW ? r2(rowX + rowW - rightPad - qualifierAreaW) : 0
    const rankFill = rankPalette[idx % Math.max(rankPalette.length, 1)] || bt.primary_color || '#0078AE'

    blocks.push({
      block_type: 'rect',
      x: rowX, y: rowY, w: rowW, h: rowH,
      fill_color: ps.row_fill_color || '#FFFFFF',
      border_color: ps.row_border_color || '#D7DEE8',
      border_width: ps.row_border_width != null ? ps.row_border_width : 0.6,
      corner_radius: ps.row_corner_radius != null ? ps.row_corner_radius : 6
    })

    blocks.push({
      block_type: 'circle',
      x: rankX, y: rankY, w: rankSize, h: rankSize,
      fill_color: rankFill,
      font_color: ps.rank_text_color || '#FFFFFF',
      text: String(item.rank != null ? item.rank : idx + 1)
    })
    blocks.push({
      block_type: 'text_box',
      x: rankX, y: rankY, w: rankSize, h: rankSize,
      text: String(item.rank != null ? item.rank : idx + 1),
      font_family: ps.rank_font_family || bt.title_font_family || 'Arial',
      font_size: ps.rank_font_size || 17,
      bold: true,
      color: ps.rank_text_color || '#FFFFFF',
      align: 'center',
      valign: 'middle'
    })

    const titleH = r2(Math.min(0.34, Math.max(0.24, rowH * 0.34)))
    const descY = r2(rowY + 0.12 + titleH)
    const descH = r2(Math.max(0.24, rowH - (descY - rowY) - 0.12))
    blocks.push({
      block_type: 'text_box',
      x: textX, y: r2(rowY + 0.1), w: textW, h: titleH,
      text: item.title || '',
      font_family: ps.title_font_family || bt.title_font_family || 'Arial',
      font_size: ps.title_font_size || 14,
      bold: true,
      color: ps.title_color || '#1F2937',
      align: 'left',
      valign: 'top'
    })
    blocks.push({
      block_type: 'text_box',
      x: textX, y: descY, w: textW, h: descH,
      text: item.description || '',
      font_family: ps.description_font_family || bt.body_font_family || 'Arial',
      font_size: ps.description_font_size || 11,
      bold: false,
      color: ps.description_color || '#374151',
      align: 'left',
      valign: 'top'
    })

    if (nonEmptyQualifiers.length) {
      const pillGap = 0.08
      const pillCount = nonEmptyQualifiers.length
      const pillH = r2(Math.min(0.28, Math.max(0.2, (rowH - 0.18 - pillGap * Math.max(0, pillCount - 1)) / Math.max(pillCount, 1))))
      nonEmptyQualifiers.forEach((q, qi) => {
        const pillY = r2(rowY + 0.12 + qi * (pillH + pillGap))
        const valueColor = qualifierPalette[qi % Math.max(qualifierPalette.length, 1)] || bt.primary_color || '#0078AE'
        blocks.push({
          block_type: 'rect',
          x: qualifierX, y: pillY, w: qualifierAreaW, h: pillH,
          fill_color: ps.qualifier_fill_color || '#EEF4E2',
          border_color: null,
          border_width: 0,
          corner_radius: 4
        })
        const label = String(q.label || '').trim()
        const value = String(q.value || '').trim()
        const pillText = label && value ? (label + ': ' + value) : (label || value)
        blocks.push({
          block_type: 'text_box',
          x: r2(qualifierX + 0.08), y: pillY, w: r2(qualifierAreaW - 0.16), h: pillH,
          text: pillText,
          font_family: ps.qualifier_label_font_family || bt.body_font_family || 'Arial',
          font_size: ps.qualifier_label_font_size || 10,
          bold: false,
          color: ps.qualifier_text_color || '#1F2937',
          align: 'center',
          valign: 'middle'
        })
      })
    }
  })
}

function _artifactToBlocks(art, blocks, bt, r2) {
  const ax = art.x || 0
  const ay = art.y || 0
  const aw = art.w || 0
  const ah = art.h || 0
  const blockStart = blocks.length

  // ── Artifact header band (if present) ────────────────────────────────────
  // header_block sits above the artifact body, already has its own x/y/w/h
  const hb          = art.header_block || null
  let   content_y   = ay   // top of the body area (after header_block)

  if (hb && hb.text) {
    const hx  = hb.x  != null ? hb.x  : ax
    const hy  = hb.y  != null ? hb.y  : ay
    const hw  = hb.w  != null ? hb.w  : aw
    const hfs = hb.font_size || 11
    const estimatedH = estimateHeaderBlockHeight(hb.text, hw, hfs)
    const hh  = Math.max(hb.h != null ? hb.h : 0.30, estimatedH)
    content_y = r2(hy + hh + 0.04)  // small gap below header band

    const headerStyle = hb.style || 'underline'
    if (headerStyle === 'brand_fill') {
      // Filled header band
      blocks.push({
        block_type:    'rect',
        x: hx, y: hy, w: hw, h: hh,
        fill_color:    hb.fill_color   || bt.primary_color || '#1A3C8F',
        border_color:  null,
        border_width:  0,
        corner_radius: hb.corner_radius || 0
      })
      blocks.push({
        block_type:  'text_box',
        x: r2(hx + 0.08), y: hy, w: r2(hw - 0.16), h: hh,
        text:        hb.text,
        font_family: hb.font_family || bt.title_font_family || 'Arial',
        font_size:   hfs,
        bold:        true,
        color:       hb.text_color || '#FFFFFF',
        align:       'left',
        valign:      'middle'
      })
    } else {
      // Underline header
      blocks.push({
        block_type:  'text_box',
        x: hx, y: hy, w: hw, h: hh,
        text:        hb.text,
        font_family: hb.font_family || bt.title_font_family || 'Arial',
        font_size:   hfs,
        bold:        true,
        color:       hb.color || bt.primary_color || '#1A3C8F',
        align:       'left',
        valign:      'top'
      })
      blocks.push({
        block_type: 'rule',
        x: hx, y: r2(hy + hh), w: hw, h: 0.03,
        color:      hb.rule_color || bt.primary_color || '#1A3C8F'
      })
    }
  }

  // ── Artifact body ─────────────────────────────────────────────────────────
  const headerEnd = blocks.length
  switch (art.type) {

    case 'chart': {
      const computed = art._computed || {}
      blocks.push({
        block_type:              'chart',
        x: ax, y: content_y, w: aw, h: r2(ay + ah - content_y),
        chart_type:              art.chart_type,
        chart_header:            art.chart_header || '',
        chart_title:             art.chart_title  || '',
        categories:              art.categories   || [],
        series:                  art.series       || [],
        dual_axis:               art.dual_axis    || false,
        secondary_series:        art.secondary_series || [],
        show_data_labels:        art.show_data_labels !== false,
        show_legend:             !!art.show_legend,
        x_label:                 art.x_label || '',
        y_label:                 art.y_label || '',
        chart_style:             art.chart_style   || {},
        series_style:            art.series_style  || [],
        brand_tokens:            { primary_color: bt.primary_color, chart_palette: bt.chart_palette },
        // Pre-computed by computeArtifactInternals — renderer reads these directly
        legend_position:         computed.legend_position        || 'none',
        data_label_size:         computed.data_label_size        || 9,
        category_label_rotation: computed.category_label_rotation || 0
      })
      break
    }

    case 'insight_text': {
      if (art.insight_mode === 'grouped') {
        _groupedInsightToBlocks(art, content_y, blocks, bt, r2)
      } else {
        _standardInsightToBlocks(art, content_y, blocks, r2)
      }
      break
    }

    case 'table': {
      const tableY = content_y
      const tableH = r2(ay + ah - content_y)
      const colWs = art.column_widths || []
      const dataRowHs = art.row_heights || []
      const headerH = art.header_row_height || 0.35

      let curX = ax
      const columnXPositions = colWs.map(cw => {
        const x = r2(curX)
        curX += (+cw || 0)
        return x
      })

      let curY = tableY
      const rowYPositions = [r2(curY)]
      curY += (+headerH || 0)
      for (const rh of dataRowHs) {
        rowYPositions.push(r2(curY))
        curY += (+rh || 0)
      }

      const headerCellFrames = colWs.map((cw, ci) => ({
        col_index: ci,
        x: r2(columnXPositions[ci] || ax),
        y: r2(tableY),
        w: r2(+cw || 0),
        h: r2(+headerH || 0)
      }))

      const bodyCellFrames = dataRowHs.map((rh, ri) =>
        colWs.map((cw, ci) => ({
          row_index: ri,
          col_index: ci,
          x: r2(columnXPositions[ci] || ax),
          y: r2(rowYPositions[ri + 1] || tableY),
          w: r2(+cw || 0),
          h: r2(+rh || 0)
        }))
      )

      blocks.push({
        block_type:         'table',
        x: ax, y: tableY, w: aw, h: tableH,
        headers:            art.headers            || [],
        rows:               art.rows               || [],
        column_widths:      colWs,
        column_x_positions: columnXPositions,
        column_types:       art.column_types        || [],
        column_alignments:  art.column_alignments   || [],
        row_heights:        dataRowHs,
        header_row_height:  headerH,
        row_y_positions:    rowYPositions,
        header_cell_frames: headerCellFrames,
        body_cell_frames:   bodyCellFrames,
        table_style:        art.table_style         || {}
      })
      break
    }

    case 'matrix': {
      _matrixToBlocks(art, content_y, blocks, bt, r2)
      break
    }

    case 'driver_tree': {
      _driverTreeToBlocks(art, content_y, blocks, bt, r2)
      break
    }

    case 'prioritization': {
      _prioritizationToBlocks(art, content_y, blocks, bt, r2)
      break
    }

    case 'cards': {
      _cardsToBlocks(art, content_y, blocks, bt, r2)
      break
    }

    case 'workflow': {
      blocks.push({
        block_type:     'workflow',
        x: ax, y: content_y, w: aw, h: r2(ay + ah - content_y),
        nodes:          art.nodes       || [],
        connections:    art.connections || [],
        workflow_style: art.workflow_style || {},
        flow_direction: art.flow_direction || '',
        workflow_type:  art.workflow_type  || ''
      })
      break
    }

    default:
      break
  }
  decorateArtifactBlocks(blocks, blockStart, headerEnd, art, 'artifact_header')
  decorateArtifactBlocks(blocks, headerEnd, blocks.length, art, 'artifact_body')
}

function _standardInsightToBlocks(art, content_y, blocks, r2) {
  const ax = art.x || 0
  const ay = art.y || 0
  const aw = art.w || 0
  const ah = art.h || 0
  const st = art.body_style || {}

  const body_y = content_y
  const body_h = r2(ay + ah - body_y)
  const sty    = art.style || {}   // fill/border live in art.style per schema

  // Container rect (fill/border)
  const hasFill   = !!sty.fill_color
  const hasBorder = !!(sty.border_color && sty.border_width)
  if (hasFill || hasBorder) {
    blocks.push({
      block_type:    'rect',
      x: ax, y: body_y, w: aw, h: body_h,
      fill_color:    sty.fill_color    || null,
      border_color:  sty.border_color  || null,
      border_width:  sty.border_width  || 0.75,
      corner_radius: sty.corner_radius || 0
    })
  }

  // Bullet list body
  blocks.push({
    block_type:   'bullet_list',
    x: ax, y: body_y, w: aw, h: body_h,
    points:       art.points || [],
    body_style:   st,
    sentiment:    art.sentiment || 'neutral'
  })
}

function _groupedInsightToBlocks(art, content_y, blocks, bt, r2) {
  const ax      = art.x || 0
  const ay      = art.y || 0
  const aw      = art.w || 0
  const ah      = art.h || 0
  const groups  = art.groups || []
  const n       = groups.length
  if (n === 0) return

  const ghs     = art.group_header_style    || {}
  const gbs     = art.group_bullet_box_style || {}
  const bsty    = art.bullet_style          || {}
  const g_gap   = art.group_gap_in          || 0.08
  const hb_gap  = art.header_to_box_gap_in  || 0.05
  const gLayout = art.group_layout          || 'rows'
  const isBadge = ghs.shape === 'circle_badge'

  const h_fill  = ghs.fill_color   || bt.primary_color || '#1A3C8F'
  const h_cr    = ghs.corner_radius || 4

  const total_content_h = r2(ay + ah - content_y)

  if (gLayout === 'rows') {
    const h_w         = ghs.w || 1.2
    const box_x       = r2(ax + h_w + hb_gap)
    const box_w       = r2(aw - h_w - hb_gap)
    const total_bullets = Math.max(1, groups.reduce((s, g) => s + (g.bullets || []).length, 0))
    const total_rh    = Math.max(0.2, total_content_h - (n - 1) * g_gap)

    let cur_y = content_y
    for (let gi = 0; gi < groups.length; gi++) {
      const g        = groups[gi]
      const nbullets = Math.max(1, (g.bullets || []).length)
      const row_h    = r2(Math.max(0.25, total_rh * (nbullets / total_bullets)))

      if (isBadge) {
        // circle_badge: circle centered vertically in the header column
        const dia      = ghs.h || 0.3
        const badge_y  = r2(cur_y + (row_h - dia) / 2)
        blocks.push({
          block_type:  'circle',
          x: ax, y: badge_y, w: dia, h: dia,
          fill_color:  h_fill,
          text:        String(gi + 1),
          font_family: ghs.font_family || bt.title_font_family || 'Arial',
          font_size:   ghs.font_size   || 10,
          font_color:  ghs.text_color  || '#FFFFFF'
        })
      } else {
        // rounded_rect: fills full h_w × row_h
        blocks.push({
          block_type: 'rect',
          x: ax, y: r2(cur_y), w: h_w, h: row_h,
          fill_color: h_fill, border_color: null, border_width: 0, corner_radius: h_cr
        })
        blocks.push({
          block_type:  'text_box',
          x: r2(ax + 0.06), y: r2(cur_y), w: r2(h_w - 0.12), h: row_h,
          text:        String(g.header || ''),
          font_family: ghs.font_family || bt.title_font_family || 'Arial',
          font_size:   ghs.font_size   || 10,
          bold:        true,
          color:       ghs.text_color  || '#FFFFFF',
          align:       'center',
          valign:      'middle'
        })
      }
      // Bullet box — rect (border only)
      if (gbs.fill_color || gbs.border_color) {
        blocks.push({
          block_type:    'rect',
          x: box_x, y: r2(cur_y), w: box_w, h: row_h,
          fill_color:    gbs.fill_color   || null,
          border_color:  gbs.border_color || null,
          border_width:  gbs.border_width || 0.75,
          corner_radius: gbs.corner_radius || 4
        })
      }
      // Bullet box — bullet list
      blocks.push({
        block_type:  'bullet_list',
        x: box_x, y: r2(cur_y), w: box_w, h: row_h,
        points:      g.bullets || [],
        body_style:  bsty,
        padding:     gbs.padding || {},
        sentiment:   art.sentiment || 'neutral'
      })

      cur_y = r2(cur_y + row_h + g_gap)
    }

  } else {
    // columns layout
    const col_w  = r2((aw - (n - 1) * g_gap) / Math.max(n, 1))
    const h_h    = ghs.h || 0.28
    const box_h  = r2(total_content_h - h_h - hb_gap)

    let cur_x = ax
    for (let gi = 0; gi < groups.length; gi++) {
      const g = groups[gi]

      if (isBadge) {
        // circle_badge: circle centered horizontally in col_w
        const dia     = h_h
        const badge_x = r2(cur_x + (col_w - dia) / 2)
        blocks.push({
          block_type:  'circle',
          x: badge_x, y: content_y, w: dia, h: dia,
          fill_color:  h_fill,
          text:        String(gi + 1),
          font_family: ghs.font_family || bt.title_font_family || 'Arial',
          font_size:   ghs.font_size   || 10,
          font_color:  ghs.text_color  || '#FFFFFF'
        })
      } else {
        // rounded_rect: spans full col_w as a header bar
        blocks.push({
          block_type: 'rect',
          x: r2(cur_x), y: content_y, w: col_w, h: h_h,
          fill_color: h_fill, border_color: null, border_width: 0, corner_radius: h_cr
        })
        blocks.push({
          block_type:  'text_box',
          x: r2(cur_x + 0.05), y: content_y, w: r2(col_w - 0.10), h: h_h,
          text:        String(g.header || ''),
          font_family: ghs.font_family || bt.title_font_family || 'Arial',
          font_size:   ghs.font_size   || 10,
          bold:        true,
          color:       ghs.text_color  || '#FFFFFF',
          align:       'center',
          valign:      'middle'
        })
      }
      const bullet_y = r2(content_y + h_h + hb_gap)
      // Bullet box — rect
      if (gbs.fill_color || gbs.border_color) {
        blocks.push({
          block_type:    'rect',
          x: r2(cur_x), y: bullet_y, w: col_w, h: box_h,
          fill_color:    gbs.fill_color   || null,
          border_color:  gbs.border_color || null,
          border_width:  gbs.border_width || 0.75,
          corner_radius: gbs.corner_radius || 4
        })
      }
      // Bullet box — bullets
      blocks.push({
        block_type: 'bullet_list',
        x: r2(cur_x), y: bullet_y, w: col_w, h: box_h,
        points:     g.bullets || [],
        body_style: bsty,
        padding:    gbs.padding || {},
        sentiment:  art.sentiment || 'neutral'
      })

      cur_x = r2(cur_x + col_w + g_gap)
    }
  }
}

function _cardsToBlocks(art, content_y, blocks, bt, r2) {
  const cards  = art.cards  || []
  const frames = art.card_frames || []   // pre-computed by computeArtifactInternals
  const cs     = art.card_style || {}
  const ts     = art.title_style || {}
  const subs   = art.subtitle_style || {}
  const bs     = art.body_style || {}
  const pad    = cs.internal_padding || 0.12
  const accentW = 0.07
  const accentGap = 0.08
  const sentimentAccent = {
    positive: bt.secondary_color || '#2D8A4E',
    negative: '#C0392B',
    neutral: bt.primary_color || '#1A3C8F'
  }
  const paletteBase = [
    bt.primary_color,
    bt.secondary_color,
    ...((bt.accent_colors || []).length ? bt.accent_colors : []),
    ...((bt.chart_palette || []).length ? bt.chart_palette : [])
  ].filter(Boolean)
  const accentPalette = [...new Set(paletteBase)]

  for (let i = 0; i < cards.length; i++) {
    const card = cards[i]
    const fr   = frames[i] || { x: art.x || 0, y: content_y, w: art.w || 0, h: art.h || 0 }
    const fx   = fr.x, fy = fr.y, fw = fr.w, fh = fr.h
    const accentColor = cards.length > 1
      ? (accentPalette[i] || accentPalette[i % Math.max(accentPalette.length, 1)] || sentimentAccent[card.sentiment] || '#1A3C8F')
      : (sentimentAccent[card.sentiment] || accentPalette[0] || '#1A3C8F')

    // Card background rect
    blocks.push({
      block_type:    'rect',
      x: fx, y: fy, w: fw, h: fh,
      fill_color:    cs.fill_color    || '#F5F5F5',
      border_color:  cs.border_color  || '#DDDDDD',
      border_width:  cs.border_width  || 0.75,
      corner_radius: 0
    })

    // Accent strip (left colour rail)
    if (accentColor) {
      blocks.push({
        block_type:    'rect',
        x: fx, y: fy, w: accentW, h: fh,
        fill_color:    accentColor,
        border_color:  null, border_width: 0,
        corner_radius: 0
      })
    }

    // Card text sections (title / subtitle / body)
    const inner_x   = r2(fx + pad + accentW + accentGap)
    const inner_y   = r2(fy + pad)
    const inner_w   = r2(Math.max(0.3, fw - (pad * 2) - accentW - accentGap))
    const inner_h   = r2(fh - 2 * pad)
    const title_h   = r2(inner_h * 0.20)
    const sub_h     = r2(inner_h * 0.42)
    const body_h    = r2(Math.max(0.18, inner_h - title_h - sub_h - 0.08))
    const titleY    = inner_y
    const subtitleY = r2(titleY + title_h + 0.03)
    const bodyY     = r2(subtitleY + sub_h + 0.05)

    if (card.title) {
      blocks.push({
        block_type: 'text_box',
        x: inner_x, y: inner_y, w: inner_w, h: title_h,
        text:       card.title,
        font_family: ts.font_family || bt.title_font_family || 'Arial',
        font_size:   ts.font_size   || 12,
        bold:        true,
        color:       ts.color || accentColor,
        align:       'left', valign: 'top'
      })
    }
    if (card.subtitle) {
      blocks.push({
        block_type: 'text_box',
        x: inner_x, y: subtitleY, w: inner_w, h: sub_h,
        text:       card.subtitle,
        font_family: subs.font_family || bt.body_font_family || 'Arial',
        font_size:   subs.font_size   || 22,
        bold:        true,
        color:       subs.color || '#111111',
        align:       'left', valign: 'middle'
      })
    }
    if (card.body) {
      blocks.push({
        block_type: 'text_box',
        x: inner_x, y: bodyY, w: inner_w, h: body_h,
        text:       card.body,
        font_family: bs.font_family || bt.body_font_family || 'Arial',
        font_size:   bs.font_size   || 9,
        bold:        false,
        color:       bs.color || '#333333',
        align:       'left', valign: 'top'
      })
    }
  }
}

function flattenToBlocks(slideSpec, brandTokens) {
  const bt     = brandTokens || {}
  const blocks = []
  const r2     = x => Math.round(x * 100) / 100

  // ── 1. Title block ────────────────────────────────────────────────────────
  const tb = slideSpec.title_block || {}
  if (tb.text) {
    blocks.push({
      block_type:  'title',
      artifact_type: 'slide',
      artifact_subtype: 'title',
      block_role: 'slide_header',
      artifact_header_text: tb.text,
      fallback_policy: {
        allow_renderer_fallback: false,
        fallback_mode: 'none',
        artifact_type: 'slide',
        artifact_subtype: 'title',
        block_role: 'slide_header',
        fallback_key: 'slide:title'
      },
      x:           tb.x           != null ? tb.x           : 0.4,
      y:           tb.y           != null ? tb.y           : 0.15,
      w:           tb.w           != null ? tb.w           : 9.2,
      h:           tb.h           != null ? tb.h           : 0.7,
      text:        tb.text,
      font_family: tb.font_family || bt.title_font_family || 'Arial',
      font_size:   tb.font_size   || 20,
      bold:        ['bold','semibold'].includes(String(tb.font_weight || 'bold').toLowerCase()),
      color:       tb.color       || bt.title_color || '#1A3C8F',
      align:       tb.align       || 'left',
      valign:      'top'
    })
  }

  // ── 2. Subtitle block ─────────────────────────────────────────────────────
  const sb = slideSpec.subtitle_block || {}
  if (sb.text) {
    blocks.push({
      block_type:  'subtitle',
      artifact_type: 'slide',
      artifact_subtype: 'subtitle',
      block_role: 'slide_subheader',
      artifact_header_text: sb.text,
      fallback_policy: {
        allow_renderer_fallback: false,
        fallback_mode: 'none',
        artifact_type: 'slide',
        artifact_subtype: 'subtitle',
        block_role: 'slide_subheader',
        fallback_key: 'slide:subtitle'
      },
      x:           sb.x           != null ? sb.x           : 0.4,
      y:           sb.y           != null ? sb.y           : 0.9,
      w:           sb.w           != null ? sb.w           : 9.2,
      h:           sb.h           != null ? sb.h           : 0.45,
      text:        sb.text,
      font_family: sb.font_family || bt.body_font_family || 'Arial',
      font_size:   sb.font_size   || 14,
      bold:        ['bold','semibold'].includes(String(sb.font_weight || '').toLowerCase()),
      color:       sb.color       || bt.body_color || '#333333',
      align:       sb.align       || 'left',
      valign:      'middle'
    })
  }

  // ── 3. Zones → Artifacts ──────────────────────────────────────────────────
  for (const zone of (slideSpec.zones || [])) {
    for (const art of (zone.artifacts || [])) {
      _artifactToBlocks(art, blocks, bt, r2)
    }
  }

  // ── 4. Global elements ────────────────────────────────────────────────────
  const ge = slideSpec.global_elements || {}

  if (ge.logo) {
    const lg = ge.logo
    blocks.push({
      block_type: 'image',
      artifact_type: 'global_element',
      artifact_subtype: 'logo',
      block_role: 'global_element',
      artifact_header_text: '',
      fallback_policy: {
        allow_renderer_fallback: false,
        fallback_mode: 'none',
        artifact_type: 'global_element',
        artifact_subtype: 'logo',
        block_role: 'global_element',
        fallback_key: 'global_element:logo'
      },
      image_role: 'logo',
      x: lg.x != null ? lg.x : 0.2,
      y: lg.y != null ? lg.y : 0.05,
      w: lg.w || 1.2,
      h: lg.h || 0.4
    })
  }
  if (ge.footer && ge.footer.text) {
    const ft = ge.footer
    blocks.push({
      block_type:  'footer',
      artifact_type: 'global_element',
      artifact_subtype: 'footer',
      block_role: 'global_element',
      artifact_header_text: ft.text,
      fallback_policy: {
        allow_renderer_fallback: false,
        fallback_mode: 'none',
        artifact_type: 'global_element',
        artifact_subtype: 'footer',
        block_role: 'global_element',
        fallback_key: 'global_element:footer'
      },
      x:           ft.x != null ? ft.x : 0.4,
      y:           ft.y != null ? ft.y : 7.3,
      w:           ft.w || 5.0,
      h:           ft.h || 0.22,
      text:        ft.text,
      font_family: ft.font_family || bt.body_font_family || 'Arial',
      font_size:   ft.font_size   || 8,
      color:       ft.color       || '#AAAAAA',
      align:       ft.align       || 'left',
      valign:      'middle'
    })
  }
  if (ge.page_number && ge.page_number.text) {
    const pn = ge.page_number
    blocks.push({
      block_type:  'page_number',
      artifact_type: 'global_element',
      artifact_subtype: 'page_number',
      block_role: 'global_element',
      artifact_header_text: pn.text || '',
      fallback_policy: {
        allow_renderer_fallback: false,
        fallback_mode: 'none',
        artifact_type: 'global_element',
        artifact_subtype: 'page_number',
        block_role: 'global_element',
        fallback_key: 'global_element:page_number'
      },
      x:           pn.x != null ? pn.x : 9.4,
      y:           pn.y != null ? pn.y : 7.3,
      w:           pn.w || 0.8,
      h:           pn.h || 0.22,
      text:        pn.text,
      font_family: pn.font_family || bt.body_font_family || 'Arial',
      font_size:   pn.font_size   || 8,
      color:       pn.color       || '#AAAAAA',
      align:       'right',
      valign:      'middle'
    })
  }

  return blocks
}


function mergeContentIntoZones(designedZones, manifestZones, brandTokens) {
  if (!designedZones || !manifestZones) return designedZones || []

  const bt = brandTokens || {}

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
        // ── Determine mode: manifest (Agent 4) is authoritative for content structure ──
        const mGroups = mArt.groups && mArt.groups.length > 0 ? mArt.groups : null
        const mPoints = mArt.points && mArt.points.length > 0 ? mArt.points : null
        const resolvedMode = mArt.insight_mode
          || (mGroups ? 'grouped' : mPoints ? 'standard' : dArt.insight_mode || 'standard')

        const heading        = mArt.heading        || mArt.insight_header || dArt.heading        || 'Key Insight'
        const insight_header = mArt.insight_header || mArt.heading        || dArt.insight_header || 'Key Insight'
        const sentiment      = mArt.sentiment      || dArt.sentiment      || 'neutral'

        // ── FLOW 2: Grouped ──────────────────────────────────────────────────────
        if (resolvedMode === 'grouped') {
          const primary   = bt.primary_color      || '#1A3C8F'
          const titleFont = bt.title_font_family  || 'Arial'
          const bodyFont  = bt.body_font_family   || 'Arial'
          const artW = dArt.w || 4
          const artH = dArt.h || 4
          const n    = (mGroups || []).length || 1

          // Layout direction: columns = headers above boxes (horizontal groups);
          //                   rows    = headers left of boxes (vertical groups)
          const gLayout = dArt.group_layout
            || (artW > artH && n <= 3 ? 'columns' : 'rows')

          // Re-use Agent 5 grouped styling when it exists; fill gaps otherwise
          const agentHasGrouped = dArt.insight_mode === 'grouped' && !!dArt.group_header_style
          const f = agentHasGrouped ? ((dArt.bullet_style || {}).font_size || 10) : 10

          // ── Content-aware header dimension calculation ───────────────────────
          // Estimate the minimum w (rows) or h (columns) needed to render each
          // group header text without character-level wrapping.
          // Uses the same approximation the renderer will use: avg char width ≈
          // font_size × 0.58 / 72 inches. Target: each header fits in ≤ 2 lines.
          const hFontSize   = 10   // header font size (pts)
          const charWIn     = hFontSize * 0.58 / 72   // avg char width in inches
          const lineHIn     = hFontSize * 1.4  / 72   // line height in inches
          const headerTexts = (mGroups || []).map(g => String(g.header || ''))

          // rows layout: fix header WIDTH so longest header fits in ≤ 2 lines
          const _minRowHeaderW = (() => {
            const maxLen = Math.max(...headerTexts.map(t => t.length), 1)
            const minW   = Math.ceil(maxLen / 2) * charWIn + 0.12  // 2-line target
            return Math.min(minW, artW * 0.35)   // cap at 35% of artifact width
          })()

          // columns layout: fix header HEIGHT so longest header (at col_w) fits
          const _minColHeaderH = (() => {
            const colW        = artW / Math.max(n, 1)   // approximate col width
            const charsPerLn  = Math.max(1, (colW - 0.10) / charWIn)
            const maxLines    = Math.max(...headerTexts.map(t =>
              Math.ceil(t.length / charsPerLn)
            ), 1)
            return maxLines * lineHIn + 0.08   // 0.08" top+bottom padding
          })()

          const group_header_style = dArt.group_header_style || {
            shape:        'rounded_rect',
            fill_color:   primary,
            text_color:   '#FFFFFF',
            font_family:  titleFont,
            font_size:    hFontSize,
            font_weight:  'bold',
            corner_radius: 4,
            w: gLayout === 'rows'
              ? r2(Math.max(_minRowHeaderW, Math.min(1.5, artW * 0.30)))
              : artW,
            h: gLayout === 'columns'
              ? r2(Math.max(_minColHeaderH, hFontSize * 1.8 / 72, artH * 0.06))
              : r2(Math.max(hFontSize * 1.8 / 72, artH * 0.06))
          }

          // If Agent 5 already set group_header_style but w/h may still be too small,
          // enforce the content-based floor on the existing values too.
          if (dArt.group_header_style) {
            if (gLayout === 'rows' && (dArt.group_header_style.w || 0) < _minRowHeaderW) {
              group_header_style.w = r2(_minRowHeaderW)
            }
            if (gLayout === 'columns' && (dArt.group_header_style.h || 0) < _minColHeaderH) {
              group_header_style.h = r2(_minColHeaderH)
            }
          }

          const group_bullet_box_style = dArt.group_bullet_box_style || {
            fill_color:   null,
            border_color: '#CCCCCC',
            border_width:  0.75,
            corner_radius: group_header_style.corner_radius || 4,
            padding: {
              top:    r2(Math.max(f * 0.8 / 72, 0.05)),
              right:  r2(Math.max(f * 1.0 / 72, 0.07)),
              bottom: r2(Math.max(f * 0.8 / 72, 0.05)),
              left:   r2(Math.max(f * 1.0 / 72, 0.07))
            }
          }

          const bullet_style = dArt.bullet_style || {
            font_family:     bodyFont,
            font_size:       f,
            font_weight:     'regular',
            color:           bt.body_color || '#111111',
            line_spacing:    1.35,
            indent_inches:   0.10,
            space_before_pt: r2(Math.max(f * 0.4, 2)),
            char:            '•'
          }

          const dimForGap = gLayout === 'rows' ? artH : artW
          const group_gap_in         = dArt.group_gap_in         || r2(Math.max(dimForGap * 0.015, 0.05))
          const header_to_box_gap_in = dArt.header_to_box_gap_in || r2(Math.max(f * 0.5 / 72, 0.03))

          return {
            ...dArt,
            insight_mode:          'grouped',
            heading,
            insight_header,
            sentiment,
            groups:                mGroups,
            points:                [],
            group_layout:          gLayout,
            group_header_style,
            group_bullet_box_style,
            bullet_style:          { ...bullet_style },
            group_gap_in:          r2(group_gap_in),
            header_to_box_gap_in:  r2(header_to_box_gap_in)
          }
        }

        // ── FLOW 1: Standard (points) ────────────────────────────────────────────
        const body_style = dArt.body_style || {
          font_family:           bt.body_font_family || 'Arial',
          font_size:             10,
          font_weight:           'regular',
          color:                 bt.body_color || '#000000',
          line_spacing:          1.3,
          indent_inches:         0.15,
          list_style:            'bullet',
          space_before_pt:       6,
          vertical_distribution: 'spread'
        }
        return {
          ...dArt,
          insight_mode:   'standard',
          heading,
          insight_header,
          sentiment,
          points:  mPoints || dArt.points || [],
          groups:  undefined,
          body_style
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

      if (t === 'matrix') {
        return {
          ...dArt,
          matrix_type:   mArt.matrix_type   || dArt.matrix_type   || '2x2',
          matrix_header: mArt.matrix_header || dArt.matrix_header || '',
          x_axis:        mArt.x_axis        || dArt.x_axis        || { label: '', low_label: '', high_label: '' },
          y_axis:        mArt.y_axis        || dArt.y_axis        || { label: '', low_label: '', high_label: '' },
          quadrants:     mArt.quadrants     || dArt.quadrants     || [],
          points:        mArt.points        || dArt.points        || []
        }
      }

      if (t === 'driver_tree') {
        return {
          ...dArt,
          tree_header: mArt.tree_header || dArt.tree_header || '',
          root:        mArt.root        || dArt.root        || { label: '', value: '' },
          branches:    mArt.branches    || dArt.branches    || []
        }
      }

      if (t === 'prioritization') {
        return {
          ...dArt,
          priority_header: mArt.priority_header || dArt.priority_header || '',
          items: (mArt.items || dArt.items || []).map(it => ({
            rank: it.rank,
            title: it.title || '',
            description: it.description || '',
            qualifiers: Array.isArray(it.qualifiers)
              ? it.qualifiers.slice(0, 2).map(q => ({ label: q?.label || '', value: q?.value || '' }))
              : [{ label: '', value: '' }, { label: '', value: '' }]
          }))
        }
      }

      return dArt
    })

    return {
      ...dZone,
      layout_hint: dZone.layout_hint || mZone.layout_hint || null,
      artifact_arrangement:
        dZone.artifact_arrangement ||
        (dZone.layout_hint || {}).artifact_arrangement ||
        mZone.artifact_arrangement ||
        (mZone.layout_hint || {}).artifact_arrangement ||
        null,
      split_hint:
        (Array.isArray(dZone.split_hint) ? dZone.split_hint : null) ||
        (Array.isArray((dZone.layout_hint || {}).split_hint) ? (dZone.layout_hint || {}).split_hint : null) ||
        (Array.isArray(mZone.split_hint) ? mZone.split_hint : null) ||
        (Array.isArray((mZone.layout_hint || {}).split_hint) ? (mZone.layout_hint || {}).split_hint : null) ||
        null,
      artifacts: mergedArtifacts
    }
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
    const inner = {
      x: r2(frame.x + frame.padding.left),
      y: r2(frame.y + frame.padding.top),
      w: r2(Math.max(0.1, frame.w - frame.padding.left - frame.padding.right)),
      h: r2(Math.max(0.1, frame.h - frame.padding.top - frame.padding.bottom))
    }
    // Find paired header placeholder: same x-column as content area, positioned just above it
    const headerPh = headerPhs.find(p =>
      Math.abs((p.x_in || 0) - (ca.x_in || 0)) < 0.15 && (p.y_in || 0) < (ca.y_in || 0)
    )
    const zoneArtifacts = zone.artifacts || []
    const singleArtifact = zoneArtifacts.length === 1
    const artifacts = zoneArtifacts.map(a => {
      const base = { ...a, placeholder_idx: ca.idx }
      if (!singleArtifact) {
        // Force downstream stacking to recompute within the resolved placeholder frame.
        return {
          ...base,
          x: null, y: null, w: null, h: null
        }
      }

      const rebound = {
        ...base,
        x: inner.x,
        y: inner.y,
        w: inner.w,
        h: inner.h
      }

      // Layout-dependent internals must be recomputed against the actual placeholder frame.
      if (rebound.type === 'cards') {
        rebound.container = { x: inner.x, y: inner.y, w: inner.w, h: inner.h }
        rebound.card_frames = []
      } else if (rebound.type === 'workflow') {
        rebound.container = { x: inner.x, y: inner.y, w: inner.w, h: inner.h }
      } else if (rebound.type === 'table') {
        rebound.column_widths = []
        rebound.row_heights = []
        rebound.header_row_height = null
      }

      return rebound
    })
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
    manifestSlide.zones || [],
    branded.brand_tokens || {}
  )

  // Layout mode: fill zone frames + artifact placeholder_idx from the layout's content_areas.
  // This runs after merge so Agent 4's artifact content is already in place.
  const layoutName = manifestSlide.selected_layout_name || designed.selected_layout_name || ''
  const isLayoutMode = !!(designed.layout_mode || layoutName)
  const brandedWithLayoutTitle = isLayoutMode && layoutName
    ? applyLayoutTitleFrames(branded, layoutName, brand)
    : branded
  const finalZones = isLayoutMode && layoutName
    ? applyLayoutZoneFrames(mergedZones, layoutName, brand)
    : mergedZones

  // Post-process: fill computed layout/sizing fields (stacking, chart, table, cards, font scaling)
  // so that generate_pptx.py can act as a pure renderer reading pre-computed values.
  computeArtifactInternals(finalZones, branded.canvas || {}, branded.brand_tokens || {})
  normalizeArtifactHeaderBands(finalZones)

  // Flatten to blocks[] — ordered, self-contained render units.
  // generate_pptx.py reads these directly when present; zones path is legacy fallback.
  brandedWithLayoutTitle.blocks = flattenToBlocks(
    { ...brandedWithLayoutTitle, zones: finalZones },
    brandedWithLayoutTitle.brand_tokens || {}
  )
  const renderIssues = validateRenderCompleteness({ ...brandedWithLayoutTitle, zones: finalZones })
  if (renderIssues.length > 0) {
    console.warn('Agent 5 -- S' + (manifestSlide.slide_number || '?') + ' render issues:', renderIssues.join('; '))
  }

  // Log merge summary
  const contentCounts = { insight_text: 0, chart: 0, cards: 0, workflow: 0, table: 0 }
  finalZones.forEach(z => (z.artifacts || []).forEach(a => {
    if (contentCounts[a.type] !== undefined) contentCounts[a.type]++
  }))
  console.log('  S' + manifestSlide.slide_number + ' merged content:',
    Object.entries(contentCounts).filter(([,n]) => n > 0).map(([t,n]) => t + ':' + n).join(' ') || 'none')

  const {
    zones: _ignoredZones,
    title_block: _ignoredTitleBlock,
    subtitle_block: _ignoredSubtitleBlock,
    ...brandedWithoutZones
  } = brandedWithLayoutTitle

  return {
    ...brandedWithoutZones,
    // Always override with manifest ground truth (Claude may drift on slide_number etc.)
    slide_number:          manifestSlide.slide_number,
    slide_type:            manifestSlide.slide_type            || designed.slide_type,
    slide_archetype:       manifestSlide.slide_archetype       || designed.slide_archetype || 'summary',
    // Layout mode fields — ground truth from Agent 4 manifest
    layout_mode:           isLayoutMode,
    selected_layout_name:  manifestSlide.selected_layout_name  || designed.selected_layout_name || '',
    section_name:          manifestSlide.section_name          || '',
    section_type:          manifestSlide.section_type          || '',
    // Slide-level content metadata
    title:            (brandedWithLayoutTitle.title_block || {}).text || manifestSlide.title    || '',
    subtitle:         (brandedWithLayoutTitle.subtitle_block || {}).text || manifestSlide.subtitle || '',
    key_message:      manifestSlide.key_message      || '',
    visual_flow_hint: manifestSlide.visual_flow_hint || '',
    speaker_note:     manifestSlide.speaker_note     || '',
    layout_name:      inferLayoutName(manifestSlide, brand),
    // Condensed structural summary for Agent 5.1 review/debug; final render contract is blocks[].
    zones_summary:    finalZones.map(z => ({
      zone_id:          z.zone_id,
      zone_role:        z.zone_role,
      narrative_weight: z.narrative_weight,
      artifact_types:   (z.artifacts || []).map(a => a.type)
    })),
    _validation_issues: issues.length > 0 ? issues : undefined,
    _render_validation_issues: renderIssues.length > 0 ? renderIssues : undefined
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

  if (!manifest || !manifest.length) {
    console.error('Agent 5 -- slideManifest is empty')
    return []
  }

  const tokens = extractBrandTokens(brand)
  console.log('Agent 5 starting -- slides:', manifest.length)
  console.log('  Primary color:', (tokens.primary_colors || [])[0] || 'none')
  console.log('  Slide size:', tokens.slide_width_inches + '" x ' + tokens.slide_height_inches + '"')
  console.log('  Title font:', (tokens.title_font || {}).family || 'default')

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
    const result = await designSlideBatch(batch, brand, b + 1)

    if (!result) {
      // Entire batch failed to parse — fall back per slide via Claude
      console.warn('Agent 5 -- batch', b + 1, 'failed entirely, running per-slide fallbacks')
      for (const ms of batch) {
        const fb = await buildFallbackDesign(ms, brand)
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
        const fb = await buildFallbackDesign(mSlide, brand)
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
          const fb = await buildFallbackDesign(mSlide, brand)
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
    (s.zones_summary || []).forEach(z =>
      (z.artifact_types || []).forEach(t => { typeCounts[t] = (typeCounts[t] || 0) + 1 })
    )
  )

  console.log('Agent 5 complete')
  console.log('  Slides:', allDesigned.length,
    '| fallback:', withFallback.length,
    '| with issues:', withIssues.length)
  console.log('  Artifact types:', JSON.stringify(typeCounts))

  return allDesigned
}
