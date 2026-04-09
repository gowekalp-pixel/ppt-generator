// AGENT 5 — OUTPUT SCHEMA
// Output JSON structure, title/subtitle sizing, zone schema.
// Assembled into AGENT5_SYSTEM by agent5.js.

const _A5_OUTPUT_SCHEMA = `*********************************************************************************
OUTPUT STRUCTURE
*********************************************************************************

Each slide must return EXACTLY:

{
  "slide_number": number,
  "slide_type": "title" | "divider" | "content",
  "layout_mode": true | false,
  "selected_layout_name": "string **” brand layout name chosen by Agent 4, or empty string",
  "canvas": {
    "width_in": number,
    "height_in": number,
    "margin": { "left": number, "right": number, "top": number, "bottom": number },
    "background": { "color": "hex" }
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
global_elements is optional **” include when appropriate.

*********************************************************************************
TITLE / SUBTITLE SIZING
*********************************************************************************

Font sizes:
- title slides: 24**“34 pt
- divider slides: 22**“30 pt
- content slides: 16**“22 pt
- subtitle is 9**“16 pt smaller than title

TITLE BLOCK HEIGHT (scratch mode only **” in template/layout mode omit x/y/w/h):
Compute h from actual line count **” do NOT use a generous fixed height:

  chars_per_line â‰ˆ (title_block.w Ã— 72) / (font_size_pt Ã— 0.52)
  n_lines = ceil(title_char_count / chars_per_line)
  title_block.h = n_lines Ã— (font_size_pt Ã— 1.35 / 72) + 0.12

  Example: 90-char title, 20pt, w=12.5"
    chars_per_line = (12.5 Ã— 72) / (20 Ã— 0.52) = 900 / 10.4 â‰ˆ 86
    n_lines = ceil(90 / 86) = 2
    h = 2 Ã— (20 Ã— 1.35 / 72) + 0.12 = 2 Ã— 0.375 + 0.12 = 0.87"

  CRITICAL: zone frame y must start at title_block.y + title_block.h + 0.20" (minimum gap)
  Never add extra buffer to title_block.h **” the renderer compacts it automatically.
  In template mode (uses_template=true) where title y/h are omitted, assume the title
  area ends at ~0.90" from the top and start all zones at â‰¥ 1.10" (0.90 + 0.20 gap).

*********************************************************************************
ZONES
*********************************************************************************

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
- each zone has 1**“2 artifacts
- title and subtitle sit OUTSIDE zones
`