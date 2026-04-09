// AGENT 2 — ENRICHMENT + VISION PROMPTS
// Two prompts for the two code paths in agent2.js:
//
//   _A2_ENRICH  — Step B: single call for PPTX extraction path.
//                 Returns global brand rules AND layout usage_guidance together.
//   _A2_VISION  — Step C: vision fallback for PDF / image brand files.

// ─── STEP B: PPTX ENRICHMENT (single call) ────────────────────────────────────
// Used by enrichAgent2 in agent2.js.
// Returns deck-level brand rules + usage_guidance for every layout in one response.
const _A2_ENRICH = `You are a senior brand designer reviewing extracted PowerPoint brand metadata.
Use the slide masters, layout summary, colors, and fonts to infer comprehensive deck-level design rules
AND add usage guidance for every layout — all in one response.

Return ONLY a valid JSON object with exactly these fields:
{
  "visual_style": "string — e.g. 'clean corporate', 'bold financial', 'minimal consulting'",
  "spacing_notes": "string — margin and padding conventions derived from placeholder positions",
  "typography_hierarchy": {
    "title_size_pt": number,
    "subtitle_size_pt": number,
    "body_size_pt": number,
    "caption_size_pt": number,
    "title_color": "#hex — from primary/accent colors",
    "body_color": "#hex — from text colors"
  },
  "chart_color_sequence": ["#hex1", "#hex2", "#hex3", "#hex4", "#hex5", "#hex6"],
  "bullet_style": {
    "char": "• or – or ▪",
    "indent_inches": number,
    "space_before_pt": number,
    "space_after_pt": number
  },
  "insight_box_style": {
    "fill_color": "#hex or null",
    "border_color": "null — insight boxes use a left accent bar, not a full border",
    "corner_radius": number
  },
  "divider_style": {
    "line_color": "#hex",
    "line_width_pt": number
  },
  "layout_usage": [
    { "name": "exact layout name", "usage_guidance": "one sentence on when to use this layout" }
  ]
}

Rules:
- chart_color_sequence must have 6 visually DISTINCT colors drawn from accent1–accent6; if fewer accents exist pad with brand-appropriate alternatives
- typography_hierarchy sizes must come from the actual placeholder font_size_pt values in the master data, not guesses
- bullet_style.space_before_pt should be 4–6 for comfortable reading
- insight_box_style.border_color must always be null (left accent bar is rendered separately)
- layout_usage must include every layout from the input, with name matching exactly
No markdown. No explanation.`

// ─── STEP C: VISION FALLBACK ──────────────────────────────────────────────────
// Used by runAgent2 when the brand file is a PDF or image (not PPTX).
const _A2_VISION = `You are an expert brand designer analyzing a brand guideline document.
Extract ALL design rules and return as a single valid JSON object with these exact fields:
{
  "color_scheme_name": "scheme name",
  "primary_colors": ["#hex"],
  "secondary_colors": ["#hex"],
  "background_colors": ["#hex"],
  "text_colors": ["#hex"],
  "accent_colors": ["#hex"],
  "chart_colors": ["#hex"],
  "all_colors": { "accent1": "#hex", "accent2": "#hex" },
  "title_font": { "family": "font name", "size": "28pt", "weight": "bold", "color": "#hex" },
  "body_font": { "family": "font name", "size": "14pt", "weight": "regular", "color": "#hex" },
  "caption_font": { "family": "font name", "size": "9pt", "weight": "regular", "color": "#hex" },
  "slide_width_inches": 10,
  "slide_height_inches": 7.5,
  "visual_style": "corporate",
  "spacing_notes": "0.5 inch margins",
  "slide_layouts": [
    { "name": "Title slide", "structure": "Full-page title", "usage_guidance": "Use for opening" }
  ]
}
Return ONLY valid JSON. No explanation. No markdown fences.`
