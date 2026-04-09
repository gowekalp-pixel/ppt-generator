// AGENT 5 — SLIDE RULES
// Brand authority, title/divider rules, layout mode vs scratch mode.
// Assembled into AGENT5_SYSTEM by agent5.js.

const _A5_SLIDE_RULES = `*********************************************************************************
BRAND GUIDELINE AUTHORITY RULE
*********************************************************************************

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

*********************************************************************************
TITLE & DIVIDER SLIDES (template mode)
*********************************************************************************

When slide_type is "title", "divider", or "thank_you" AND uses_template is true:
The master provides ALL visual elements **” background, logo, decorations, footer.
Agent 6 places text directly into the master's title/subtitle placeholders.

Output ONLY:
- title_block: { text: "Thank You", font_family, font_size, font_weight, color } **” NO x/y/w/h; for thank_you use exactly "Thank You" as text
- subtitle_block: null (always null for thank_you)
- zones: []
- global_elements: {}
- canvas.background: null
- layout_mode: true

*********************************************************************************
ALL CONTENT SLIDES **” FIXED ELEMENTS (template mode)
*********************************************************************************

When slide_type is "content" AND uses_template is true, regardless of layout selection:
The master template positions slide title, subtitle, footer, and page number.

ALWAYS apply these rules for ALL content slides in template mode:
- title_block: text, font_family, font_size, font_weight, color ONLY **” omit x/y/w/h
- subtitle_block: same **” omit x/y/w/h
- global_elements: {} **” master handles footer and page number automatically
- canvas.background: null

*********************************************************************************
CONTENT SLIDES **” LAYOUT MODE (selected_layout_name is non-empty)
*********************************************************************************

When uses_template is true AND selected_layout_name is non-empty:
The pipeline automatically maps each zone to its content area slot in the named layout.
Your job is CONTENT QUALITY and VISUAL STYLE **” not positioning.

Rules:
1. Set layout_mode: true
2. For each zone: set frame to null **” the pipeline fills frame from the layout's content_areas
3. Do NOT set placeholder_idx on artifacts **” the pipeline assigns the real PPTX placeholder idx
4. For each artifact (except cards): add header_block
   - layout_map[selected_layout_name].ph_count > 2 (multi-slot layout):
     â†’ set header_block.placeholder_ref: true
     â†’ x/y/w/h may be null (pipeline positions from placeholder)
   - ph_count â‰¤ 2: compute header_block at top of zone area; h = 0.30
5. Focus on brand-compliant styling: chart colors from chart_color_sequence, correct fonts,
   card_frames, insight_text body_style, table column_widths, workflow_style

*********************************************************************************
CONTENT SLIDES **” SCRATCH MODE (selected_layout_name is empty)
*********************************************************************************

When uses_template is false OR selected_layout_name is empty:
Compute all coordinates from layout_hint splits.
- Set layout_mode: false
- For thank_you slides in scratch mode: zones: [], subtitle_block: null, title_block text = "Thank You" with full x/y/w/h centered on slide
- zones: compute full frame coordinates from layout_hint splits
- Artifacts: compute x/y/w/h within zone bounds
- Artifact headers: compute header_block at top of zone inner bounds; shrink artifact area
- global_elements: include footer and page_number when uses_template is false
- canvas.background: set when uses_template is false
`