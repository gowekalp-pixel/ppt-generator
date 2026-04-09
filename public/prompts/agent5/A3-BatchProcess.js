// AGENT 5 — BATCH USER PROMPT: STATIC INSTRUCTIONS BLOCK
// Used by designSlideBatch() in agent5.js.
// __SLIDE_COUNT__ is replaced at runtime with batchManifest.length.
// Part of the user message (not the system prompt) — assembled by agent5.js.

const _A5_BATCH_INSTRUCTIONS = `INSTRUCTIONS:
- Process ONLY these __SLIDE_COUNT__ slides
- Apply brand design tokens exactly
- Compute exact coordinates for every element (2 decimal places)
- FULLY specify all artifacts including all style sub-objects
- chart: must have chart_style and series_style[]
- workflow: must have workflow_style, nodes[] with x/y/w/h, connections[] with path[]
- table: must have table_style, column_widths[], column_types[], column_alignments[], header_row_height, row_heights[] (cell positions/frames are computed automatically)
- comparison_table: must have comparison_style, criteria[], options[], recommended_option
- initiative_map: must have initiative_style, dimension_labels[], initiatives[]
- profile_card_set: must have profile_style, profiles[], layout_direction
- risk_register: must have risk_style and severity_levels[]; tags[].tone controls chip color; each severity_level has pip_levels (numeric total scale, e.g. 5); pips[].intensity is numeric (filled blocks out of pip_levels)
- cards: must have card_style, card_frames[] with x/y/w/h per card
- matrix: must have matrix_style plus semantic fields from Agent 4 (x_axis, y_axis, quadrants[id/title/primary_message/tone], points[label/short_label/quadrant_id/x/y/emphasis]); quadrant has NO secondary_message; points have NO primary_message or secondary_message; x/y are numeric 0–100
- driver_tree: must have tree_style plus semantic fields from Agent 4 (root, branches)
- prioritization: must have priority_style plus semantic fields from Agent 4 (items[], qualifiers[])
- insight_text (standard mode): must have insight_mode:"standard", style, heading_style, body_style
- insight_text (grouped mode):  must have insight_mode:"grouped", heading_style, group_layout, group_header_style, group_bullet_box_style, bullet_style, group_gap_in, header_to_box_gap_in
- charts: include final legend_position, data_label_size, category_label_rotation, and series styling; stat_bar must have rows[], column_headers[] (array with display_type per column), and annotation_style{}
- workflows: include final node geometry, connection paths, node_inner_padding, and external_label_gap
- tables: include column_widths, column_types, column_alignments, header_row_height, row_heights, and cell_padding (do NOT compute column_x_positions, row_y_positions, header_cell_frames, body_cell_frames — these are computed automatically)
- comparison_table / initiative_map / profile_card_set / risk_register are flattened locally into rect/text blocks, so emphasize rounded rows, semantic fills, and explicit labels rather than native table behavior
- matrix: include final matrix_style; preserve x_axis, y_axis, quadrants (id/title/primary_message/tone only — no secondary_message), and points (label/short_label/quadrant_id/x/y numeric 0–100/emphasis — no primary_message/secondary_message) for block flattening
- driver_tree: include final tree_style and preserve root/branches for block flattening
- prioritization: include final priority_style and preserve ranked items/qualifiers for block flattening
- non-chart/table artifacts are flattened into primitive blocks locally, so their geometry and style must be complete and render-ready
- do NOT return blocks[]; return only the designed slide spec and artifact internals
- Return a valid JSON array of exactly __SLIDE_COUNT__ slide objects`
