// ─── AGENT 5 — OUTPUT VALIDATION ─────────────────────────────────────────────
// Validates LLM output from Agent 5 at two points:
//   1. validateDesignedSlide  — raw Claude output (zones / artifacts schema)
//   2. validateRenderCompleteness — flattened blocks after normaliseDesignedSlide
//
// When adding a new artifact type:
//   1. Add the type string to supportedArtifactTypes in validateDesignedSlide.
//   2. Add required-field checks for the new type in the forEach loop below.
//   3. No changes needed in validateRenderCompleteness (it validates geometry only).
//
// Depends on agent5.js (loaded first via <script> tag):
//   normalizeArtifactType, normalizeChartSubtype — used in validateDesignedSlide
//   getZoneInnerBounds — used in validateRenderCompleteness
// Note: clamp, rectWithin, getZoneInnerBounds stay in agent5.js — they are
//       also used by normaliseDesignedSlide (normalization, not just validation).

// ─── GEOMETRY HELPERS (validation-only) ──────────────────────────────────────

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

// ─── SCHEMA VALIDATOR — raw LLM output ───────────────────────────────────────
// Add a new artifact type? Add to supportedArtifactTypes + required-field checks below.

function validateDesignedSlide(slide) {
  const issues = []
  const supportedArtifactTypes = new Set([
    'chart',
    'stat_bar',
    'insight_text',
    'table',
    'comparison_table',
    'initiative_map',
    'profile_card_set',
    'risk_register',
    'matrix',
    'driver_tree',
    'prioritization',
    'cards',
    'workflow'
  ])

  if (!slide.canvas)            issues.push('missing canvas')
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
      const normalizedType = normalizeArtifactType(a.type, a.chart_type)
      const normalizedChartType = normalizeChartSubtype(a.type, a.chart_type)
      if (!a.type)                                     issues.push(p + ': missing type')
      if (a.type && !supportedArtifactTypes.has(normalizedType)) issues.push(p + ': unsupported artifact type ' + a.type)
      if (normalizedType === 'chart'    && !a.chart_style)     issues.push(p + ': chart missing chart_style')
      if (normalizedType === 'chart'    && !a.series_style)    issues.push(p + ': chart missing series_style')
      if (normalizedType === 'chart'    && a.chart_style && a.chart_style.legend_position == null) issues.push(p + ': chart missing legend_position')
      if (normalizedType === 'chart'    && a.chart_style && a.chart_style.data_label_size == null) issues.push(p + ': chart missing data_label_size')
      if (normalizedType === 'chart'    && a.chart_style && a.chart_style.category_label_rotation == null) issues.push(p + ': chart missing category_label_rotation')
      if (normalizedType === 'stat_bar' && !Array.isArray(a.rows)) issues.push(p + ': stat_bar missing rows[]')
      if (normalizedType === 'stat_bar' && !Array.isArray(a.column_headers)) issues.push(p + ': stat_bar column_headers must be an array')
      if (normalizedType === 'stat_bar' && !a.artifact_header && !a.stat_header) issues.push(p + ': stat_bar missing artifact_header')
      if (normalizedType === 'stat_bar' && !a.annotation_style) issues.push(p + ': stat_bar missing annotation_style')
      if (normalizedType === 'chart'    && normalizedChartType === 'pie' && Array.isArray(a.series_style) && Array.isArray(a.categories) && a.series_style.length !== a.categories.length) issues.push(p + ': pie chart series_style.length (' + a.series_style.length + ') must equal categories.length (' + a.categories.length + ')')
      if (normalizedType === 'chart'    && normalizedChartType === 'group_pie' && Array.isArray(a.series_style) && Array.isArray(a.categories) && a.series_style.length !== a.categories.length) issues.push(p + ': group_pie series_style.length (' + a.series_style.length + ') must equal categories.length (' + a.categories.length + ') — one style entry per slice')
      if (normalizedType === 'chart'    && normalizedChartType === 'group_pie' && Array.isArray(a.series) && (a.series.length < 2 || a.series.length > 8)) issues.push(p + ': group_pie series (entities) must be 2–8, got ' + (a.series || []).length)
      if (normalizedType === 'workflow' && !a.nodes?.length)   issues.push(p + ': workflow missing nodes')
      if (normalizedType === 'workflow' && !a.workflow_style)  issues.push(p + ': workflow missing workflow_style')
      if (normalizedType === 'workflow' && a.workflow_style && a.workflow_style.node_inner_padding == null) issues.push(p + ': workflow missing node_inner_padding')
      if (normalizedType === 'workflow' && a.workflow_style && a.workflow_style.external_label_gap == null) issues.push(p + ': workflow missing external_label_gap')
      if (normalizedType === 'workflow' && (a.connections || []).some(c => !Array.isArray(c.path) || c.path.length < 2)) issues.push(p + ': workflow connection missing path')
      if (normalizedType === 'table'    && !a.table_style)     issues.push(p + ': table missing table_style')
      if (normalizedType === 'table'    && !a.column_widths)   issues.push(p + ': table missing column_widths')
      if (normalizedType === 'table'    && !a.row_heights)     issues.push(p + ': table missing row_heights')
      if (normalizedType === 'table'    && !a.column_types)    issues.push(p + ': table missing column_types')
      if (normalizedType === 'table'    && !a.column_alignments) issues.push(p + ': table missing column_alignments')
      if (normalizedType === 'table'    && a.table_style && a.table_style.cell_padding == null) issues.push(p + ': table missing cell_padding')
      if (normalizedType === 'comparison_table' && !Array.isArray(a.columns)) issues.push(p + ': comparison_table missing columns[]')
      if (normalizedType === 'comparison_table' && !Array.isArray(a.rows)) issues.push(p + ': comparison_table missing rows[]')
      if (normalizedType === 'comparison_table' && !a.comparison_style) issues.push(p + ': comparison_table missing comparison_style')
      if (normalizedType === 'initiative_map' && !Array.isArray(a.column_headers)) issues.push(p + ': initiative_map missing column_headers[]')
      if (normalizedType === 'initiative_map' && !Array.isArray(a.rows)) issues.push(p + ': initiative_map missing rows[]')
      if (normalizedType === 'initiative_map' && !a.initiative_style) issues.push(p + ': initiative_map missing initiative_style')
      if (normalizedType === 'profile_card_set' && !Array.isArray(a.profiles)) issues.push(p + ': profile_card_set missing profiles')
      if (normalizedType === 'profile_card_set' && !a.profile_style) issues.push(p + ': profile_card_set missing profile_style')
      if (normalizedType === 'risk_register' && !Array.isArray(a.severity_levels)) issues.push(p + ': risk_register missing severity_levels[]')
      if (normalizedType === 'risk_register' && !a.risk_style) issues.push(p + ': risk_register missing risk_style')
      if (normalizedType === 'cards'    && !a.card_frames?.length) issues.push(p + ': cards missing card_frames')
      if (normalizedType === 'cards'    && !a.card_style)      issues.push(p + ': cards missing card_style')
      if (normalizedType === 'cards'    && !a.cards_layout)    issues.push(p + ': cards missing cards_layout')
      if (normalizedType === 'cards'    && !a.container)       issues.push(p + ': cards missing container')
      if (normalizedType === 'matrix'   && !a.matrix_style)    issues.push(p + ': matrix missing matrix_style')
      if (normalizedType === 'matrix'   && !a.x_axis?.label)   issues.push(p + ': matrix missing x_axis.label')
      if (normalizedType === 'matrix'   && !a.y_axis?.label)   issues.push(p + ': matrix missing y_axis.label')
      if (normalizedType === 'matrix'   && (a.quadrants || []).length !== 4) issues.push(p + ': matrix must define 4 quadrants')
      if (normalizedType === 'matrix'   && !(a.points || []).length) issues.push(p + ': matrix missing points')
      if (normalizedType === 'matrix') {
        const validQids = new Set(['q1','q2','q3','q4'])
        ;(a.points || []).forEach((pt, pi) => {
          if (!pt.quadrant_id || !validQids.has(String(pt.quadrant_id).toLowerCase())) issues.push(p + `: matrix point[${pi}] missing valid quadrant_id`)
          if (typeof pt.x !== 'number') issues.push(p + `: matrix point[${pi}] x must be numeric 0–100`)
          if (typeof pt.y !== 'number') issues.push(p + `: matrix point[${pi}] y must be numeric 0–100`)
        })
      }
      if (normalizedType === 'driver_tree' && !a.tree_style)   issues.push(p + ': driver_tree missing tree_style')
      if (normalizedType === 'driver_tree' && !a.root?.label && !a.root?.node_label)  issues.push(p + ': driver_tree missing root.label/root.node_label')
      if (normalizedType === 'driver_tree' && !(a.branches || []).length) issues.push(p + ': driver_tree missing branches')
      if (normalizedType === 'prioritization' && !a.priority_style) issues.push(p + ': prioritization missing priority_style')
      if (normalizedType === 'prioritization' && !(a.items || []).length) issues.push(p + ': prioritization missing items')
      if (normalizedType === 'prioritization' && (a.items || []).some(it => it.rank == null || !String(it.title || '').trim())) issues.push(p + ': prioritization items require rank and title')
      if (normalizedType === 'insight_text' && !a.heading_style) issues.push(p + ': insight_text missing heading_style')
      if (normalizedType === 'insight_text' && a.insight_mode === 'grouped' && !a.group_header_style) issues.push(p + ': grouped insight_text missing group_header_style')
      if (normalizedType === 'insight_text' && a.insight_mode === 'grouped' && !a.group_bullet_box_style) issues.push(p + ': grouped insight_text missing group_bullet_box_style')
      if (normalizedType === 'insight_text' && a.insight_mode === 'grouped' && !a.bullet_style) issues.push(p + ': grouped insight_text missing bullet_style')
      if (normalizedType === 'insight_text' && a.insight_mode === 'grouped' && a.group_gap_in == null) issues.push(p + ': grouped insight_text missing group_gap_in')
      if (normalizedType === 'insight_text' && a.insight_mode === 'grouped' && a.header_to_box_gap_in == null) issues.push(p + ': grouped insight_text missing header_to_box_gap_in')
      if (normalizedType === 'insight_text' && (!a.insight_mode || a.insight_mode === 'standard') && !a.body_style) issues.push(p + ': insight_text missing body_style')
      if (normalizedType === 'insight_text' && (!a.insight_mode || a.insight_mode === 'standard') && a.body_style && a.body_style.list_style == null) issues.push(p + ': insight_text missing list_style')
      if (normalizedType === 'insight_text' && (!a.insight_mode || a.insight_mode === 'standard') && a.body_style && a.body_style.line_spacing == null) issues.push(p + ': insight_text missing line_spacing')
      if (normalizedType === 'insight_text' && (!a.insight_mode || a.insight_mode === 'standard') && a.body_style && a.body_style.indent_inches == null) issues.push(p + ': insight_text missing indent_inches')
      if (normalizedType === 'insight_text' && (!a.insight_mode || a.insight_mode === 'standard') && a.body_style && a.body_style.space_before_pt == null) issues.push(p + ': insight_text missing space_before_pt')
    })
  })

  return issues
}

// ─── RENDER COMPLETENESS VALIDATOR — post-flatten blocks ─────────────────────
// Runs after normaliseDesignedSlide. Checks geometry, block completeness,
// overlap, and canvas fill ratio. No artifact-type-specific logic here.

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
    if (['title', 'subtitle', 'footer', 'page_number', 'image', 'chart', 'table', 'bullet_list', 'rect', 'text_box', 'rule', 'circle', 'line', 'icon_badge'].includes(b.block_type)) {
      if (!b.artifact_type) issues.push(p + ': missing artifact_type')
      if (!b.artifact_subtype) issues.push(p + ': missing artifact_subtype')
      if (!b.fallback_policy) issues.push(p + ': missing fallback_policy')
      if (!b.block_role) issues.push(p + ': missing block_role')
    }
    if (b.block_role && /^artifact_/.test(b.block_role) && !b.artifact_id) {
      issues.push(p + ': missing artifact_id')
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
      if (b.table_fit_failed) issues.push(p + ': table block failed fit validation')
    }
    if (b.block_type === 'line') {
      if (b.x1 == null || b.y1 == null || b.x2 == null || b.y2 == null) issues.push(p + ': line block missing endpoints')
    }
  })

  const overlapTypes = new Set(['rect', 'text_box', 'bullet_list', 'circle', 'chart', 'table'])
  const artifactBlocks = (slide.blocks || []).filter(b =>
    b &&
    b.artifact_id &&
    /^artifact_/.test(String(b.block_role || '')) &&
    overlapTypes.has(String(b.block_type || ''))
  )
  for (let i = 0; i < artifactBlocks.length; i++) {
    for (let j = i + 1; j < artifactBlocks.length; j++) {
      const a = artifactBlocks[i]
      const b = artifactBlocks[j]
      if (a.artifact_id === b.artifact_id) continue
      if (rectsOverlap(a, b, 0.01)) {
        issues.push('artifact blocks overlap: ' + a.artifact_id + ' & ' + b.artifact_id)
      }
    }
  }

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
