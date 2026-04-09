// ─── AGENT 4 — SLIDE CONTENT ARCHITECT ────────────────────────────────────────
// Input:  state.outline    — presentation brief / slide plan from Agent 3
//         state.contentB64 — original PDF or source document
// Output: slideManifest    — flat JSON array, one object per slide
//
// Agent 4 decides WHAT each slide is trying to say, WHAT zones (messaging arcs)
// it needs, and WHAT artifacts sit inside each zone.
// Agent 5 decides final coordinates, styling, and rendering.
//
// Key concepts:
//   Zone    = a messaging arc — one coherent argument unit
//   Artifact= the visual expression inside a zone (chart, table, workflow, etc.)
//   Each slide: max 4 zones, title/subtitle outside zones
//   Each zone:  max 2 artifacts

// ─── PROMPT TEXT ──────────────────────────────────────────────────────────────
// Phase fragments loaded via <script> tags (see index.html / test-agent4.html):
//   prompts/agent4/P0-SystemHeader.js     → _A4_HEADER
//   prompts/agent4/P1-SlideZoning.js      → _A4_PHASE1
//   prompts/agent4/P2-ZoneContent.js      → _A4_PHASE2
//   prompts/agent4/P3-ArtifactSelection.js → _A4_PHASE3  ← edit when adding artifact types
//   prompts/agent4/P4-SlideLayout.js      → _A4_PHASE4
//   prompts/agent4/P5-DeckAssembly.js     → _A4_PHASE5  (procedural binding only)
//   prompts/agent4/A1-ArtifactSchema.js   → _A4_OUTPUT_SCHEMA  ← edit when changing schemas
//   prompts/agent4/A3-QualityCheck.js     → _A4_QUALITY_GATES
//
// Repair / fallback logic → repair/agent4-repair.js
//   AGENT4_NARRATIVE_CONSTRAINTS, buildRepairPrompt, repairSlide
//
// Full system prompt — used by writeSlideBatch.
// All phases run conceptually within Claude's single response per batch:
//   P1 → zone count/structure planning
//   P2 → zone content derivation
//   P3 → artifact selection
//   P4 → layout finalization
//   P5 → output assembly rules (slide-type fields, normalization)
//   A1 → JSON output schema
//   A3 → pre-output quality gates
const AGENT4_SYSTEM = [
  _A4_HEADER,
  _A4_PHASE1,
  _A4_PHASE2,
  _A4_PHASE3,
  _A4_PHASE4,
  _A4_PHASE5,
  _A4_OUTPUT_SCHEMA,
  _A4_QUALITY_GATES,
].join('\n\n')

// Targeted system prompt — used by repairSlide and add_artifact.
// Zones already exist; no zone planning or layout finalization needed.
// Only artifact rules, output schema, and quality gates are relevant.
const AGENT4_REPAIR_SYSTEM = [
  _A4_HEADER,
  _A4_PHASE3,
  _A4_OUTPUT_SCHEMA,
  _A4_QUALITY_GATES,
].join('\n\n')

// AGENT4_NARRATIVE_CONSTRAINTS, buildRepairPrompt, repairSlide
// → moved to repair/agent4-repair.js

// ─── BATCH USER PROMPT BUILDER ────────────────────────────────────────────────
// Builds the per-batch user message for writeSlideBatch.
function buildBatchPrompt(briefSummary, layoutNames, batchPlan, batchNum, registryLine, keyMsgLines, compactBatchPlan, hasLayouts) {
  return `PRESENTATION BRIEF:
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

${_A4_BATCH_INSTRUCTIONS.replace('__SLIDE_COUNT__', batchPlan.length)}`
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
// VALIDATION → repair/agent4-validate.js
// validateArtifact, validateReasoningArtifactUsage, validateWorkflowUsage,
// validateZoneForStructureSlot, validateZoneStructureRules,
// validateWorkflowArtifactAndZone, validateStructuralPatternRules,
// validateSlideArtifactMix, hasPlaceholderContent,
// enforceReasoningArtifactUsage, enforceStructuralPatternRules,
// ZONE_STRUCTURE_LIBRARY + all zone/artifact helpers
// ═══════════════════════════════════════════════════════════════════════════════

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
    const valid = ['full', 'dominant', 'co-equal', 'compact']
    const v = String(a.artifact_coverage_hint).toLowerCase().trim()
    a.artifact_coverage_hint = valid.includes(v) ? v : undefined
  }

  if (t === 'stat_bar') {
    if (!a.stat_header) a.stat_header = a.artifact_header || a.chart_header || ''
    if (!a.artifact_header) a.artifact_header = a.stat_header || ''
    if (!a.stat_decision) a.stat_decision = a.chart_decision || ''
    if (!a.rows) a.rows = []
    if (!a.annotation_style) a.annotation_style = 'trailing'
    if (Array.isArray(a.column_headers)) {
      // New flexible schema: normalise cells
      if (a.scale_UL != null) a.scale_UL = +a.scale_UL || null
      // Normalise per-column scale_LL and scale_UL
      a.column_headers = a.column_headers.map(col => ({
        ...col,
        ...(col?.scale_LL != null ? { scale_LL: Math.max(0, +col.scale_LL) } : {}),
        ...(col?.scale_UL != null ? { scale_UL: +col.scale_UL } : {})
      }))
      // Auto-fix: ensure every "bar" column is immediately followed by a "normal" column
      const fixedCols = []
      for (let i = 0; i < a.column_headers.length; i++) {
        const col = a.column_headers[i]
        fixedCols.push(col)
        if (col?.display_type === 'bar') {
          const next = a.column_headers[i + 1]
          if (!next || next.display_type !== 'normal') {
            // Inject a companion normal column for the bar value
            fixedCols.push({ id: String(col.id) + '_val', value: '', display_type: 'normal' })
          }
        }
      }
      a.column_headers = fixedCols
      // Collect injected companion col ids (bar_id → companion_id)
      const injectedCompanions = {}
      for (let i = 0; i < fixedCols.length - 1; i++) {
        if (fixedCols[i].display_type === 'bar' && fixedCols[i + 1].id === String(fixedCols[i].id) + '_val') {
          injectedCompanions[String(fixedCols[i].id)] = fixedCols[i + 1].id
        }
      }
      a.rows = a.rows.map((row, idx) => {
        const normCells = Array.isArray(row?.cells) ? row.cells.map(cell => ({
          col_id: String(cell?.col_id ?? ''),
          value:  String(cell?.value ?? '')
        })) : []
        // For any auto-injected companion, copy bar value if companion cell is absent
        for (const [barId, companionId] of Object.entries(injectedCompanions)) {
          const hasCompanion = normCells.some(c => c.col_id === companionId)
          if (!hasCompanion) {
            const barCell = normCells.find(c => c.col_id === barId)
            normCells.push({ col_id: companionId, value: barCell ? barCell.value : '' })
          }
        }
        return {
          row_id:    row?.row_id ?? (idx + 1),
          row_focus: row?.row_focus === 'Y' ? 'Y' : 'N',
          cells:     normCells
        }
      })
    }
  }

  if (t === 'chart') {
    if (!a.categories) a.categories = []
    if (!a.series) a.series = []
    if (!a.chart_type) a.chart_type = 'bar'
    if (!a.chart_title) a.chart_title = ''
    if (!a.chart_header) a.chart_header = a.artifact_header || ''
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
    if (!a.insight_header) a.insight_header = a.artifact_header || ''
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
    if (!a.workflow_header) a.workflow_header = a.artifact_header || ''
  }

  if (t === 'table') {
    if (!a.headers) a.headers = []
    if (!a.rows) a.rows = []
    if (!a.title) a.title = ''
    if (!a.table_header) a.table_header = a.artifact_header || ''
  }

  if (t === 'matrix') {
    if (!a.matrix_type) a.matrix_type = '2x2'
    if (!a.matrix_header) a.matrix_header = a.artifact_header || ''
    if (!a.x_axis) a.x_axis = { label: '', low_label: '', high_label: '' }
    if (!a.y_axis) a.y_axis = { label: '', low_label: '', high_label: '' }
    if (!a.quadrants) a.quadrants = []
    if (!a.points) a.points = []
  }

  if (t === 'driver_tree') {
    if (!a.tree_header) a.tree_header = a.artifact_header || ''
    if (!a.root) a.root = { node_label: '', primary_message: '', secondary_message: '' }
    if (!a.branches) a.branches = []
  }

  if (t === 'prioritization') {
    if (!a.priority_header) a.priority_header = a.artifact_header || ''
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

function zoneRoleToWeight(zoneRole) {
  // Derives narrative_weight from zone_role — never read from LLM output.
  // PRIMARY and CO-PRIMARY map to 'primary'; everything else maps to 'secondary'.
  const r = String(zoneRole || '').toLowerCase()
  return /^primary|co.?primary|primary_proof|^summary|^recommendation|^process/.test(r)
    ? 'primary'
    : 'secondary'
}

function normaliseZone(z) {
  if (!z) return null
  const zoneSplit = z.zone_split || (z.layout_hint || {}).split || 'full'
  const artifactArrangement = (z.layout_hint || {}).artifact_arrangement || z.artifact_arrangement || null
  const artifactSplitHint = Array.isArray(z.artifact_split_hint)
    ? z.artifact_split_hint
    : (Array.isArray((z.layout_hint || {}).split_hint) ? (z.layout_hint || {}).split_hint : (Array.isArray(z.split_hint) ? z.split_hint : null))
  const zoneRole = z.zone_role || 'primary'
  return {
    zone_id:          z.zone_id          || 'z1',
    zone_slot:        z.zone_slot        || '',
    zone_role:        zoneRole,
    message_objective:z.message_objective|| '',
    narrative_weight: zoneRoleToWeight(zoneRole),
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

  // Structural slides (title, divider, thank_you) must never have zones — enforce unconditionally.
  if (slideType !== 'title' && slideType !== 'divider' && slideType !== 'thank_you') {
    if (slide.zones && Array.isArray(slide.zones) && slide.zones.length > 0) {
      zones = slide.zones.map(normaliseZone).filter(Boolean)
    } else {
      // Build default fallback zones (2-zone: chart left 60% + insight right 40%)
      zones = [
        { zone_id: 'z1', zone_role: 'primary',
          message_objective: 'Evidence supporting the claim',
          layout_hint: { split: 'left_60' },
          artifacts: [{ type: 'chart', chart_type: 'bar', categories: [], series: [] }] },
        { zone_id: 'z2', zone_role: 'secondary',
          message_objective: 'Interpretation and implication of the evidence',
          layout_hint: { split: 'right_40' },
          artifacts: [{ type: 'insight_text', artifact_header: 'Key Insight', points: [], sentiment: 'neutral' }] }
      ].map(normaliseZone).filter(Boolean)
    }

    // Cap at 4 zones
    zones = zones.slice(0, 4)
  }

  const normalized = {
    slide_number:                 slide.slide_number                 || plan.slide_number,
    slide_type:                   slideType,
    narrative_role:               slide.narrative_role               || plan.narrative_role  || '',
    zone_structure:               slide.zone_structure               || '',
    selected_layout_name:         slide.selected_layout_name         || '',
    title:                        slide.title                        || plan.slide_title_draft || ('Slide ' + plan.slide_number),
    subtitle:                     slide.subtitle                     || '',
    key_message:                  slide.key_message                  || '',
    visual_flow_hint:             slide.visual_flow_hint             || '',
    context_from_previous_slide:  slide.context_from_previous_slide  || '',
    zones:                        zones,
    speaker_note:                 (slideType === 'content' ? (slide.speaker_note || plan.strategic_objective || '') : '')
  }
  return applyZoneStructureMetadata(enforceStructuralPatternRules(enforceReasoningArtifactUsage(normalized)))
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
    subtitle:              s.subtitle              || '',
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

  const prompt = buildBatchPrompt(briefSummary, layoutNames, batchPlan, batchNum, registryLine, keyMsgLines, compactBatchPlan, hasLayouts)

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
  if (type === 'stat_bar') return 90
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

// zoneHasOnlyCards, zoneCardCount, isCompactCardsZone → repair/agent4-validate.js

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

  // insight_text and table are always secondary — cap their share at 30% when paired with a primary artifact
  const secondaryOnlyTypes = ['insight_text', 'table']
  if (secondaryOnlyTypes.includes(firstType) && !secondaryOnlyTypes.includes(secondType)) firstShare = Math.min(firstShare, 30)
  if (secondaryOnlyTypes.includes(secondType) && !secondaryOnlyTypes.includes(firstType)) secondShare = Math.min(secondShare, 30)
  if (secondaryOnlyTypes.includes(firstType) && !secondaryOnlyTypes.includes(secondType)) secondShare = 100 - firstShare
  if (secondaryOnlyTypes.includes(secondType) && !secondaryOnlyTypes.includes(firstType)) firstShare = 100 - secondShare

  // stat_bar minimum height enforcement: 30% + (rows-2)*11% of zone.
  // Assumes zone is full-height (conservative — better to over-allocate than clip rows).
  const statBarMinShare = (art) => {
    const n = Array.isArray(art?.rows) ? Math.min(art.rows.length, 8) : 2
    return Math.min(100, 30 + Math.max(0, n - 2) * 11)
  }
  if (firstType === 'stat_bar') firstShare = Math.max(firstShare, statBarMinShare(artifacts[0]))
  if (secondType === 'stat_bar') secondShare = Math.max(secondShare, statBarMinShare(artifacts[1]))
  if (firstType === 'stat_bar' && secondType !== 'stat_bar') secondShare = 100 - firstShare
  if (secondType === 'stat_bar' && firstType !== 'stat_bar') firstShare = 100 - secondShare

  const numericCoverage = artifacts.map((_, idx) => {
    if (artifacts.length === 2) return idx === 0 ? firstShare : secondShare
    if (idx === 0) return firstShare
    const rem = Math.max(0, 100 - firstShare)
    return rem / Math.max(artifacts.length - 1, 1)
  })
  const normalizedCoverage = numericCoverage.map(v => Math.round(v * 100) / 100)
  const coverageToken = (share) => share >= 65 ? 'dominant' : share >= 45 ? 'co-equal' : 'compact'
  const artifactsWithCoverage = artifacts.map((art, idx) => ({
    ...art,
    artifact_coverage_hint: coverageToken(normalizedCoverage[idx])
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
      const colHeaders = artifact.column_headers
      if (Array.isArray(colHeaders)) {
        // New flexible schema
        return {
          type: 'stat_bar',
          artifact_header: artifact.artifact_header || artifact.stat_header || artifact.chart_header || '',
          stat_decision: artifact.stat_decision || artifact.chart_insight || '',
          annotation_style: artifact.annotation_style || 'trailing',
          ...(artifact.scale_UL != null ? { scale_UL: artifact.scale_UL } : {}),
          column_headers: colHeaders.map(c => ({
            id: String(c?.id ?? ''),
            value: String(c?.value ?? ''),
            display_type: c?.display_type || 'text',
            ...(c?.scale_LL != null ? { scale_LL: +c.scale_LL } : {}),
            ...(c?.scale_UL != null ? { scale_UL: +c.scale_UL } : {})
          })),
          rows: Array.isArray(artifact.rows) ? artifact.rows.map((row, idx) => ({
            row_id: row?.row_id ?? (idx + 1),
            row_focus: row?.row_focus === 'Y' ? 'Y' : 'N',
            cells: Array.isArray(row?.cells) ? row.cells.map(cell => ({
              col_id: String(cell?.col_id ?? ''),
              value: String(cell?.value ?? '')
            })) : []
          })) : [],
          ...coverage
        }
      }
      // column_headers is not an array — return empty shell; repair loop will catch it
      return { type: 'stat_bar', artifact_header: artifact.artifact_header || '', column_headers: [], rows: [], annotation_style: 'trailing', ...coverage }
    }
    if (type === 'cards') {
      return {
        type: 'cards',
        artifact_header: artifact.artifact_header || artifact.artifact_header_text || '',
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
        nodes: Array.isArray(artifact.nodes) ? artifact.nodes.map(node => ({
          id: node?.id || '',
          node_label: node?.node_label || node?.label || '',
          primary_message: node?.primary_message || node?.value || '',
          secondary_message: node?.secondary_message || node?.description || '',
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
      const semToNum = { low: 25, medium: 50, high: 75 }
      const normalizedQuadrants = (Array.isArray(artifact.quadrants) ? artifact.quadrants : []).map((q, i) => ({
        id: q.id || `q${i + 1}`,
        title: q.title || q.name || '',
        primary_message: q.primary_message || '',
        tone: q.tone || 'neutral'
        // secondary_message intentionally omitted — moved to paired insight_text
      }))
      const normalizedPoints = (Array.isArray(artifact.points) ? artifact.points : []).map(pt => {
        const rawX = pt.x; const rawY = pt.y
        const xNum = typeof rawX === 'number' ? rawX : (semToNum[String(rawX || 'medium').toLowerCase()] ?? 50)
        const yNum = typeof rawY === 'number' ? rawY : (semToNum[String(rawY || 'medium').toLowerCase()] ?? 50)
        return {
          label: pt.label || '',
          short_label: pt.short_label || '',
          quadrant_id: pt.quadrant_id || '',
          x: xNum,
          y: yNum,
          emphasis: pt.emphasis || 'medium'
          // primary_message / secondary_message intentionally omitted — in paired insight_text
        }
      })
      return {
        type: 'matrix',
        matrix_type: artifact.matrix_type || '2x2',
        matrix_header: artifact.matrix_header || artifact.artifact_header || '',
        x_axis: artifact.x_axis || { label: '', low_label: '', high_label: '' },
        y_axis: artifact.y_axis || { label: '', low_label: '', high_label: '' },
        quadrants: normalizedQuadrants,
        points: normalizedPoints,
        ...coverage
      }
    }
    if (type === 'driver_tree') {
      const root = artifact.root || {}
      return {
        type: 'driver_tree',
        tree_header: artifact.tree_header || '',
        root: {
          node_label:        root.node_label        || root.label || '',
          primary_message:   root.primary_message   || root.value || '',
          secondary_message: root.secondary_message || ''
        },
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
      // ── New flat schema: columns[] + rows[].cells[{value, subtext, tone}] ──
      if (Array.isArray(artifact.columns) && artifact.columns.length) {
        return {
          type: 'comparison_table',
          artifact_header: artifact.artifact_header || '',
          columns: artifact.columns.map(c => String(c || '')),
          rows: (Array.isArray(artifact.rows) ? artifact.rows : []).map(r => ({
            is_recommended: !!r?.is_recommended,
            badge: r?.badge || undefined,
            cells: (Array.isArray(r?.cells) ? r.cells : []).map(cell => ({
              value:     cell?.value     != null ? String(cell.value)     : undefined,
              icon_type: cell?.icon_type != null ? String(cell.icon_type) : undefined,
              subtext:   cell?.subtext   != null ? String(cell.subtext)   : undefined,
              tone:      cell?.tone      || 'neutral'
            }))
          })),
          ...coverage
        }
      }
      return {
        type: 'comparison_table',
        artifact_header: artifact.artifact_header || artifact.comparison_header || artifact.table_header || '',
        columns: Array.isArray(artifact.columns) ? artifact.columns : [],
        rows: Array.isArray(artifact.rows) ? artifact.rows.map(r => ({
          id: r?.id || '',
          is_recommended: r?.is_recommended || false,
          badge: r?.badge || '',
          cells: Array.isArray(r?.cells) ? r.cells.map(cell => ({
            value:     cell?.value     != null ? String(cell.value)     : undefined,
            icon_type: cell?.icon_type != null ? String(cell.icon_type) : undefined,
            subtext:   cell?.subtext   != null ? String(cell.subtext)   : undefined,
            tone:      cell?.tone      || 'neutral'
          })) : []
        })) : [],
        recommended_row_id: artifact.recommended_row_id || undefined,
        ...coverage
      }
    }
    if (type === 'initiative_map') {
      const normTag = t => typeof t === 'string'
        ? { label: t, tone: 'neutral' }
        : { label: String(t?.label || t?.text || t?.name || ''), tone: t?.tone || 'neutral' }
      return {
        type: 'initiative_map',
        artifact_header: artifact.artifact_header || artifact.initiative_header || artifact.table_header || '',
        column_headers: Array.isArray(artifact.column_headers) ? artifact.column_headers.map(c => ({ id: c?.id || '', label: c?.label || '' })) : [],
        rows: Array.isArray(artifact.rows) ? artifact.rows.map(r => ({
          id: r?.id || '',
          initiative_name: r?.initiative_name || r?.name || '',
          initiative_subtitle: r?.initiative_subtitle || r?.subtitle || undefined,
          cells: Array.isArray(r?.cells) ? r.cells.map(cell => ({
            column_id: cell?.column_id || '',
            primary_message: cell?.primary_message || '',
            secondary_message: cell?.secondary_message || undefined,
            tags: Array.isArray(cell?.tags) ? cell.tags.map(normTag) : undefined,
            cell_tone: cell?.cell_tone || 'neutral'
          })) : []
        })) : [],
        ...coverage
      }
    }
    if (type === 'risk_register') {
      return {
        type: 'risk_register',
        risk_header: artifact.risk_header || artifact.table_header || '',
        severity_levels: Array.isArray(artifact.severity_levels) ? artifact.severity_levels.map((lvl, li) => ({
          id: lvl.id || `level_${li + 1}`,
          label: lvl.label || '',
          tone: String(lvl.tone || lvl.severity || 'medium').toLowerCase(),
          pip_levels: typeof lvl.pip_levels === 'number' ? Math.max(1, Math.round(lvl.pip_levels)) : 5,
          item_details: (Array.isArray(lvl.item_details) ? lvl.item_details : []).map(item => ({
            primary_message: item.primary_message || item.risk_title || item.title || '',
            secondary_message: item.secondary_message || item.risk_detail || item.detail || '',
            tags: (Array.isArray(item.tags) ? item.tags : []).map(t => ({
              value: String(t.value || t.label || ''),
              tone: String(t.tone || 'neutral').toLowerCase()
            })),
            pips: (Array.isArray(item.pips) ? item.pips : []).map(p => ({
              label: String(p.label || p.value || ''),
              intensity: typeof p.intensity === 'number' ? p.intensity : p.intensity
            }))
          }))
        })) : [],
        ...coverage
      }
    }
    if (type === 'profile_card_set') {
      return {
        type: 'profile_card_set',
        profile_header: artifact.profile_header || artifact.artifact_header || artifact.artifact_header_text || '',
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
          })) : []
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
// MAIN RUNNER
// ═══════════════════════════════════════════════════════════════════════════════

async function runAgent4(state) {
  const brief          = state.outline
  const contentB64     = state.contentB64
  const brand          = state.brandRulebook || {}
  // Optional: when set, only process these slide numbers and merge into existingSlides
  const targetNumbers  = Array.isArray(state.slideNumbers) && state.slideNumbers.length > 0
    ? new Set(state.slideNumbers.map(Number))
    : null
  const existingSlides = Array.isArray(state.existingSlides) ? state.existingSlides : null
  const isPartialRun   = !!(targetNumbers && existingSlides)

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
  const allContentPlan = slidePlan.filter(s => s.slide_type === 'content')
  const structuralPlan = slidePlan.filter(s => s.slide_type !== 'content')

  // In a partial run, only process the requested slide numbers
  const contentPlan  = isPartialRun
    ? allContentPlan.filter(s => targetNumbers.has(s.slide_number))
    : allContentPlan
  const contentCount = contentPlan.length

  if (isPartialRun) {
    console.log('Agent 4 starting — PARTIAL RUN for slides:', [...targetNumbers].sort((a,b)=>a-b).join(', '))
  } else {
    console.log('Agent 4 starting — content slides:', contentCount, '| structural slides from Agent 3:', structuralPlan.length)
  }
  console.log('  Brand layouts total:', totalLayouts, '| content layouts:', layoutNames.length,
    layoutNames.length >= 5 ? '→ layout mode (Agent 4 selects per slide)' : '→ zone-split mode')

  // In a partial run, start from a copy of the existing deck; structural slides are already present.
  // In a full run, build structural slides from Agent 3 plan.
  let allSlides = isPartialRun
    ? existingSlides.map(s => ({ ...s }))
    : structuralPlan.map(plan => normaliseSlide({
        slide_number: plan.slide_number,
        slide_type:   plan.slide_type,
        title:        plan.slide_title_draft || '',
        subtitle:     plan.subtitle          || '',
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
      batch.forEach(plan => {
        const normalised = normaliseSlide({}, plan)
        const idx = allSlides.findIndex(s => s.slide_number === plan.slide_number)
        if (idx >= 0) allSlides[idx] = normalised
        else allSlides.push(normalised)
      })
    } else {
      result.forEach(s => {
        // Content slide — match to plan entry by slide_number
        const plan = batch.find(p => p.slide_number === s.slide_number) || batch[0]
        const normalised = normaliseSlide(s, plan)
        // Replace existing entry (partial run) or append (full run)
        const idx = allSlides.findIndex(existing => existing.slide_number === normalised.slide_number)
        if (idx >= 0) allSlides[idx] = normalised
        else allSlides.push(normalised)

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
          const normalised = normaliseSlide({}, plan)
          const idx = allSlides.findIndex(s => s.slide_number === plan.slide_number)
          if (idx >= 0) allSlides[idx] = normalised
          else allSlides.push(normalised)
        }
      })
    }

  }

  // Sort all slides by slide_number — structural + content batches may arrive out of order
  allSlides.sort((a, b) => (a.slide_number || 0) - (b.slide_number || 0))

  // Layout-mode enforcement: when 5+ content layouts exist every content slide must
  // have selected_layout_name set.  Claude sometimes misses this — fill gaps here.
  // In a partial run, only enforce on the slides we just generated.
  const hasLayouts = layoutNames.length >= 5
  const shouldEnforce = (s) => !isPartialRun || targetNumbers.has(s.slide_number)
  if (hasLayouts) {
    allSlides = allSlides.map(s => {
      if (!shouldEnforce(s)) return s
      if (s.slide_type === 'content' && !hasCompatibleLayout(s, layoutNames)) {
        console.log('  No compatible layout for slide', s.slide_number, '→ using scratch splits')
        return assignScratchSplits(s)
      }
      if (s.slide_type === 'content' && !s.selected_layout_name) {
        const assigned = pickBestLayout(s, layoutNames)
        console.log('  Auto-assigned layout for slide', s.slide_number, '→', assigned)
        return { ...s, selected_layout_name: assigned }
      }
      return s
    })
  } else {
    allSlides = allSlides.map(s => shouldEnforce(s) ? assignScratchSplits(s) : s)
  }

  // Validate and repair — in a partial run, only check the slides we just generated
  const slidesToCheck = isPartialRun
    ? allSlides.filter(s => targetNumbers.has(s.slide_number))
    : allSlides
  const failed = slidesToCheck.filter(s => hasPlaceholderContent(s))
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
      if (!shouldEnforce(s)) return s
      if (s.slide_type === 'content' && !hasCompatibleLayout(s, layoutNames)) {
        console.log('  Post-repair: no compatible layout for slide', s.slide_number, '→ using scratch splits')
        return assignScratchSplits(s)
      }
      if (s.slide_type === 'content' && !s.selected_layout_name) {
        const assigned = pickBestLayout(s, layoutNames)
        console.log('  Post-repair layout assignment for slide', s.slide_number, '→', assigned)
        return { ...s, selected_layout_name: assigned }
      }
      return s
    })
  } else {
    allSlides = allSlides.map(s => shouldEnforce(s) ? assignScratchSplits(s) : s)
  }

  // Summary log
  const artifactTypes = {}
  let totalZones = 0
  let totalArtifacts = 0

  allSlides.forEach(s => {
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
  console.log('  Artifact types:', JSON.stringify(artifactTypes))
  console.log('  Placeholder remaining:', allSlides.filter(s => hasPlaceholderContent(s)).length)

  return allSlides.map(pruneAgent4SlideForOutput)
}
