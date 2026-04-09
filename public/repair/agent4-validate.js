// ─── AGENT 4 — OUTPUT VALIDATION ─────────────────────────────────────────────
// Validates LLM output from Agent 4 before it reaches the repair or finalise step.
//
// When adding a new artifact type:
//   1. Add a case block in validateArtifact() for the new type's schema rules.
//   2. If the new type has solo/pairing restrictions, update:
//      isSoloOnlyArtifact(), isReasoningArtifact(), validateStructuralPatternRules()
//   3. hasPlaceholderContent() is the gate — it calls everything above automatically.
//
// Flow (called from runAgent4 in agent4.js):
//   writeSlideBatch → [each slide] → hasPlaceholderContent → repairSlide (if needed)
//
// Depends on: nothing outside this file (fully self-contained).

// ─── PER-ARTIFACT TYPE HELPERS ────────────────────────────────────────────────

function artifactType(artifact) {
  return String((artifact || {}).type || '').toLowerCase()
}

function artifactCardCount(artifact) {
  return artifactType(artifact) === 'cards' ? ((artifact.cards || []).length || 0) : 0
}

function artifactWorkflowNodeCount(artifact) {
  return artifactType(artifact) === 'workflow' ? ((artifact.nodes || []).length || 0) : 0
}

function artifactWorkflowLevelCount(artifact) {
  if (artifactType(artifact) !== 'workflow') return 0
  const levels = new Set((artifact.nodes || []).map(n => Number.isFinite(+n?.level) ? +n.level : null).filter(v => v != null))
  return levels.size
}

function artifactDriverTreeNodeCount(artifact) {
  if (artifactType(artifact) !== 'driver_tree') return 0
  const branches = artifact.branches || []
  return 1 + branches.length + branches.reduce((sum, b) => sum + (((b && b.children) || []).length || 0), 0)
}

function artifactTableShape(artifact) {
  return artifactType(artifact) === 'table'
    ? { cols: ((artifact.headers || []).length || 0), rows: ((artifact.rows || []).length || 0) }
    : { cols: 0, rows: 0 }
}

function artifactPrioritizationCount(artifact) {
  return artifactType(artifact) === 'prioritization' ? ((artifact.items || []).length || 0) : 0
}

function artifactInsightGroupCount(artifact) {
  return artifactType(artifact) === 'insight_text' ? ((artifact.groups || []).length || 0) : 0
}

function isReasoningArtifact(artifact) {
  return ['matrix', 'driver_tree', 'prioritization'].includes(artifactType(artifact))
}

function isSoloOnlyArtifact(artifact) {
  const t = artifactType(artifact)
  if (['prioritization', 'risk_register', 'profile_card_set'].includes(t)) return true
  if (t === 'workflow') {
    const wfType = String(artifact?.workflow_type || '').toLowerCase()
    return ['process_flow', 'timeline'].includes(wfType)
  }
  return false
}

function isSparseCardsArtifact(artifact) {
  return artifactType(artifact) === 'cards' && artifactCardCount(artifact) > 0 && artifactCardCount(artifact) <= 2
}

function isDenseSoloArtifact(artifact) {
  const t = artifactType(artifact)
  if (t === 'cards') return artifactCardCount(artifact) >= 8
  if (t === 'matrix') return true
  if (t === 'driver_tree') return artifactDriverTreeNodeCount(artifact) >= 6
  if (t === 'prioritization') return artifactPrioritizationCount(artifact) >= 5
  if (t === 'insight_text') return artifactInsightGroupCount(artifact) >= 3
  if (t === 'table') {
    const shape = artifactTableShape(artifact)
    return shape.cols >= 5 && shape.rows >= 10
  }
  if (t === 'workflow') return artifactWorkflowNodeCount(artifact) >= 6
  return false
}

function isSubstantialArtifact(artifact) {
  const t = artifactType(artifact)
  if (t === 'insight_text') {
    return ((artifact.points || []).length || 0) >= 3 || artifactInsightGroupCount(artifact) >= 2
  }
  if (t === 'workflow') return artifactWorkflowNodeCount(artifact) >= 4
  if (t === 'table') {
    const shape = artifactTableShape(artifact)
    return shape.cols >= 4 || shape.rows >= 6
  }
  if (t === 'chart') return ((artifact.categories || []).length || 0) >= 4
  if (t === 'cards') return artifactCardCount(artifact) >= 4
  if (t === 'prioritization') return artifactPrioritizationCount(artifact) >= 4
  if (t === 'driver_tree') return artifactDriverTreeNodeCount(artifact) >= 5
  if (t === 'matrix') return true
  return false
}

// ─── ZONE CARD HELPERS ────────────────────────────────────────────────────────

function zoneHasOnlyCards(zone) {
  const arts = zone?.artifacts || []
  return arts.length > 0 && arts.every(a => String((a || {}).type || '').toLowerCase() === 'cards')
}

function zoneCardCount(zone) {
  const cardArtifacts = (zone?.artifacts || []).filter(a => String((a || {}).type || '').toLowerCase() === 'cards')
  return cardArtifacts.reduce((sum, art) => sum + ((art.cards || []).length || 0), 0)
}

function isCompactCardsZone(zone) {
  return zoneHasOnlyCards(zone) && zoneCardCount(zone) > 0 && zoneCardCount(zone) <= 2
}

// ─── ZONE STRUCTURE LIBRARY ───────────────────────────────────────────────────

const ZONE_STRUCTURE_LIBRARY = {
  ZS01_single_full: {
    zoneCount: 1,
    slots: ['primary_full'],
    scratchSplits: ['full']
  },
  ZS02_stacked_equal: {
    zoneCount: 2,
    slots: ['top_primary', 'bottom_secondary'],
    scratchSplits: ['top_50', 'bottom_50']
  },
  ZS03_side_by_side_equal: {
    zoneCount: 2,
    slots: ['left_primary', 'right_secondary'],
    scratchSplits: ['left_50', 'right_50']
  },
  ZS04_left_dominant_right_stack: {
    zoneCount: 3,
    slots: ['left_dominant', 'right_top_support', 'right_bottom_support'],
    scratchSplits: ['left_60', 'top_right_50', 'bottom_full']
  },
  ZS05_right_dominant_left_stack: {
    zoneCount: 3,
    slots: ['left_top_support', 'left_bottom_support', 'right_dominant'],
    scratchSplits: ['top_left_50', 'bottom_full', 'right_60']
  },
  ZS06_top_full_bottom_two: {
    zoneCount: 3,
    slots: ['top_anchor', 'bottom_left_support', 'bottom_right_support'],
    scratchSplits: ['top_35', 'left_50', 'right_50']
  },
  ZS07_top_two_bottom_dominant: {
    zoneCount: 3,
    slots: ['top_left_support', 'top_right_support', 'bottom_dominant'],
    scratchSplits: ['top_left_50', 'top_right_50', 'bottom_60']
  },
  ZS08_quad_grid: {
    zoneCount: 4,
    slots: ['top_left', 'top_right', 'bottom_left', 'bottom_right'],
    scratchSplits: ['tl', 'tr', 'bl', 'br']
  },
  ZS09_left_dominant_right_triptych: {
    zoneCount: 4,
    slots: ['left_dominant', 'right_top_support', 'right_mid_support', 'right_bottom_support'],
    scratchSplits: ['left_60', 'top_right_50', 'tr', 'br']
  },
  ZS10_top_full_bottom_three: {
    zoneCount: 4,
    slots: ['top_anchor', 'bottom_left_support', 'bottom_mid_support', 'bottom_right_support'],
    scratchSplits: ['top_35', 'bl', 'bottom_full', 'br']
  },
  ZS11_three_rows_equal: {
    zoneCount: 3,
    slots: ['row_1', 'row_2', 'row_3'],
    scratchSplits: ['top_33', 'top_33', 'bottom_34']
  },
  ZW01_three_columns_equal: {
    zoneCount: 3,
    canvasFamily: 'wide',
    slots: ['col_1', 'col_2', 'col_3'],
    scratchSplits: ['left_33', 'left_33', 'right_34']
  },
  ZW02_three_columns_right_stack: {
    zoneCount: 4,
    canvasFamily: 'wide',
    slots: ['left_col', 'mid_col', 'right_top', 'right_bottom'],
    scratchSplits: ['left_33', 'left_33', 'top_right_50', 'br']
  },
  ZW03_three_columns_left_stack: {
    zoneCount: 4,
    canvasFamily: 'wide',
    slots: ['left_top', 'left_bottom', 'mid_col', 'right_col'],
    scratchSplits: ['top_left_50', 'bl', 'left_33', 'right_34']
  },
  ZW04_four_columns_equal: {
    zoneCount: 4,
    canvasFamily: 'wide',
    slots: ['col_1', 'col_2', 'col_3', 'col_4'],
    scratchSplits: ['tl', 'tr', 'bl', 'br']
  }
}

function zoneStructureDef(id) {
  return ZONE_STRUCTURE_LIBRARY[id] || null
}

// ─── ZONE PROFILE ─────────────────────────────────────────────────────────────

function zoneStructureArtifactProfile(zone) {
  const arts = zone?.artifacts || []
  return {
    count: arts.length,
    types: arts.map(artifactType),
    hasInsight: arts.some(a => artifactType(a) === 'insight_text'),
    hasReasoning: arts.some(isReasoningArtifact),
    hasWorkflow: arts.some(a => artifactType(a) === 'workflow'),
    hasChart: arts.some(a => artifactType(a) === 'chart'),
    hasTable: arts.some(a => artifactType(a) === 'table'),
    hasCards: arts.some(a => artifactType(a) === 'cards'),
    denseSolo: arts.length === 1 && isDenseSoloArtifact(arts[0]),
    compactCards: arts.some(isSparseCardsArtifact),
    groupedInsightOnly: arts.length === 1 && artifactType(arts[0]) === 'insight_text' && artifactInsightGroupCount(arts[0]) >= 3
  }
}

function isDominantSlot(slotName) {
  return /dominant|primary|anchor|full/.test(String(slotName || '').toLowerCase())
}

function isSupportSlot(slotName) {
  return /support|secondary|top_|bottom_|left_|right_|row_|col_/.test(String(slotName || '').toLowerCase()) && !isDominantSlot(slotName)
}

// ─── ZONE STRUCTURE VALIDATORS ────────────────────────────────────────────────

function validateZoneForStructureSlot(zone, slotName, structureId) {
  const arts = zone?.artifacts || []
  const profile = zoneStructureArtifactProfile(zone)
  if (!arts.length) return false

  if (structureId === 'ZS01_single_full') {
    if (profile.count === 1) {
      return profile.denseSolo || profile.groupedInsightOnly
    }
    if (profile.count === 2) {
      const key = zoneArtifactPairKey(zone)
      return new Set([
        'chart+insight_text',
        'driver_tree+insight_text',
        'insight_text+matrix',
        'insight_text+prioritization',
        'insight_text+workflow',
        'insight_text+table'
      ]).has(key)
    }
    return false
  }

  if (structureId === 'ZS08_quad_grid' || structureId === 'ZW04_four_columns_equal') {
    if (profile.count !== 1) return false
    if (profile.hasReasoning || profile.hasTable) return false
    return true
  }

  if (profile.hasReasoning) {
    if (!isDominantSlot(slotName)) return false
    if (profile.count > 2) return false
    return arts.every(a => isReasoningArtifact(a) || artifactType(a) === 'insight_text')
  }

  if (profile.hasCards && profile.compactCards && isDominantSlot(slotName)) return false
  if (profile.hasTable && isSupportSlot(slotName) && profile.count > 1) return false
  if (profile.hasWorkflow && isSupportSlot(slotName)) {
    const wf = arts.find(a => artifactType(a) === 'workflow')
    const wfType = String(wf?.workflow_type || '').toLowerCase()
    if (['process_flow', 'timeline'].includes(wfType)) return false
  }

  return true
}

function inferZoneStructure(slide) {
  const zones = slide?.zones || []
  const count = zones.length
  const allArtifacts = zones.flatMap(z => z.artifacts || [])
  const hasReasoning = allArtifacts.some(isReasoningArtifact)
  const hasCompactCards = zones.some(isCompactCardsZone)
  const hasWideWorkflow = zones.some(z => (z.artifacts || []).some(a => {
    const t = artifactType(a)
    const wfType = String(a?.workflow_type || '').toLowerCase()
    return t === 'workflow' && ['process_flow', 'timeline'].includes(wfType)
  }))
  const allSolo = zones.every(z => (z.artifacts || []).length === 1)

  if (count <= 1) return 'ZS01_single_full'
  if (count === 2) {
    if (hasReasoning || hasWideWorkflow) return 'ZS02_stacked_equal'
    return 'ZS03_side_by_side_equal'
  }
  if (count === 3) {
    if (hasCompactCards) return 'ZS06_top_full_bottom_two'
    if (hasWideWorkflow) return 'ZS07_top_two_bottom_dominant'
    return 'ZS04_left_dominant_right_stack'
  }
  if (count === 4) {
    if (allSolo) return 'ZS08_quad_grid'
    return 'ZS09_left_dominant_right_triptych'
  }
  return 'ZS08_quad_grid'
}

function applyZoneStructureMetadata(slide) {
  if (!slide || slide.slide_type !== 'content') return slide
  const structureId = slide.zone_structure || inferZoneStructure(slide)
  const def = zoneStructureDef(structureId) || zoneStructureDef(inferZoneStructure(slide))
  const zones = (slide.zones || []).slice(0, def?.zoneCount || 4).map((zone, idx) => ({
    ...zone,
    zone_slot: zone.zone_slot || def?.slots?.[idx] || `slot_${idx + 1}`
  }))
  return {
    ...slide,
    zone_structure: def ? structureId : inferZoneStructure(slide),
    zones
  }
}

function validateZoneStructureRules(slide) {
  if (!slide || slide.slide_type !== 'content') return true
  const structureId = slide.zone_structure || inferZoneStructure(slide)
  const def = zoneStructureDef(structureId)
  const zones = slide.zones || []
  if (!def) return false
  if (zones.length !== def.zoneCount) return false
  return zones.every((zone, idx) => validateZoneForStructureSlot(zone, zone.zone_slot || def.slots[idx], structureId))
}

function parseZoneSplitForConstraints(zone) {
  const split = String(zone?.zone_split || zone?.layout_hint?.split || 'full').toLowerCase()
  if (split === 'full' || split === 'bottom_full') return { fullWidth: true, fullHeight: split === 'full', widthFrac: 1, heightFrac: split === 'full' ? 1 : 0.5 }
  if (split === 'top_full') return { fullWidth: true, fullHeight: false, widthFrac: 1, heightFrac: 0.5 }
  const m = split.match(/^(left|right|top|bottom)_(\d{1,3})$/)
  if (m) {
    const side = m[1]
    const pct = Math.max(1, Math.min(99, parseInt(m[2], 10) || 0)) / 100
    if (side === 'left' || side === 'right') return { fullWidth: false, fullHeight: true, widthFrac: pct, heightFrac: 1 }
    return { fullWidth: true, fullHeight: false, widthFrac: 1, heightFrac: pct }
  }
  if (['top_left_50', 'top_right_50', 'tl', 'tr', 'bl', 'br'].includes(split)) return { fullWidth: false, fullHeight: false, widthFrac: 0.5, heightFrac: 0.5 }
  return { fullWidth: false, fullHeight: false, widthFrac: 0.5, heightFrac: 0.5 }
}

function validateWorkflowArtifactAndZone(workflowArtifact, zone) {
  if (artifactType(workflowArtifact) !== 'workflow') return true
  const workflowType = String(workflowArtifact.workflow_type || '').toLowerCase()
  const flow = String(workflowArtifact.flow_direction || '').toLowerCase()
  const nodeCount = artifactWorkflowNodeCount(workflowArtifact)
  const levelCount = artifactWorkflowLevelCount(workflowArtifact)
  const zoneShape = parseZoneSplitForConstraints(zone)

  if (workflowType === 'information_flow') return false

  if (workflowType === 'process_flow') {
    if (flow !== 'left_to_right') return false
    if (nodeCount < 4) return false
    if (!zoneShape.fullWidth) return false
    if (zoneShape.heightFrac < 0.5) return false
  }

  if (workflowType === 'hierarchy') {
    if (flow !== 'top_down_branching') return false
    if (levelCount < 3) return false
    if (zoneShape.widthFrac < 0.5) return false
    if (!zoneShape.fullHeight) return false
  }

  if (workflowType === 'decomposition') {
    if (!['left_to_right', 'top_to_bottom', 'top_down_branching'].includes(flow)) return false
    if (nodeCount > 3 && flow === 'left_to_right' && !zoneShape.fullWidth) return false
    if (nodeCount > 3 && ['top_to_bottom', 'top_down_branching'].includes(flow) && !zoneShape.fullHeight) return false
  }

  if (workflowType === 'timeline') {
    if (flow !== 'left_to_right') return false
    if (nodeCount < 4) return false
    if (!zoneShape.fullWidth) return false
    if (zoneShape.heightFrac < 0.5) return false
  }

  return true
}

function zoneArtifactPairKey(zone) {
  return ((zone?.artifacts || []).map(artifactType).sort().join('+'))
}

function slideHasDominantReasoningArtifact(slide) {
  const zones = slide?.zones || []
  const reasoningZones = zones.filter(z => (z.artifacts || []).some(isReasoningArtifact))
  if (!reasoningZones.length) return false
  return reasoningZones.some(z => (z.narrative_weight || '').toLowerCase() === 'primary')
}

// ─── STRUCTURAL PATTERN VALIDATOR ─────────────────────────────────────────────
// Add a new artifact type? Update the allowed pair lists below as needed.

function validateStructuralPatternRules(slide) {
  if (!slide || slide.slide_type !== 'content') return true
  const zones = slide.zones || []
  const zoneCount = zones.length
  const counts = zones.map(z => (z.artifacts || []).length)
  const totalArtifacts = counts.reduce((sum, n) => sum + n, 0)
  const allArtifacts = zones.flatMap(z => z.artifacts || [])
  const reasoningCount = allArtifacts.filter(isReasoningArtifact).length

  if (zoneCount === 1 && totalArtifacts === 1) {
    return isDenseSoloArtifact(allArtifacts[0])
  }

  if (zoneCount === 1 && totalArtifacts === 2) {
    const arts = zones[0].artifacts || []
    if (arts.some(isSoloOnlyArtifact)) return false
    const types = arts.map(artifactType)
    const hasInsight = types.includes('insight_text')
    const pair = types.slice().sort().join('+')
    if (pair === 'cards+insight_text') return artifactCardCount(arts.find(a => artifactType(a) === 'cards')) >= 4
    if (['chart+insight_text', 'workflow+insight_text', 'table+insight_text', 'matrix+insight_text', 'driver_tree+insight_text'].includes(pair)) {
      const nonInsight = arts.find(a => artifactType(a) !== 'insight_text')
      return hasInsight && isSubstantialArtifact(nonInsight)
    }
    if (pair === 'cards+workflow') {
      const cards = arts.find(a => artifactType(a) === 'cards')
      const workflow = arts.find(a => artifactType(a) === 'workflow')
      return artifactCardCount(cards) > 0 && artifactCardCount(cards) <= 2 && artifactWorkflowNodeCount(workflow) >= 4
    }
    if (pair === 'chart+table') return isSubstantialArtifact(arts[0]) || isSubstantialArtifact(arts[1])
    return false
  }

  if (zoneCount === 2 && totalArtifacts === 2) {
    if (!counts.every(n => n === 1)) return false
    if (reasoningCount > 1) return false
    const sparsePrimaryCards = zones.some(z => (z.narrative_weight || '').toLowerCase() === 'primary' && (z.artifacts || []).some(isSparseCardsArtifact))
    return !sparsePrimaryCards
  }

  if (zoneCount === 2 && totalArtifacts === 3) {
    if (counts.slice().sort((a, b) => a - b).join(',') !== '1,2') return false
    const pairedZone = zones.find(z => (z.artifacts || []).length === 2)
    const otherZone = zones.find(z => (z.artifacts || []).length === 1)
    const pairedArts = pairedZone?.artifacts || []
    if (pairedArts.some(isSoloOnlyArtifact)) return false
    const pairedReasoning = pairedArts.filter(isReasoningArtifact)
    const pairKey = zoneArtifactPairKey(pairedZone)
    const allowedPairs = new Set([
      'chart+insight_text',
      'insight_text+workflow',
      'insight_text+table',
      'driver_tree+insight_text',
      'insight_text+matrix',
      'cards+insight_text'
    ])
    if (pairedReasoning.length > 0 && pairedArts.some(a => !isReasoningArtifact(a) && artifactType(a) !== 'insight_text')) return false
    if ((otherZone?.artifacts || []).some(isReasoningArtifact) && pairedReasoning.length > 0) return false
    if (pairKey === 'cards+insight_text' && artifactCardCount(pairedArts.find(a => artifactType(a) === 'cards')) < 4) return false
    if (pairKey !== 'chart+table' && pairKey !== 'cards+workflow' && !allowedPairs.has(pairKey)) return false
    if (pairKey === 'cards+workflow') {
      const cards = pairedArts.find(a => artifactType(a) === 'cards')
      const workflow = pairedArts.find(a => artifactType(a) === 'workflow')
      if (!(artifactCardCount(cards) > 0 && artifactCardCount(cards) <= 2 && artifactWorkflowNodeCount(workflow) >= 4)) return false
    }
    if (pairKey === 'chart+table') {
      if (!pairedArts.every(isSubstantialArtifact)) return false
    }
    if ((otherZone?.artifacts || []).some(isReasoningArtifact) && reasoningCount > 1) return false
    return true
  }

  if (zoneCount === 2 && totalArtifacts === 4) {
    if (!counts.every(n => n === 2)) return false
    if (reasoningCount > 0) return false
    return zones.every(z => {
      const arts = z.artifacts || []
      if (arts.some(isSoloOnlyArtifact)) return false
      const types = arts.map(artifactType)
      const pair = types.slice().sort().join('+')
      return ['chart+insight_text', 'workflow+insight_text', 'insight_text+table', 'cards+insight_text'].includes(pair)
    })
  }

  if (zoneCount === 3 && totalArtifacts === 3) {
    if (!counts.every(n => n === 1)) return false
    const sparsePrimaryCards = zones.some(z => (z.narrative_weight || '').toLowerCase() === 'primary' && (z.artifacts || []).some(isSparseCardsArtifact))
    const hasPrimary = zones.some(z => (z.narrative_weight || '').toLowerCase() === 'primary')
    const hasSecondaryOrSupporting = zones.some(z => ['secondary', 'supporting'].includes((z.narrative_weight || '').toLowerCase()))
    return !sparsePrimaryCards
      && hasPrimary
      && hasSecondaryOrSupporting
  }

  if (zoneCount === 3 && totalArtifacts === 4) {
    if (counts.slice().sort((a, b) => a - b).join(',') !== '1,1,2') return false
    const pairedZones = zones.filter(z => (z.artifacts || []).length === 2)
    if (pairedZones.length !== 1) return false
    const pairedArts = pairedZones[0].artifacts || []
    if (pairedArts.some(isSoloOnlyArtifact)) return false
    if (pairedArts.some(isReasoningArtifact)) {
      if ((pairedZones[0].narrative_weight || '').toLowerCase() !== 'primary') return false
      if (pairedArts.some(a => !isReasoningArtifact(a) && artifactType(a) !== 'insight_text')) return false
    }
    if (reasoningCount > 0 && !slideHasDominantReasoningArtifact(slide)) return false
    return true
  }

  if (zoneCount === 3 && totalArtifacts >= 5) return false

  if (zoneCount === 4 && totalArtifacts === 4) {
    return counts.every(n => n === 1)
  }

  if (zoneCount === 4 && totalArtifacts >= 5) {
    if (reasoningCount > 0) return false
    return counts.every(n => n <= 2)
  }

  return true
}

// ─── ENFORCEMENT (normalisation-time fixes) ───────────────────────────────────
// These run during normaliseSlide to silently correct small structural issues
// before the validation gate. They do not make Claude API calls.

function enforceReasoningArtifactUsage(slide) {
  const reasoningTypes = new Set(['matrix', 'driver_tree', 'prioritization'])
  // These reasoning types must stand alone — no companion artifact permitted
  const reasoningSoloOnly = new Set(['prioritization'])
  if (!slide || slide.slide_type !== 'content') return slide
  const zones = (slide.zones || []).map(zone => {
    const arts = zone.artifacts || []
    if (arts.length <= 1) return zone

    // Reasoning artifacts: matrix/driver_tree/prioritization
    const reasoningArts = arts.filter(a => reasoningTypes.has(String(a.type || '').toLowerCase()))
    if (reasoningArts.length) {
      const primaryReasoning = reasoningArts[0]
      const primaryType = String(primaryReasoning.type || '').toLowerCase()
      if (reasoningSoloOnly.has(primaryType)) {
        return { ...zone, narrative_weight: 'primary', artifacts: [primaryReasoning] }
      }
      // matrix / driver_tree: keep one insight_text companion
      const insightArts = arts.filter(a => String(a.type || '').toLowerCase() === 'insight_text')
      return { ...zone, narrative_weight: 'primary', artifacts: [primaryReasoning].concat(insightArts.slice(0, 1)) }
    }

    // Non-reasoning solo-only types: risk_register, profile_card_set, process_flow/timeline workflow
    const soloArt = arts.find(isSoloOnlyArtifact)
    if (soloArt) {
      return { ...zone, artifacts: [soloArt] }
    }

    return zone
  })
  return { ...slide, zones }
}

function enforceStructuralPatternRules(slide) {
  if (!slide || slide.slide_type !== 'content') return slide
  const zones = (slide.zones || []).map(zone => {
    const arts = zone.artifacts || []
    const role = String(zone.zone_role || '').toLowerCase()
    // insight_text cannot be the sole artifact in a primary or co-primary zone.
    // It is always a secondary/supporting element — downgrade the zone role to 'secondary'.
    if (arts.length === 1 && String(arts[0]?.type || '').toLowerCase() === 'insight_text') {
      if (/^primary|co.?primary/.test(role)) {
        return { ...zone, zone_role: 'secondary' }
      }
    }
    return zone
  })
  return { ...slide, zones }
}

// ─── SLIDE-LEVEL VALIDATORS ───────────────────────────────────────────────────

function validateReasoningArtifactUsage(slide) {
  const reasoningTypes = new Set(['matrix', 'driver_tree', 'prioritization'])
  const zones = slide.zones || []
  for (const zone of zones) {
    const arts = zone.artifacts || []
    const reasoningArts = arts.filter(a => reasoningTypes.has((a.type || '').toLowerCase()))
    if (!reasoningArts.length) continue
    if ((zone.narrative_weight || 'primary') !== 'primary') return false
    if (reasoningArts.length > 1) return false
    if (arts.some(a => !reasoningTypes.has((a.type || '').toLowerCase()) && (a.type || '').toLowerCase() !== 'insight_text')) return false
  }
  return true
}

function validateWorkflowUsage(slide) {
  const zones = slide?.zones || []
  for (const zone of zones) {
    for (const artifact of (zone.artifacts || [])) {
      if (artifactType(artifact) === 'workflow' && !validateWorkflowArtifactAndZone(artifact, zone)) return false
    }
  }
  return true
}

function validateSlideArtifactMix(slide) {
  if (slide.slide_type !== 'content') return true
  const artifacts = []
  ;(slide.zones || []).forEach(zone => (zone.artifacts || []).forEach(art => artifacts.push((art.type || '').toLowerCase())))
  if (!artifacts.length) return false
  if (artifacts.every(t => t === 'cards')) return false
  if (!validateStructuralPatternRules(slide)) return false
  if (!validateZoneStructureRules(slide)) return false
  return true
}

// ─── PER-ARTIFACT SCHEMA VALIDATOR ────────────────────────────────────────────
// Add a case block here when adding a new artifact type.

function validateArtifact(artifact) {
  const t = (artifact.type || '').toLowerCase()

  if (t === 'chart') {
    const cats = artifact.categories || []
    const series = artifact.series || []
    const chartType = (artifact.chart_type || '').toLowerCase()

    // group_pie has its own validation rules
    if (chartType === 'group_pie') {
      if (cats.length < 2) return { valid: false, reason: 'group_pie needs 2+ categories (slices), got ' + cats.length }
      if (cats.length > 7) return { valid: false, reason: 'group_pie exceeds HARD MAX 7 slices, got ' + cats.length + ' — convert to clustered_bar' }
      if (series.length < 2) return { valid: false, reason: 'group_pie needs 2+ series (entities/pies), got ' + series.length + ' — convert to single pie' }
      if (series.length > 8) return { valid: false, reason: 'group_pie exceeds HARD MAX 8 entities, got ' + series.length + ' — convert to table or clustered_bar' }
      for (const s of series) {
        if ((s.values || []).length !== cats.length) {
          return { valid: false, reason: 'group_pie series "' + s.name + '" values count mismatch: ' + (s.values||[]).length + ' vs ' + cats.length }
        }
        if ((s.values || []).every(v => v === 0)) {
          return { valid: false, reason: 'group_pie series "' + s.name + '" has all-zero values' }
        }
      }
      return { valid: true }
    }

    // Need 2+ categories (3+ ideally but 2 is minimum after normalisation)
    if (cats.length < 2) return { valid: false, reason: 'chart needs 2+ categories, got ' + cats.length }

    if (!series.length) return { valid: false, reason: 'chart has no series' }

    // Values must match categories
    for (const s of series) {
      if ((s.values || []).length !== cats.length) {
        return { valid: false, reason: 'series values count mismatch: ' + (s.values||[]).length + ' vs ' + cats.length }
      }
    }

    // All-zero series
    for (const s of series) {
      if ((s.values || []).every(v => v === 0)) {
        return { valid: false, reason: 'chart series has all-zero values' }
      }
    }

    // clustered_bar needs 2 series
    if (artifact.chart_type === 'clustered_bar' && series.length < 2) {
      artifact.chart_type = 'bar' // auto-fix
    }

    return { valid: true }
  }

  if (t === 'stat_bar') {
    const rows = artifact.rows || []
    if (rows.length < 2) return { valid: false, reason: 'stat_bar needs 2+ rows' }
    const header = artifact.artifact_header || artifact.stat_header || artifact.chart_header || ''
    if (String(header).trim().length < 3) return { valid: false, reason: 'stat_bar missing artifact_header' }
    if (!artifact.annotation_style) return { valid: false, reason: 'stat_bar missing annotation_style' }
    const colHeaders = artifact.column_headers
    if (Array.isArray(colHeaders)) {
      // New flexible schema: rows have cells arrays
      const barCols = colHeaders.filter(c => c?.display_type === 'bar')
      if (barCols.length < 1) return { valid: false, reason: 'stat_bar must have at least one column with display_type "bar"' }
      if (barCols.length > 3) return { valid: false, reason: 'stat_bar supports at most 3 bar columns' }
      // Enforce: every "bar" column must be immediately followed by a "normal" column
      for (let bi = 0; bi < colHeaders.length; bi++) {
        if (colHeaders[bi]?.display_type === 'bar') {
          if (bi === colHeaders.length - 1 || colHeaders[bi + 1]?.display_type !== 'normal') {
            return { valid: false, reason: `stat_bar bar column "${colHeaders[bi].id}" must be immediately followed by a "normal" column` }
          }
        }
      }
      // Validate bar column cell values in each row
      for (const barCol of barCols) {
        const barColId = String(barCol.id)
        for (const row of rows) {
          if (!Array.isArray(row?.cells) || row.cells.length === 0) return { valid: false, reason: 'stat_bar row missing cells array' }
          const barCell = row.cells.find(c => String(c?.col_id) === barColId)
          if (!barCell) return { valid: false, reason: `stat_bar row missing cell for bar column "${barColId}"` }
          if (!Number.isFinite(+barCell.value)) return { valid: false, reason: `stat_bar bar column "${barColId}" cell value must be numeric` }
        }
        if (rows.every(r => {
          const barCell = (r?.cells || []).find(c => String(c?.col_id) === barColId)
          return (+barCell?.value || 0) === 0
        })) return { valid: false, reason: `stat_bar bar column "${barColId}" has all-zero values` }
      }
    } else {
      return { valid: false, reason: 'stat_bar column_headers must be an array — use the current schema with display_type per column' }
    }
    return { valid: true }
  }

  if (t === 'insight_text') {
    const groups = artifact.groups || []
    const points = artifact.points || []
    if (groups.length > 0) {
      // Grouped mode: valid if at least one group has bullets
      const hasBullets = groups.some(g => (g.bullets || []).length > 0)
      if (!hasBullets) return { valid: false, reason: 'insight_text grouped has no bullets in any group' }
      return { valid: true }
    }
    if (!points.length) return { valid: false, reason: 'insight_text has no points' }
    if (points.every(p => !p || p.trim().length < 5)) return { valid: false, reason: 'insight_text has only placeholder points' }
    return { valid: true }
  }

  if (t === 'cards') {
    const cards = artifact.cards || []
    if (!cards.length) return { valid: false, reason: 'cards has no items' }
    return { valid: true }
  }

  if (t === 'workflow') {
    const nodes = artifact.nodes || []
    if (!nodes.length) return { valid: false, reason: 'workflow has no nodes' }
    if (String(artifact.workflow_type || '').toLowerCase() === 'information_flow') return { valid: false, reason: 'information_flow is not allowed' }
    return { valid: true }
  }

  if (t === 'table') {
    if (!(artifact.headers || []).length) return { valid: false, reason: 'table has no headers' }
    if (!(artifact.rows || []).length) return { valid: false, reason: 'table has no rows' }
    return { valid: true }
  }

  if (t === 'matrix') {
    if (!artifact.x_axis?.label || !artifact.y_axis?.label) return { valid: false, reason: 'matrix missing axis labels' }
    if ((artifact.quadrants || []).length !== 4) return { valid: false, reason: 'matrix must define 4 quadrants' }
    if (!(artifact.points || []).length) return { valid: false, reason: 'matrix has no points' }
    const validQids = new Set(['q1','q2','q3','q4'])
    for (const p of artifact.points || []) {
      if (!p.quadrant_id || !validQids.has(p.quadrant_id)) return { valid: false, reason: `matrix point "${p.label}" missing valid quadrant_id` }
      if (typeof p.x !== 'number' || p.x < 0 || p.x > 100) return { valid: false, reason: `matrix point "${p.label}" x must be a number 0–100` }
      if (typeof p.y !== 'number' || p.y < 0 || p.y > 100) return { valid: false, reason: `matrix point "${p.label}" y must be a number 0–100` }
    }
    return { valid: true }
  }

  if (t === 'driver_tree') {
    if (!artifact.root?.node_label) return { valid: false, reason: 'driver_tree missing root node_label' }
    if (!(artifact.branches || []).length) return { valid: false, reason: 'driver_tree has no branches' }
    return { valid: true }
  }

  if (t === 'prioritization') {
    if (!(artifact.items || []).length) return { valid: false, reason: 'prioritization has no items' }
    if ((artifact.items || []).some(i => i.rank == null || !i.title)) return { valid: false, reason: 'prioritization items missing rank/title' }
    return { valid: true }
  }

  if (t === 'comparison_table') {
    if (!Array.isArray(artifact.columns) || !artifact.columns.length) return { valid: false, reason: 'comparison_table missing columns[]' }
    if (!Array.isArray(artifact.rows) || !artifact.rows.length) return { valid: false, reason: 'comparison_table missing rows[]' }
    if (artifact.rows.some(r => !(r.cells || []).length)) return { valid: false, reason: 'comparison_table row missing cells' }
    return { valid: true }
  }

  if (t === 'initiative_map') {
    if (!(artifact.column_headers || []).length) return { valid: false, reason: 'initiative_map missing column_headers[]' }
    if (!(artifact.rows || []).length) return { valid: false, reason: 'initiative_map missing rows[]' }
    return { valid: true }
  }

  if (t === 'risk_register') {
    if (!Array.isArray(artifact.severity_levels) || !artifact.severity_levels.length) return { valid: false, reason: 'risk_register missing severity_levels[]' }
    return { valid: true }
  }

  if (t === 'profile_card_set') {
    if (!(artifact.profiles || []).length) return { valid: false, reason: 'profile_card_set has no profiles' }
    if ((artifact.profiles || []).some(p => !p.entity_name)) return { valid: false, reason: 'profile_card_set profile missing entity_name' }
    return { valid: true }
  }

  return { valid: true }
}

// ─── VALIDATION GATE ──────────────────────────────────────────────────────────
// Entry point called by runAgent4 after each batch to identify slides needing repair.

function hasPlaceholderContent(slide) {
  if (slide.slide_type !== 'content') return false
  if (!slide.zones || !slide.zones.length) return true
  if (!slide.key_message || slide.key_message.trim().length < 10) return true
  if (!validateReasoningArtifactUsage(slide)) return true
  if (!validateWorkflowUsage(slide)) return true
  if (!validateSlideArtifactMix(slide)) return true

  for (const zone of slide.zones) {
    for (const artifact of (zone.artifacts || [])) {
      const check = validateArtifact(artifact)
      if (!check.valid) return true
    }
  }
  return false
}
