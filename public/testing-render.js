(function () {
  const ARTIFACT_CATALOG = [
    { type: 'insight_text', label: 'Insight Text', subtypes: ['standard', 'grouped'] },
    { type: 'chart', label: 'Chart', subtypes: ['bar', 'horizontal_bar', 'clustered_bar', 'line', 'pie', 'donut', 'group_pie'] },
    { type: 'stat_bar', label: 'Stat Bar', subtypes: ['stat_bar'] },
    { type: 'cards', label: 'Cards', subtypes: ['row', 'column', 'grid'] },
    { type: 'workflow', label: 'Workflow', subtypes: ['process_flow', 'hierarchy', 'decomposition', 'timeline'] },
    { type: 'table', label: 'Table', subtypes: ['table'] },
    { type: 'comparison_table', label: 'Comparison Table', subtypes: ['comparison_table'] },
    { type: 'initiative_map', label: 'Initiative Map', subtypes: ['initiative_map'] },
    { type: 'profile_card_set', label: 'Profile Card Set', subtypes: ['horizontal', 'grid'] },
    { type: 'risk_register', label: 'Risk Register', subtypes: ['risk_register'] },
    { type: 'matrix', label: 'Matrix', subtypes: ['2x2'] },
    { type: 'driver_tree', label: 'Driver Tree', subtypes: ['driver_tree'] },
    { type: 'prioritization', label: 'Prioritization', subtypes: ['ranked_list'] }
  ]

  const state = {
    manifestSlides: [],
    agent5Slides: [],
    brandTokens: {},
    selectedSlideNumber: null,
    overrides: {},
    rebuiltSlides: {},
    outputObject: null
  }

  const $ = (id) => document.getElementById(id)

  function deepClone(value) { return value == null ? value : JSON.parse(JSON.stringify(value)) }

  function localNormalizeArtifactType(type, chartType) {
    if (typeof normalizeArtifactType === 'function') return normalizeArtifactType(type, chartType)
    const t = String(type || '').toLowerCase()
    if (t === 'stat_bar' || t === 'star_bar') return 'stat_bar'
    return t || 'unknown'
  }

  function localArtifactSubtype(artifact) {
    const type = localNormalizeArtifactType(artifact?.type, artifact?.chart_type)
    if (type === 'chart') return String(artifact?.chart_type || 'bar').toLowerCase()
    if (type === 'insight_text') return String(artifact?.insight_mode || 'standard').toLowerCase()
    if (type === 'workflow') return String(artifact?.workflow_type || 'process_flow').toLowerCase()
    if (type === 'cards') return String(artifact?.cards_layout || 'column').toLowerCase()
    if (type === 'profile_card_set') return String(artifact?.layout_direction || 'horizontal').toLowerCase()
    if (type === 'matrix') return String(artifact?.matrix_type || '2x2').toLowerCase()
    return type
  }

  function artifactHeaderText(artifact, zone, slide) {
    return (
      artifact?.stat_header || artifact?.chart_header || artifact?.insight_header ||
      artifact?.workflow_header || artifact?.table_header || artifact?.comparison_header ||
      artifact?.initiative_header || artifact?.profile_header || artifact?.risk_header ||
      artifact?.matrix_header || artifact?.tree_header || artifact?.priority_header ||
      artifact?.heading || zone?.message_objective || slide?.key_message || slide?.title || 'Artifact Test'
    )
  }

  function safeParseJSON(raw) {
    const cleaned = String(raw || '').replace(/```json\s*/gi, '').replace(/```\s*/g, '').trim()
    return JSON.parse(cleaned)
  }

  function parseAgent4Input(raw) {
    const parsed = safeParseJSON(raw)
    if (Array.isArray(parsed)) return parsed
    if (parsed && Array.isArray(parsed.slides)) return parsed.slides
    if (parsed && parsed.slide_number) return [parsed]
    throw new Error('Agent 4 input must be a slide array, a single slide object, or an object with slides[].')
  }

  function parseAgent5Input(raw) {
    const parsed = safeParseJSON(raw)
    if (parsed && Array.isArray(parsed.slides)) return { brand_tokens: parsed.brand_tokens || {}, slides: parsed.slides }
    if (Array.isArray(parsed)) return { brand_tokens: {}, slides: parsed }
    if (parsed && parsed.slide_number) return { brand_tokens: {}, slides: [parsed] }
    throw new Error('Agent 5 input must be an object with slides[] or a slide array.')
  }

  function fileToText(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader()
      reader.onload = () => resolve(String(reader.result || ''))
      reader.onerror = reject
      reader.readAsText(file)
    })
  }

  function showError(message) {
    $('error-box').textContent = message
    $('error-box').classList.add('show')
    $('success-box').classList.remove('show')
  }

  function showSuccess(message) {
    $('success-box').textContent = message
    $('success-box').classList.add('show')
    $('error-box').classList.remove('show')
  }

  function resetMessages() {
    $('error-box').classList.remove('show')
    $('success-box').classList.remove('show')
  }

  function getCatalogEntry(type) { return ARTIFACT_CATALOG.find((entry) => entry.type === type) || null }
  function catalogSubtypeOptions(type) { return (getCatalogEntry(type)?.subtypes || [type]).slice() }

  function initialSelectionsForSlide(slide) {
    const selections = {}
    ;(slide?.zones || []).forEach((zone, zi) => {
      ;(zone.artifacts || []).forEach((artifact, ai) => {
        selections[`${zi}-${ai}`] = {
          type: localNormalizeArtifactType(artifact?.type, artifact?.chart_type),
          subtype: localArtifactSubtype(artifact)
        }
      })
    })
    return selections
  }

  function ensureSlideOverride(slideNumber) {
    if (!state.overrides[slideNumber]) {
      const manifestSlide = state.manifestSlides.find((slide) => slide.slide_number === slideNumber)
      state.overrides[slideNumber] = initialSelectionsForSlide(manifestSlide || {})
    }
    return state.overrides[slideNumber]
  }

  function estimateTextHeight(text, widthIn, fontSizePt, lineHeight) {
    const textStr = String(text || '').trim()
    if (!textStr) return 0
    const usableWidth = Math.max(0.3, Number(widthIn) || 0.3)
    const fontSize = Math.max(7, Number(fontSizePt) || 10)
    const charsPerLine = Math.max(8, Math.floor((usableWidth * 72) / (fontSize * 0.52)))
    const words = textStr.split(/\s+/).filter(Boolean)
    let lines = 1
    let lineLen = 0
    words.forEach((word) => {
      const nextLen = lineLen === 0 ? word.length : lineLen + 1 + word.length
      if (nextLen <= charsPerLine) lineLen = nextLen
      else { lines += 1; lineLen = word.length }
    })
    return lines * (fontSize * (lineHeight || 1.25) / 72)
  }

  function deriveTitleBlock(agent5Slide, bt) {
    if (agent5Slide?.title_block?.text) return deepClone(agent5Slide.title_block)
    const titleBlock = (agent5Slide?.blocks || []).find((block) =>
      block.block_role === 'slide_header' || (block.artifact_type === 'slide' && block.artifact_subtype === 'title')
    )
    if (titleBlock) {
      return {
        text: titleBlock.text || agent5Slide.title || '',
        x: titleBlock.x,
        y: titleBlock.y,
        w: titleBlock.w,
        h: titleBlock.h,
        font_family: titleBlock.font_family || bt.title_font_family || 'Arial',
        font_size: titleBlock.font_size || 20,
        font_weight: titleBlock.bold ? 'bold' : 'regular',
        color: titleBlock.color || bt.title_color || bt.primary_color || '#1F299C',
        align: titleBlock.align || 'left',
        valign: titleBlock.valign || 'top',
        wrap: true
      }
    }
    const title = agent5Slide?.title || ''
    const y = 0.6
    const h = Math.max(0.45, estimateTextHeight(title, 12.59, 20, 1.3) + 0.12)
    return {
      text: title,
      x: 0.37,
      y,
      w: 12.59,
      h,
      font_family: bt.title_font_family || 'Arial',
      font_size: 20,
      font_weight: 'bold',
      color: bt.title_color || bt.primary_color || '#1F299C',
      align: 'left',
      valign: 'top',
      wrap: true
    }
  }

  function deriveSubtitleBlock(agent5Slide, titleBlock, bt) {
    if (agent5Slide?.subtitle_block?.text) return deepClone(agent5Slide.subtitle_block)
    const subtitleBlock = (agent5Slide?.blocks || []).find((block) =>
      block.block_role === 'slide_subheader' || (block.artifact_type === 'slide' && block.artifact_subtype === 'subtitle')
    )
    if (subtitleBlock) {
      return {
        text: subtitleBlock.text || agent5Slide.subtitle || '',
        x: subtitleBlock.x,
        y: subtitleBlock.y,
        w: subtitleBlock.w,
        h: subtitleBlock.h,
        font_family: subtitleBlock.font_family || bt.body_font_family || 'Arial',
        font_size: subtitleBlock.font_size || 11,
        font_weight: subtitleBlock.bold ? 'bold' : 'regular',
        color: subtitleBlock.color || bt.caption_color || '#787CA9',
        align: subtitleBlock.align || 'left',
        valign: subtitleBlock.valign || 'middle',
        wrap: true
      }
    }
    const subtitle = agent5Slide?.subtitle || ''
    if (!subtitle) return null
    const y = (titleBlock?.y || 0.6) + (titleBlock?.h || 0.57) + 0.08
    const h = Math.max(0.26, estimateTextHeight(subtitle, 12.59, 11, 1.2) + 0.08)
    return {
      text: subtitle,
      x: 0.37,
      y,
      w: 12.59,
      h,
      font_family: bt.body_font_family || 'Arial',
      font_size: 11,
      font_weight: 'regular',
      color: bt.caption_color || '#787CA9',
      align: 'left',
      valign: 'middle',
      wrap: true
    }
  }

  function fallbackZoneFrame(index, total, canvas) {
    const margin = canvas?.margin || {}
    const left = Number(margin.left != null ? margin.left : 0.37)
    const right = Number(margin.right != null ? margin.right : 0.37)
    const top = Number(margin.top != null ? margin.top : 0.6)
    const bottom = Number(margin.bottom != null ? margin.bottom : 0.37)
    const width = Number(canvas?.width_in || 13.33)
    const height = Number(canvas?.height_in || 7.5)
    const usableW = width - left - right
    const usableH = height - top - bottom - 1.1
    const startY = top + 1.2
    if (total <= 1) return { x: left, y: startY, w: usableW, h: usableH }
    const gap = 0.22
    const frameW = (usableW - gap * (total - 1)) / total
    return { x: left + index * (frameW + gap), y: startY, w: frameW, h: usableH }
  }

  function resolveZoneFrame(manifestZone, designedZone, index, total, canvas) {
    return deepClone(designedZone?.frame || manifestZone?.frame || fallbackZoneFrame(index, total, canvas))
  }

  function makeDemoManifestArtifact(type, subtype, originalArtifact, zone, slide) {
    const header = artifactHeaderText(originalArtifact, zone, slide)
    const shared = {
      type,
      artifact_coverage_hint: originalArtifact?.artifact_coverage_hint || zone?.message_objective || ''
    }

    if (type === 'insight_text') {
      if (subtype === 'grouped') {
        return {
          ...shared,
          type: 'insight_text',
          insight_mode: 'grouped',
          insight_header: header,
          groups: [
            { header: 'Signal', bullets: ['Demand is concentrated and measurable.', 'The proof base is strong enough to compare formats.'] },
            { header: 'Test', bullets: ['Use this zone to validate grouped insight rendering.', 'Swap headers, bullets, or count to stress the layout.'] },
            { header: 'Outcome', bullets: ['This output is synthetic and only meant for artifact rendering.', 'Content quality is intentionally secondary here.'] }
          ],
          sentiment: 'neutral'
        }
      }
      return {
        ...shared,
        type: 'insight_text',
        insight_mode: 'standard',
        insight_header: header,
        heading: header,
        points: [
          'This is a local rendering test for the selected artifact.',
          'The app keeps slide context but swaps the artifact structure.',
          'Use the regenerated Agent 5 response to validate Agent 6 rendering.'
        ],
        sentiment: 'neutral'
      }
    }

    if (type === 'chart') {
      const chartType = subtype || 'bar'
      if (chartType === 'group_pie') {
        return {
          ...shared,
          type: 'chart',
          chart_type: 'group_pie',
          chart_header: header,
          chart_title: '',
          categories: ['High', 'Medium', 'Low'],
          series: [
            { name: 'Channel A', values: [45, 35, 20], unit: 'percent' },
            { name: 'Channel B', values: [30, 40, 30], unit: 'percent' },
            { name: 'Channel C', values: [55, 25, 20], unit: 'percent' }
          ],
          x_label: '',
          y_label: ''
        }
      }
      if (chartType === 'pie' || chartType === 'donut') {
        return {
          ...shared,
          type: 'chart',
          chart_type: chartType,
          chart_header: header,
          chart_title: '',
          categories: ['Core', 'Growth', 'Risk'],
          series: [{ name: 'Share', values: [52, 31, 17], unit: 'percent' }],
          x_label: '',
          y_label: ''
        }
      }
      if (chartType === 'clustered_bar') {
        return {
          ...shared,
          type: 'chart',
          chart_type: 'clustered_bar',
          chart_header: header,
          chart_title: '',
          categories: ['Q1', 'Q2', 'Q3', 'Q4'],
          series: [
            { name: 'Base', values: [18, 24, 21, 29], unit: 'count' },
            { name: 'Target', values: [22, 26, 25, 32], unit: 'count' }
          ],
          x_label: 'Quarter',
          y_label: 'Value'
        }
      }
      if (chartType === 'line') {
        return {
          ...shared,
          type: 'chart',
          chart_type: 'line',
          chart_header: header,
          chart_title: '',
          categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May'],
          series: [{ name: 'Trend', values: [10, 14, 13, 18, 22], unit: 'count' }],
          x_label: 'Month',
          y_label: 'Value'
        }
      }
      return {
        ...shared,
        type: 'chart',
        chart_type: chartType,
        chart_header: header,
        chart_title: '',
        categories: ['A', 'B', 'C', 'D'],
        series: [{ name: 'Series 1', values: [42, 34, 27, 19], unit: 'count' }],
        x_label: chartType === 'horizontal_bar' ? 'Value' : 'Category',
        y_label: chartType === 'horizontal_bar' ? 'Category' : 'Value'
      }
    }

    if (type === 'stat_bar') {
      return {
        ...shared,
        type: 'stat_bar',
        stat_header: header,
        stat_decision: 'Synthetic stat bar for local rendering validation.',
        column_headers: { label: 'Item', metric: 'Metric', value: 'Value', annotation: 'Note' },
        annotation_style: 'trailing',
        rows: [
          { id: 'row_1', label: 'Alpha', value: 1.2, display_value: '1.2', annotation: 'Primary highlight', highlight: true, bar_color: '#D7E4BF' },
          { id: 'row_2', label: 'Beta', value: 8.9, display_value: '8.9', annotation: 'Strong benchmark', highlight: false },
          { id: 'row_3', label: 'Gamma', value: 10.4, display_value: '10.4', annotation: 'Secondary case', highlight: false },
          { id: 'row_4', label: 'Delta', value: 12.0, display_value: '12.0', annotation: 'Upper bound', highlight: false }
        ]
      }
    }

    if (type === 'cards') {
      return {
        ...shared,
        type: 'cards',
        cards_layout: subtype || 'column',
        cards: [
          { title: 'Metric One', subtitle: '42%', body: 'Primary card body for render testing.', sentiment: 'positive' },
          { title: 'Metric Two', subtitle: '18%', body: 'Secondary card body for render testing.', sentiment: 'neutral' },
          { title: 'Metric Three', subtitle: '7%', body: 'Tertiary card body for render testing.', sentiment: 'negative' }
        ]
      }
    }

    if (type === 'workflow') {
      const workflowType = subtype || 'process_flow'
      const flowDirection = workflowType === 'hierarchy' ? 'top_to_bottom' : 'left_to_right'
      if (workflowType === 'hierarchy') {
        return {
          ...shared,
          type: 'workflow',
          workflow_type: 'hierarchy',
          flow_direction: flowDirection,
          workflow_header: header,
          nodes: [
            { id: 'n1', label: 'Root', value: 'Owner', description: 'Top node', level: 1 },
            { id: 'n2', label: 'Branch A', value: 'Step', description: 'First branch', level: 2 },
            { id: 'n3', label: 'Branch B', value: 'Step', description: 'Second branch', level: 2 },
            { id: 'n4', label: 'Leaf A1', value: 'Action', description: 'Leaf node', level: 3 },
            { id: 'n5', label: 'Leaf B1', value: 'Action', description: 'Leaf node', level: 3 }
          ],
          connections: [
            { from: 'n1', to: 'n2', type: 'arrow' },
            { from: 'n1', to: 'n3', type: 'arrow' },
            { from: 'n2', to: 'n4', type: 'arrow' },
            { from: 'n3', to: 'n5', type: 'arrow' }
          ]
        }
      }
      return {
        ...shared,
        type: 'workflow',
        workflow_type: workflowType,
        flow_direction: workflowType === 'timeline' ? 'left_to_right' : flowDirection,
        workflow_header: header,
        nodes: [
          { id: 'n1', label: 'Discover', value: '01', description: 'Start the test flow', level: 1 },
          { id: 'n2', label: 'Shape', value: '02', description: 'Convert to selected artifact', level: 1 },
          { id: 'n3', label: 'Render', value: '03', description: 'Flatten to blocks locally', level: 1 }
        ],
        connections: [
          { from: 'n1', to: 'n2', type: 'arrow' },
          { from: 'n2', to: 'n3', type: 'arrow' }
        ]
      }
    }

    if (type === 'table') {
      return {
        ...shared,
        type: 'table',
        table_header: header,
        headers: ['Option', 'Value', 'Status'],
        rows: [
          ['Alpha', '42', 'Good'],
          ['Beta', '31', 'Watch'],
          ['Gamma', '18', 'Low']
        ],
        highlight_rows: [0],
        note: 'Synthetic table for local artifact rendering.'
      }
    }

    if (type === 'comparison_table') {
      return {
        ...shared,
        type: 'comparison_table',
        comparison_header: header,
        criteria: ['Fit', 'Speed', 'Complexity'],
        options: [
          {
            name: 'Option A',
            ratings: [
              { criterion: 'Fit', rating: 'yes' },
              { criterion: 'Speed', rating: 'partial' },
              { criterion: 'Complexity', rating: 'no' }
            ]
          },
          {
            name: 'Option B',
            ratings: [
              { criterion: 'Fit', rating: 'yes' },
              { criterion: 'Speed', rating: 'yes' },
              { criterion: 'Complexity', rating: 'partial' }
            ]
          }
        ],
        recommended_option: 'Option B'
      }
    }

    if (type === 'initiative_map') {
      return {
        ...shared,
        type: 'initiative_map',
        initiative_header: header,
        dimension_labels: [
          { id: 'now', label: 'Now' },
          { id: 'next', label: 'Next' }
        ],
        initiatives: [
          {
            name: 'Data foundation',
            subtitle: 'Platform',
            placements: [
              { lane_id: 'now', title: 'Schema cleanup', chips: ['ETL', 'Keys'], outcome: 'Weeks 1-2' },
              { lane_id: 'next', title: 'Monitoring', chips: ['QA'], outcome: 'Weeks 3-4' }
            ]
          },
          {
            name: 'Artifact validation',
            subtitle: 'Design',
            placements: [
              { lane_id: 'now', title: 'Block test', chips: ['Agent 5'], outcome: 'Immediate' }
            ]
          }
        ]
      }
    }

    if (type === 'profile_card_set') {
      return {
        ...shared,
        type: 'profile_card_set',
        profile_header: header,
        layout_direction: subtype === 'grid' ? 'grid' : 'horizontal',
        profiles: [
          {
            entity_name: 'Persona A',
            subtitle: 'Primary segment',
            badge_text: 'Core',
            secondary_items: [
              { label: 'Need', value: 'Clarity' },
              { label: 'Market', value: ['Metro', 'Tier 1'], representation_type: 'chip_list' }
            ]
          },
          {
            entity_name: 'Persona B',
            subtitle: 'Secondary segment',
            badge_text: 'Pilot',
            secondary_items: [
              { label: 'Need', value: 'Speed' },
              { label: 'Market', value: ['Online'], representation_type: 'chip_list' }
            ]
          }
        ]
      }
    }

    if (type === 'risk_register') {
      return {
        ...shared,
        type: 'risk_register',
        risk_header: header,
        show_mitigation: true,
        risks: [
          { severity: 'high', description: 'Renderer mismatch across artifact families.', mitigation: 'Keep blocks canonical.', owner: 'Eng', status: 'Open' },
          { severity: 'medium', description: 'Synthetic content may overflow narrow zones.', mitigation: 'Adjust sample density.', owner: 'Design', status: 'Watch' },
          { severity: 'low', description: 'Non-critical cosmetic drift in preview.', mitigation: 'Validate with Agent 6.', owner: 'QA', status: 'Contained' }
        ]
      }
    }

    if (type === 'matrix') {
      return {
        ...shared,
        type: 'matrix',
        matrix_type: subtype || '2x2',
        matrix_header: header,
        x_axis: { label: 'Impact', low_label: 'Low', high_label: 'High' },
        y_axis: { label: 'Feasibility', low_label: 'Low', high_label: 'High' },
        quadrants: [
          { id: 'q1', name: 'Monitor', insight: 'Low impact, high effort.' },
          { id: 'q2', name: 'Invest', insight: 'High impact, high feasibility.' },
          { id: 'q3', name: 'Ignore', insight: 'Low impact, low feasibility.' },
          { id: 'q4', name: 'Test', insight: 'High impact, uncertain effort.' }
        ],
        points: [
          { label: 'Artifact A', x: 'high', y: 'high' },
          { label: 'Artifact B', x: 'medium', y: 'low' },
          { label: 'Artifact C', x: 'low', y: 'medium' }
        ]
      }
    }

    if (type === 'driver_tree') {
      return {
        ...shared,
        type: 'driver_tree',
        tree_header: header,
        root: { label: 'Primary outcome', value: '100%' },
        branches: [
          {
            label: 'Driver A',
            value: '40%',
            children: [
              { label: 'Leaf A1', value: '18%' },
              { label: 'Leaf A2', value: '22%' }
            ]
          },
          {
            label: 'Driver B',
            value: '60%',
            children: [
              { label: 'Leaf B1', value: '35%' },
              { label: 'Leaf B2', value: '25%' }
            ]
          }
        ]
      }
    }

    if (type === 'prioritization') {
      return {
        ...shared,
        type: 'prioritization',
        priority_header: header,
        items: [
          { rank: 1, title: 'Stabilize render contract', description: 'Keep block output deterministic.', qualifiers: ['High impact', 'Low effort'] },
          { rank: 2, title: 'Swap artifact families', description: 'Validate flattening coverage.', qualifiers: ['Medium effort'] },
          { rank: 3, title: 'Compare Agent 6 output', description: 'Use regenerated JSON as test input.', qualifiers: ['Fast feedback'] }
        ]
      }
    }

    return {
      ...shared,
      type: 'insight_text',
      insight_mode: 'standard',
      insight_header: header,
      heading: header,
      points: ['Fallback artifact placeholder.']
    }
  }

  function buildZoneShells(modifiedManifestSlide, originalAgent5Slide, bt) {
    const manifestZones = modifiedManifestSlide.zones || []
    const designedZones = originalAgent5Slide?.zones || []
    const shellZones = manifestZones.map((zone, zi) => {
      const designedZone = designedZones[zi] || null
      return {
        zone_id: zone.zone_id || `z${zi + 1}`,
        zone_role: zone.zone_role || 'supporting_evidence',
        narrative_weight: zone.narrative_weight || 'secondary',
        message_objective: zone.message_objective || '',
        layout_hint: zone.layout_hint || null,
        padding: zone.padding || null,
        frame: resolveZoneFrame(zone, designedZone, zi, manifestZones.length, originalAgent5Slide?.canvas),
        artifacts: (zone.artifacts || []).map((artifact) => {
          if (typeof buildSafeArtifactShell === 'function') return buildSafeArtifactShell(artifact, bt)
          return deepClone(artifact)
        })
      }
    })
    return typeof mergeContentIntoZones === 'function' ? mergeContentIntoZones(shellZones, manifestZones, bt) : shellZones
  }

  function buildSlideFromSelections(slideNumber) {
    const manifestSlide = state.manifestSlides.find((slide) => slide.slide_number === slideNumber)
    const originalSlide = state.agent5Slides.find((slide) => slide.slide_number === slideNumber)
    if (!manifestSlide || !originalSlide) return null

    const overrides = ensureSlideOverride(slideNumber)
    const modifiedManifestSlide = deepClone(manifestSlide)

    ;(modifiedManifestSlide.zones || []).forEach((zone, zi) => {
      ;(zone.artifacts || []).forEach((artifact, ai) => {
        const key = `${zi}-${ai}`
        const selection = overrides[key] || {
          type: localNormalizeArtifactType(artifact?.type, artifact?.chart_type),
          subtype: localArtifactSubtype(artifact)
        }
        zone.artifacts[ai] = makeDemoManifestArtifact(selection.type, selection.subtype, artifact, zone, modifiedManifestSlide)
      })
    })

    const bt = deepClone(state.brandTokens || {})
    const titleBlock = deriveTitleBlock(originalSlide, bt)
    const subtitleBlock = deriveSubtitleBlock(originalSlide, titleBlock, bt)

    const rebuiltSlide = {
      slide_number: originalSlide.slide_number,
      slide_type: originalSlide.slide_type || modifiedManifestSlide.slide_type || 'content',
      slide_archetype: originalSlide.slide_archetype || modifiedManifestSlide.slide_archetype || 'dashboard',
      layout_mode: false,
      selected_layout_name: '',
      canvas: deepClone(originalSlide.canvas || modifiedManifestSlide.canvas),
      global_elements: deepClone(originalSlide.global_elements || {}),
      title: titleBlock?.text || originalSlide.title || modifiedManifestSlide.title || '',
      subtitle: subtitleBlock?.text || originalSlide.subtitle || modifiedManifestSlide.subtitle || '',
      key_message: modifiedManifestSlide.key_message || originalSlide.key_message || '',
      speaker_note: modifiedManifestSlide.speaker_note || originalSlide.speaker_note || '',
      title_block: titleBlock,
      subtitle_block: subtitleBlock,
      zones: buildZoneShells(modifiedManifestSlide, originalSlide, bt)
    }

    if ((!rebuiltSlide.zones || rebuiltSlide.zones.length === 0) && typeof buildScratchZoneFrames === 'function') {
      rebuiltSlide.zones = buildScratchZoneFrames(modifiedManifestSlide.zones || [], rebuiltSlide)
    }

    if (typeof computeArtifactInternals === 'function') computeArtifactInternals(rebuiltSlide.zones || [], rebuiltSlide.canvas || {}, bt)
    if (typeof normalizeArtifactHeaderBands === 'function') normalizeArtifactHeaderBands(rebuiltSlide.zones || [])

    ;(rebuiltSlide.zones || []).forEach((zone, zi) => {
      ;(zone.artifacts || []).forEach((artifact, ai) => {
        if (!artifact._artifact_id) artifact._artifact_id = `s${slideNumber}_z${zi}_a${ai}`
      })
    })

    rebuiltSlide.blocks = typeof sanitizeBlocks === 'function' && typeof flattenToBlocks === 'function'
      ? sanitizeBlocks(flattenToBlocks(rebuiltSlide, bt), rebuiltSlide)
      : []

    rebuiltSlide.zones_summary = (rebuiltSlide.zones || []).map((zone) => ({
      zone_id: zone.zone_id,
      zone_role: zone.zone_role,
      narrative_weight: zone.narrative_weight,
      artifact_types: (zone.artifacts || []).map((artifact) => localNormalizeArtifactType(artifact?.type, artifact?.chart_type))
    }))

    const validationIssues = typeof validateDesignedSlide === 'function' ? validateDesignedSlide(rebuiltSlide) : []
    const renderIssues = typeof validateRenderCompleteness === 'function' ? validateRenderCompleteness(rebuiltSlide) : []
    rebuiltSlide._testing_render = { validation_issues: validationIssues, render_issues: renderIssues }
    return rebuiltSlide
  }

  function rebuildOutputObject() {
    const slides = state.agent5Slides.map((slide) => state.rebuiltSlides[slide.slide_number] || slide)
    state.outputObject = { brand_tokens: deepClone(state.brandTokens || {}), slides }
    $('output-json').textContent = JSON.stringify(state.outputObject, null, 2)
  }

  function updateStatsForSlide(slide) {
    $('stat-blocks').textContent = String((slide?.blocks || []).length)
    const artifactCount = (slide?.zones || []).reduce((sum, zone) => sum + ((zone.artifacts || []).length || 0), 0)
    $('stat-artifacts').textContent = String(artifactCount)
    const issueCount = ((slide?._testing_render?.validation_issues || []).length + (slide?._testing_render?.render_issues || []).length)
    $('stat-issues').textContent = String(issueCount)
  }

  function renderPreview(slide) {
    const canvas = $('preview-canvas')
    const ctx = canvas.getContext('2d')
    ctx.clearRect(0, 0, canvas.width, canvas.height)
    ctx.fillStyle = slide?.canvas?.background?.color || '#FFFFFF'
    ctx.fillRect(0, 0, canvas.width, canvas.height)

    const scaleX = canvas.width / (Number(slide?.canvas?.width_in || 13.33) || 13.33)
    const scaleY = canvas.height / (Number(slide?.canvas?.height_in || 7.5) || 7.5)
    const sx = (value) => Number(value || 0) * scaleX
    const sy = (value) => Number(value || 0) * scaleY

    function roundedRect(x, y, w, h, r) {
      const radius = Math.max(0, Math.min(r || 0, Math.min(w, h) / 2))
      ctx.beginPath()
      ctx.moveTo(x + radius, y)
      ctx.arcTo(x + w, y, x + w, y + h, radius)
      ctx.arcTo(x + w, y + h, x, y + h, radius)
      ctx.arcTo(x, y + h, x, y, radius)
      ctx.arcTo(x, y, x + w, y, radius)
      ctx.closePath()
    }

    ;(slide?.blocks || []).forEach((block) => {
      const type = block.block_type
      if (type === 'rect') {
        const x = sx(block.x), y = sy(block.y), w = sx(block.w), h = sy(block.h)
        roundedRect(x, y, w, h, Number(block.corner_radius || 0) * Math.min(scaleX, scaleY))
        ctx.fillStyle = block.fill_color || 'rgba(0,0,0,0)'
        ctx.fill()
        if (block.border_color && Number(block.border_width || 0) > 0) {
          ctx.strokeStyle = block.border_color
          ctx.lineWidth = Math.max(1, Number(block.border_width || 0))
          ctx.stroke()
        }
        return
      }
      if (type === 'rule' || type === 'line') {
        ctx.beginPath()
        if (type === 'rule') {
          ctx.moveTo(sx(block.x), sy(block.y))
          ctx.lineTo(sx((block.x || 0) + (block.w || 0)), sy(block.y))
        } else {
          ctx.moveTo(sx(block.x1), sy(block.y1))
          ctx.lineTo(sx(block.x2), sy(block.y2))
        }
        ctx.strokeStyle = block.color || '#64748b'
        ctx.lineWidth = Math.max(1, Number(block.line_width || 1))
        ctx.stroke()
        return
      }
      if (type === 'circle') {
        const x = sx(block.x), y = sy(block.y), w = sx(block.w), h = sy(block.h)
        ctx.beginPath()
        ctx.ellipse(x + w / 2, y + h / 2, w / 2, h / 2, 0, 0, Math.PI * 2)
        ctx.fillStyle = block.fill_color || '#cbd5e1'
        ctx.fill()
        if (block.text) {
          ctx.fillStyle = block.font_color || '#ffffff'
          ctx.font = `bold ${Math.max(10, h * 0.45)}px Arial`
          ctx.textAlign = 'center'
          ctx.textBaseline = 'middle'
          ctx.fillText(String(block.text), x + w / 2, y + h / 2)
        }
        return
      }
      if (type === 'text_box' || type === 'title' || type === 'subtitle') {
        const x = sx(block.x), y = sy(block.y), w = sx(block.w), h = sy(block.h)
        ctx.fillStyle = block.color || '#111827'
        ctx.font = `${block.bold ? 'bold ' : ''}${Math.max(10, Number(block.font_size || 10) * 1.45)}px Arial`
        ctx.textAlign = block.align === 'center' ? 'center' : block.align === 'right' ? 'right' : 'left'
        ctx.textBaseline = 'top'
        const words = String(block.text || '').split(/\s+/)
        const lines = []
        let current = ''
        const maxWidth = Math.max(20, w)
        words.forEach((word) => {
          const candidate = current ? `${current} ${word}` : word
          if (ctx.measureText(candidate).width <= maxWidth || !current) current = candidate
          else { lines.push(current); current = word }
        })
        if (current) lines.push(current)
        const lineHeight = Math.max(14, Number(block.font_size || 10) * 1.6)
        const anchorX = block.align === 'center' ? x + w / 2 : block.align === 'right' ? x + w : x
        lines.slice(0, Math.max(1, Math.floor(h / lineHeight))).forEach((line, idx) => {
          ctx.fillText(line, anchorX, y + idx * lineHeight)
        })
        return
      }
      if (type === 'bullet_list') {
        const x = sx(block.x), y = sy(block.y), h = sy(block.h)
        const style = block.body_style || {}
        ctx.fillStyle = style.color || '#111827'
        ctx.font = `${style.font_weight === 'bold' ? 'bold ' : ''}${Math.max(9, Number(style.font_size || 9) * 1.35)}px Arial`
        ctx.textAlign = 'left'
        ctx.textBaseline = 'top'
        const lineHeight = Math.max(13, Number(style.font_size || 9) * 1.6)
        ;(block.points || []).forEach((point, idx) => {
          const py = y + idx * lineHeight * 1.15
          if (py > y + h - lineHeight) return
          ctx.fillText(`• ${String(point)}`, x + 6, py)
        })
        return
      }
      if (type === 'chart' || type === 'table') {
        const x = sx(block.x), y = sy(block.y), w = sx(block.w), h = sy(block.h)
        ctx.fillStyle = '#f8fafc'
        ctx.fillRect(x, y, w, h)
        ctx.strokeStyle = '#cbd5e1'
        ctx.lineWidth = 1
        ctx.strokeRect(x, y, w, h)
        ctx.fillStyle = '#475569'
        ctx.font = 'bold 14px Arial'
        ctx.textAlign = 'center'
        ctx.textBaseline = 'middle'
        ctx.fillText(type.toUpperCase(), x + w / 2, y + h / 2)
      }
    })
  }

  function renderSlideMeta(slideNumber) {
    const manifestSlide = state.manifestSlides.find((slide) => slide.slide_number === slideNumber)
    const currentSlide = state.rebuiltSlides[slideNumber] || state.agent5Slides.find((slide) => slide.slide_number === slideNumber)
    $('slide-meta').innerHTML = [
      ['Slide', `S${slideNumber}`],
      ['Type', currentSlide?.slide_type || manifestSlide?.slide_type || '-'],
      ['Archetype', currentSlide?.slide_archetype || manifestSlide?.slide_archetype || '-'],
      ['Title', currentSlide?.title || manifestSlide?.title || '-']
    ].map(([label, value]) => `<div class="meta-block"><div class="meta-label">${label}</div><div class="meta-value">${value}</div></div>`).join('')
  }

  function renderSlideList() {
    $('slide-list').innerHTML = state.manifestSlides.map((slide) => {
      const active = slide.slide_number === state.selectedSlideNumber ? 'active' : ''
      return `<button class="slide-pill ${active}" data-slide="${slide.slide_number}">S${slide.slide_number}</button>`
    }).join('')
    ;[...document.querySelectorAll('.slide-pill')].forEach((button) => {
      button.addEventListener('click', () => {
        state.selectedSlideNumber = Number(button.dataset.slide)
        ensureSlideOverride(state.selectedSlideNumber)
        if (!state.rebuiltSlides[state.selectedSlideNumber]) {
          state.rebuiltSlides[state.selectedSlideNumber] = buildSlideFromSelections(state.selectedSlideNumber)
          rebuildOutputObject()
        }
        renderSelectedSlide()
      })
    })
  }

  function selectorOptionsHtml(options, selectedValue) {
    return options.map((value) => {
      const selected = String(value) === String(selectedValue) ? 'selected' : ''
      return `<option value="${value}" ${selected}>${value}</option>`
    }).join('')
  }

  function renderZones() {
    const slide = state.manifestSlides.find((item) => item.slide_number === state.selectedSlideNumber)
    const overrides = ensureSlideOverride(state.selectedSlideNumber)
    $('zones-wrap').innerHTML = ((slide?.zones || []).map((zone, zi) => {
      const artifactsHtml = (zone.artifacts || []).map((artifact, ai) => {
        const key = `${zi}-${ai}`
        const selection = overrides[key]
        const currentType = localNormalizeArtifactType(artifact?.type, artifact?.chart_type)
        const currentSubtype = localArtifactSubtype(artifact)
        const typeOptions = ARTIFACT_CATALOG.map((entry) => entry.type)
        const subtypeOptions = catalogSubtypeOptions(selection?.type || currentType)
        return `
          <div class="artifact-row">
            <div class="artifact-label">Artifact ${ai + 1}</div>
            <div class="artifact-current">Agent 4 original: <strong>${currentType}</strong>${currentSubtype && currentSubtype !== currentType ? ` / ${currentSubtype}` : ''}</div>
            <div class="selector-grid">
              <select data-role="type" data-key="${key}">
                ${selectorOptionsHtml(typeOptions, selection?.type || currentType)}
              </select>
              <select data-role="subtype" data-key="${key}">
                ${selectorOptionsHtml(subtypeOptions, selection?.subtype || currentSubtype || subtypeOptions[0])}
              </select>
            </div>
          </div>
        `
      }).join('')
      return `
        <div class="zone-card">
          <div class="zone-head">
            <div class="zone-title-row">
              <div class="zone-title">${zone.zone_id || `Zone ${zi + 1}`}</div>
              <div class="zone-badge">${zone.zone_role || 'zone'}</div>
            </div>
            <div class="zone-copy">${zone.message_objective || 'No message objective provided.'}</div>
            <div class="zone-frame">${JSON.stringify(zone.frame || {})}</div>
          </div>
          ${artifactsHtml || '<div class="artifact-row">No artifacts</div>'}
        </div>
      `
    }).join('')) || '<div class="zone-card"><div class="zone-head"><div class="zone-copy">This slide has no Agent 4 zones.</div></div></div>'

    ;[...document.querySelectorAll('select[data-role="type"]')].forEach((select) => {
      select.addEventListener('change', (event) => {
        const key = event.target.dataset.key
        const slideOverrides = ensureSlideOverride(state.selectedSlideNumber)
        const type = event.target.value
        slideOverrides[key] = { type, subtype: catalogSubtypeOptions(type)[0] }
        state.rebuiltSlides[state.selectedSlideNumber] = buildSlideFromSelections(state.selectedSlideNumber)
        rebuildOutputObject()
        renderSelectedSlide()
      })
    })
    ;[...document.querySelectorAll('select[data-role="subtype"]')].forEach((select) => {
      select.addEventListener('change', (event) => {
        const key = event.target.dataset.key
        const slideOverrides = ensureSlideOverride(state.selectedSlideNumber)
        slideOverrides[key] = { ...(slideOverrides[key] || {}), subtype: event.target.value }
        state.rebuiltSlides[state.selectedSlideNumber] = buildSlideFromSelections(state.selectedSlideNumber)
        rebuildOutputObject()
        renderSelectedSlide()
      })
    })
  }

  function renderSelectedSlide() {
    renderSlideList()
    renderSlideMeta(state.selectedSlideNumber)
    renderZones()
    const slide = state.rebuiltSlides[state.selectedSlideNumber] || buildSlideFromSelections(state.selectedSlideNumber)
    state.rebuiltSlides[state.selectedSlideNumber] = slide
    rebuildOutputObject()
    updateStatsForSlide(slide)
    renderPreview(slide)
  }

  function loadTestingRender() {
    resetMessages()
    try {
      const manifestSlides = parseAgent4Input($('agent4-input').value)
      const agent5Spec = parseAgent5Input($('agent5-input').value)
      const agent5Slides = agent5Spec.slides || []
      if (!manifestSlides.length) throw new Error('Agent 4 input parsed, but no slides were found.')
      if (!agent5Slides.length) throw new Error('Agent 5 input parsed, but no slides were found.')

      const manifestNumbers = new Set(manifestSlides.map((slide) => slide.slide_number))
      const sharedSlides = agent5Slides.filter((slide) => manifestNumbers.has(slide.slide_number))
      if (!sharedSlides.length) throw new Error('Agent 4 and Agent 5 do not share any slide_number values.')

      state.manifestSlides = manifestSlides.filter((slide) => sharedSlides.some((item) => item.slide_number === slide.slide_number))
      state.agent5Slides = sharedSlides
      state.brandTokens = agent5Spec.brand_tokens || {}
      state.selectedSlideNumber = state.manifestSlides[0].slide_number
      state.overrides = {}
      state.rebuiltSlides = {}
      ensureSlideOverride(state.selectedSlideNumber)
      state.rebuiltSlides[state.selectedSlideNumber] = buildSlideFromSelections(state.selectedSlideNumber)
      rebuildOutputObject()

      $('results').classList.add('show')
      renderSelectedSlide()
      showSuccess(`Loaded ${state.manifestSlides.length} shared slides. TestingRender is using local Agent 5 helpers only.`)
    } catch (error) {
      showError(error.message || String(error))
    }
  }

  function resetApp() {
    state.manifestSlides = []
    state.agent5Slides = []
    state.brandTokens = {}
    state.selectedSlideNumber = null
    state.overrides = {}
    state.rebuiltSlides = {}
    state.outputObject = null
    $('results').classList.remove('show')
    $('slide-list').innerHTML = ''
    $('slide-meta').innerHTML = ''
    $('zones-wrap').innerHTML = ''
    $('output-json').textContent = ''
    updateStatsForSlide(null)
    $('preview-canvas').getContext('2d').clearRect(0, 0, $('preview-canvas').width, $('preview-canvas').height)
    resetMessages()
  }

  async function bindFileInput(inputId, textareaId) {
    const input = $(inputId)
    input.addEventListener('change', async (event) => {
      const file = event.target.files && event.target.files[0]
      if (!file) return
      try {
        $(textareaId).value = await fileToText(file)
        showSuccess(`Loaded ${file.name}`)
      } catch (error) {
        showError(`Could not read ${file.name}: ${error.message}`)
      }
    })
  }

  function copyOutput() {
    if (!state.outputObject) return
    navigator.clipboard.writeText(JSON.stringify(state.outputObject, null, 2))
      .then(() => showSuccess('Updated Agent 5 JSON copied to clipboard.'))
      .catch((error) => showError(`Copy failed: ${error.message}`))
  }

  document.addEventListener('DOMContentLoaded', function () {
    $('load-btn').addEventListener('click', loadTestingRender)
    $('reset-btn').addEventListener('click', resetApp)
    $('copy-btn').addEventListener('click', copyOutput)
    bindFileInput('agent4-file', 'agent4-input')
    bindFileInput('agent5-file', 'agent5-input')
  })
})()
