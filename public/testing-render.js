(function () {
  const ARTIFACT_CATALOG = [
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
    agent5SlidesEnhanced: {},
    brandTokens: {},
    selectedSlideNumber: null,
    overrides: {},
    rebuiltSlides: {},
    outputObject: null
  }

  const $ = (id) => document.getElementById(id)

  function deepClone(value) { return value == null ? value : JSON.parse(JSON.stringify(value)) }
  function round2(value) { return Math.round(Number(value) * 100) / 100 }

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
        const normalizedType = localNormalizeArtifactType(artifact?.type, artifact?.chart_type)
        if (normalizedType === 'insight_text') return
        selections[`${zi}-${ai}`] = {
          type: normalizedType,
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

  function uniqStrings(values, fallbackPrefix) {
    const seen = new Set()
    return (values || []).map((value, idx) => {
      const raw = String(value == null ? `${fallbackPrefix || 'Item'} ${idx + 1}` : value).trim() || `${fallbackPrefix || 'Item'} ${idx + 1}`
      let candidate = raw
      let suffix = 2
      while (seen.has(candidate.toLowerCase())) {
        candidate = `${raw} ${suffix}`
        suffix += 1
      }
      seen.add(candidate.toLowerCase())
      return candidate
    })
  }

  function artifactTextLines(artifact, zone, slide) {
    const lines = []
    const push = (value) => {
      const text = String(value == null ? '' : value).trim()
      if (text) lines.push(text)
    }
    push(artifact?.heading)
    push(artifact?.insight_header)
    ;(artifact?.points || []).forEach(push)
    ;(artifact?.groups || []).forEach((group) => {
      push(group?.header)
      ;(group?.bullets || []).forEach(push)
    })
    ;(artifact?.cards || []).forEach((card) => {
      push(card?.title)
      push(card?.subtitle)
      push(card?.body)
    })
    ;(artifact?.rows || []).forEach((row) => {
      if (Array.isArray(row)) row.forEach(push)
      else {
        push(row?.label)
        push(row?.display_value)
        push(row?.annotation)
      }
    })
    ;(artifact?.headers || []).forEach(push)
    ;(artifact?.criteria || []).forEach(push)
    ;(artifact?.options || []).forEach((option) => {
      push(option?.name)
      ;(option?.ratings || []).forEach((rating) => {
        push(rating?.criterion)
        push(rating?.rating)
      })
    })
    ;(artifact?.initiatives || []).forEach((item) => {
      push(item?.name)
      push(item?.subtitle)
      ;(item?.placements || []).forEach((placement) => {
        push(placement?.title)
        push(placement?.outcome)
        ;(placement?.chips || []).forEach(push)
      })
    })
    ;(artifact?.profiles || []).forEach((profile) => {
      push(profile?.entity_name)
      push(profile?.subtitle)
      push(profile?.badge_text)
      ;(profile?.secondary_items || []).forEach((item) => {
        push(item?.label)
        if (Array.isArray(item?.value)) item.value.forEach(push)
        else push(item?.value)
      })
    })
    ;(artifact?.risks || []).forEach((risk) => {
      push(risk?.severity)
      push(risk?.description)
      push(risk?.mitigation)
      push(risk?.owner)
      push(risk?.status)
    })
    push(artifact?.root?.label)
    push(artifact?.root?.value)
    ;(artifact?.branches || []).forEach((branch) => {
      push(branch?.label)
      push(branch?.value)
      ;(branch?.children || []).forEach((child) => {
        push(child?.label)
        push(child?.value)
      })
    })
    ;(artifact?.items || []).forEach((item) => {
      push(item?.title)
      push(item?.description)
      ;(item?.qualifiers || []).forEach(push)
    })
    ;(artifact?.nodes || []).forEach((node) => {
      push(node?.label)
      push(node?.value)
      push(node?.description)
    })
    ;(artifact?.categories || []).forEach(push)
    ;(artifact?.series || []).forEach((series) => {
      push(series?.name)
      ;(series?.values || []).forEach((value) => push(value))
    })
    ;(artifact?.quadrants || []).forEach((quadrant) => {
      push(quadrant?.name)
      push(quadrant?.insight)
    })
    ;(artifact?.points || []).forEach((point) => {
      if (typeof point === 'string') return
      push(point?.label)
    })
    push(zone?.message_objective)
    push(slide?.key_message)
    return uniqStrings(lines, 'Item')
  }

  function artifactNumericSeries(artifact) {
    if (artifact?.type === 'chart' && Array.isArray(artifact?.series) && artifact.series.length) {
      const categories = uniqStrings(artifact.categories || artifact.series[0]?.values?.map((_, idx) => `Series ${idx + 1}`) || [], 'Item')
      const series = artifact.series.map((entry, idx) => ({
        name: String(entry?.name || `Series ${idx + 1}`),
        values: (entry?.values || []).map((value, vi) => Number.isFinite(Number(value)) ? Number(value) : vi + 1),
        unit: entry?.unit || 'count'
      }))
      return { categories, series }
    }
    if (artifact?.type === 'stat_bar') {
      const rows = artifact.rows || []
      return {
        categories: uniqStrings(rows.map((row, idx) => row?.label || `Item ${idx + 1}`), 'Item'),
        series: [{
          name: (artifact.column_headers && (artifact.column_headers.metric || artifact.column_headers.value)) || 'Value',
          values: rows.map((row, idx) => Number.isFinite(Number(row?.value)) ? Number(row.value) : idx + 1),
          unit: 'count'
        }]
      }
    }
    if (artifact?.type === 'table') {
      const rows = artifact.rows || []
      return {
        categories: uniqStrings(rows.map((row, idx) => Array.isArray(row) ? row[0] : `Row ${idx + 1}`), 'Row'),
        series: [{
          name: (artifact.headers && artifact.headers[1]) || 'Value',
          values: rows.map((row, idx) => {
            const value = Array.isArray(row) ? row[1] : null
            return Number.isFinite(Number(value)) ? Number(value) : idx + 1
          }),
          unit: 'count'
        }]
      }
    }
    const labels = artifactTextLines(artifact).slice(0, 5)
    return {
      categories: uniqStrings(labels.length ? labels : ['A', 'B', 'C'], 'Item'),
      series: [{
        name: 'Value',
        values: (labels.length ? labels : ['A', 'B', 'C']).map((_, idx) => idx + 1),
        unit: 'count'
      }]
    }
  }

  function deriveProfileCardsFromArtifact(originalArtifact, lines, standardPoints) {
    const statBarRows = deriveStructuredStatBarRows(originalArtifact, lines, standardPoints)
    if (statBarRows) {
      const costLabel = originalArtifact?.column_headers?.value || 'Value'
      const annotationLabel = originalArtifact?.column_headers?.annotation || 'Use case'
      const metricLabel = originalArtifact?.column_headers?.metric || 'Metric'
      return statBarRows.map((row, idx) => ({
        entity_name: row.name || `Profile ${idx + 1}`,
        subtitle: metricLabel,
        badge_text: row.metricText || `Value ${idx + 1}`,
        secondary_items: [
          { label: costLabel, value: row.cost || 'Sample' },
          { label: annotationLabel, value: row.useCase || row.annotation || 'Sample' }
        ]
      }))
    }

    if (originalArtifact?.type === 'table' && Array.isArray(originalArtifact?.rows) && originalArtifact.rows.length) {
      const headers = Array.isArray(originalArtifact?.headers) ? originalArtifact.headers : []
      return originalArtifact.rows.slice(0, 6).map((row, idx) => {
        const cells = Array.isArray(row) ? row : []
        return {
          entity_name: String(cells[0] || `Profile ${idx + 1}`),
          subtitle: headers[1] || 'Value',
          badge_text: String(cells[1] || ''),
          secondary_items: cells.slice(2, 5).map((cell, ci) => ({
            label: headers[ci + 2] || `Field ${ci + 1}`,
            value: String(cell || '')
          })).filter((item) => item.value)
        }
      })
    }

    return null
  }

  function parseStatBarAnnotationParts(annotation) {
    const raw = String(annotation || '').trim()
    if (!raw) {
      return { cost: '', useCase: '' }
    }
    const costMatch = raw.match(/₹\s*[\d,.]+(?:\s*\/\s*[A-Za-z]+)?/i)
    const cost = costMatch ? costMatch[0].replace(/\s+/g, '') : ''
    let useCase = raw
    if (costMatch) {
      useCase = `${raw.slice(0, costMatch.index)} ${raw.slice((costMatch.index || 0) + costMatch[0].length)}`.trim()
      useCase = useCase.replace(/^[;,\-\s]+|[;,\-\s]+$/g, '').trim()
    }
    if (!useCase && raw.includes(';')) {
      useCase = raw.split(';').slice(1).join(';').trim()
    }
    return { cost, useCase }
  }

  function deriveStructuredStatBarRows(originalArtifact, lines, standardPoints) {
    if (!(originalArtifact?.type === 'stat_bar' && Array.isArray(originalArtifact?.rows) && originalArtifact.rows.length)) {
      return null
    }

    return originalArtifact.rows.slice(0, 6).map((row, idx) => {
      const annotation = String(row?.annotation || standardPoints[idx] || lines[idx + 1] || '').trim()
      const parsed = parseStatBarAnnotationParts(annotation)
      const numericValue = Number(row?.value)
      const metricText = String(
        row?.display_value ||
        `${row?.value ?? ''}${row?.unit ? `${row.unit}` : ''}`.trim() ||
        `Value ${idx + 1}`
      ).trim() || `Value ${idx + 1}`
      return {
        id: row?.id || `row_${idx + 1}`,
        rank: idx + 1,
        name: String(row?.label || `Item ${idx + 1}`),
        metricValue: Number.isFinite(numericValue) ? numericValue : null,
        metricText,
        cost: parsed.cost || 'Sample',
        useCase: parsed.useCase || annotation || 'Sample',
        annotation: annotation || 'Sample',
        highlight: row?.highlight === true
      }
    })
  }

  function comparisonRatingFromRank(rank, total, invert) {
    const safeRank = Math.max(1, Number(rank) || 1)
    const safeTotal = Math.max(1, Number(total) || 1)
    const effectiveRank = invert ? (safeTotal - safeRank + 1) : safeRank
    if (effectiveRank <= 1) return 'yes'
    if (effectiveRank === 2) return 'partial'
    return 'no'
  }

  function riskSeverityFromRow(row, idx, total) {
    if (row?.highlight) return 'low'
    if ((row?.metricValue || 0) >= 99.5) return 'low'
    if ((row?.metricValue || 0) >= 98.5) return 'medium'
    return idx >= total - 1 ? 'high' : 'medium'
  }

  function buildConvertedManifestArtifact(type, subtype, originalArtifact, zone, slide) {
    const header = artifactHeaderText(originalArtifact, zone, slide)
    const shared = {
      type,
      artifact_coverage_hint: originalArtifact?.artifact_coverage_hint || zone?.message_objective || ''
    }
    const lines = artifactTextLines(originalArtifact, zone, slide)
    const metrics = artifactNumericSeries(originalArtifact)
    const standardPoints = uniqStrings((originalArtifact?.points || []).length ? originalArtifact.points : lines.slice(0, 4), 'Point').slice(0, 5)
    const groupedSource = Array.isArray(originalArtifact?.groups) && originalArtifact.groups.length
      ? originalArtifact.groups.map((group, idx) => ({
          header: String(group?.header || `Group ${idx + 1}`),
          bullets: uniqStrings(group?.bullets || [], 'Point').slice(0, 3)
        }))
      : [
          { header: 'Signal', bullets: standardPoints.slice(0, 2) },
          { header: 'Evidence', bullets: standardPoints.slice(2, 4) },
          { header: 'Implication', bullets: standardPoints.slice(4, 6) }
        ].filter((group) => group.bullets.length)
    const cardItems = Array.isArray(originalArtifact?.cards) && originalArtifact.cards.length
      ? originalArtifact.cards
      : standardPoints.slice(0, 4).map((point, idx) => ({
          title: lines[idx] || `Card ${idx + 1}`,
          subtitle: idx === 0 ? 'Primary' : `Item ${idx + 1}`,
          body: point,
          sentiment: originalArtifact?.sentiment || 'neutral'
        }))

    if (type === 'insight_text') {
      if (subtype === 'grouped') {
        return {
          ...shared,
          type: 'insight_text',
          insight_mode: 'grouped',
          insight_header: header,
          heading: header,
          groups: groupedSource,
          sentiment: originalArtifact?.sentiment || 'neutral'
        }
      }
      return {
        ...shared,
        type: 'insight_text',
        insight_mode: 'standard',
        insight_header: header,
        heading: header,
        points: standardPoints,
        sentiment: originalArtifact?.sentiment || 'neutral'
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
          categories: metrics.categories.slice(0, 5),
          series: metrics.series.slice(0, 3).map((series, idx) => ({
            name: series.name || `Series ${idx + 1}`,
            values: series.values.slice(0, Math.min(5, metrics.categories.length)),
            unit: series.unit || 'count'
          })),
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
          categories: metrics.categories.slice(0, 6),
          series: [{
            name: metrics.series[0]?.name || 'Share',
            values: (metrics.series[0]?.values || []).slice(0, Math.min(6, metrics.categories.length)),
            unit: metrics.series[0]?.unit || 'count'
          }],
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
          categories: metrics.categories.slice(0, 6),
          series: (metrics.series.length > 1 ? metrics.series.slice(0, 3) : [
            metrics.series[0],
            {
              name: `${metrics.series[0]?.name || 'Value'} Target`,
              values: (metrics.series[0]?.values || []).slice(0, 6).map((value) => Math.round(Number(value || 0) * 1.15 * 100) / 100),
              unit: metrics.series[0]?.unit || 'count'
            }
          ]).map((series, idx) => ({
            name: series?.name || `Series ${idx + 1}`,
            values: (series?.values || []).slice(0, Math.min(6, metrics.categories.length)),
            unit: series?.unit || 'count'
          })),
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
          categories: metrics.categories.slice(0, 6),
          series: [{
            name: metrics.series[0]?.name || 'Trend',
            values: (metrics.series[0]?.values || []).slice(0, Math.min(6, metrics.categories.length)),
            unit: metrics.series[0]?.unit || 'count'
          }],
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
        categories: metrics.categories.slice(0, 8),
        series: (metrics.series.length ? metrics.series : [{ name: 'Series 1', values: [1, 2, 3], unit: 'count' }]).slice(0, 3).map((series, idx) => ({
          name: series?.name || `Series ${idx + 1}`,
          values: (series?.values || []).slice(0, Math.min(8, metrics.categories.length)),
          unit: series?.unit || 'count'
        })),
        x_label: chartType === 'horizontal_bar' ? 'Value' : 'Category',
        y_label: chartType === 'horizontal_bar' ? 'Category' : 'Value'
      }
    }

    if (type === 'stat_bar') {
      return {
        ...shared,
        type: 'stat_bar',
        stat_header: header,
        stat_decision: zone?.message_objective || slide?.key_message || header,
        column_headers: { label: 'Item', metric: 'Metric', value: 'Value', annotation: 'Note' },
        annotation_style: 'trailing',
        rows: metrics.categories.slice(0, 6).map((label, idx) => {
          const value = metrics.series[0]?.values?.[idx]
          return {
            id: `row_${idx + 1}`,
            label,
            value: Number.isFinite(Number(value)) ? Number(value) : idx + 1,
            display_value: String(Number.isFinite(Number(value)) ? Number(value) : idx + 1),
            annotation: standardPoints[idx] || lines[idx + 1] || header,
            highlight: idx === 0,
            bar_color: idx === 0 ? '#D7E4BF' : undefined
          }
        })
      }
    }

    if (type === 'cards') {
      return {
        ...shared,
        type: 'cards',
        cards_layout: subtype || 'column',
        cards: cardItems.slice(0, 6)
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
            { id: 'n1', label: lines[0] || header, value: 'Root', description: standardPoints[0] || '', level: 1 },
            { id: 'n2', label: lines[1] || 'Branch A', value: 'Branch', description: standardPoints[1] || '', level: 2 },
            { id: 'n3', label: lines[2] || 'Branch B', value: 'Branch', description: standardPoints[2] || '', level: 2 },
            { id: 'n4', label: lines[3] || 'Leaf A1', value: 'Action', description: standardPoints[3] || '', level: 3 },
            { id: 'n5', label: lines[4] || 'Leaf B1', value: 'Action', description: standardPoints[4] || '', level: 3 }
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
          { id: 'n1', label: lines[0] || 'Discover', value: '01', description: standardPoints[0] || '', level: 1 },
          { id: 'n2', label: lines[1] || 'Shape', value: '02', description: standardPoints[1] || '', level: 1 },
          { id: 'n3', label: lines[2] || 'Render', value: '03', description: standardPoints[2] || '', level: 1 }
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
        headers: ['Item', 'Value', 'Note'],
        rows: metrics.categories.slice(0, 6).map((label, idx) => [
          label,
          String(metrics.series[0]?.values?.[idx] ?? idx + 1),
          standardPoints[idx] || lines[idx + 1] || ''
        ]),
        highlight_rows: [0].filter((_, idx) => idx < Math.min(1, metrics.categories.length)),
        note: zone?.message_objective || slide?.key_message || ''
      }
    }

    if (type === 'comparison_table') {
      const statBarRows = deriveStructuredStatBarRows(originalArtifact, lines, standardPoints)
      if (statBarRows) {
        const metricCriterion = originalArtifact?.column_headers?.metric || 'Metric'
        const costCriterion = originalArtifact?.column_headers?.value || 'Value'
        const noteCriterion = originalArtifact?.column_headers?.annotation || 'Use case'
        const costs = statBarRows.map((row, idx) => ({
          idx,
          value: Number(String(row.cost || '').replace(/[^\d.]/g, ''))
        }))
        const rankedByCost = costs
          .map((entry) => ({ ...entry, value: Number.isFinite(entry.value) ? entry.value : Number.POSITIVE_INFINITY }))
          .sort((a, b) => a.value - b.value)
        const costRanks = new Map(rankedByCost.map((entry, rank) => [entry.idx, rank + 1]))

        return {
          ...shared,
          type: 'comparison_table',
          comparison_header: header,
          criteria: [metricCriterion, costCriterion, noteCriterion],
          options: statBarRows.map((row, idx) => ({
            name: row.name,
            ratings: [
              { criterion: metricCriterion, rating: comparisonRatingFromRank(idx + 1, statBarRows.length, false), note: row.metricText },
              { criterion: costCriterion, rating: comparisonRatingFromRank(costRanks.get(idx) || idx + 1, statBarRows.length, false), note: row.cost || 'Sample' },
              {
                criterion: noteCriterion,
                rating: row.highlight || /benchmark|best|workhorse/i.test(row.useCase) ? 'yes' : (/premium/i.test(row.useCase) ? 'partial' : 'no'),
                note: row.useCase || 'Sample'
              }
            ]
          })),
          recommended_option: (statBarRows.find((row) => row.highlight) || statBarRows[0] || {}).name || 'Sample'
        }
      }

      return {
        ...shared,
        type: 'comparison_table',
        comparison_header: header,
        criteria: uniqStrings((originalArtifact?.criteria || lines.slice(0, 3)), 'Criteria').slice(0, 4),
        options: [
          {
            name: lines[3] || 'Option A',
            ratings: uniqStrings((originalArtifact?.criteria || lines.slice(0, 3)), 'Criteria').slice(0, 4).map((criterion, idx) => ({
              criterion,
              rating: ['yes', 'partial', 'no'][idx % 3]
            }))
          },
          {
            name: lines[4] || 'Option B',
            ratings: uniqStrings((originalArtifact?.criteria || lines.slice(0, 3)), 'Criteria').slice(0, 4).map((criterion, idx) => ({
              criterion,
              rating: ['yes', 'yes', 'partial', 'no'][idx % 4]
            }))
          }
        ],
        recommended_option: lines[4] || 'Option B'
      }
    }

    if (type === 'initiative_map') {
      const statBarRows = deriveStructuredStatBarRows(originalArtifact, lines, standardPoints)
      if (statBarRows) {
        const metricLabel = originalArtifact?.column_headers?.metric || 'Metric'
        const costLabel = originalArtifact?.column_headers?.value || 'Value'
        const noteLabel = originalArtifact?.column_headers?.annotation || 'Use case'
        return {
          ...shared,
          type: 'initiative_map',
          initiative_header: header,
          dimension_labels: [
            { id: 'metric', label: metricLabel },
            { id: 'cost', label: costLabel },
            { id: 'use_case', label: noteLabel }
          ],
          initiatives: statBarRows.map((row) => ({
            name: row.name,
            subtitle: row.highlight ? 'Highlighted row' : 'Converted from stat bar row',
            placements: [
              { lane_id: 'metric', title: row.metricText || 'Sample', chips: [row.highlight ? 'Highlighted' : 'Standard'], outcome: 'Sample' },
              { lane_id: 'cost', title: row.cost || 'Sample', chips: ['Cost'], outcome: 'Sample' },
              { lane_id: 'use_case', title: row.useCase || 'Sample', chips: ['Use case'], outcome: row.metricText || 'Sample' }
            ]
          }))
        }
      }

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
            name: lines[0] || 'Primary initiative',
            subtitle: lines[1] || 'Current workstream',
            placements: [
              { lane_id: 'now', title: standardPoints[0] || 'Current step', chips: lines.slice(2, 4), outcome: 'Current' },
              { lane_id: 'next', title: standardPoints[1] || 'Next step', chips: lines.slice(4, 6), outcome: 'Next' }
            ]
          },
          {
            name: lines[2] || 'Secondary initiative',
            subtitle: lines[3] || 'Follow-up',
            placements: [
              { lane_id: 'now', title: standardPoints[2] || 'Immediate action', chips: lines.slice(6, 8), outcome: 'Immediate' }
            ]
          }
        ]
      }
    }

    if (type === 'profile_card_set') {
      const derivedProfiles = deriveProfileCardsFromArtifact(originalArtifact, lines, standardPoints)
      return {
        ...shared,
        type: 'profile_card_set',
        profile_header: header,
        layout_direction: subtype === 'grid' ? 'grid' : 'horizontal',
        profiles: (derivedProfiles || cardItems.slice(0, 4).map((card, idx) => ({
          entity_name: card.title || `Profile ${idx + 1}`,
          subtitle: card.subtitle || `Segment ${idx + 1}`,
          badge_text: idx === 0 ? 'Core' : 'Test',
          secondary_items: [
            { label: 'Insight', value: card.body || '' },
            { label: 'Signals', value: lines.slice(idx * 2, idx * 2 + 2), representation_type: 'chip_list' }
          ]
        }))).map((profile) => ({
          ...profile,
          secondary_items: Array.isArray(profile.secondary_items) ? profile.secondary_items.slice(0, 3) : []
        }))
      }
    }

    if (type === 'risk_register') {
      const statBarRows = deriveStructuredStatBarRows(originalArtifact, lines, standardPoints)
      if (statBarRows) {
        return {
          ...shared,
          type: 'risk_register',
          risk_header: header,
          show_mitigation: true,
          risks: statBarRows.map((row, idx) => ({
            severity: riskSeverityFromRow(row, idx, statBarRows.length),
            title: row.name,
            detail: `On-time: ${row.metricText || 'Sample'} | Avg cost: ${row.cost || 'Sample'}`,
            mitigation: row.useCase || 'Sample',
            owner: 'Sample',
            status: row.highlight ? 'Contained' : 'Open',
            likelihood: row.highlight ? 1 : Math.min(3, idx + 1),
            impact: row.highlight ? 1 : Math.max(1, Math.min(3, statBarRows.length - idx))
          }))
        }
      }

      return {
        ...shared,
        type: 'risk_register',
        risk_header: header,
        show_mitigation: true,
        risks: standardPoints.slice(0, 4).map((point, idx) => ({
          severity: ['high', 'medium', 'low', 'medium'][idx % 4],
          description: point,
          mitigation: lines[idx + 1] || 'Review and mitigate.',
          owner: ['Ops', 'Design', 'QA', 'Eng'][idx % 4],
          status: ['Open', 'Watch', 'Contained', 'Open'][idx % 4]
        }))
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
          { id: 'q1', name: 'Monitor', insight: standardPoints[0] || 'Low impact, high effort.' },
          { id: 'q2', name: 'Invest', insight: standardPoints[1] || 'High impact, high feasibility.' },
          { id: 'q3', name: 'Ignore', insight: standardPoints[2] || 'Low impact, low feasibility.' },
          { id: 'q4', name: 'Test', insight: standardPoints[3] || 'High impact, uncertain effort.' }
        ],
        points: lines.slice(0, 4).map((line, idx) => ({
          label: line,
          x: ['high', 'medium', 'low', 'medium'][idx % 4],
          y: ['high', 'low', 'medium', 'high'][idx % 4]
        }))
      }
    }

    if (type === 'driver_tree') {
      return {
        ...shared,
        type: 'driver_tree',
        tree_header: header,
        root: { label: lines[0] || header, value: '100%' },
        branches: [
          {
            label: lines[1] || 'Driver A',
            value: '40%',
            children: [
              { label: lines[3] || 'Leaf A1', value: '18%' },
              { label: lines[4] || 'Leaf A2', value: '22%' }
            ]
          },
          {
            label: lines[2] || 'Driver B',
            value: '60%',
            children: [
              { label: lines[5] || 'Leaf B1', value: '35%' },
              { label: lines[6] || 'Leaf B2', value: '25%' }
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
        items: standardPoints.slice(0, 5).map((point, idx) => ({
          rank: idx + 1,
          title: lines[idx] || `Priority ${idx + 1}`,
          description: point,
          qualifiers: [
            { label: idx === 0 ? 'High impact' : 'Medium impact', value: '' },
            { label: idx % 2 === 0 ? 'Low effort' : 'Medium effort', value: '' }
          ]
        }))
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

  function inferOriginalArtifactId(originalSlide, zi, ai) {
    return originalSlide?.zones?.[zi]?.artifacts?.[ai]?._artifact_id || `s${originalSlide?.slide_number}_z${zi}_a${ai}`
  }

  function rectFromBlock(block) {
    if (!block) return null
    if ([block.x, block.y, block.w, block.h].every((value) => value != null)) {
      return { x: Number(block.x), y: Number(block.y), w: Number(block.w), h: Number(block.h) }
    }
    if ([block.x1, block.y1, block.x2, block.y2].every((value) => value != null)) {
      const x1 = Number(block.x1); const x2 = Number(block.x2)
      const y1 = Number(block.y1); const y2 = Number(block.y2)
      return { x: Math.min(x1, x2), y: Math.min(y1, y2), w: Math.abs(x2 - x1), h: Math.abs(y2 - y1) }
    }
    return null
  }

  function unionRects(rects) {
    const valid = (rects || []).filter((rect) => rect && [rect.x, rect.y, rect.w, rect.h].every((value) => Number.isFinite(Number(value))))
    if (!valid.length) return null
    const minX = Math.min(...valid.map((rect) => rect.x))
    const minY = Math.min(...valid.map((rect) => rect.y))
    const maxX = Math.max(...valid.map((rect) => rect.x + rect.w))
    const maxY = Math.max(...valid.map((rect) => rect.y + rect.h))
    return {
      x: Math.round(minX * 100) / 100,
      y: Math.round(minY * 100) / 100,
      w: Math.round(Math.max(0.1, maxX - minX) * 100) / 100,
      h: Math.round(Math.max(0.1, maxY - minY) * 100) / 100
    }
  }

  function collectOriginalArtifactGeometry(originalSlide, zi, ai) {
    const artifactId = inferOriginalArtifactId(originalSlide, zi, ai)
    const blocks = (originalSlide?.blocks || []).filter((block) => block.artifact_id === artifactId)
    const bodyBounds = unionRects(blocks.filter((block) => block.block_role === 'artifact_body').map(rectFromBlock))
    const allBounds = unionRects(blocks.map(rectFromBlock))
    const zoneArtifact = originalSlide?.zones?.[zi]?.artifacts?.[ai]
    const artifactBounds = bodyBounds || allBounds || (
      zoneArtifact && [zoneArtifact.x, zoneArtifact.y, zoneArtifact.w, zoneArtifact.h].every((value) => value != null)
        ? { x: zoneArtifact.x, y: zoneArtifact.y, w: zoneArtifact.w, h: zoneArtifact.h }
        : null
    )
    const headerBounds = unionRects(blocks.filter((block) => block.block_role === 'artifact_header').map(rectFromBlock))
    // allBounds covers header + body — the true footprint of this artifact in the Agent 5 output
    return { artifactId, artifactBounds, headerBounds, allBounds }
  }

  function applyOriginalGeometry(rebuiltSlide, originalSlide) {
    ;(rebuiltSlide?.zones || []).forEach((zone, zi) => {
      const originalZone = originalSlide?.zones?.[zi]
      if (originalZone?.frame) zone.frame = deepClone(originalZone.frame)
      ;(zone.artifacts || []).forEach((artifact, ai) => {
        const geometry = collectOriginalArtifactGeometry(originalSlide, zi, ai)
        if (geometry.artifactBounds) {
          artifact.x = geometry.artifactBounds.x
          artifact.y = geometry.artifactBounds.y
          artifact.w = geometry.artifactBounds.w
          artifact.h = geometry.artifactBounds.h
          if (artifact.type === 'cards' || artifact.type === 'workflow') artifact.container = deepClone(geometry.artifactBounds)
        }
        if (artifact.header_block && geometry.headerBounds) {
          artifact.header_block = {
            ...artifact.header_block,
            x: geometry.headerBounds.x,
            y: geometry.headerBounds.y,
            w: geometry.headerBounds.w,
            h: geometry.headerBounds.h
          }
        }
      })
    })
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

  // Returns the Agent 5 slide enhanced with _artifact_ids and blocks (generated if missing).
  // Result is cached in state.agent5SlidesEnhanced so geometry lookups always work.
  function getEnhancedAgent5Slide(slideNumber) {
    if (state.agent5SlidesEnhanced[slideNumber]) return state.agent5SlidesEnhanced[slideNumber]
    const original = state.agent5Slides.find((s) => s.slide_number === slideNumber)
    if (!original) return null
    const enhanced = deepClone(original)
    const bt = deepClone(state.brandTokens || {})
    // Stamp _artifact_id on every artifact so geometry lookups are consistent
    ;(enhanced.zones || []).forEach((zone, zi) => {
      ;(zone.artifacts || []).forEach((artifact, ai) => {
        if (!artifact._artifact_id) artifact._artifact_id = `s${enhanced.slide_number}_z${zi}_a${ai}`
      })
    })
    // Generate blocks from agent5 zones if the original output didn't include them
    if (!Array.isArray(enhanced.blocks) || !enhanced.blocks.length) {
      if (typeof computeArtifactInternals === 'function') computeArtifactInternals(enhanced.zones || [], enhanced.canvas || {}, bt)
      if (typeof normalizeArtifactHeaderBands === 'function') normalizeArtifactHeaderBands(enhanced.zones || [])
      enhanced.blocks = typeof flattenToBlocks === 'function'
        ? (typeof sanitizeBlocks === 'function'
            ? sanitizeBlocks(flattenToBlocks(enhanced, bt), enhanced)
            : (flattenToBlocks(enhanced, bt) || []))
        : []
    }
    state.agent5SlidesEnhanced[slideNumber] = enhanced
    return enhanced
  }

  // Clamps a single block's x/y/w/h so it sits entirely within bounds.
  function clampBlockToBounds(block, bounds) {
    if (!bounds) return block
    const bx = Number(bounds.x), by = Number(bounds.y)
    const bx2 = bx + Number(bounds.w), by2 = by + Number(bounds.h)
    const fields = []
    if (block.x != null && block.w != null) fields.push(['x', 'w', bx, bx2])
    if (block.y != null && block.h != null) fields.push(['y', 'h', by, by2])
    if (block.x1 != null && block.x2 != null) {
      block.x1 = Math.max(bx, Math.min(bx2, Number(block.x1)))
      block.x2 = Math.max(bx, Math.min(bx2, Number(block.x2)))
    }
    if (block.y1 != null && block.y2 != null) {
      block.y1 = Math.max(by, Math.min(by2, Number(block.y1)))
      block.y2 = Math.max(by, Math.min(by2, Number(block.y2)))
    }
    fields.forEach(([pos, size, lo, hi]) => {
      const start = Math.max(lo, Math.min(hi, Number(block[pos])))
      const end = Math.min(hi, start + Number(block[size]))
      block[pos] = round2(start)
      block[size] = round2(Math.max(0, end - start))
    })
    return block
  }

  // Runs the Agent 5 helper pipeline on a single artifact and returns its primitive blocks.
  // The artifact must already have x/y/w/h set to the target bounding box.
  // clampBounds (optional): every generated block is clamped to stay within this rect.
  function generateArtifactBlocks(artifact, artifactId, zoneFrame, canvas, bt, clampBounds) {
    let processed = deepClone(artifact)
    if (typeof buildSafeArtifactShell === 'function') {
      processed = buildSafeArtifactShell(processed, bt) || processed
    }
    processed._artifact_id = artifactId

    const tempZone = {
      zone_id: 'tz1',
      zone_role: 'primary',
      narrative_weight: 'primary',
      message_objective: artifact.artifact_coverage_hint || '',
      frame: deepClone(zoneFrame),
      artifacts: [processed]
    }

    // mergeContentIntoZones enriches the artifact shell with zone-aware styling
    if (typeof mergeContentIntoZones === 'function') {
      const merged = mergeContentIntoZones([deepClone(tempZone)], [{ artifacts: [deepClone(artifact)] }], bt)
      if (merged?.[0]?.artifacts?.[0]) {
        tempZone.artifacts = merged[0].artifacts
        tempZone.frame = merged[0].frame || tempZone.frame
      }
    }
    // Re-stamp the id — mergeContentIntoZones may have replaced the artifact object
    if (tempZone.artifacts[0]) tempZone.artifacts[0]._artifact_id = artifactId

    const tempSlide = {
      slide_number: 9999,
      slide_type: 'content',
      slide_archetype: 'dashboard',
      layout_mode: false,
      selected_layout_name: '',
      canvas: deepClone(canvas) || { width_in: 13.33, height_in: 7.5, margin: { left: 0.37, right: 0.37, top: 0.6, bottom: 0.37 } },
      title_block: null,
      subtitle_block: null,
      title: '',
      subtitle: '',
      key_message: '',
      global_elements: {},
      zones: [tempZone],
      blocks: []
    }

    if (typeof computeArtifactInternals === 'function') computeArtifactInternals(tempSlide.zones, tempSlide.canvas, bt)
    if (typeof normalizeArtifactHeaderBands === 'function') normalizeArtifactHeaderBands(tempSlide.zones)
    // Re-stamp again after internal layout calculations
    if (tempSlide.zones[0]?.artifacts?.[0]) tempSlide.zones[0].artifacts[0]._artifact_id = artifactId

    if (typeof flattenToBlocks !== 'function') return []
    const rawBlocks = flattenToBlocks(tempSlide, bt)
    const allBlocks = typeof sanitizeBlocks === 'function' ? sanitizeBlocks(rawBlocks || [], tempSlide) : (rawBlocks || [])

    // Strip any slide-level title/subtitle blocks — we only want the artifact's own blocks
    const artifactBlocks = allBlocks.filter((b) =>
      b.block_role !== 'slide_header' &&
      b.block_role !== 'slide_subheader' &&
      b.artifact_type !== 'slide' &&
      b.artifact_subtype !== 'title' &&
      b.artifact_subtype !== 'subtitle'
    )

    // Ensure every returned block is tagged with the correct artifact_id, then clamp
    artifactBlocks.forEach((b) => {
      b.artifact_id = artifactId
      if (clampBounds) clampBlockToBounds(b, clampBounds)
    })
    return artifactBlocks
  }

  function buildSlideFromSelections(slideNumber) {
    const manifestSlide = state.manifestSlides.find((slide) => slide.slide_number === slideNumber)
    const originalSlide = getEnhancedAgent5Slide(slideNumber)
    if (!manifestSlide || !originalSlide) return null

    const overrides = ensureSlideOverride(slideNumber)
    const bt = deepClone(state.brandTokens || {})

    // Start from the original Agent 5 output as the immutable base
    const rebuiltSlide = deepClone(originalSlide)
    if (!Array.isArray(rebuiltSlide.blocks)) rebuiltSlide.blocks = []

    // For each non-insight artifact, replace blocks only when the user changed its type/subtype
    ;(manifestSlide.zones || []).forEach((zone, zi) => {
      ;(zone.artifacts || []).forEach((artifact, ai) => {
        const manifestType = localNormalizeArtifactType(artifact?.type, artifact?.chart_type)
        if (manifestType === 'insight_text') return

        const key = `${zi}-${ai}`
        const selection = overrides[key]
        if (!selection) return

        // If still the same as the manifest default, keep original Agent 5 blocks untouched
        const manifestSubtype = localArtifactSubtype(artifact)
        if (selection.type === manifestType && selection.subtype === manifestSubtype) return

        // ── User changed this artifact — surgical replacement ──────────────────
        const artifactId = inferOriginalArtifactId(originalSlide, zi, ai)
        const geometry = collectOriginalArtifactGeometry(originalSlide, zi, ai)

        // allBounds = union of ALL blocks for this artifact_id in Agent 5 output
        // (header + body together). This is the true on-slide footprint we must stay within.
        // Fall back through: zone frame inner area → computed default.
        const zoneFrame = originalSlide?.zones?.[zi]?.frame ||
          fallbackZoneFrame(zi, manifestSlide.zones.length, originalSlide.canvas)
        const pad = zoneFrame.padding || {}
        const fallbackBounds = {
          x: round2((zoneFrame.x || 0) + (pad.left || 0)),
          y: round2((zoneFrame.y || 0) + (pad.top || 0)),
          w: round2(Math.max(0.5, (zoneFrame.w || 0) - (pad.left || 0) - (pad.right || 0))),
          h: round2(Math.max(0.5, (zoneFrame.h || 0) - (pad.top || 0) - (pad.bottom || 0)))
        }

        // Full footprint (incl. header) — used as the layout frame AND as the clamp fence
        const footprint = geometry.allBounds || geometry.artifactBounds || fallbackBounds
        // Body-only bounds (what the artifact itself occupies, excluding header)
        const bounds = geometry.artifactBounds || footprint

        // Use footprint as the zone frame passed to generateArtifactBlocks so
        // computeArtifactInternals places everything within the old artifact's slot
        const tightZoneFrame = { ...zoneFrame, x: footprint.x, y: footprint.y, w: footprint.w, h: footprint.h, padding: {} }

        // Build the new artifact data from Agent 4 (with sample data where needed)
        const newArtifact = buildConvertedManifestArtifact(selection.type, selection.subtype, artifact, zone, manifestSlide)
        // Stamp body bounds so the artifact itself starts within the old footprint
        newArtifact.x = bounds.x
        newArtifact.y = bounds.y
        newArtifact.w = bounds.w
        newArtifact.h = bounds.h
        if (selection.type === 'cards' || selection.type === 'workflow') newArtifact.container = deepClone(bounds)
        newArtifact._artifact_id = artifactId

        // Generate blocks, layout within tightZoneFrame, then clamp every block to footprint
        const newBlocks = generateArtifactBlocks(newArtifact, artifactId, tightZoneFrame, originalSlide.canvas, bt, footprint)

        // Swap out old artifact blocks for the new ones; everything else is unchanged
        rebuiltSlide.blocks = [
          ...rebuiltSlide.blocks.filter((b) => b.artifact_id !== artifactId),
          ...newBlocks
        ]

        // Update the zone artifact spec in the rebuilt slide
        if (rebuiltSlide.zones?.[zi]) {
          if (!Array.isArray(rebuiltSlide.zones[zi].artifacts)) rebuiltSlide.zones[zi].artifacts = []
          rebuiltSlide.zones[zi].artifacts[ai] = newArtifact
        }
      })
    })

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

  function formatRect(rect) {
    if (!rect) return 'unavailable'
    return `x:${rect.x}, y:${rect.y}, w:${rect.w}, h:${rect.h}`
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
        roundedRect(x, y, w, h, Number(block.corner_radius || 0) / 72 * Math.min(scaleX, scaleY))
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

  function renderArtifactControls() {
    const slide = state.manifestSlides.find((item) => item.slide_number === state.selectedSlideNumber)
    const overrides = ensureSlideOverride(state.selectedSlideNumber)
    const originalSlide = getEnhancedAgent5Slide(state.selectedSlideNumber)
    const artifactCards = []
    ;(slide?.zones || []).forEach((zone, zi) => {
      ;(zone.artifacts || []).forEach((artifact, ai) => {
        const currentType = localNormalizeArtifactType(artifact?.type, artifact?.chart_type)
        if (currentType === 'insight_text') return
        const key = `${zi}-${ai}`
        const selection = overrides[key]
        const currentSubtype = localArtifactSubtype(artifact)
        const typeOptions = ARTIFACT_CATALOG.map((entry) => entry.type)
        const subtypeOptions = catalogSubtypeOptions(selection?.type || currentType)
        const geometry = collectOriginalArtifactGeometry(originalSlide, zi, ai)
        artifactCards.push(`
          <div class="artifact-card">
            <div class="artifact-head">
              <div class="artifact-title-row">
                <div class="artifact-title">Artifact ${artifactCards.length + 1}</div>
                <div class="artifact-badge">${currentType}</div>
              </div>
              <div class="artifact-current">Agent 4 original: <strong>${currentType}</strong>${currentSubtype && currentSubtype !== currentType ? ` / ${currentSubtype}` : ''}</div>
              <div class="artifact-current">Agent 5 rendered bounds: <strong>${formatRect(geometry.artifactBounds)}</strong></div>
            </div>
            <div class="artifact-row">
              <div class="selector-grid">
                <select data-role="type" data-key="${key}">
                  ${selectorOptionsHtml(typeOptions, selection?.type || currentType)}
                </select>
                <select data-role="subtype" data-key="${key}">
                  ${selectorOptionsHtml(subtypeOptions, selection?.subtype || currentSubtype || subtypeOptions[0])}
                </select>
              </div>
            </div>
          </div>
        `)
      })
    })
    $('artifact-list').innerHTML = artifactCards.join('') || '<div class="artifact-card"><div class="artifact-head"><div class="artifact-copy">This slide has no swappable non-insight artifacts.</div></div></div>'

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
    renderArtifactControls()
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
    state.agent5SlidesEnhanced = {}
    state.brandTokens = {}
    state.selectedSlideNumber = null
    state.overrides = {}
    state.rebuiltSlides = {}
    state.outputObject = null
    $('results').classList.remove('show')
    $('slide-list').innerHTML = ''
    $('slide-meta').innerHTML = ''
    $('artifact-list').innerHTML = ''
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
