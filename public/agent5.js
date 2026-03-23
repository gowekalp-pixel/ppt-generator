// ─── AGENT 5 — DESIGN DIRECTOR ────────────────────────────────────────────────
// Input:  state.slideManifest  — slide content from Agent 4
//         state.brandRulebook  — colors, fonts, layouts from Agent 2
//         state.outline        — presentationBrief from Agent 3
//                                (document_type, narrative_flow, tone)
//
// Output: finalSpec — flat JSON array, one fully specified slide object per slide
//
// Agent 5 thinks like a seasoned management consultant designer:
// - Numbers always get a visual (chart, stat, table)
// - Layout selection driven by presentation type AND content type
// - Brand rules applied precisely
// - Visual type can be overridden if Agent 4's choice doesn't fit the layout
// - Every slide passes the "10-second test" — message clear at a glance

// ─── PRESENTATION TYPE PROFILES ──────────────────────────────────────────────
// These drive default layout and visual preferences per document type

const PRESENTATION_PROFILES = {
  financial: {
    description: 'Financial report, P&L, balance sheet, cash flow, performance review',
    keywords: ['financial', 'finance', 'revenue', 'profit', 'loss', 'balance', 'cash', 'ebitda', 'margin', 'cost', 'budget', 'forecast'],
    layout_preferences: {
      data_table:      'data-dense layout with full content area — use for financial statements',
      stat_callout:    'headline layout — large numbers dominate, minimal text',
      chart_bar:       'content layout with chart area — comparisons across periods or segments',
      chart_line:      'content layout — trend over time, growth trajectory',
      chart_waterfall: 'content layout — P&L bridge, variance analysis, cost buildup',
      bullet_list:     'text layout — qualitative commentary, management discussion',
      three_column:    'three-column layout — three key drivers, three risks, three actions'
    },
    title_style: 'insight-led — title states the financial conclusion not the topic',
    default_visual: 'chart_bar',
    table_threshold: 'use table when 4+ line items need to be shown with multiple attributes',
    number_rule: 'ALL numbers must be visualised — no numbers sitting in bullet points'
  },

  strategic: {
    description: 'Strategy, market entry, growth plan, competitive positioning, transformation',
    keywords: ['strategy', 'strategic', 'growth', 'market', 'competitive', 'position', 'vision', 'transformation', 'initiative', 'roadmap'],
    layout_preferences: {
      three_column:  'three-column layout — three pillars, three options, three priorities',
      two_column:    'two-column layout — current vs future, problem vs solution',
      process_flow:  'full-width layout — roadmap, phased approach, implementation steps',
      quote_callout: 'full-width layout — governing thought, key strategic insight',
      icon_cards:    'card layout — strategic initiatives, capability areas',
      bullet_list:   'text layout — strategic rationale, implications, risks',
      stat_callout:  'headline layout — market size, opportunity, key metrics'
    },
    title_style: 'hypothesis-led — title states the strategic argument',
    default_visual: 'three_column',
    table_threshold: 'use table only for option comparison matrices or capability assessments',
    number_rule: 'Key metrics visualised as stat callouts — market size, growth rate, share'
  },

  market_research: {
    description: 'Market analysis, competitive landscape, customer research, industry trends',
    keywords: ['market', 'research', 'competition', 'competitor', 'customer', 'industry', 'trend', 'survey', 'segment', 'analysis'],
    layout_preferences: {
      chart_bar:    'content layout — market share, competitive comparison',
      chart_line:   'content layout — market growth trends',
      stat_callout: 'headline layout — market size, CAGR, penetration rate',
      two_column:   'two-column layout — competitive positioning, strengths vs weaknesses',
      bullet_list:  'text layout — market dynamics, customer insights, implications',
      data_table:   'content layout — competitive feature matrix, market sizing table'
    },
    title_style: 'finding-led — title states what the research found',
    default_visual: 'chart_bar',
    table_threshold: 'use table for competitive feature matrices or market sizing',
    number_rule: 'Market metrics always visualised — charts for trends, stats for headlines'
  },

  operational: {
    description: 'Operations review, process improvement, performance management, KPI tracking',
    keywords: ['operational', 'operations', 'process', 'efficiency', 'kpi', 'performance', 'metric', 'production', 'supply', 'quality'],
    layout_preferences: {
      data_table:   'content layout — KPI tracking table, operational metrics',
      chart_bar:    'content layout — performance vs target, departmental comparison',
      chart_line:   'content layout — trend monitoring, process metrics over time',
      process_flow: 'full-width layout — process maps, improvement steps',
      stat_callout: 'headline layout — key operational metrics, achievement vs target',
      two_column:   'two-column layout — issues vs actions, current vs improved state',
      bullet_list:  'text layout — root cause analysis, recommendations'
    },
    title_style: 'fact-led — title states the operational finding or status',
    default_visual: 'data_table',
    table_threshold: 'tables widely used — operational data is naturally tabular',
    number_rule: 'All KPIs and metrics visualised — charts for trends, tables for dashboards'
  },

  general: {
    description: 'General business presentation, mixed content',
    keywords: [],
    layout_preferences: {},
    title_style: 'insight-led',
    default_visual: 'bullet_list',
    table_threshold: 'use table when structure is genuinely needed',
    number_rule: 'Numbers should be visualised where possible'
  }
}


// ─── DETECT PRESENTATION TYPE ─────────────────────────────────────────────────
function detectPresentationType(brief) {
  if (!brief) return 'general'

  const docType = (brief.document_type || '').toLowerCase()
  const flow    = (brief.narrative_flow || '').toLowerCase()
  const combined = docType + ' ' + flow

  for (const [type, profile] of Object.entries(PRESENTATION_PROFILES)) {
    if (type === 'general') continue
    for (const keyword of profile.keywords) {
      if (combined.includes(keyword)) {
        console.log('Agent 5 — detected presentation type:', type, '(matched keyword:', keyword + ')')
        return type
      }
    }
  }

  return 'general'
}


// ─── SELECT BEST LAYOUT ───────────────────────────────────────────────────────
function selectLayout(slide, availableLayouts, presentationType, profile) {
  if (!availableLayouts || availableLayouts.length === 0) return null

  const vt        = (slide.visual_type || '').toLowerCase()
  const slideType = (slide.slide_type  || '').toLowerCase()
  const sType     = (slide.section_type|| '').toLowerCase()

  // Title slides — always use "Title slide" layout
  if (slideType === 'title') {
    return availableLayouts.find(l =>
      l.name.toLowerCase().includes('title slide') ||
      l.type === 'title'
    ) || availableLayouts[0]
  }

  // Divider slides — always use "Section divider" layout
  if (slideType === 'divider') {
    return availableLayouts.find(l =>
      l.name.toLowerCase().includes('section') ||
      l.name.toLowerCase().includes('divider') ||
      l.type === 'obj'
    ) || availableLayouts[1]
  }

  // Content slides — match visual type to layout structure
  // Three-column visuals need a 3-column layout
  if (['three_column', 'icon_cards'].includes(vt)) {
    const found = availableLayouts.find(l =>
      l.name.toLowerCase().includes('3 across') ||
      l.name.toLowerCase().includes('three') ||
      (l.structure || '').toLowerCase().includes('3-column')
    )
    if (found) return found
  }

  // Two-column visuals
  if (['two_column', 'chart_bar', 'chart_line', 'chart_waterfall'].includes(vt)) {
    const found = availableLayouts.find(l =>
      l.name.toLowerCase().includes('2 across') ||
      l.name.toLowerCase().includes('1 across') ||
      (l.structure || '').toLowerCase().includes('column')
    )
    if (found) return found
  }

  // Data tables need full-width layout
  if (vt === 'data_table') {
    const found = availableLayouts.find(l =>
      l.name.toLowerCase().includes('body text') ||
      l.name.toLowerCase().includes('1 across') ||
      (l.ph_count || 0) >= 2
    )
    if (found) return found
  }

  // Stat callout, quote — use single wide layout
  if (['stat_callout', 'quote_callout'].includes(vt)) {
    const found = availableLayouts.find(l =>
      l.name.toLowerCase().includes('1 across') ||
      l.name.toLowerCase().includes('body text') ||
      l.name.toLowerCase().includes('title only')
    )
    if (found) return found
  }

  // Process flow — needs full-width
  if (vt === 'process_flow') {
    const found = availableLayouts.find(l =>
      l.name.toLowerCase().includes('1 across') ||
      l.name.toLowerCase().includes('body text')
    )
    if (found) return found
  }

  // Default — body text layout
  return availableLayouts.find(l =>
    l.name.toLowerCase().includes('body text') ||
    l.name.toLowerCase().includes('1 across') ||
    l.type === 'obj'
  ) || availableLayouts.find(l => !['title slide', 'section divider', 'blank'].some(x => l.name.toLowerCase().includes(x)))
  || availableLayouts[0]
}


// ─── REVIEW AND OVERRIDE VISUAL TYPE ─────────────────────────────────────────
// Agent 5 acts as senior reviewer — can override Agent 4's visual choice

function reviewVisualType(slide, profile, presentationType) {
  const vt      = (slide.visual_type || '').toLowerCase()
  const content = slide.content || {}
  const st      = (slide.slide_type || '').toLowerCase()

  if (st === 'title')   return 'title_slide'
  if (st === 'divider') return 'divider_slide'

  // Table override — only if genuinely needed
  if (vt === 'data_table') {
    const rows = content.rows || []
    if (rows.length < 3 && (content.headers || []).length < 3) {
      console.log('Agent 5 — overriding data_table (too small) →', slide.visual_type, '→ stat_callout')
      return 'stat_callout'
    }
  }

  // Numbers in bullet list — override to stat_callout if financial
  if (vt === 'bullet_list' && presentationType === 'financial') {
    const bullets = content.bullets || []
    const hasNumbers = bullets.some(b => /\d/.test(b))
    if (hasNumbers && bullets.length <= 3) {
      console.log('Agent 5 — overriding bullet_list with numbers → stat_callout')
      return 'stat_callout'
    }
  }

  // Three_column with < 3 columns — downgrade
  if (vt === 'three_column') {
    const cols = content.columns || []
    if (cols.length < 3) {
      console.log('Agent 5 — overriding three_column (only', cols.length, 'columns) → bullet_list')
      return 'bullet_list'
    }
  }

  // Process flow with < 2 steps — downgrade
  if (vt === 'process_flow') {
    const steps = content.steps || []
    if (steps.length < 2) {
      return 'bullet_list'
    }
  }

  return vt
}


// ─── BUILD BRAND SPEC FOR A SLIDE ────────────────────────────────────────────
function buildBrandSpec(slide, brand, presentationType) {
  const slideType = (slide.slide_type || '').toLowerCase()
  const isDark    = slideType === 'title' || slideType === 'divider'

  const primary   = (brand.primary_colors    || ['#0F2FB5'])[0]
  const secondary = (brand.secondary_colors  || ['#FF8E00'])[0]
  const bgColor   = (brand.background_colors || ['#FFFFFF'])[0]
  const textColor = (brand.text_colors       || ['#000000'])[0]
  const titleFont = (brand.title_font || {}).family || 'Arial'
  const bodyFont  = (brand.body_font  || {}).family || 'Arial'

  // Financial presentations — slightly tighter font sizes for data density
  const titleSize = presentationType === 'financial' ? '22pt' : '24pt'
  const bodySize  = presentationType === 'financial' ? '12pt' : '14pt'
  const dataSize  = '11pt'

  return {
    background_color: isDark ? primary  : bgColor,
    title_color:      isDark ? '#FFFFFF': primary,
    title_font:       titleFont,
    title_size:       titleSize,
    body_font:        bodyFont,
    body_size:        bodySize,
    body_color:       isDark ? '#FFFFFF': textColor,
    accent_color:     secondary,
    data_font_size:   dataSize,
    chart_colors:     brand.chart_colors || [primary, secondary],
    all_colors:       brand.all_colors   || {}
  }
}


// ─── MAIN AGENT 5 SYSTEM PROMPT ──────────────────────────────────────────────
const AGENT5_SYSTEM = `You are a Design Director at a top-tier management consulting firm.
You have built hundreds of board-level presentations across financial, strategic, market research and operational contexts.

You will receive:
1. A slide manifest (from the content writer) — what each slide says
2. A brand rulebook — colors, fonts, available slide layouts
3. A presentation brief — document type, narrative flow, governing thought

Your job is to finalize the design specification for every slide.

You must return a flat JSON array — one object per slide — with EXACTLY these fields per slide:

{
  "slide_number": number,
  "slide_type": "title" | "divider" | "content",
  "section_name": "string",
  "section_type": "string",
  "layout_name": "string — exact name of the layout from available layouts",
  "title": "string — insight-led title, not topic title",
  "subtitle": "string — for title slides only",
  "key_message": "string — the single takeaway",
  "visual_type": "string — your final visual type decision",
  "content": { ... same structure as input, refined if needed ... },
  "brand": {
    "background_color": "#hex",
    "title_color": "#hex",
    "title_font": "font name",
    "title_size": "pt",
    "body_font": "font name",
    "body_size": "pt",
    "body_color": "#hex",
    "accent_color": "#hex",
    "data_font_size": "pt",
    "chart_colors": ["#hex", "#hex"]
  },
  "design_notes": "string — brief note on why this visual/layout was chosen",
  "speaker_note": "string"
}

DESIGN RULES — follow these strictly:

NUMBERS:
- ALL numbers must be visualised — charts, stat callouts, or tables
- Never leave numbers sitting in bullet points
- Financial data → chart_bar (comparisons), chart_line (trends), chart_waterfall (bridges), stat_callout (headlines), data_table (statements)
- Market data → chart_bar or stat_callout
- Operational KPIs → stat_callout or data_table

TABLES:
- Only use data_table when: (a) 4+ items need comparison across 3+ attributes, OR (b) it is a financial statement
- For 2-3 items or simple lists → use bullet_list or stat_callout instead
- Maximum 6 rows for readability on a single slide

LAYOUTS:
- Match layout to content complexity — use the exact layout names from the available list
- Title slides → "Title slide" layout only
- Divider slides → "Section divider" layout only  
- Three-column content → layout with 3-column structure
- Full-width text or single chart → single content area layout

TITLES:
- Every content slide title must state the INSIGHT, not the topic
- Wrong: "Revenue Analysis" | Right: "Revenue grew 18% driven by Product X"
- Wrong: "Market Overview" | Right: "Market growing at 22% CAGR with significant headroom"

PRESENTATION TYPE IMPACT:
- Financial → data-dense, precise, evidence-first. Charts and tables dominate. Tight layout.
- Strategic → flow-based, persuasive, argument-first. Three-columns, process flows, quote callouts.
- Market research → finding-led, mixed visual. Charts for data, text for implications.
- Operational → fact-based, action-oriented. Tables for KPIs, process flows for steps.

BRAND:
- Title and divider slides: primary brand color background, white text
- Content slides: white background, primary color for titles only
- Accent/secondary color: use only for highlights, key numbers, callout borders — sparingly
- Never use more than 2 colors on a single content slide

Return ONLY a valid JSON array. No explanation. No markdown fences.`


// ─── MAIN RUNNER ─────────────────────────────────────────────────────────────
async function runAgent5(state) {
  const manifest  = state.slideManifest
  const brand     = state.brandRulebook
  const brief     = state.outline

  console.log('Agent 5 starting')
  console.log('  Slides to design:', manifest.length)
  console.log('  Available layouts:', (brand.slide_layouts || []).length)

  // Detect presentation type
  const presentationType = detectPresentationType(brief)
  const profile          = PRESENTATION_PROFILES[presentationType] || PRESENTATION_PROFILES.general
  console.log('  Presentation type:', presentationType)
  console.log('  Profile:', profile.description)

  // Pre-process — review visual types before sending to Claude
  const reviewedManifest = manifest.map(slide => ({
    ...slide,
    visual_type: reviewVisualType(slide, profile, presentationType)
  }))

  const overrides = manifest.filter((s,i) => s.visual_type !== reviewedManifest[i].visual_type)
  if (overrides.length > 0) {
    console.log('Agent 5 — visual type overrides applied:', overrides.length, 'slides')
  }

  // Build layout list for Claude — names and structures only (no placeholder coords)
  const layoutSummary = (brand.slide_layouts || []).map(l => ({
    name:           l.name,
    structure:      l.structure,
    usage_guidance: l.usage_guidance || '',
    ph_count:       l.ph_count
  }))

  const messages = [{
    role: 'user',
    content: `PRESENTATION TYPE: ${presentationType.toUpperCase()}
Profile: ${profile.description}
Number rule: ${profile.number_rule}
Table rule: ${profile.table_threshold}

PRESENTATION BRIEF SUMMARY:
- Document type: ${(brief || {}).document_type || '—'}
- Governing thought: ${(brief || {}).governing_thought || '—'}
- Narrative flow: ${(brief || {}).narrative_flow || '—'}
- Tone: ${(brief || {}).tone || 'professional'}
- Data heavy: ${(brief || {}).data_heavy ? 'yes' : 'no'}

BRAND RULEBOOK:
- Primary colors: ${(brand.primary_colors || []).join(', ')}
- Secondary colors: ${(brand.secondary_colors || []).join(', ')}
- Background: ${(brand.background_colors || ['#FFFFFF'])[0]}
- Title font: ${(brand.title_font || {}).family || 'Arial'}
- Body font: ${(brand.body_font || {}).family || 'Arial'}
- Chart colors: ${(brand.chart_colors || []).join(', ')}

AVAILABLE SLIDE LAYOUTS (use EXACT names):
${JSON.stringify(layoutSummary, null, 2)}

SLIDE MANIFEST (${reviewedManifest.length} slides — finalize all of them):
${JSON.stringify(reviewedManifest, null, 2)}

Produce the final design specification for all ${reviewedManifest.length} slides.
Apply brand rules, select the best layout from the available list, finalize visual types.
Return ONLY the JSON array.`
  }]

  console.log('Agent 5 — calling Claude for final design pass...')
  const raw    = await callClaude(AGENT5_SYSTEM, messages, 4000)
  console.log('Agent 5 — response length:', raw.length, 'chars')

  let finalSpec = safeParseJSON(raw, null)

  // ── Validate ──────────────────────────────────────────────────────────────
  if (!Array.isArray(finalSpec) || finalSpec.length === 0) {
    console.warn('Agent 5 — Claude parse failed, applying brand rules manually')
    return applyBrandManually(reviewedManifest, brand, brief, presentationType, profile)
  }

  if (finalSpec.length < manifest.length) {
    console.warn('Agent 5 — got', finalSpec.length, 'slides, expected', manifest.length, '— using manual fallback')
    return applyBrandManually(reviewedManifest, brand, brief, presentationType, profile)
  }

  // ── Enrich — ensure every slide has brand spec ────────────────────────────
  finalSpec = finalSpec.map((slide, i) => {
    const source = reviewedManifest[i] || {}

    // If Claude didn't include brand block — build it
    if (!slide.brand || !slide.brand.background_color) {
      slide.brand = buildBrandSpec(source, brand, presentationType)
    }

    // If Claude didn't select a layout — select one
    if (!slide.layout_name) {
      const layout = selectLayout(source, brand.slide_layouts || [], presentationType, profile)
      slide.layout_name = layout ? layout.name : 'Body text KM'
    }

    // Preserve content from Agent 4 if Claude lost it
    if (!slide.content || Object.keys(slide.content).length === 0) {
      slide.content = source.content || {}
    }

    return slide
  })

  // ── Log summary ───────────────────────────────────────────────────────────
  const layoutsUsed = [...new Set(finalSpec.map(s => s.layout_name).filter(Boolean))]
  const vtUsed      = [...new Set(finalSpec.map(s => s.visual_type).filter(Boolean))]

  console.log('Agent 5 complete')
  console.log('  Final slides:', finalSpec.length)
  console.log('  Layouts used:', layoutsUsed.join(', '))
  console.log('  Visual types:', vtUsed.join(', '))
  console.log('  Presentation type applied:', presentationType)

  return finalSpec
}


// ─── MANUAL BRAND APPLICATION FALLBACK ───────────────────────────────────────
// Used when Claude's response cannot be parsed

function applyBrandManually(manifest, brand, brief, presentationType, profile) {
  console.log('Agent 5 — applying brand manually to', manifest.length, 'slides')

  return manifest.map(slide => {
    const finalVT  = reviewVisualType(slide, profile, presentationType)
    const layout   = selectLayout(slide, brand.slide_layouts || [], presentationType, profile)
    const brandSpec = buildBrandSpec(slide, brand, presentationType)

    return {
      slide_number:  slide.slide_number,
      slide_type:    slide.slide_type,
      section_name:  slide.section_name,
      section_type:  slide.section_type,
      layout_name:   layout ? layout.name : 'Body text KM',
      title:         slide.title,
      subtitle:      slide.subtitle || '',
      key_message:   slide.key_message || '',
      visual_type:   finalVT,
      content:       slide.content || {},
      brand:         brandSpec,
      design_notes:  'Brand applied manually — ' + presentationType + ' presentation type, ' + finalVT + ' visual',
      speaker_note:  slide.speaker_note || ''
    }
  })
}
