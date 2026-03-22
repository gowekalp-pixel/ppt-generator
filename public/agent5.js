// ─── AGENT 5 — BRAND + CONTENT MERGER ────────────────────────────────────────
// Input:  state.slideManifest from Agent 4, state.brandRulebook from Agent 2
// Output: finalSpec — JSON array with full design spec per slide
// Makes ONE Claude API call

const AGENT5_SYSTEM = `You are a senior presentation designer.

You will receive slide content and brand rules. Your job is to produce a final design specification for every slide.

Return a JSON array where each object has EXACTLY these fields:
{
  "slide_number": 1,
  "type": "title",
  "title": "slide title",
  "subtitle": "subtitle if applicable",
  "bullets": ["bullet 1", "bullet 2"],
  "visual_type": "text",
  "background_color": "#FFFFFF",
  "title_color": "#1A3C6E",
  "title_font": "Calibri",
  "title_size": "32pt",
  "body_font": "Calibri",
  "body_size": "18pt",
  "body_color": "#333333",
  "accent_color": "#F4A300",
  "speaker_note": "presenter note"
}

Rules:
- Use ONLY colors from the brand rules — do not invent new colors
- Title slides: background = primary brand color, text = white
- Divider slides: background = primary brand color, text = white  
- Content slides: background = background brand color (usually white), title = primary color
- accent_color = secondary brand color for highlights, bullets, icons
- Keep ALL titles and bullets from the slide content — do not shorten or remove any
- slide_number, type, title, bullets must match the input exactly
- Return ONLY a valid JSON array. No explanation. No markdown fences.`

async function runAgent5(state) {
  console.log('Agent 5 starting — merging', state.slideManifest.length, 'slides with brand rules')

  const messages = [{
    role: 'user',
    content: `BRAND RULES:\n${JSON.stringify(state.brandRulebook, null, 2)}\n\nSLIDE CONTENT (${state.slideManifest.length} slides):\n${JSON.stringify(state.slideManifest, null, 2)}\n\nProduce the final branded slide specification as a JSON array. Include ALL ${state.slideManifest.length} slides.`
  }]

  const raw = await callClaude(AGENT5_SYSTEM, messages, 4000)

  const result = safeParseJSON(raw, state.slideManifest)

  if (!Array.isArray(result) || result.length === 0) {
    console.warn('Agent 5 — parse failed, using Agent 4 manifest with brand colors applied')
    return applyBrandFallback(state.slideManifest, state.brandRulebook)
  }

  // If Claude returned fewer slides than we have, fall back
  if (result.length < state.slideManifest.length) {
    console.warn('Agent 5 — got', result.length, 'slides but expected', state.slideManifest.length)
    console.warn('Agent 5 — applying brand colors manually to full manifest')
    return applyBrandFallback(state.slideManifest, state.brandRulebook)
  }

  console.log('Agent 5 — final spec ready:', result.length, 'slides')
  return result
}

// ─── FALLBACK: Apply brand colors manually if Agent 5 Claude call fails ───────
function applyBrandFallback(slideManifest, brand) {
  const primary    = (brand.primary_colors    || ['#1A3C6E'])[0]
  const secondary  = (brand.secondary_colors  || ['#F4A300'])[0]
  const bgColor    = (brand.background_colors || ['#FFFFFF'])[0]
  const titleFont  = (brand.title_font  || {}).family || 'Calibri'
  const titleSize  = (brand.title_font  || {}).size   || '32pt'
  const bodyFont   = (brand.body_font   || {}).family || 'Calibri'
  const bodySize   = (brand.body_font   || {}).size   || '18pt'
  const bodyColor  = (brand.body_font   || {}).color  || '#333333'

  return slideManifest.map(slide => ({
    slide_number:     slide.slide_number,
    type:             slide.type,
    title:            slide.title,
    subtitle:         slide.subtitle     || '',
    bullets:          slide.bullets      || [],
    visual_type:      slide.visual_type  || 'text',
    background_color: (slide.type === 'title' || slide.type === 'divider') ? primary : bgColor,
    title_color:      (slide.type === 'title' || slide.type === 'divider') ? '#FFFFFF' : primary,
    title_font:       titleFont,
    title_size:       titleSize,
    body_font:        bodyFont,
    body_size:        bodySize,
    body_color:       bodyColor,
    accent_color:     secondary,
    speaker_note:     slide.speaker_note || ''
  }))
}
