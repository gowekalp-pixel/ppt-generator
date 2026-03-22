// ─── AGENT 2 — BRAND GUIDELINE PARSER ────────────────────────────────────────
// Input:  state (brandB64, brandExt), brandContent from Agent 1
// Output: brandRulebook object (colors, fonts, layouts)
// Makes ONE Claude API call

const AGENT2_SYSTEM = `You are an expert brand designer and design systems consultant.
Extract ALL brand design rules from the provided content and return as a single valid JSON object.

Return exactly these fields:
{
  "primary_colors": ["#hexcode"],
  "secondary_colors": ["#hexcode"],
  "background_colors": ["#hexcode"],
  "title_font": { "family": "font name", "size": "32pt", "weight": "bold", "color": "#hexcode" },
  "body_font": { "family": "font name", "size": "18pt", "weight": "regular", "color": "#hexcode" },
  "caption_font": { "family": "font name", "size": "12pt", "weight": "regular", "color": "#hexcode" },
  "slide_width": 1280,
  "slide_height": 720,
  "logo_position": "top-right",
  "layout_patterns": ["title slide", "two-column", "bullet list"],
  "visual_style": "corporate",
  "spacing_notes": "0.5 inch margins",
  "chart_colors": ["#hexcode"]
}

Rules:
- All hex codes must include the # prefix
- If exact hex values are not visible, infer from color names (e.g. "navy blue" = #1A3C6E)
- Return ONLY valid JSON. No explanation. No markdown fences. No extra text.`

const AGENT2_FALLBACK = {
  primary_colors:    ['#1A3C6E'],
  secondary_colors:  ['#F4A300'],
  background_colors: ['#FFFFFF'],
  title_font:   { family: 'Calibri', size: '32pt', weight: 'bold',    color: '#1A3C6E' },
  body_font:    { family: 'Calibri', size: '18pt', weight: 'regular', color: '#333333' },
  caption_font: { family: 'Calibri', size: '12pt', weight: 'regular', color: '#666666' },
  slide_width:  1280,
  slide_height: 720,
  logo_position: 'top-right',
  layout_patterns: ['title slide', 'two-column', 'bullet list'],
  visual_style: 'corporate',
  spacing_notes: '0.5 inch margins, 0.3 inch between blocks',
  chart_colors: ['#1A3C6E', '#F4A300', '#22C55E']
}

async function runAgent2(state, brandContent) {
  console.log('Agent 2 starting — brand content type:', brandContent.slice(0, 20))

  let messages

  if (brandContent === '__PDF__') {
    messages = [{
      role: 'user',
      content: [
        { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: state.brandB64 } },
        { type: 'text', text: 'Extract all brand design rules from this PDF document. Return JSON only.' }
      ]
    }]
  } else if (brandContent === '__IMAGE__') {
    const mime = state.brandExt === 'png' ? 'image/png' : 'image/jpeg'
    messages = [{
      role: 'user',
      content: [
        { type: 'image', source: { type: 'base64', media_type: mime, data: state.brandB64 } },
        { type: 'text', text: 'Extract all brand design rules from this slide image. Return JSON only.' }
      ]
    }]
  } else {
    messages = [{
      role: 'user',
      content: 'Brand guideline content extracted from PPTX:\n\n' + brandContent + '\n\nReturn brand rules as JSON only.'
    }]
  }

  const raw    = await callClaude(AGENT2_SYSTEM, messages, 1000)
  const result = safeParseJSON(raw, AGENT2_FALLBACK)

  // Validate — make sure critical fields exist
  if (!result.primary_colors || !result.primary_colors.length) {
    console.warn('Agent 2 — missing primary_colors, using fallback')
    result.primary_colors = AGENT2_FALLBACK.primary_colors
  }
  if (!result.title_font || !result.title_font.family) {
    console.warn('Agent 2 — missing title_font, using fallback')
    result.title_font = AGENT2_FALLBACK.title_font
  }
  if (!result.body_font || !result.body_font.family) {
    console.warn('Agent 2 — missing body_font, using fallback')
    result.body_font = AGENT2_FALLBACK.body_font
  }

  return result
}
