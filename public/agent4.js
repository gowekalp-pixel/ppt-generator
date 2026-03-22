// ─── AGENT 4 — DETAILED SLIDE CONTENT WRITER ─────────────────────────────────
// Input:  state.outline from Agent 3, state.slideCount
// Output: slideManifest — JSON array, one object per slide
// Makes ONE Claude API call

const AGENT4_SYSTEM = `You are a senior business writer creating content for a board-level presentation.

You will receive a slide outline. Expand EVERY slide into full content.

Return a JSON array where each object has EXACTLY these fields:
{
  "slide_number": 1,
  "type": "title",
  "title": "exact slide title",
  "subtitle": "subtitle text for title slides, empty string for others",
  "bullets": ["bullet 1", "bullet 2", "bullet 3"],
  "visual_type": "text",
  "speaker_note": "one sentence for the presenter"
}

Rules for "type" field:
- Use "title" for the first slide only
- Use "divider" for section separator slides (they get no bullets)
- Use "content" for all other slides

Rules for "visual_type" field — pick the most appropriate:
- "text" — simple bullet list (default)
- "three-column" — exactly 3 parallel points
- "two-column" — two groups of related points
- "table" — comparative or structured data (format bullets as "Label: Value")
- "quote" — a single key insight or statistic as a callout
- "icons" — 3 distinct categories or features

Rules for "bullets" field:
- Title slides: empty array []
- Divider slides: empty array []
- Content slides: 3 to 5 bullets, each a complete sentence, specific and data-driven
- Each bullet must be meaningful — no generic filler

IMPORTANT:
- Return ONLY a valid JSON array
- No explanation before or after
- No markdown fences
- Every slide from the outline must appear in the output
- slide_number must match the outline exactly`

async function runAgent4(state) {
  console.log('Agent 4 starting — expanding', state.slideCount, 'slides')

  const messages = [{
    role: 'user',
    content: `Expand this presentation outline into full slide content. Return a JSON array with one object per slide. Every slide must be included.\n\nOUTLINE:\n${state.outline}`
  }]

  const raw = await callClaude(AGENT4_SYSTEM, messages, 3000)

  const fallback = [{
    slide_number: 1,
    type:         'title',
    title:        'Presentation',
    subtitle:     '',
    bullets:      [],
    visual_type:  'text',
    speaker_note: 'Opening slide'
  }]

  const result = safeParseJSON(raw, fallback)

  // Validate it's an array with content
  if (!Array.isArray(result)) {
    console.error('Agent 4 — result is not an array:', typeof result)
    return fallback
  }

  if (result.length === 0) {
    console.error('Agent 4 — empty array returned')
    return fallback
  }

  console.log('Agent 4 — slides generated:', result.length)

  // Ensure every slide has required fields
  return result.map((slide, i) => ({
    slide_number: slide.slide_number || (i + 1),
    type:         slide.type         || 'content',
    title:        slide.title        || 'Slide ' + (i + 1),
    subtitle:     slide.subtitle     || '',
    bullets:      Array.isArray(slide.bullets) ? slide.bullets : [],
    visual_type:  slide.visual_type  || 'text',
    speaker_note: slide.speaker_note || ''
  }))
}
