// ─── AGENT 3 — CONTENT ANALYST & STRUCTURE BUILDER ───────────────────────────
// Input:  state (contentB64, slideCount)
// Output: outline string (numbered slide list)
// Makes ONE Claude API call — reads the content PDF and creates structure

async function runAgent3(state) {
  console.log('Agent 3 starting — slide count:', state.slideCount)

  const system = `You are a management consultant structuring a board-level presentation for senior leadership.

Read the provided content carefully and create a logical presentation outline.

Rules:
- Total slides must be EXACTLY ${state.slideCount} slides — no more, no less
- Slide 1 is always a Title slide
- Include 2-3 section divider slides between major topics
- Last slide is always Next Steps or Conclusion
- Each slide gets a clear, specific title (not generic)
- Think about narrative flow: context → problem → insight → recommendation → next steps

Return ONLY a numbered list. Each line must follow this EXACT format:
Slide N: [Title] — [one line description of what this slide covers]

Example:
Slide 1: India MSME Lending Landscape — Title slide with presentation name and date
Slide 2: Executive Summary — Three key takeaways from the analysis
Slide 3: [Section] Market Context — Divider slide introducing market section`

  const messages = [{
    role: 'user',
    content: [
      {
        type: 'document',
        source: { type: 'base64', media_type: 'application/pdf', data: state.contentB64 }
      },
      {
        type: 'text',
        text: `Read this document carefully. Create a ${state.slideCount} slide presentation outline for senior management. Follow the exact format specified. Make titles specific to the actual content, not generic.`
      }
    ]
  }]

  const outline = await callClaude(system, messages, 2000)

  // Validate slide count
  const slideLines = outline.match(/^Slide \d+:/gm) || []
  console.log('Agent 3 — slides in outline:', slideLines.length, 'requested:', state.slideCount)

  if (slideLines.length === 0) {
    throw new Error('Agent 3 returned an empty outline. Check your content PDF.')
  }

  return outline
}
