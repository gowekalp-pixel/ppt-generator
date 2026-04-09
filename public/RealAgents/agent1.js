// ─── AGENT 1 — INPUT PACKAGER ─────────────────────────────────────────────────
// Input:  state (brandFile, brandB64, brandExt, contentB64)
// Output: brandContent string (text extracted from PPTX, or flag for PDF/image)
// No Claude API call — pure file handling

async function runAgent1(state) {
  console.log('Agent 1 starting — file type:', state.brandExt)

  let brandContent = ''

  if (state.brandExt === 'pptx' || state.brandExt === 'ppt') {
    try {
      brandContent = await extractPptxText(state.brandB64)
      if (!brandContent || brandContent.length < 20) {
        throw new Error('Extracted text too short — likely an image-based PPTX')
      }
      console.log('Agent 1 — PPTX text extracted:', brandContent.length, 'chars')
    } catch (e) {
      console.warn('Agent 1 — PPTX extraction failed, using filename fallback:', e.message)
      brandContent = 'Brand PPTX file: ' + state.brandFile.name +
        '. Apply standard corporate brand guidelines with professional colors and fonts.'
    }
  } else if (state.brandExt === 'pdf') {
    brandContent = '__PDF__'
    console.log('Agent 1 — PDF brand file detected, will send to Claude vision')
  } else if (['png', 'jpg', 'jpeg'].includes(state.brandExt)) {
    brandContent = '__IMAGE__'
    console.log('Agent 1 — Image brand file detected, will send to Claude vision')
  } else {
    brandContent = '__PDF__'
    console.warn('Agent 1 — Unknown file type, treating as PDF')
  }

  console.log('Agent 1 complete — brandContent type:', brandContent.slice(0, 30))
  return brandContent
}

// ─── PPTX TEXT EXTRACTOR (uses JSZip loaded via CDN) ─────────────────────────
async function extractPptxText(b64) {
  const binary = atob(b64)
  const bytes  = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)

  const zip = await JSZip.loadAsync(bytes.buffer)

  // Get slide files sorted by number
  const slideFiles = Object.keys(zip.files)
    .filter(n => /^ppt\/slides\/slide\d+\.xml$/.test(n))
    .sort((a, b) => {
      return parseInt(a.match(/\d+/)[0]) - parseInt(b.match(/\d+/)[0])
    })

  let extracted = ''

  // Extract text from each slide
  for (const sf of slideFiles.slice(0, 12)) {
    const xml   = await zip.files[sf].async('string')
    const texts = [...xml.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)]
      .map(m => m[1].trim()).filter(Boolean)
    if (texts.length) extracted += '\n[' + sf + ']\n' + texts.join('\n')
  }

  // Extract theme colors
  const themeFile = zip.files['ppt/theme/theme1.xml']
  if (themeFile) {
    const themeXml  = await themeFile.async('string')
    const hexColors = [...new Set(
      [...themeXml.matchAll(/val="([0-9A-Fa-f]{6})"/g)].map(m => '#' + m[1])
    )].slice(0, 20)
    if (hexColors.length) extracted += '\n\n[Theme colors]\n' + hexColors.join(', ')
  }

  // Extract font names from layouts
  const layoutFiles = Object.keys(zip.files)
    .filter(n => /ppt\/slideLayouts\/slideLayout\d+\.xml/.test(n))
    .slice(0, 3)
  for (const lf of layoutFiles) {
    const xml   = await zip.files[lf].async('string')
    const fonts = [...new Set(
      [...xml.matchAll(/typeface="([^"+][^"]*)"/g)].map(m => m[1])
    )]
    if (fonts.length) extracted += '\n[Fonts] ' + fonts.join(', ')
  }

  return extracted.trim().slice(0, 9000)
}
