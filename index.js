require('dotenv').config()
const express = require('express')
const fs   = require('fs')
const path = require('path')
const app = express()

app.use(express.json({ limit: '30mb' }))
app.use(express.static('public'))

// Root redirect — public/index.html no longer exists; main UI is in Test-Frontend/
app.get('/', (_req, res) => res.redirect('/Test-Frontend/'))

// Wire Vercel serverless handlers as local Express routes
const claudeHandler = require('./api/claude')
app.post('/api/claude', claudeHandler)
app.post('/api/extract-brand', (_req, res) => {
  res.status(501).json({
    error: 'The local Express dev server does not host /api/extract-brand. Run the app via `vercel dev` so the Python endpoint is available.'
  })
})
app.post('/api/generate-pptx', (_req, res) => {
  res.status(501).json({
    error: 'The local Express dev server does not host /api/generate-pptx. Run the app via `vercel dev` so the Python endpoint is available.'
  })
})

// ─── INJECT NEW ARTIFACT TYPE ─────────────────────────────────────────────────
// Patches change-management replica files. Never touches production prompts/.
// Targets:
//   change-management/prompts/agent4/P3-ArtifactSelection.js  (modular replica)
//   change-management/prompts/agent4/A1-ArtifactSchema.js     (modular replica)
//   change-management/agent4-R.js                             (monolithic replica)
// Body: { artifact_type, display_name, description, p3:{...}, a1:{...} }
app.post('/api/inject-artifact', (req, res) => {
  const def = req.body
  if (!def || !def.artifact_type || !def.p3 || !def.a1) {
    return res.status(400).json({ error: 'Missing required fields: artifact_type, p3, a1' })
  }

  const CM = path.join(__dirname, 'public', 'change-management')
  const a4Path = path.join(CM, 'agent4-R.js')
  const p3Path = path.join(CM, 'prompts', 'agent4', 'P3-ArtifactSelection.js')
  const a1Path = path.join(CM, 'prompts', 'agent4', 'A1-ArtifactSchema.js')

  let a4, p3, a1
  try {
    a4 = fs.readFileSync(a4Path, 'utf8')
    p3 = fs.readFileSync(p3Path, 'utf8')
    a1 = fs.readFileSync(a1Path, 'utf8')
  } catch (err) {
    return res.status(500).json({ error: 'Could not read replica files: ' + err.message })
  }

  try {
    fs.writeFileSync(a4Path + '.bak', a4, 'utf8')
    fs.writeFileSync(p3Path + '.bak', p3, 'utf8')
    fs.writeFileSync(a1Path + '.bak', a1, 'utf8')
  } catch (_) { /* non-fatal */ }

  const steps = {}
  const type  = def.artifact_type
  const p3d   = def.p3
  const a1d   = def.a1

  // Guard: skip if already present
  if (p3.includes(`  ${type}\n`) || p3.includes(`  ${type}\r\n`)) {
    return res.status(409).json({ error: `Artifact "${type}" already exists in P3-ArtifactSelection.js.` })
  }

  // ════════════════════════════════════════════════════════════════════════════
  // P3-ArtifactSelection.js replica  (Unix \n)
  // ════════════════════════════════════════════════════════════════════════════

  // P3-1. Add to AVAILABLE ARTIFACT TYPES list (after risk_register, before matrix)
  const p3_type_anchor = '  risk_register\n  matrix'
  if (p3.includes(p3_type_anchor)) {
    p3 = p3.replace(p3_type_anchor, `  risk_register\n  ${type}\n  matrix`)
    steps.p3_type = true
  } else {
    steps.p3_type = false
  }

  // P3-2. Add to PRIMARY zone permitted list (after risk_register, before cards)
  if (p3d.can_be_primary) {
    const p3_primary_anchor = '    - risk_register\n    - cards'
    if (p3.includes(p3_primary_anchor)) {
      p3 = p3.replace(p3_primary_anchor, `    - risk_register\n    - ${type}\n    - cards`)
      steps.p3_primary = true
    } else {
      steps.p3_primary = false
    }
  } else {
    steps.p3_primary = true  // not applicable — skipped
  }

  // P3-3. Add pairing rule (after risk_register pairing, before matrix pairing)
  const p3_pairing_anchor = '  risk_register                             no second artifact permitted\n  matrix'
  if (p3d.pairing_rule && p3.includes(p3_pairing_anchor)) {
    p3 = p3.replace(
      p3_pairing_anchor,
      `  risk_register                             no second artifact permitted\n${p3d.pairing_rule}\n  matrix`
    )
    steps.p3_pairing = true
  } else {
    steps.p3_pairing = false
  }

  // P3-4. Add selection indicator (after risk_register indicator, before CARDS SELECTION)
  const p3_selection_anchor = '  NEVER use plain table when severity-by-row is the primary signal.\n\n\u2500\u2500\u2500 CARDS SELECTION'
  if (p3d.selection_indicator && p3.includes(p3_selection_anchor)) {
    p3 = p3.replace(
      p3_selection_anchor,
      `  NEVER use plain table when severity-by-row is the primary signal.\n\n${p3d.selection_indicator}\n\n\u2500\u2500\u2500 CARDS SELECTION`
    )
    steps.p3_selection = true
  } else {
    steps.p3_selection = false
  }

  // ════════════════════════════════════════════════════════════════════════════
  // A1-ArtifactSchema.js replica  (Unix \n)
  // ════════════════════════════════════════════════════════════════════════════

  const a1_anchor = '  NEVER use plain table when severity-by-row is the primary signal.\n`'
  if (a1d.schema_snippet && a1.includes(a1_anchor)) {
    const block = '\n' + a1d.schema_snippet + '\n' + (a1d.schema_usage_notes || '') + '\n'
    a1 = a1.replace(a1_anchor, `  NEVER use plain table when severity-by-row is the primary signal.\n${block}\``)
    steps.a1_schema = true
  } else {
    steps.a1_schema = false
  }

  // ════════════════════════════════════════════════════════════════════════════
  // agent4-R.js monolithic replica  (Windows \r\n)
  // ════════════════════════════════════════════════════════════════════════════

  // A4-1. Add to AVAILABLE ARTIFACT TYPES list
  const a4_type_anchor = '  risk_register\r\n  matrix'
  if (a4.includes(a4_type_anchor)) {
    a4 = a4.replace(a4_type_anchor, `  risk_register\r\n  ${type}\r\n  matrix`)
    steps.a4_type = true
  } else {
    steps.a4_type = false
  }

  // A4-2. Add to PRIMARY zone permitted list
  if (p3d.can_be_primary) {
    const a4_primary_anchor = '    - risk_register\r\n    - cards'
    if (a4.includes(a4_primary_anchor)) {
      a4 = a4.replace(a4_primary_anchor, `    - risk_register\r\n    - ${type}\r\n    - cards`)
      steps.a4_primary = true
    } else {
      steps.a4_primary = false
    }
  } else {
    steps.a4_primary = true
  }

  // A4-3. Add pairing rule
  const a4_pairing_anchor = '  risk_register                             no second artifact permitted\r\n  matrix'
  if (p3d.pairing_rule && a4.includes(a4_pairing_anchor)) {
    a4 = a4.replace(
      a4_pairing_anchor,
      `  risk_register                             no second artifact permitted\r\n${p3d.pairing_rule}\r\n  matrix`
    )
    steps.a4_pairing = true
  } else {
    steps.a4_pairing = false
  }

  // A4-4. Add density rule
  const a4_density_anchor = '    risk_register:     No compact; standard <6 risks;  dense >=6\r\n\r\n  FAMILY 6'
  if (p3d.density_rule && a4.includes(a4_density_anchor)) {
    a4 = a4.replace(
      a4_density_anchor,
      `    risk_register:     No compact; standard <6 risks;  dense >=6\r\n${p3d.density_rule}\r\n\r\n  FAMILY 6`
    )
    steps.a4_density = true
  } else {
    steps.a4_density = false
  }

  // A4-5. Add schema snippet + usage notes
  const a4_schema_anchor = '  NEVER use plain table when severity-by-row is the primary signal.\r\n\r\n\u2500\u2500\u2500 CARDS SELECTION'
  if (a1d.schema_snippet && a4.includes(a4_schema_anchor)) {
    const eol = '\r\n'
    const block = eol + a1d.schema_snippet + eol + (a1d.schema_usage_notes || '') + eol
    a4 = a4.replace(
      a4_schema_anchor,
      `  NEVER use plain table when severity-by-row is the primary signal.${eol}${block}${eol}\u2500\u2500\u2500 CARDS SELECTION`
    )
    steps.a4_schema = true
  } else {
    steps.a4_schema = false
  }

  // ── WRITE REPLICA FILES ──────────────────────────────────────────────────────
  try {
    fs.writeFileSync(a4Path, a4, 'utf8')
    fs.writeFileSync(p3Path, p3, 'utf8')
    fs.writeFileSync(a1Path, a1, 'utf8')
  } catch (err) {
    const restore = (src) => { try { fs.writeFileSync(src, fs.readFileSync(src + '.bak', 'utf8'), 'utf8') } catch (_) {} }
    ;[a4Path, p3Path, a1Path].forEach(restore)
    return res.status(500).json({ error: 'Failed to write replica files: ' + err.message })
  }

  const allOk = Object.values(steps).every(Boolean)
  const failedSteps = Object.entries(steps).filter(([,v]) => !v).map(([k]) => k)

  res.json({
    success: allOk,
    steps,
    warnings: failedSteps.length
      ? `Some injection points were not found and were skipped: ${failedSteps.join(', ')}. Manual insertion may be needed.`
      : undefined
  })
})

app.use((err, req, res, next) => {
  if (err && err.type === 'entity.too.large') {
    return res.status(413).json({
      error: 'Request payload too large',
      detail: 'The local server JSON body limit is 30mb. Reduce batch size or payload size if this persists.'
    })
  }
  if (err) {
    return res.status(err.status || 500).json({ error: err.message || 'Server error' })
  }
  next()
})

app.listen(3000, () => {
  console.log('PPT Generator running on http://localhost:3000')
})
