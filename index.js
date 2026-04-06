require('dotenv').config()
const express = require('express')
const fs   = require('fs')
const path = require('path')
const app = express()

app.use(express.json({ limit: '30mb' }))
app.use(express.static('public'))

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
// Patches agent4.js + agent5.js to register a new artifact type.
// Body: { artifact_type, display_name, description, agent4:{...}, agent5:{...} }
app.post('/api/inject-artifact', (req, res) => {
  const def = req.body
  if (!def || !def.artifact_type || !def.agent4 || !def.agent5) {
    return res.status(400).json({ error: 'Missing required fields: artifact_type, agent4, agent5' })
  }

  const a4Path = path.join(__dirname, 'public', 'agent4.js')
  const a5Path = path.join(__dirname, 'public', 'agent5.js')

  // Back up originals
  const a4PathBak = a4Path + '.bak'
  const a5PathBak = a5Path + '.bak'

  let a4, a5
  try {
    a4 = fs.readFileSync(a4Path, 'utf8')
    a5 = fs.readFileSync(a5Path, 'utf8')
  } catch (err) {
    return res.status(500).json({ error: 'Could not read agent files: ' + err.message })
  }

  // Write backups before touching originals
  try {
    fs.writeFileSync(a4PathBak, a4, 'utf8')
    fs.writeFileSync(a5PathBak, a5, 'utf8')
  } catch (_) { /* non-fatal */ }

  const steps = {}
  const type   = def.artifact_type
  const a4data = def.agent4
  const a5data = def.agent5

  // Guard: skip if already injected
  if (a4.includes(`  ${type}\n`) && a4.includes(type + ':')) {
    return res.status(409).json({ error: `Artifact "${type}" already exists in agent4.js.` })
  }

  // ── AGENT 4 INJECTIONS ──────────────────────────────────────────────────────

  // 1. Add to AVAILABLE ARTIFACT TYPES list (after risk_register, before matrix)
  const a4_type_anchor = '  risk_register\n  matrix'
  if (a4.includes(a4_type_anchor)) {
    a4 = a4.replace(a4_type_anchor, `  risk_register\n  ${type}\n  matrix`)
    steps.a4_type = true
  } else {
    steps.a4_type = false
  }

  // 2. Add to PRIMARY zone permitted list (after risk_register, before cards)
  if (a4data.can_be_primary) {
    const a4_primary_anchor = '    - risk_register\n    - cards'
    if (a4.includes(a4_primary_anchor)) {
      a4 = a4.replace(a4_primary_anchor, `    - risk_register\n    - ${type}\n    - cards`)
      steps.a4_primary = true
    } else {
      steps.a4_primary = false
    }
  } else {
    steps.a4_primary = true  // skipped — not applicable
  }

  // 3. Add pairing rule (after risk_register pairing line, before matrix pairing line)
  const a4_pairing_anchor = '  risk_register                             no second artifact permitted\n  matrix'
  if (a4data.pairing_rule && a4.includes(a4_pairing_anchor)) {
    a4 = a4.replace(
      a4_pairing_anchor,
      `  risk_register                             no second artifact permitted\n${a4data.pairing_rule}\n  matrix`
    )
    steps.a4_pairing = true
  } else {
    steps.a4_pairing = false
  }

  // 4. Add density rule (after risk_register density entry, before FAMILY 6)
  const a4_density_anchor = '    risk_register:     No compact; standard <6 risks;  dense >=6\n\n  FAMILY 6'
  if (a4data.density_rule && a4.includes(a4_density_anchor)) {
    a4 = a4.replace(
      a4_density_anchor,
      `    risk_register:     No compact; standard <6 risks;  dense >=6\n${a4data.density_rule}\n\n  FAMILY 6`
    )
    steps.a4_density = true
  } else {
    steps.a4_density = false
  }

  // 5. Add schema snippet + usage notes (before PRE-OUTPUT QUALITY GATES separator)
  const a4_schema_anchor = '  NEVER use plain table when severity-by-row is the primary signal.\n\n\u2501\u2501\u2501'
  if (a4data.schema_snippet && a4.includes(a4_schema_anchor)) {
    const schemaBlock = `\n${a4data.schema_snippet}\n${a4data.schema_usage_notes || ''}\n`
    a4 = a4.replace(a4_schema_anchor, `  NEVER use plain table when severity-by-row is the primary signal.\n${schemaBlock}\n\u2501\u2501\u2501`)
    steps.a4_schema = true
  } else {
    steps.a4_schema = false
  }

  // ── AGENT 5 INJECTIONS ──────────────────────────────────────────────────────

  // 6. Register in supportedArtifactTypes Set (after 'workflow')
  const a5_type_anchor = `    'workflow'\n  ])`
  if (a5.includes(a5_type_anchor)) {
    a5 = a5.replace(a5_type_anchor, `    'workflow',\n    '${type}'\n  ])`)
    steps.a5_type = true
  } else {
    steps.a5_type = false
  }

  // 7. Add switch case entry (before default: break at end of artifact switch)
  const a5_case_anchor = `    case 'workflow': {\n      _workflowToBlocks(art, content_y, blocks, bt, r2)\n      break\n    }\n\n    default:`
  if (a5data.case_entry && a5.includes(a5_case_anchor)) {
    a5 = a5.replace(
      a5_case_anchor,
      `    case 'workflow': {\n      _workflowToBlocks(art, content_y, blocks, bt, r2)\n      break\n    }\n\n${a5data.case_entry}\n\n    default:`
    )
    steps.a5_case = true
  } else {
    steps.a5_case = false
  }

  // 8. Add flattening function (before MAIN RUNNER section)
  const a5_fn_anchor = `// \u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\u2550\n// MAIN RUNNER`
  if (a5data.function_code && a5.includes(a5_fn_anchor)) {
    a5 = a5.replace(
      a5_fn_anchor,
      `${a5data.function_code}\n\n${a5_fn_anchor}`
    )
    steps.a5_fn = true
  } else {
    // Fallback: append before the last closing lines if anchor not found
    if (a5data.function_code) {
      const fallbackAnchor = '\nasync function runAgent5('
      if (a5.includes(fallbackAnchor)) {
        a5 = a5.replace(fallbackAnchor, `\n${a5data.function_code}\n\nasync function runAgent5(`)
        steps.a5_fn = true
      } else {
        steps.a5_fn = false
      }
    } else {
      steps.a5_fn = false
    }
  }

  // ── WRITE BACK ──────────────────────────────────────────────────────────────
  try {
    fs.writeFileSync(a4Path, a4, 'utf8')
    fs.writeFileSync(a5Path, a5, 'utf8')
  } catch (err) {
    // Restore backups
    try { fs.writeFileSync(a4Path, fs.readFileSync(a4PathBak, 'utf8'), 'utf8') } catch (_) {}
    try { fs.writeFileSync(a5Path, fs.readFileSync(a5PathBak, 'utf8'), 'utf8') } catch (_) {}
    return res.status(500).json({ error: 'Failed to write agent files: ' + err.message })
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
