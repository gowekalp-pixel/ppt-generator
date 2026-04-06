// ─── AGENT: ADD ARTIFACT ─────────────────────────────────────────────────────
// Takes a user-uploaded representative image and metadata about a new artifact
// type, calls Claude Vision to analyze it, and generates:
//   - JSON schema snippet for agent4's schema catalogue
//   - Phase 3 selection constraints for agent4
//   - Phase 4 layout/density rules for agent4
//   - Agent 5 flattening function (JavaScript)
// On user approval, calls /api/inject-artifact to patch agent4.js + agent5.js.

/* ─── CONTEXT SNIPPETS fed to Claude ─────────────────────────────────────── */

const AA_SCHEMA_EXAMPLE = `
stat_bar:
  {
    "type": "stat_bar",
    "artifact_header": "string — the one-line insight the ranking proves",
    "annotation_style": "inline" | "trailing",
    "column_headers": {
      "label": "string",
      "metric": "string",
      "value": "string",
      "annotation": "string"
    },
    "rows": [
      {
        "id": "string",
        "label": "string — left-side row label",
        "value": number,
        "unit": "string",
        "display_value": "string — optional preformatted label",
        "annotation": "string — right-side qualifier text",
        "annotation_representation": "text" | "pill",
        "bar_color": "string — optional override hex",
        "highlight": true | false
      }
    ]
  }
  stat_bar usage:
  - Replace a plain table with stat_bar when rows are ranked entities, one numeric metric drives the comparison, and one short annotation per row adds meaning.
  - highlight: set explicitly — never infer from rank.
`

const AA_FUNCTION_EXAMPLE = `
// Example of a minimal flattening function (from _prioritizationToBlocks):
// Signature: function _XxxToBlocks(art, content_y, blocks, bt, r2)
//   art        — the artifact object from the designed slide spec
//   content_y  — y-coordinate where content starts (below artifact header)
//   blocks     — array to push primitive render blocks into
//   bt         — brand tokens: bt.primary_color, bt.secondary_color,
//                bt.accent_colors[], bt.chart_palette[], bt.body_font_family,
//                bt.body_color, bt.caption_color
//   r2(v)      — rounds v to 2 decimal places
//
// Block types you can push:
//   { block_type:'rect',  x,y,w,h, fill_color, border_color, border_width, corner_radius }
//   { block_type:'text',  x,y,w,h, text, font_family, font_size, font_weight, color, align, valign, wrap }
//   { block_type:'line',  x1,y1,x2,y2, color, width }
//
// Coordinates are in inches. (0,0) = top-left of slide.
`

const AA_PAIRING_EXAMPLES = `
  Charts (Except pie, donut and Group Pie)  Insight_text (any subtype), max 30% of zone size
  stat_bar                                  insight_text standard (callout only — 1–2 points), max 25% of zone size
  risk_register                             no second artifact permitted
  matrix                                    insight_text grouped, max 30% of zone size
`

/* ─── MAIN ENTRY POINT ───────────────────────────────────────────────────── */

async function runAgentAddArtifact(inputs) {
  // inputs: { imageB64, imageMime, artifactName, description,
  //           canBePrimary, secondaryArtifact, densities:{compact,standard,dense},
  //           narrativeRoles, onProgress }
  const log = inputs.onProgress || (() => {})

  log('Sending to Claude for analysis…')
  const raw = await _callClaude(inputs)

  log('Parsing response…')
  const parsed = _parseResponse(raw)

  return parsed
}

/* ─── CLAUDE CALL ────────────────────────────────────────────────────────── */

async function _callClaude(inputs) {
  const { imageB64, imageMime, artifactName, description,
          canBePrimary, secondaryArtifact, densities, narrativeRoles } = inputs

  const densityList = Object.entries(densities || {})
    .filter(([, v]) => v).map(([k]) => k).join(', ') || 'standard'

  const systemPrompt = `You are an expert presentation design system architect.
You are adding a NEW artifact type to an AI-powered PowerPoint generation system.
The system has two key agents:
  - Agent 4 (Content Architect): selects artifact types, defines density, and populates data per zone.
  - Agent 5 (Layout Engine): flattens each artifact into primitive render blocks (rect / text / line).

EXISTING ARTIFACT TYPES (do NOT reuse these names):
  chart, stat_bar, insight_text, cards, profile_card_set,
  workflow, table, comparison_table, initiative_map, risk_register,
  matrix, driver_tree, prioritization

SCHEMA FORMAT EXAMPLE (how agent4 catalogs each artifact):
${AA_SCHEMA_EXAMPLE}

FLATTENING FUNCTION GUIDE (how agent5 converts artifact → blocks):
${AA_FUNCTION_EXAMPLE}

PAIRING RULES FORMAT (how secondary artifact pairing is described):
${AA_PAIRING_EXAMPLES}

DENSITY FAMILIES — FAMILY 5B is for structured display artifacts:
  stat_bar:          No compact; standard 4–5 rows;   dense 6–8 rows
  comparison_table:  No compact; standard ≤4 opt × ≤4 crit;  dense larger
  initiative_map:    No compact; standard ≤5 rows × ≤4 dims;  dense larger
  profile_card_set:  compact 2;  standard 3–5;  dense 6+
  risk_register:     No compact; standard <6 risks;  dense >=6

NARRATIVE ROLES (agent3 assigns these, agent4 uses them to choose artifacts):
  summary, explainer_to_summary, drill_down, segmentation, trend_analysis,
  waterfall_decomposition, benchmark_comparison, exception_highlight, validation,
  context_setter, problem_statement, risk_assessment, scenario_analysis,
  option_evaluation, recommendations, methodology_note, additional_information,
  transition_narrative

OUTPUT: Respond with a single JSON object wrapped in \`\`\`json ... \`\`\` fences.
No explanation text outside the JSON block.`

  const userMessage = `Analyze the artifact shown in this image and generate a complete definition.

USER-PROVIDED CONTEXT:
  Artifact name hint: "${artifactName}"
  Description: "${description}"
  Can be PRIMARY zone artifact: ${canBePrimary ? 'Yes' : 'No'}
  Secondary artifact pairing (if primary): "${secondaryArtifact || 'none'}"
  Supported density levels: ${densityList}
  Best narrative roles: "${narrativeRoles || 'any'}"

Generate this exact JSON structure:

{
  "artifact_type": "snake_case_name (≤20 chars, no spaces)",
  "display_name": "Human Readable Name",
  "description": "One clear sentence describing what this artifact visualises and when to use it",

  "agent4": {
    "can_be_primary": true or false,
    "type_list_entry": "  artifact_type_name",
    "pairing_rule": "  artifact_type_name    [describe permitted secondary artifact and max % of zone, or 'no second artifact permitted']",
    "density_rule": "    artifact_type_name:  [compact rule]; [standard rule]; [dense rule]  (use 'No compact' if compact not supported)",
    "selection_guidance": "  - Display Name - [2-3 sentence guidance on when to choose this over a plain table or other artifact. Pattern: 'Replace X when Y. Use it only when Z.']",
    "schema_snippet": "artifact_type_name:\n  {\n    \"type\": \"artifact_type_name\",\n    \"artifact_header\": \"string — ...\",\n    [all fields with types and inline comments]\n  }",
    "schema_usage_notes": "  artifact_type_name usage:\n  - [3-5 bullet rules about how to populate this artifact correctly]",
    "gate_check": "  [ ] [Describe any hard constraint that should be checked in a PRE-OUTPUT GATE, e.g. minimum/maximum counts or required fields]"
  },

  "agent5": {
    "function_name": "_artifactTypeNameToBlocks",
    "function_code": "function _artifactTypeNameToBlocks(art, content_y, blocks, bt, r2) {\n  // ... complete working JavaScript ...\n}",
    "case_entry": "    case 'artifact_type_name': {\n      _artifactTypeNameToBlocks(art, content_y, blocks, bt, r2)\n      break\n    }"
  }
}

IMPORTANT for function_code:
- Use art.x, art.y, art.w, art.h for bounding box (inches)
- content_y is y-start after header; compute available height as: r2((art.y + art.h) - content_y)
- Use bt.primary_color, bt.secondary_color, bt.accent_colors, bt.chart_palette for brand colors
- Use bt.body_font_family for font, bt.body_color for default text color
- Push rect/text/line blocks into the blocks array
- Keep the function complete and working (no TODOs or placeholders)`

  const body = {
    system: systemPrompt,
    max_tokens: 8000,
    messages: [{
      role: 'user',
      content: [
        {
          type: 'image',
          source: { type: 'base64', media_type: imageMime || 'image/png', data: imageB64 }
        },
        { type: 'text', text: userMessage }
      ]
    }]
  }

  const res = await fetch('/api/claude', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(body)
  })
  if (!res.ok) {
    const err = await res.json().catch(() => ({}))
    throw new Error(`Claude API error ${res.status}: ${err.error || res.statusText}`)
  }
  const data = await res.json()
  return (data.content?.[0]?.text || '').trim()
}

/* ─── RESPONSE PARSER ────────────────────────────────────────────────────── */

function _parseResponse(rawText) {
  // Extract JSON from ```json ... ``` fences
  let jsonStr = null
  const fenced = rawText.match(/```json\s*([\s\S]*?)```/)
  if (fenced) {
    jsonStr = fenced[1].trim()
  } else {
    // Fallback: find outermost { }
    const start = rawText.indexOf('{')
    const end   = rawText.lastIndexOf('}')
    if (start !== -1 && end > start) jsonStr = rawText.slice(start, end + 1)
  }
  if (!jsonStr) throw new Error('Claude response did not contain a JSON block.')

  let parsed
  try {
    parsed = JSON.parse(jsonStr)
  } catch (e) {
    throw new Error('Claude response JSON is malformed: ' + e.message)
  }

  // Validate required fields
  const required = ['artifact_type', 'display_name', 'description', 'agent4', 'agent5']
  for (const f of required) {
    if (!parsed[f]) throw new Error(`Claude response missing required field: ${f}`)
  }
  return parsed
}

/* ─── INJECTION CALL ─────────────────────────────────────────────────────── */

async function injectArtifact(definition) {
  const res = await fetch('/api/inject-artifact', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(definition)
  })
  const data = await res.json()
  if (!res.ok) throw new Error(data.error || `Injection failed (${res.status})`)
  return data
}
