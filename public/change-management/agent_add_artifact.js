// ─── AGENT: ADD ARTIFACT ─────────────────────────────────────────────────────
// Takes a user-uploaded image + name + basic description.
// Calls Claude Vision to analyze the image and generate:
//   - P3-ArtifactSelection entries: type list, primary eligibility, pairing rule,
//     density rule, and selection indicator text
//   - A1-ArtifactSchema entry: full JSON schema + usage notes
// On user approval, calls /api/inject-artifact to patch the change-management
// replica files (P3-ArtifactSelection.js, A1-ArtifactSchema.js, agent4-R.js).
// Agent 5 flattening is a separate developer step.

/* ─── SCHEMA PHILOSOPHY — shown verbatim to Claude ───────────────────────── */

// These are real, current schemas from agent4.js. They demonstrate the exact
// format Claude must follow when generating a new schema.

const AA_SCHEMA_EXAMPLES = `
════════════════════════════════════════════════════════
SCHEMA CATALOGUE PHILOSOPHY — READ CAREFULLY
════════════════════════════════════════════════════════

Each schema entry in agent4 has FOUR parts:
  1. TYPE NAME LINE:  the snake_case artifact type followed by a colon
  2. JSON BLOCK:      the actual schema, with inline example values and comments on every field
  3. USAGE NOTES:     2-indent bullet rules, covering population logic, hard limits, anti-patterns
  4. NEVER statement: explicit "NEVER use X when Y" anti-pattern sentence

SCHEMA WRITING RULES — non-negotiable:
  a) Every field must show a concrete example or enumeration, not just a type.
     BAD:   "severity": "string"
     GOOD:  "severity": "critical" | "high" | "medium" | "low"
     BAD:   "value": number
     GOOD:  "value": 8.1  ← or a numeric string like "8.1%" if that is what the renderer expects
  b) Every string field has a trailing inline comment:  "field": "example value — ≤N words; what it does"
  c) Allowed values are ALWAYS shown as:  "option1" | "option2" | "option3"
  d) Optional fields are called out:  "field": "value — OPTIONAL; omit if …"
  e) Hard limits (word counts, item counts) go in the comment, not in prose.
     Example:  "primary_message": "Revenue dropped 18% YoY — bold headline ≤8 words"
  f) Arrays always show one representative element; comment after the closing bracket if needed.
  g) Deprecated field names must be called out with: IMPORTANT: Do NOT use [old_name] — use [new_name] instead.
  h) Size rules (zone width/height requirements) go in the usage notes, not the schema block.
  i) COLUMN PAIRING RULES go in a dedicated sub-block if the artifact uses column_headers.

════════════════════════════════════════════════════════
EXAMPLE 1 — stat_bar  (columnar table with horizontal bars)
════════════════════════════════════════════════════════

stat_bar:
  IMPORTANT: Use ONLY the columnar schema below. Do NOT use the old flat "rows[].label/value/bar_color" format — it is deprecated and will fail validation.
  {
    "type": "stat_bar",
    "artifact_header": "North delivers highest margin despite lowest volume — bold proof headline ≤12 words",
    "annotation_style": "inline" | "trailing",
    "scale_UL": null,
    "column_headers": [
      { "id": "col1", "value": "Region",      "display_type": "text"   },
      { "id": "col2", "value": "On-Time Rate", "display_type": "bar",   "scale_UL": 100 },
      { "id": "col3", "value": "",             "display_type": "normal" },
      { "id": "col4", "value": "Trend",        "display_type": "text"   }
    ],
    "rows": [
      {
        "row_id": 1,
        "row_focus": "Y" | "N",
        "cells": [
          { "col_id": "col1", "value": "North" },
          { "col_id": "col2", "value": "92"    },
          { "col_id": "col3", "value": "92%"   },
          { "col_id": "col4", "value": "↑ 4pp YoY" }
        ]
      }
    ]
  }
  column_headers display_type rules:
    "text"   — plain text column (entity name / label / annotation). First text col = left label; others = trailing text.
    "bar"    — renders as a proportional horizontal bar. 1–3 bar columns allowed. Cell value must be a numeric string.
               Each bar column must carry "scale_UL": number (e.g. 100 for percentages, realistic max for ₹/count columns).
               Optionally add "scale_LL": number to override the auto lower limit (default: 50% of column minimum).
               Bar fill fraction = (value − scale_LL) / (scale_UL − scale_LL).
    "normal" — secondary display value, right-aligned (e.g. formatted metric "₹8.2" or "16.5 Days").
  COLUMN PAIRING RULE: Every "bar" column MUST be immediately followed by a "normal" column with value:"" (empty header).
  Never place a "bar" column last or adjacent to another "bar" column.
  row_focus "Y": visually highlights the row — set explicitly; do NOT mark all rows "Y"; use ≤2 per artifact.
  annotation_style default: "trailing". scale_UL: set per bar column, NOT at artifact level.
  SIZE RULES:
    1 bar column → zone width ≥ 50% of slide width
    2 bar columns → zone width ≥ 75% of slide width
    3 bar columns → zone width = 100% (full width)
    2 rows → zone height ≥ 30% slide height; each extra row adds ~11% (8 rows = 100%)
    Minimum 2 rows, maximum 8 rows.
  stat_bar usage:
  - Replace a plain table with stat_bar when rows are ranked entities, one numeric metric drives the comparison, and one short annotation per row adds meaning.
  - Use stat_bar only when the board needs ordered scanning, not cross-row / cross-column lookup.
  - row_focus: set explicitly from data — never infer from rank position.
  - Never use stat_bar for time-series data (use a line chart) or for composition/share data (use a bar or pie chart).
  NEVER use plain table when rows are ranked entities with a single dominant metric.

════════════════════════════════════════════════════════
EXAMPLE 2 — comparison_table  (options × criteria verdict grid)
════════════════════════════════════════════════════════

comparison_table:
  IMPORTANT: Use ONLY the flat schema below. Do NOT use old field names column_headers[], criteria[], options[] — they are deprecated and will cause validation failure.
  {
    "type": "comparison_table",
    "artifact_header": "Direct settlement is the dominant option across cost, speed and compliance — ≤12 words",
    "columns": ["Option", "Avg. Cost", "Turnaround", "Compliance Fit"],
    "rows": [
      {
        "is_recommended": true,
        "badge": "Recommended",
        "cells": [
          { "value": "Direct Settlement", "icon_type": null, "subtext": null,            "tone": "label"    },
          { "value": "8.1%",              "icon_type": null, "subtext": "3× lower",      "tone": "positive" },
          { "value": null,                "icon_type": "check", "subtext": "Same day",   "tone": "positive" },
          { "value": null,                "icon_type": "cross", "subtext": "Gap noted",  "tone": "negative" }
        ]
      }
    ]
  }
  Cell rules — each data cell (cells[1..n]) has EITHER value OR icon_type, never both:
  - value:     metric/percentage/currency text → rendered as a colored pill (e.g. "8.1%", "₹2,547"). Set to null when cell conveys a verdict.
  - icon_type: named icon → rendered inside a colored circle. Set to null when cell shows a metric.
               valid values: check | cross | partial | arrow_up | arrow_down | arrow_right | star | warning | diamond | chevron
  - tone:      "positive"=green | "negative"=red | "neutral"=grey | "label"=cells[0] only (the option name).
  - subtext:   ≤6-word annotation below the value or icon; null if not needed.
  - Exactly one row must have is_recommended: true.
  - Never put "Yes" / "No" / "✓" in value — use icon_type:"check" / "cross" instead.
  comparison_table usage:
  - Use when the slide's message is a clear winner or ranking across multiple evaluation criteria.
  - Columns are criteria; rows are options. Every option must be evaluated on every criterion.
  - badge: short winner label shown as a pill on the recommended row (e.g. "Recommended", "Best fit").
  - Do NOT use when options have no clear winner or when data is purely factual (use initiative_map).
  NEVER use plain table for option-vs-criteria data.

════════════════════════════════════════════════════════
EXAMPLE 3 — initiative_map  (parallel workstreams × structured dimensions)
════════════════════════════════════════════════════════

initiative_map:
  {
    "type": "initiative_map",
    "artifact_header": "Three city clusters drive 80% of GMV — each requires a distinct playbook — ≤12 words",
    "column_headers": [
      { "id": "initiative", "label": "Initiative"            },
      { "id": "c1",         "label": "Phase 1 — Immediate"  },
      { "id": "c2",         "label": "Phase 2 — Next Qtr"   }
    ],
    "rows": [
      {
        "id": "row1",
        "initiative_name": "Tier-1 City Scale — 2–5 words",
        "initiative_subtitle": "Optional muted sub-label — OPTIONAL; omit if initiative_name is self-explanatory; ≤5 words; must NOT repeat words from initiative_name",
        "cells": [
          {
            "column_id": "c1",
            "primary_message": "~₹0.3–0.4Cr revenue — headline fact ≤5 words; suppressed when tags[] non-empty",
            "secondary_message": "9,009 orders at ₹0.26Cr — OPTIONAL distinct data point ≤7 words; omit if redundant",
            "tags": [
              { "label": "Bangalore", "tone": "primary" | "secondary" | "neutral" }
            ],
            "cell_tone": "primary" | "secondary" | "neutral"
          }
        ]
      }
    ]
  }
  CRITICAL rendering rule: when tags[] is non-empty, primary_message is suppressed — tags ARE the primary visual signal.
  Put the key metric into secondary_message (≤7 words); leave primary_message "".
  Typical tag-free cell: primary_message="~₹0.3–0.4Cr revenue", secondary_message="" (omit unless distinct).
  Typical tagged cell:   tags=[{label:"Bangalore",tone:"primary"}], primary_message="", secondary_message="9,009 orders at ₹0.26Cr".
  initiative_map usage:
  - Each row is one initiative / workstream; columns are structured dimensions (phase, owner, KPI, status, city cluster).
  - DENSITY RULE: every cell must fit in ~0.7" of height — be ruthless with brevity.
  - Tags are optional pills (city names, owners, priority bands). Max 3 tags per cell.
  - Do NOT use when rows have a rank order (use prioritization). Do NOT use for process steps (use workflow).
  NEVER use when the message is ranked ordering or step-by-step dependency.

════════════════════════════════════════════════════════
EXAMPLE 4 — risk_register  (severity-banded risk list)
════════════════════════════════════════════════════════

risk_register:
  {
    "type": "risk_register",
    "risk_header": "Four systemic risks require board-level escalation — one-line framing; NOT artifact_header",
    "severity_levels": [
      {
        "id": "level_1",
        "label": "Critical severity — immediate action required — band heading shown in colored strip",
        "tone": "critical" | "high" | "medium" | "low",
        "pip_levels": 5,
        "item_details": [
          {
            "primary_message": "Punjab Dairy NPA — bold risk headline ≤8 words",
            "secondary_message": "₹17.66L outstanding; provision coverage only 40% — supporting evidence ≤18 words",
            "tags": [
              { "value": "Credit",   "tone": "neutral" | "positive" | "negative" | "warning" },
              { "value": "Q2 2025",  "tone": "neutral" }
            ],
            "pips": [
              { "label": "Likelihood", "intensity": 4 },
              { "label": "Impact",     "intensity": 5 }
            ]
          }
        ]
      }
    ]
  }
  Do NOT populate artifact_header — risk_register uses risk_header as its own internal header.
  tone drives band fill, dot color, and pip fill: critical=red | high=orange | medium=amber | low=gray.
  pip_levels: set once per severity_level (3 or 5). intensity = filled blocks out of pip_levels.
  Order severity_levels worst-to-best: critical → high → medium → low. Omit unused levels.
  primary_message: bold title ≤8 words. secondary_message: evidence line ≤18 words.
  tags[]: 1–3 short chips per item; value ≤2 words; tone drives chip color.
  pips[]: 1–3 dimension assessments (e.g. Likelihood, Impact). intensity set from data.
  risk_register usage:
  - Use when rows are named risks / issues / exceptions and each item carries severity + owner + mitigation status.
  - Use when the board's first need is severity-banded risk exposure, not flat KPI monitoring.
  - Never reduce to a plain table — the severity band IS the primary data channel.
  - Each severity_level.label should be action-oriented: "Critical — immediate escalation" not just "Critical".
  NEVER use plain table when severity-by-row is the primary signal.
`

/* ─── MAIN ENTRY POINT ───────────────────────────────────────────────────── */

async function runAgentAddArtifact(inputs) {
  // inputs: { imageB64, imageMime, artifactName, description, onProgress }
  // Claude analyzes the image to infer primary eligibility, density, pairings,
  // narrative roles, schema, and flattening function automatically.
  const log = inputs.onProgress || (() => {})

  log('Sending to Claude for analysis…')
  const raw = await _callClaude(inputs)

  log('Parsing response…')
  const parsed = _parseResponse(raw)

  return parsed
}

/* ─── CLAUDE CALL ────────────────────────────────────────────────────────── */

async function _callClaude(inputs) {
  const { imageB64, imageMime, artifactName, description } = inputs

  const systemPrompt = `You are an expert presentation design system architect.
Your task is to analyze a new visual artifact (shown in an image) and generate the exact text
needed to register it in two prompt files:

  1. P3-ArtifactSelection.js — governs WHEN Agent 4 selects this artifact type
  2. A1-ArtifactSchema.js    — governs HOW Agent 4 populates this artifact type

EXISTING ARTIFACT TYPES (do NOT reuse these names):
  chart, stat_bar, insight_text, cards, profile_card_set,
  workflow, table, comparison_table, initiative_map, risk_register,
  matrix, driver_tree, prioritization

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FILE 1 — P3-ArtifactSelection.js  (WHAT YOU MUST GENERATE)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

P3 has three places that need a new entry:

A) AVAILABLE ARTIFACT TYPES list — one line:
   Example:  "  risk_register"

B) PERMITTED SECOND ARTIFACT PAIRINGS table — one line:
   Format:   "  artifact_type_name    [secondary description or 'no second artifact permitted']"
   Examples:
     "  stat_bar                                  insight_text standard (callout only — 1–2 points), max 25% of zone size"
     "  risk_register                             no second artifact permitted"

C) "Other artifacts SELECTION INDICATORS" section — a paragraph block:
   Format (match exactly):
     "  artifact_type_name
     [One-paragraph description: what data it encodes, what makes it unique, what problem it solves]
     Use when: [specific condition that triggers this artifact — not just 'data is available'].
     NEVER use [alternative artifact] when [this artifact's defining condition]."

   Real example for risk_register:
     "  risk_register
     A severity-encoded list of risks or issues where row background color IS the
     primary data channel (not decoration). Multiple rows may have different severity
     levels simultaneously — this is what distinguishes it from highlight_rows on a
     plain table (which marks only one row at a time).
     Use when: risks or issues must be shown with per-row severity variation AND
     the board's eye must be directed to critical items without reading every cell.
     NEVER use plain table when severity-by-row is the primary signal."

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FILE 2 — A1-ArtifactSchema.js  (WHAT YOU MUST GENERATE)
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

A1 holds the JSON schema that Agent 4 must follow when populating this artifact.
Each entry has FOUR parts:
  1. TYPE NAME LINE:  artifact_type_name:
  2. Optional IMPORTANT deprecation note (only if the artifact has old/alternate field names)
  3. JSON BLOCK — every field with a concrete example value and inline comment
  4. USAGE NOTES — bullet rules + a NEVER statement

${AA_SCHEMA_EXAMPLES}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
IMAGE ANALYSIS INSTRUCTIONS
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

Before writing anything:
1. Study the image carefully. What data dimensions does it encode? How are elements
   arranged spatially? What visual encoding is used (color, size, position, shape, label)?
2. Infer from the image:
   - Is it data-rich enough to be a PRIMARY zone artifact (standalone, 50–100% of zone)?
   - What secondary artifact (if any) could pair alongside it?
   - What density levels does it naturally support?
3. Use the user-provided description only as a hint — your visual analysis takes precedence.
4. Write description as one precise sentence derived from the image.

PAIRING FORMAT REFERENCE:
  Charts (except pie/donut/group_pie)   insight_text (any subtype), max 30% of zone size
  stat_bar                              insight_text standard (callout only — 1–2 points), max 25% of zone size
  comparison_table                      insight_text standard, max 30% of zone size
  risk_register                         no second artifact permitted

DENSITY FORMAT REFERENCE (FAMILY 5B — structured display):
  stat_bar:          No compact; standard 4–5 rows;   dense 6–8 rows
  comparison_table:  No compact; standard ≤4 opt × ≤4 crit;  dense larger
  risk_register:     No compact; standard <6 risks;  dense >=6

OUTPUT: Respond with a single JSON object wrapped in \`\`\`json ... \`\`\` fences.
No explanation text outside the JSON block.`

  const userMessage = `Analyze the artifact shown in this image. Decipher its visual layout, then generate
the exact entries needed for P3-ArtifactSelection.js and A1-ArtifactSchema.js.

USER-PROVIDED HINTS (image analysis takes precedence):
  Artifact name hint:  "${artifactName}"
  Basic description:   "${description || '(none — infer entirely from the image)'}"

Generate this exact JSON structure. Every field is required.

{
  "artifact_type":  "snake_case identifier ≤20 chars — derived from the image",
  "display_name":   "Human Readable Name",
  "description":    "One precise sentence: what it visualises AND the deciding condition for choosing it",

  "p3": {
    "can_be_primary": true | false,

    "type_list_entry": "  artifact_type_name",

    "pairing_rule": "  artifact_type_name    [secondary artifact permitted and max % of zone, or 'no second artifact permitted']",

    "density_rule": "    artifact_type_name:  [compact rule or 'No compact']; standard [threshold]; dense [threshold]",

    "selection_indicator": "  artifact_type_name\\n  [one paragraph: what it encodes, what makes it unique]\\n  Use when: [specific triggering condition].\\n  NEVER use [alternative] when [this artifact's defining condition]."
  },

  "a1": {
    "schema_snippet": "FULL schema entry in the EXACT 4-part format shown in the examples above.\\nMUST include:\\n  1. TYPE NAME LINE: artifact_type_name:\\n  2. JSON BLOCK — every field has a concrete example value + inline comment\\n     BAD: \\"severity\\": \\"string\\"   GOOD: \\"severity\\": \\"critical\\" | \\"high\\" | \\"medium\\" | \\"low\\"\\n     Every string field ends with: \\"example — ≤N words; what it means\\"\\n     Optional fields end with: \\"example — OPTIONAL; omit if …\\"\\n     Arrays show one fully-filled representative element.\\n  3. SIZE RULES sub-block: minimum zone width % and height %\\n  4. Column pairing rules (if artifact uses column_headers)",

    "schema_usage_notes": "  artifact_type_name usage:\\n  - [rule 1: what each structural field means — concrete]\\n  - [rule 2: hard limit with ≤/≥]\\n  - [CRITICAL rendering rule if any]\\n  - Typical population example: [concrete Typical X: ...]\\n  - Do NOT use X when Y — use [other artifact] instead.\\n  NEVER use [alternative] when [this artifact's defining condition]."
  }
}

QUALITY CHECKLIST — verify before emitting:
  [ ] p3.selection_indicator follows the exact 4-line format (name, paragraph, Use when, NEVER)
  [ ] p3.pairing_rule and p3.density_rule use the exact spacing/format from the reference examples
  [ ] a1.schema_snippet has NO bare "string" or "number" — every field has a concrete example
  [ ] every string field in the JSON block has a trailing " — ≤N words; what it does" comment
  [ ] optional fields have "— OPTIONAL; omit if …"
  [ ] enumerations are written as "a" | "b" | "c" directly in the value
  [ ] SIZE RULES sub-block is present
  [ ] a1.schema_usage_notes ends with a NEVER statement`

  const body = {
    system: systemPrompt,
    max_tokens: 10000,
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

  // Validate required top-level fields
  for (const f of ['artifact_type', 'display_name', 'description', 'p3', 'a1']) {
    if (!parsed[f]) throw new Error(`Claude response missing required field: ${f}`)
  }
  // Validate p3 sub-fields
  for (const f of ['type_list_entry', 'pairing_rule', 'density_rule', 'selection_indicator']) {
    if (!parsed.p3[f]) throw new Error(`Claude response missing p3.${f}`)
  }
  // Validate a1 sub-fields
  for (const f of ['schema_snippet', 'schema_usage_notes']) {
    if (!parsed.a1[f]) throw new Error(`Claude response missing a1.${f}`)
  }
  return parsed
}

/* ─── VALIDATE & REFINE (called from Regenerate button) ─────────────────── */
// Takes the currently edited definition (no image needed), sends it to Claude
// for validation and correction, then returns the corrected definition in the
// same JSON shape as runAgentAddArtifact.

async function validateAndRefineArtifact(editedDef) {
  const raw = await _callClaudeValidate(editedDef)
  return _parseResponse(raw)
}

async function _callClaudeValidate(editedDef) {
  const systemPrompt = `You are an expert presentation design system architect.
The user has manually edited a new artifact definition intended for a PPT generation system.
Your job is to validate and refine the edited content — fixing format errors, spelling mistakes,
incorrect field references, and schema quality issues — WITHOUT changing the user's intent.

WHAT TO VALIDATE AND FIX:
  1. artifact_type       must be snake_case, ≤20 chars, no spaces
  2. display_name        human readable, title case
  3. description         one precise sentence — no fluff
  4. p3.type_list_entry  must be exactly "  artifact_type" (two leading spaces)
  5. p3.pairing_rule     must match: "  artifact_type    [secondary desc or 'no second artifact permitted']"
  6. p3.density_rule     must match: "    artifact_type:  No compact; standard …;  dense …"
  7. p3.selection_indicator  must follow the exact 4-line block:
       "  artifact_type_name
       [one paragraph: what it encodes, what makes it unique]
       Use when: [specific triggering condition].
       NEVER use [alternative artifact] when [this artifact's defining condition]."
  8. a1.schema_snippet   FULL 4-part entry:
       - TYPE NAME LINE: artifact_type_name:
       - JSON BLOCK — every field must have a concrete example value + trailing inline comment
         BAD:   "severity": "string"
         GOOD:  "severity": "critical" | "high" | "medium" | "low"
         Every optional field: "field": "value — OPTIONAL; omit if …"
       - SIZE RULES sub-block (zone width/height requirements)
       - NEVER statement
  9. a1.schema_usage_notes  bullet rules ending with a NEVER statement

EXISTING ARTIFACT TYPES (artifact_type must NOT duplicate these):
  chart, stat_bar, insight_text, cards, profile_card_set,
  workflow, table, comparison_table, initiative_map, risk_register,
  matrix, driver_tree, prioritization

${AA_SCHEMA_EXAMPLES}

OUTPUT: Respond with a single corrected JSON object wrapped in \`\`\`json ... \`\`\` fences.
Preserve the user's intent. No explanation text outside the JSON block.`

  const userMessage = `Validate and refine this edited artifact definition.
Fix all format errors, spelling mistakes, incorrect field references, bare type values in schemas.
Preserve the user's intent and direction.

EDITED DEFINITION:
\`\`\`json
${JSON.stringify(editedDef, null, 2)}
\`\`\`

Return the corrected definition in the same JSON structure.`

  const body = {
    system: systemPrompt,
    max_tokens: 10000,
    messages: [{ role: 'user', content: userMessage }]
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

/* ─── AGENT 5: GENERATE FLATTENED SCHEMA ─────────────────────────────────── */
// Takes the Agent 4 definition (post-inject) and generates the Agent 5
// rendering schema: either a native PowerPoint chart mapping or a full
// custom primitive-rendering spec.

async function generateAgent5Schema(a4def) {
  const raw = await _callClaudeAgent5(a4def)
  return _parseAgent5Response(raw)
}

async function _callClaudeAgent5(a4def) {
  const systemPrompt = `You are an expert PowerPoint presentation rendering engineer.

Agent 4 produces structured data for artifacts. Agent 5 converts that data into PowerPoint rendering specs.
Your job: given a NEW Agent 4 artifact schema, decide whether it is a **native chart** or a **custom rendered artifact**, then generate the correct Agent 5 schema entry.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
DECISION RULE: native chart vs custom
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  NATIVE CHART if:
    - The artifact's primary data encoding is a standard PowerPoint chart type
      (bar, column, line, scatter, area, pie, donut, waterfall, funnel, bubble, radar)
    - The data consists of series + categories, not free-form boxes or custom shapes
    - A standard PowerPoint chart engine can render it faithfully without custom shape code
  → is_native_chart: true
  → native_chart_type: the exact PowerPoint chart type (e.g. "bar", "column", "line", "scatter", "waterfall")
  → schema_entry: Produce a schema block IDENTICAL in structure to the existing "chart" type schema.
    Embed a comment at the top of the schema block explaining what chart_type value to use
    and any special Agent 5 overrides for this artifact. Reference the existing "chart" schema —
    do NOT invent a parallel structure.

  CUSTOM RENDERER if:
    - The artifact uses custom shapes, colored bands, pip indicators, tag chips, row/column grids,
      icon cells, severity fills, or any visual encoding that PowerPoint charts cannot express
    - Even if it shows data in a table-like form, if Agent 4 drives the layout, it is custom
  → is_native_chart: false
  → native_chart_type: null
  → schema_entry: Generate a FULL custom schema block. Study the existing custom types for format:
    stat_bar, comparison_table, initiative_map, risk_register, profile_card_set.
    Pattern for custom schema:
      - A _style object with ALL visual tokens Agent 5 must decide (fonts, colors, sizes, gaps)
        that cannot be inferred from Agent 4 data alone
      - A header_block field (null or structure)
      - Rules section explaining what layout JS computes vs what Agent 5 must specify
      - A MICRO-LAYOUT OWNERSHIP note saying which fields Agent 5 must set explicitly
      - Content fields (rows, columns, cells, etc.) come from Agent 4 manifest —
        Agent 5 NEVER duplicates them; Agent 5 only adds the style skin

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
FORMAT OF schema_entry
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
The schema_entry string is inserted verbatim into A1-FlattenedArtifactSchema.js.
It must be a complete section block in this format:

*********************************************************************************
NN. ARTIFACT_DISPLAY_NAME (in caps)
*********************************************************************************

{
  "type": "artifact_type_name",
  "x": number, "y": number, "w": number, "h": number,
  ...all style fields with concrete examples...
  "header_block": null or { ... }
}

Rules:
- [rule 1: what JS layout engine computes — do NOT set in Agent 5]
- [rule 2: what Agent 5 must set explicitly]
- ...

*********************************************************************************
MICRO-LAYOUT OWNERSHIP:
- [list fields Agent 5 must explicitly set]

Numbering: use "XX" — the integration script will not renumber; the developer will fix.

OUTPUT: Respond with a single JSON object wrapped in \`\`\`json ... \`\`\` fences. No text outside.`

  const userMessage = `Generate the Agent 5 rendering schema for this new artifact.

AGENT 4 ARTIFACT TYPE:   ${a4def.artifact_type}
AGENT 4 DISPLAY NAME:    ${a4def.display_name}
AGENT 4 DESCRIPTION:     ${a4def.description}

AGENT 4 SCHEMA (A1-ArtifactSchema.js entry):
${a4def.a1.schema_snippet}

AGENT 4 USAGE NOTES:
${a4def.a1.schema_usage_notes}

Decide:
  1. Is this a native PowerPoint chart type? (true/false)
  2. If yes, which chart_type value maps to it?
  3. Generate the complete schema_entry block.

Return this exact JSON:
{
  "is_native_chart": true | false,
  "native_chart_type": "chart_type_name or null",
  "description": "One sentence: how Agent 5 renders this artifact",
  "schema_entry": "FULL schema block string — see format above"
}`

  const body = {
    system: systemPrompt,
    max_tokens: 8000,
    messages: [{ role: 'user', content: userMessage }]
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

function _parseAgent5Response(rawText) {
  let jsonStr = null
  const fenced = rawText.match(/```json\s*([\s\S]*?)```/)
  if (fenced) {
    jsonStr = fenced[1].trim()
  } else {
    const start = rawText.indexOf('{')
    const end   = rawText.lastIndexOf('}')
    if (start !== -1 && end > start) jsonStr = rawText.slice(start, end + 1)
  }
  if (!jsonStr) throw new Error('Claude response did not contain a JSON block.')

  let parsed
  try { parsed = JSON.parse(jsonStr) }
  catch (e) { throw new Error('Claude Agent 5 response JSON is malformed: ' + e.message) }

  for (const f of ['is_native_chart', 'description', 'schema_entry']) {
    if (parsed[f] === undefined || parsed[f] === null || parsed[f] === '') {
      throw new Error(`Agent 5 Claude response missing required field: ${f}`)
    }
  }
  return parsed
}

/* ─── AGENT 5: VALIDATE & REFINE ────────────────────────────────────────── */
async function validateAndRefineAgent5Schema(editedA5, a4def) {
  const systemPrompt = `You are an expert PowerPoint presentation rendering engineer.
The user has manually edited an Agent 5 schema entry for a new artifact.
Validate and fix format errors, schema quality issues, missing fields, or incorrect structure.
Preserve the user's intent.

Rules to enforce:
- _style objects must have concrete color hex values (e.g. "hex" is not valid — use #RRGGBB format tokens)
- All font_size fields must be numbers, not strings
- header_block must be present as null or a fully-formed object
- If is_native_chart is true, the schema_entry must reference the existing "chart" schema structure
- If is_native_chart is false, the schema_entry must include a _style block, header_block, and Rules + MICRO-LAYOUT OWNERSHIP sections
- schema_entry must start with the *** section header and end with a *** separator line

OUTPUT: Respond with a single JSON object in \`\`\`json ... \`\`\` fences.`

  const userMessage = `Validate and refine this Agent 5 schema definition.

ARTIFACT TYPE:  ${a4def.artifact_type}
IS NATIVE CHART: ${editedA5.is_native_chart}
NATIVE CHART TYPE: ${editedA5.native_chart_type || 'null'}

EDITED SCHEMA ENTRY:
${editedA5.schema_entry}

Return:
{
  "is_native_chart": ${editedA5.is_native_chart},
  "native_chart_type": ${editedA5.native_chart_type ? '"' + editedA5.native_chart_type + '"' : 'null'},
  "description": "${editedA5.description || ''}",
  "schema_entry": "corrected schema block"
}`

  const body = {
    system: systemPrompt,
    max_tokens: 8000,
    messages: [{ role: 'user', content: userMessage }]
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
  return _parseAgent5Response((data.content?.[0]?.text || '').trim())
}

/* ─── AGENT 5: INJECTION CALL ────────────────────────────────────────────── */
async function injectAgent5Artifact(payload) {
  const res = await fetch('/api/inject-agent5-artifact', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
  })
  const data = await _safeJson(res, 'Agent 5 injection')
  if (!res.ok) throw new Error(data.error || `Agent 5 injection failed (${res.status})`)
  return data
}

/* ─── INJECTION CALL ─────────────────────────────────────────────────────── */

async function injectArtifact(definition) {
  const res = await fetch('/api/inject-artifact', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(definition)
  })
  const data = await _safeJson(res, 'Injection')
  if (!res.ok) throw new Error(data.error || `Injection failed (${res.status})`)
  return data
}

/* ─── SAFE JSON PARSE ────────────────────────────────────────────────────── */
// Prevents "Unexpected token" crash when the server returns HTML (e.g. a 404
// page from Vercel when the endpoint doesn't exist as a serverless function).
async function _safeJson(res, label) {
  const text = await res.text()
  try {
    return JSON.parse(text)
  } catch (_) {
    // Surface the raw response so the user knows what went wrong
    const preview = text.slice(0, 120).replace(/\s+/g, ' ')
    throw new Error(
      `${label} endpoint returned non-JSON (HTTP ${res.status}). ` +
      `This feature requires the local dev server — run "node index.js" instead of deploying to Vercel. ` +
      `Server said: "${preview}"`
    )
  }
}
