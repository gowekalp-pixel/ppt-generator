// ─── AGENT 5 — REPAIR / FALLBACK ─────────────────────────────────────────────
// This file is intentionally separate from agent5.js so fallback design logic
// lives in one place.
//
// When adding a new artifact type:
//   1. buildMinimalSafeSlide in agent5.js handles the safe shell — add a case
//      in buildSafeArtifactShell if the new type needs a custom empty template.
//   2. validateFallbackStructure checks zone/artifact count and type matching —
//      no changes needed for new types (it checks types generically).
//   3. Edit prompts/agent5/A6-fallback-system.js if fallback prompt language changes.
//
// Depends on (defined in agent5.js, must be fully loaded before pipeline runs):
//   extractBrandTokens, callClaude, safeParseJSON,
//   validateDesignedSlide, buildMinimalSafeSlide, artifactSignatureType,
//   AGENT5_FALLBACK_SYSTEM

// ─── ZONE/ARTIFACT SIGNATURE ──────────────────────────────────────────────────
// Returns a 2-D array: [ [artifactType, ...], ... ] for each zone.
// Used by validateFallbackStructure to compare manifest vs Claude output.
function manifestZoneArtifactSignature(slideLike) {
  return (slideLike?.zones || []).map(z => (z.artifacts || []).map(a => artifactSignatureType(a)))
}

// ─── FALLBACK STRUCTURE VALIDATOR ─────────────────────────────────────────────
// Checks that a Claude-produced fallback slide matches the zone/artifact
// structure declared in the manifest. Returns an array of issue strings.
function validateFallbackStructure(candidate, manifestSlide) {
  const issues = []
  const candSig = manifestZoneArtifactSignature(candidate)
  const manifestSig = manifestZoneArtifactSignature(manifestSlide)
  if (candSig.length !== manifestSig.length) {
    issues.push('zone count mismatch vs manifest')
    return issues
  }
  for (let zi = 0; zi < manifestSig.length; zi++) {
    const mArts = manifestSig[zi]
    const cArts = candSig[zi] || []
    if (cArts.length !== mArts.length) {
      issues.push('z' + zi + ': artifact count mismatch vs manifest')
      continue
    }
    for (let ai = 0; ai < mArts.length; ai++) {
      if (String(cArts[ai] || '') !== String(mArts[ai] || '')) {
        issues.push('z' + zi + '.a' + ai + ': artifact type mismatch vs manifest (' + cArts[ai] + ' != ' + mArts[ai] + ')')
      }
    }
  }
  return issues
}

// ─── FALLBACK DESIGN RUNNER ───────────────────────────────────────────────────
// Called when the primary designSlideBatch call produces a structurally invalid
// slide. Makes a single-slide Claude call with a more focused fallback system
// prompt. Falls back to buildMinimalSafeSlide if Claude also fails.
async function buildFallbackDesign(manifestSlide, brand) {
  console.log('Agent 5 -- fallback Claude call for S' + manifestSlide.slide_number)

  const tokens = extractBrandTokens(brand)

  const prompt =
    'BRAND DESIGN TOKENS:\n' +
    JSON.stringify(tokens, null, 2) +
    '\n\nSLIDE TO REBUILD:\n' +
    JSON.stringify(manifestSlide, null, 2) +
    '\n\n' + _A5_FALLBACK_INSTRUCTIONS

  try {
    const raw    = await callClaude(AGENT5_FALLBACK_SYSTEM, [{ role: 'user', content: prompt }], 5000)
    const parsed = safeParseJSON(raw, null)

    // Claude may return an array with one item or a bare object
    const candidate = Array.isArray(parsed) ? parsed[0] : parsed

    if (candidate && typeof candidate === 'object' && candidate.canvas && candidate.zones) {
      const issues = validateDesignedSlide(candidate)
      const signatureIssues = validateFallbackStructure(candidate, manifestSlide)
      if (issues.length === 0 && signatureIssues.length === 0) {
        console.log('Agent 5 -- fallback Claude succeeded for S' + manifestSlide.slide_number)
        return candidate
      }
      console.warn('Agent 5 -- fallback Claude still has issues for S' + manifestSlide.slide_number + ':', issues.concat(signatureIssues).join('; '))
    }
  } catch (e) {
    console.warn('Agent 5 -- fallback Claude call failed for S' + manifestSlide.slide_number + ':', e.message)
  }

  // Last resort: minimal structurally valid object derived purely from brand tokens
  // No hardcoded content — pull everything from manifest and brand
  return buildMinimalSafeSlide(manifestSlide, tokens)
}
