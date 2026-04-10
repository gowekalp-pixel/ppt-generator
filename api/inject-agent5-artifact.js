// Vercel serverless stub for /api/inject-agent5-artifact
// File-system writes are not supported in Vercel's serverless environment.
// Run "node index.js" locally to use this endpoint.
module.exports = async (_req, res) => {
  res.setHeader('Access-Control-Allow-Origin', '*')
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type')
  if (_req.method === 'OPTIONS') return res.status(200).end()

  res.status(501).json({
    error:
      'The /api/inject-agent5-artifact endpoint writes to the local filesystem and cannot run on Vercel. ' +
      'Start the local dev server with "node index.js" and open http://localhost:3000/change-management/add-artifact.html to use this feature.'
  })
}
