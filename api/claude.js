const Anthropic = require('anthropic')

module.exports = async (req, res) => {

  res.setHeader('Access-Control-Allow-Origin', '*')
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type')

  if (req.method === 'OPTIONS') {
    return res.status(200).end()
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' })
  }

  // Check API key exists
  if (!process.env.ANTHROPIC_API_KEY) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY is not set in environment variables' })
  }

  try {
    const client = new Anthropic({
      apiKey: process.env.ANTHROPIC_API_KEY
    })

    const { system, messages, max_tokens } = req.body

    if (!messages || !Array.isArray(messages)) {
      return res.status(400).json({ error: 'messages array is required' })
    }

    const response = await client.messages.create({
      model:      'claude-sonnet-4-20250514',
      max_tokens: max_tokens || 1500,
      system:     system || '',
      messages
    })

    return res.status(200).json(response)

  } catch (error) {
    console.error('Claude API error:', error)
    return res.status(500).json({
      error:   error.message || 'Unknown error',
      type:    error.constructor.name,
      details: error.status ? 'HTTP ' + error.status : ''
    })
  }
}