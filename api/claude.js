 const Anthropic = require('anthropic')

const client = new Anthropic({
  apiKey: process.env.ANTHROPIC_API_KEY
})

module.exports = async (req, res) => {

  res.setHeader('Access-Control-Allow-Origin', '*')
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS')
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type')

  if (req.method === 'OPTIONS') {
    return res.status(200).end()
  }

  try {
    const { system, messages, max_tokens } = req.body

    const response = await client.messages.create({
      model: 'claude-sonnet-4-20250514',
      max_tokens: max_tokens || 1000,
      system,
      messages
    })

    res.status(200).json(response)

  } catch (error) {
    console.error('Claude API error:', error)
    res.status(500).json({ error: error.message })
  }
}