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

  if (!process.env.ANTHROPIC_API_KEY) {
    return res.status(500).json({ error: 'ANTHROPIC_API_KEY not set' })
  }

  try {
    const { system, messages, max_tokens } = req.body

    const MAX_RETRIES = 4
    const BASE_DELAY_MS = 2000   // 2s, 4s, 8s, 16s

    let lastStatus = 500
    let lastData   = null

    for (let attempt = 0; attempt <= MAX_RETRIES; attempt++) {
      if (attempt > 0) {
        const delay = BASE_DELAY_MS * Math.pow(2, attempt - 1)
        console.log(`Overloaded (529) — retry ${attempt}/${MAX_RETRIES} after ${delay}ms`)
        await new Promise(resolve => setTimeout(resolve, delay))
      }

      const response = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'Content-Type':      'application/json',
          'x-api-key':         process.env.ANTHROPIC_API_KEY,
          'anthropic-version': '2023-06-01',
          'anthropic-beta':    'pdfs-2024-09-25'
        },
        body: JSON.stringify({
          model:      'claude-sonnet-4-6',
          max_tokens: max_tokens || 1500,
          system:     system || '',
          messages
        })
      })

      lastStatus = response.status
      lastData   = await response.json()

      if (response.ok) {
        return res.status(200).json(lastData)
      }

      // Only retry on 529 Overloaded; surface all other errors immediately
      if (response.status !== 529) break
    }

    return res.status(lastStatus).json({
      error: lastData?.error?.message || 'Anthropic API error'
    })

  } catch (error) {
    console.error('Error:', error)
    return res.status(500).json({ error: error.message })
  }
}