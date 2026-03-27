require('dotenv').config()
const express = require('express')
const app = express()

app.use(express.json({ limit: '30mb' }))
app.use(express.static('public'))

// Wire Vercel serverless handlers as local Express routes
const claudeHandler = require('./api/claude')
app.post('/api/claude', claudeHandler)

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
