require('dotenv').config()
const express = require('express')
const app = express()

app.use(express.json({ limit: '10mb' }))
app.use(express.static('public'))

// Wire Vercel serverless handlers as local Express routes
const claudeHandler = require('./api/claude')
app.post('/api/claude', claudeHandler)

app.listen(3000, () => {
  console.log('PPT Generator running on http://localhost:3000')
})
