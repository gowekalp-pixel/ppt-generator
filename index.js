const express = require('express')
const app = express()

app.use(express.json())
app.use(express.static('public'))

app.listen(3000, () => {
  console.log('PPT Generator running on http://localhost:3000')
})
