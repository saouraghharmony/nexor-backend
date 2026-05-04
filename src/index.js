const express = require('express')
const cors = require('cors')
require('dotenv').config()

const app = express()
const PORT = process.env.PORT || 8000

app.use(cors({ origin: 'http://localhost:3000' }))
app.use(express.json())

app.use('/api/auth', require('./routes/auth'))
app.use('/api/users', require('./routes/users'))
app.use('/api/roles', require('./routes/roles'))
app.use('/api/permissions', require('./routes/permissions'))
app.use('/api/templates', require('./routes/templates'))
app.use('/api/reports', require('./routes/reports'))


app.get('/', (req, res) => res.json({ message: 'NEXOR RDP Backend is running' }))

app.listen(PORT, () => console.log(`Server running on http://localhost:${PORT}`))