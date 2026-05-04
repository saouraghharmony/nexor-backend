const multer = require('multer')
const path = require('path')
const fs = require('fs')

// Create uploads folder if it doesn't exist
const uploadsDir = path.join(__dirname, '../../uploads')
if (!fs.existsSync(uploadsDir)) {
  fs.mkdirSync(uploadsDir, { recursive: true })
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, uploadsDir),
  filename: (req, file, cb) => cb(null, `${Date.now()}-${file.originalname}`)
})

const fileFilter = (req, file, cb) => {
  const allowed = ['.xlsx', '.xls']
  const ext = path.extname(file.originalname).toLowerCase()
  if (allowed.includes(ext)) cb(null, true)
  else cb(new Error('Only Excel files are allowed'))
}

const upload = multer({ storage, fileFilter })

module.exports = upload