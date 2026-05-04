const router = require('express').Router()
const auth = require('../middleware/authMiddleware')
const perm = require('../middleware/permissionMiddleware')
const upload = require('../middleware/upload')
const { generateReport, getReports, deleteReport } = require('../controllers/reportsController')

router.get('/', auth, perm('View History'), getReports)
router.post('/generate', auth, perm('Create Report'), upload.single('file'), generateReport)
router.delete('/:id', auth, perm('Delete Report'), deleteReport)

module.exports = router