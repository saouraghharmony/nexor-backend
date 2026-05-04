const router = require('express').Router()
const auth = require('../middleware/authMiddleware')
const { getPermissions } = require('../controllers/permissionsController')
router.get('/', auth, getPermissions)
module.exports = router