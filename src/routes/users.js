const router = require('express').Router()
const auth = require('../middleware/authMiddleware')
const perm = require('../middleware/permissionMiddleware')
const { getUsers, createUser, updateUser, deleteUser } = require('../controllers/usersController')
router.get('/', auth, perm('View Users'), getUsers)
router.post('/', auth, perm('Create User'), createUser)
router.put('/:id', auth, perm('Edit User'), updateUser)
router.delete('/:id', auth, perm('Delete User'), deleteUser)
module.exports = router