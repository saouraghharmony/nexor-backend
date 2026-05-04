const pool = require('../db')

const checkPermission = (permissionName) => {
  return async (req, res, next) => {
    try {
      const { roleId } = req.user
      const result = await pool.query(
        `SELECT rp.id FROM role_permissions rp
         JOIN permissions p ON rp.permission_id = p.id
         WHERE rp.role_id = $1 AND p.name = $2`,
        [roleId, permissionName]
      )
      if (result.rows.length === 0) {
        return res.status(403).json({ message: `Access denied. You need the "${permissionName}" permission.` })
      }
      next()
    } catch (error) {
      console.error('Permission check error:', error)
      res.status(500).json({ message: 'Internal server error' })
    }
  }
}

module.exports = checkPermission