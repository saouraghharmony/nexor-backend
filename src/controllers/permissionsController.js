const pool = require('../db')

const getPermissions = async (req, res) => {
  try {
    const result = await pool.query('SELECT * FROM permissions ORDER BY category, id ASC')
    const grouped = {}
    result.rows.forEach(perm => {
      if (!grouped[perm.category]) grouped[perm.category] = []
      grouped[perm.category].push(perm.name)
    })
    const permissionGroups = Object.entries(grouped).map(([category, items]) => ({ category, items }))
    res.json(permissionGroups)
  } catch (error) {
    console.error('Get permissions error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

module.exports = { getPermissions }