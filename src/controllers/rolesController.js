const pool = require('../db')

const getRoles = async (req, res) => {
  try {
    const result = await pool.query('SELECT * FROM roles ORDER BY id ASC')
    res.json(result.rows)
  } catch (error) {
    console.error('Get roles error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

const createRole = async (req, res) => {
  try {
    const { name, description } = req.body
    const existing = await pool.query('SELECT id FROM roles WHERE name = $1', [name])
    if (existing.rows.length > 0) return res.status(400).json({ message: 'Role already exists' })
    const result = await pool.query('INSERT INTO roles (name, description) VALUES ($1, $2) RETURNING *', [name, description])
    res.status(201).json({ message: 'Role created successfully', role: result.rows[0] })
  } catch (error) {
    console.error('Create role error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

const deleteRole = async (req, res) => {
  try {
    const { id } = req.params
    const result = await pool.query('DELETE FROM roles WHERE id = $1 RETURNING id', [id])
    if (result.rows.length === 0) return res.status(404).json({ message: 'Role not found' })
    res.json({ message: 'Role deleted successfully' })
  } catch (error) {
    console.error('Delete role error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

const getRolePermissions = async (req, res) => {
  try {
    const { id } = req.params
    const result = await pool.query(
      `SELECT permissions.name, permissions.category
       FROM role_permissions
       JOIN permissions ON role_permissions.permission_id = permissions.id
       WHERE role_permissions.role_id = $1`,
      [id]
    )
    res.json(result.rows)
  } catch (error) {
    console.error('Get role permissions error:', error.message)
    res.status(500).json({ message: error.message })
  }
}

const saveRolePermissions = async (req, res) => {
  try {
    const { id } = req.params
    const { permissions } = req.body
    await pool.query('DELETE FROM role_permissions WHERE role_id = $1', [id])
    if (!permissions || permissions.length === 0) return res.json({ message: 'Permissions cleared' })
    for (const permissionName of permissions) {
      const perm = await pool.query('SELECT id FROM permissions WHERE name = $1', [permissionName])
      if (perm.rows.length > 0) {
        await pool.query('INSERT INTO role_permissions (role_id, permission_id) VALUES ($1, $2)', [id, perm.rows[0].id])
      }
    }
    res.json({ message: 'Permissions saved successfully' })
  } catch (error) {
    console.error('Save permissions error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

module.exports = { getRoles, createRole, deleteRole, getRolePermissions, saveRolePermissions }