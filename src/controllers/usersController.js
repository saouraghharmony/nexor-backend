const pool = require('../db')
const bcrypt = require('bcryptjs')

const getUsers = async (req, res) => {
  try {
    const result = await pool.query(
      `SELECT users.id, users.full_name, users.email, users.status, users.role_id, roles.name as role
       FROM users JOIN roles ON users.role_id = roles.id ORDER BY users.created_at DESC`
    )
    res.json(result.rows)
  } catch (error) {
    console.error('Get users error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

const createUser = async (req, res) => {
  try {
    const { fullName, email, password, roleId, isActive } = req.body
    const existing = await pool.query('SELECT id FROM users WHERE email = $1', [email])
    if (existing.rows.length > 0) return res.status(400).json({ message: 'Email already in use' })
    const passwordHash = await bcrypt.hash(password || 'changeme123', 10)
    const result = await pool.query(
      `INSERT INTO users (full_name, email, password_hash, role_id, status) VALUES ($1, $2, $3, $4, $5) RETURNING id, full_name, email, role_id, status`,
      [fullName, email, passwordHash, roleId, isActive ? 'active' : 'inactive']
    )
    res.status(201).json({ message: 'User created successfully', user: result.rows[0] })
  } catch (error) {
    console.error('Create user error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

const updateUser = async (req, res) => {
  try {
    const { id } = req.params
    const { fullName, email, roleId, isActive, password } = req.body
    let passwordHash
    if (password && password.trim().length > 0) {
      passwordHash = await bcrypt.hash(password, 10)
    } else {
      const existing = await pool.query('SELECT password_hash FROM users WHERE id = $1', [id])
      if (existing.rows.length === 0) return res.status(404).json({ message: 'User not found' })
      passwordHash = existing.rows[0].password_hash
    }
    const result = await pool.query(
      `UPDATE users SET full_name = $1, email = $2, role_id = $3, status = $4, password_hash = $5, updated_at = NOW() WHERE id = $6 RETURNING id, full_name, email, role_id, status`,
      [fullName, email, roleId, isActive ? 'active' : 'inactive', passwordHash, id]
    )
    if (result.rows.length === 0) return res.status(404).json({ message: 'User not found' })
    res.json({ message: 'User updated successfully', user: result.rows[0] })
  } catch (error) {
    console.error('Update user error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

const deleteUser = async (req, res) => {
  try {
    const { id } = req.params
    const result = await pool.query('DELETE FROM users WHERE id = $1 RETURNING id', [id])
    if (result.rows.length === 0) return res.status(404).json({ message: 'User not found' })
    res.json({ message: 'User deleted successfully' })
  } catch (error) {
    console.error('Delete user error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

module.exports = { getUsers, createUser, updateUser, deleteUser }