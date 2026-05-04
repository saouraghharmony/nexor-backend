const bcrypt = require('bcryptjs')
const jwt = require('jsonwebtoken')
const pool = require('../db')

const register = async (req, res) => {
  try {
    const { fullName, email, password, roleId } = req.body
    const existing = await pool.query('SELECT id FROM users WHERE email = $1', [email])
    if (existing.rows.length > 0) return res.status(400).json({ message: 'Email already in use' })
    const passwordHash = await bcrypt.hash(password, 10)
    const result = await pool.query(
      `INSERT INTO users (full_name, email, password_hash, role_id) VALUES ($1, $2, $3, $4) RETURNING id, full_name, email, role_id, status`,
      [fullName, email, passwordHash, roleId || 2]
    )
    res.status(201).json({ message: 'User created successfully', user: result.rows[0] })
  } catch (error) {
    console.error('Register error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

const login = async (req, res) => {
  try {
    const { email, password } = req.body
    const result = await pool.query(
      `SELECT users.*, roles.name as role_name FROM users JOIN roles ON users.role_id = roles.id WHERE users.email = $1`,
      [email]
    )
    if (result.rows.length === 0) return res.status(401).json({ message: 'Invalid email or password' })
    const user = result.rows[0]
    if (user.status !== 'active') return res.status(403).json({ message: 'Account is inactive' })
    const isValid = await bcrypt.compare(password, user.password_hash)
    if (!isValid) return res.status(401).json({ message: 'Invalid email or password' })
    const token = jwt.sign({ userId: user.id, roleId: user.role_id }, process.env.JWT_SECRET, { expiresIn: '7d' })
    res.json({
      message: 'Login successful',
      token,
      user: { id: user.id, fullName: user.full_name, email: user.email, role: user.role_name, roleId: user.role_id, status: user.status }
    })
  } catch (error) {
    console.error('Login error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

module.exports = { register, login }