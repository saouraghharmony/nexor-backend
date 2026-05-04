const pool = require('../db')

const getTemplates = async (req, res) => {
  try {
    const result = await pool.query(
      `SELECT templates.*, users.full_name as created_by_name
       FROM templates LEFT JOIN users ON templates.created_by = users.id
       ORDER BY templates.created_at DESC`
    )
    res.json(result.rows)
  } catch (error) {
    console.error('Get templates error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

const getTemplate = async (req, res) => {
  try {
    const { id } = req.params
    const result = await pool.query(
      `SELECT templates.*, users.full_name as created_by_name
       FROM templates LEFT JOIN users ON templates.created_by = users.id
       WHERE templates.id = $1`,
      [id]
    )
    if (result.rows.length === 0) return res.status(404).json({ message: 'Template not found' })
    res.json(result.rows[0])
  } catch (error) {
    console.error('Get template error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

const createTemplate = async (req, res) => {
  try {
    const { name, description, settings } = req.body
    const result = await pool.query(
      `INSERT INTO templates (name, description, settings, created_by) VALUES ($1, $2, $3, $4) RETURNING *`,
      [name, description, JSON.stringify(settings || {}), req.user.userId]
    )
    res.status(201).json({ message: 'Template created successfully', template: result.rows[0] })
  } catch (error) {
    console.error('Create template error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

const updateTemplate = async (req, res) => {
  try {
    const { id } = req.params
    const { name, description, settings } = req.body
    const result = await pool.query(
      `UPDATE templates SET name = $1, description = $2, settings = $3, updated_at = NOW() WHERE id = $4 RETURNING *`,
      [name, description, JSON.stringify(settings || {}), id]
    )
    if (result.rows.length === 0) return res.status(404).json({ message: 'Template not found' })
    res.json({ message: 'Template updated successfully', template: result.rows[0] })
  } catch (error) {
    console.error('Update template error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

const deleteTemplate = async (req, res) => {
  try {
    const { id } = req.params
    const result = await pool.query('DELETE FROM templates WHERE id = $1 RETURNING id', [id])
    if (result.rows.length === 0) return res.status(404).json({ message: 'Template not found' })
    res.json({ message: 'Template deleted successfully' })
  } catch (error) {
    console.error('Delete template error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

module.exports = { getTemplates, getTemplate, createTemplate, updateTemplate, deleteTemplate }