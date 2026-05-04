const pool = require('../db')
const XLSX = require('xlsx')
const axios = require('axios')
const path = require('path')
const fs = require('fs')
const {
  Document, Packer, Paragraph, TextRun, ImageRun,
  AlignmentType, PageBreak, HeadingLevel
} = require('docx')

/*
  Helper: download image from URL and return as buffer.
  Returns null if download fails for any reason.
*/
const downloadImage = async (url) => {
  try {
    if (!url || typeof url !== 'string') return null
    const response = await axios.get(url, {
      responseType: 'arraybuffer',
      timeout: 8000,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    })
    return Buffer.from(response.data)
  } catch {
    return null
  }
}

/*
  Helper: get safe image dimensions that fit in the document.
*/
const getImageDimensions = (buffer) => {
  try {
    if (buffer[0] === 0x89 && buffer[1] === 0x50) {
      const width = buffer.readUInt32BE(16)
      const height = buffer.readUInt32BE(20)
      if (width > 0 && height > 0) {
        const maxWidth = 580
        const ratio = Math.min(maxWidth / width, 1)
        return { width: Math.round(width * ratio), height: Math.round(height * ratio) }
      }
    }
    return { width: 580, height: 350 }
  } catch {
    return { width: 580, height: 350 }
  }
}

/*
  Helper: convert hex color to docx format (no #)
*/
const toDocxColor = (color) => {
  if (!color) return '0f172a'
  return color.replace('#', '')
}

/*
  Helper: build text runs with keyword highlighting
*/
const buildTextRuns = (text, keywordList, baseStyle) => {
  if (!text) return []
  if (keywordList.length === 0) {
    return [new TextRun({ text, ...baseStyle })]
  }

  const escapedKeywords = keywordList.map(k =>
    k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  )
  const regex = new RegExp(`(${escapedKeywords.join('|')})`, 'gi')
  const parts = text.split(regex)

  return parts.map(part => {
    const isKeyword = keywordList.includes(part.toLowerCase())
    return new TextRun({
      text: part,
      ...baseStyle,
      highlight: isKeyword ? 'yellow' : undefined,
      bold: isKeyword ? true : baseStyle.bold,
    })
  })
}

/*
  generateReport: main function
  1. Reads Excel file
  2. Filters articles by keywords
  3. Downloads images
  4. Applies template settings
  5. Generates Word document
  6. Returns .docx file
*/
const generateReport = async (req, res) => {
  try {
    const {
      templateId, title, subtitle,
      keywords, format, reportShape, reportDate
    } = req.body

    if (!title) return res.status(400).json({ message: 'Report title is required' })

    // ── Fetch template settings ──────────────────────────────
    let templateSettings = {}
    if (templateId) {
      const result = await pool.query(
        'SELECT settings FROM templates WHERE id = $1', [templateId]
      )
      if (result.rows.length > 0 && result.rows[0].settings) {
        templateSettings = result.rows[0].settings
      }
    }

    const cover = templateSettings.coverPage || {}
    const article = templateSettings.articlePage || {}
    const headerS = templateSettings.header || {}
    const footerS = templateSettings.footer || {}

    // ── Read Excel file ──────────────────────────────────────
    let articles = []

    if (req.file) {
      const workbook = XLSX.readFile(req.file.path)
      const sheetName = workbook.SheetNames[0]
      const sheet = workbook.Sheets[sheetName]
      const rows = XLSX.utils.sheet_to_json(sheet)

      // Filter successful extractions only
      articles = rows.filter(row =>
        row.ExtractionStatus === 'Extraction réussie' ||
        row.ExtractionStatus === 'Success' ||
        !row.ExtractionStatus
      )

      // Filter by keywords if provided
      if (keywords && keywords.trim()) {
        const keywordList = keywords.split(',').map(k => k.trim().toLowerCase()).filter(k => k)
        articles = articles.filter(row => {
          const text = `${row.Title || ''} ${row.ArticleText || ''}`.toLowerCase()
          return keywordList.some(kw => text.includes(kw))
        })
      }

      if (articles.length === 0) {
        if (req.file) fs.unlinkSync(req.file.path)
        return res.status(400).json({ message: 'No articles found matching your keywords' })
      }
    }

    // Parse keyword list for highlighting
    const keywordList = keywords
      ? keywords.split(',').map(k => k.trim().toLowerCase()).filter(k => k)
      : []

    // ── Build document children ──────────────────────────────
    const children = []

    // ── COVER PAGE ───────────────────────────────────────────
    const coverAlign =
      cover.alignment === 'Left' ? AlignmentType.LEFT :
      cover.alignment === 'Right' ? AlignmentType.RIGHT :
      AlignmentType.CENTER

    children.push(new Paragraph({ children: [], spacing: { before: 2000 } }))

    children.push(new Paragraph({
      alignment: coverAlign,
      spacing: { before: 400, after: 300 },
      children: [
        new TextRun({
          text: title,
          bold: cover.bold !== false,
          size: (cover.titleSize || 48) * 2,
          color: toDocxColor(cover.titleColor),
          font: cover.fontFamily || 'Arial',
        })
      ]
    }))

    if (subtitle) {
      children.push(new Paragraph({
        alignment: coverAlign,
        spacing: { after: 300 },
        children: [
          new TextRun({
            text: subtitle,
            italics: true,
            size: (cover.subtitleSize || 24) * 2,
            color: toDocxColor(cover.subtitleColor),
            font: cover.fontFamily || 'Arial',
          })
        ]
      }))
    }

    // Date
    const dateStr = (() => {
      const now = new Date(reportDate || Date.now())
      const fmt = cover.dateFormat || 'MMM DD, YYYY'
      if (fmt === 'DD/MM/YYYY') {
        return `${String(now.getDate()).padStart(2,'0')}/${String(now.getMonth()+1).padStart(2,'0')}/${now.getFullYear()}`
      }
      if (fmt === 'YYYY-MM-DD') {
        return `${now.getFullYear()}-${String(now.getMonth()+1).padStart(2,'0')}-${String(now.getDate()).padStart(2,'0')}`
      }
      return now.toLocaleDateString('en-US', { month: 'long', day: '2-digit', year: 'numeric' })
    })()

    children.push(new Paragraph({
      alignment: coverAlign,
      children: [
        new TextRun({
          text: dateStr,
          size: 24,
          color: toDocxColor(cover.dateColor) || '94a3b8',
          font: cover.fontFamily || 'Arial',
        })
      ]
    }))

    // Page break after cover
    children.push(new Paragraph({ children: [new PageBreak()] }))

    // ── TABLE OF CONTENTS ────────────────────────────────────
    if (articles.length > 0) {
      children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 400, after: 300 },
        children: [new TextRun({ text: 'Table of Contents', bold: true, size: 36 })]
      }))

      articles.forEach((a, i) => {
        children.push(new Paragraph({
          spacing: { after: 120 },
          children: [
            new TextRun({
              text: `${i + 1}.  ${a.Title || 'Untitled'}`,
              size: 22,
              font: 'Arial',
            })
          ]
        }))
      })

      children.push(new Paragraph({ children: [new PageBreak()] }))
    }

    // ── ARTICLES ─────────────────────────────────────────────
    for (const a of articles) {

      // Article title
      children.push(new Paragraph({
        spacing: { before: 300, after: 200 },
        children: [
          new TextRun({
            text: a.Title || 'Untitled',
            bold: article.titleBold !== false,
            italics: article.titleItalic || false,
            size: (article.titleSize || 28) * 2,
            color: toDocxColor(article.titleColor),
            font: article.titleFontFamily || 'Arial',
          })
        ]
      }))

      // Source + metadata
      const meta = [
        a.SiteName && `Source: ${a.SiteName}`,
        a.Author && `Author: ${a.Author}`,
        a.Date && `Date: ${a.Date}`,
      ].filter(Boolean).join('  |  ')

      if (meta) {
        children.push(new Paragraph({
          spacing: { after: 200 },
          children: [
            new TextRun({
              text: meta,
              size: (article.sourceSize || 12) * 2,
              color: toDocxColor(article.sourceColor) || '64748b',
              font: article.sourceFontFamily || 'Arial',
              bold: article.sourceBold || false,
              italics: article.sourceItalic || false,
            })
          ]
        }))
      }

      // Image (top position)
      if (a.Image && article.imagePosition !== 'Bottom') {
        const imageBuffer = await downloadImage(a.Image)
        if (imageBuffer) {
          try {
            const dims = getImageDimensions(imageBuffer)
            // Apply image size percentage
            const sizeRatio = (article.imageSize || 100) / 100
            children.push(new Paragraph({
              spacing: { after: 200 },
              children: [
                new ImageRun({
                  data: imageBuffer,
                  transformation: {
                    width: Math.round(dims.width * sizeRatio),
                    height: Math.round(dims.height * sizeRatio),
                  }
                })
              ]
            }))
          } catch {
            // Skip image silently
          }
        }
      }

      // Article body with keyword highlighting
      if (a.ArticleText) {
        const paragraphs = a.ArticleText
          .split('\n')
          .map(p => p.trim())
          .filter(p => p.length > 0)

        const bodyStyle = {
          size: (article.bodySize || 11) * 2,
          color: toDocxColor(article.bodyColor) || '374151',
          font: article.bodyFontFamily || 'Arial',
          bold: article.bodyBold || false,
          italics: article.bodyItalic || false,
        }

        for (const para of paragraphs) {
          const runs = buildTextRuns(para, keywordList, bodyStyle)
          children.push(new Paragraph({
            spacing: { after: 160, line: Math.round((article.bodyLineHeight || 1.5) * 240) },
            children: runs
          }))
        }
      }

      // Image (bottom position)
      if (a.Image && article.imagePosition === 'Bottom') {
        const imageBuffer = await downloadImage(a.Image)
        if (imageBuffer) {
          try {
            const dims = getImageDimensions(imageBuffer)
            const sizeRatio = (article.imageSize || 100) / 100
            children.push(new Paragraph({
              spacing: { before: 200 },
              children: [
                new ImageRun({
                  data: imageBuffer,
                  transformation: {
                    width: Math.round(dims.width * sizeRatio),
                    height: Math.round(dims.height * sizeRatio),
                  }
                })
              ]
            }))
          } catch {
            // Skip image silently
          }
        }
      }

      // Page break between articles
      children.push(new Paragraph({ children: [new PageBreak()] }))
    }

    // ── GENERATE WORD DOCUMENT ───────────────────────────────
    const doc = new Document({
      sections: [{
        properties: {},
        children,
      }]
    })

    const buffer = await Packer.toBuffer(doc)

    // ── SAVE TO DATABASE ─────────────────────────────────────
    await pool.query(
      `INSERT INTO reports
        (name, template_id, created_by, keywords, links_count, format, report_shape, report_date, status)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8, 'success')`,
      [
        title,
        templateId || null,
        req.user.userId,
        keywords || '',
        articles.length,
        format || 'word',
        reportShape || 'combined',
        reportDate || new Date()
      ]
    )

    // ── CLEAN UP UPLOADED FILE ───────────────────────────────
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path)
    }

    // ── SEND FILE TO BROWSER ─────────────────────────────────
    const fileName = `${title.replace(/\s+/g, '_')}_${Date.now()}.docx`
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`)
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    res.send(buffer)

  } catch (error) {
    console.error('Generate report error:', error)
    if (req.file && fs.existsSync(req.file.path)) {
      fs.unlinkSync(req.file.path)
    }
    res.status(500).json({ message: 'Failed to generate report: ' + error.message })
  }
}

const getReports = async (req, res) => {
  try {
    const result = await pool.query(
      `SELECT reports.*, users.full_name as created_by_name, templates.name as template_name
       FROM reports
       LEFT JOIN users ON reports.created_by = users.id
       LEFT JOIN templates ON reports.template_id = templates.id
       ORDER BY reports.created_at DESC`
    )
    res.json(result.rows)
  } catch (error) {
    console.error('Get reports error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

const deleteReport = async (req, res) => {
  try {
    const { id } = req.params
    const result = await pool.query(
      'DELETE FROM reports WHERE id = $1 RETURNING id', [id]
    )
    if (result.rows.length === 0) {
      return res.status(404).json({ message: 'Report not found' })
    }
    res.json({ message: 'Report deleted successfully' })
  } catch (error) {
    console.error('Delete report error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

module.exports = { generateReport, getReports, deleteReport }