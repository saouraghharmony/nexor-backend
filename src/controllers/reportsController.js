const pool = require('../db')
const XLSX = require('xlsx')
const axios = require('axios')
const path = require('path')
const fs = require('fs')
const {
  Document, Packer, Paragraph, TextRun, ImageRun,
  HeadingLevel, AlignmentType, PageBreak, TableOfContents,
  StyleLevel
} = require('docx')

/*
  Helper: download image from URL and return as buffer.
  Returns null if download fails.
*/
const downloadImage = async (url) => {
  try {
    const response = await axios.get(url, {
      responseType: 'arraybuffer',
      timeout: 10000,
      headers: { 'User-Agent': 'Mozilla/5.0' }
    })
    return Buffer.from(response.data)
  } catch {
    return null
  }
}

/*
  Helper: get image dimensions safely.
  Returns default dimensions if image is invalid.
*/
const getImageDimensions = (buffer) => {
  try {
    // Check PNG signature
    if (buffer[0] === 0x89 && buffer[1] === 0x50) {
      const width = buffer.readUInt32BE(16)
      const height = buffer.readUInt32BE(20)
      if (width > 0 && height > 0) {
        const maxWidth = 600
        const ratio = Math.min(maxWidth / width, 1)
        return { width: Math.round(width * ratio), height: Math.round(height * ratio) }
      }
    }
    return { width: 600, height: 400 }
  } catch {
    return { width: 600, height: 400 }
  }
}

/*
  generateReport: main function that:
  1. Reads the uploaded Excel file
  2. Filters articles by keywords
  3. Downloads images
  4. Generates a Word document using template settings
  5. Saves the file and returns it
*/
const generateReport = async (req, res) => {
  try {
    const {
      templateId,
      title,
      subtitle,
      keywords,
      format,
      reportShape,
      reportDate
    } = req.body

    // Check if file was uploaded
    if (!req.file) {
      return res.status(400).json({ message: 'Excel file is required' })
    }

    // Read Excel file
    const workbook = XLSX.readFile(req.file.path)
    const sheetName = workbook.SheetNames[0]
    const sheet = workbook.Sheets[sheetName]
    const rows = XLSX.utils.sheet_to_json(sheet)

    // Filter rows by keywords if provided
    let articles = rows.filter(row => row.ExtractionStatus === 'Extraction réussie')

    if (keywords && keywords.trim()) {
      const keywordList = keywords.split(',').map(k => k.trim().toLowerCase())
      articles = articles.filter(row => {
        const text = `${row.Title} ${row.ArticleText}`.toLowerCase()
        return keywordList.some(kw => text.includes(kw))
      })
    }

    if (articles.length === 0) {
      return res.status(400).json({ message: 'No articles found matching your keywords' })
    }

    // Fetch template settings from database
    let templateSettings = {}
    if (templateId) {
      const templateResult = await pool.query('SELECT settings FROM templates WHERE id = $1', [templateId])
      if (templateResult.rows.length > 0) {
        templateSettings = templateResult.rows[0].settings || {}
      }
    }

    const coverPage = templateSettings.coverPage || {}
    const headerSettings = templateSettings.header || {}
    const footerSettings = templateSettings.footer || {}
    const articlePage = templateSettings.articlePage || {}

    // Build document sections
    const children = []

    // ── COVER PAGE ──────────────────────────────────────
    children.push(
      new Paragraph({
        text: title || 'Annual Report',
        heading: HeadingLevel.TITLE,
        alignment: AlignmentType.CENTER,
        spacing: { before: 3000, after: 400 },
        run: {
          size: Math.min((coverPage.titleSize || 48) * 2, 96),
          bold: coverPage.bold || false,
          color: (coverPage.titleColor || '#0f172a').replace('#', ''),
          font: coverPage.fontFamily || 'Arial',
        }
      })
    )

    if (subtitle) {
      children.push(
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 400 },
          children: [
            new TextRun({
              text: subtitle,
              italics: true,
              size: (coverPage.subtitleSize || 24) * 2,
              color: (coverPage.subtitleColor || '#475569').replace('#', ''),
              font: coverPage.fontFamily || 'Arial',
            })
          ]
        })
      )
    }

    children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [
          new TextRun({
            text: new Date(reportDate || Date.now()).toLocaleDateString('en-US', { month: 'long', day: '2-digit', year: 'numeric' }),
            size: 24,
            color: (coverPage.dateColor || '94a3b8').replace('#', ''),
          })
        ]
      })
    )

    // Page break after cover
    children.push(new Paragraph({ children: [new PageBreak()] }))

    // ── TABLE OF CONTENTS ──────────────────────────────
    children.push(
      new Paragraph({
        text: 'Table of Contents',
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 400, after: 400 },
      })
    )

    // Build TOC entries manually
    articles.forEach((article, index) => {
      children.push(
        new Paragraph({
          spacing: { after: 120 },
          children: [
            new TextRun({
              text: `${index + 1}. ${article.Title || 'Untitled'}`,
              size: 24,
              font: 'Arial',
            })
          ]
        })
      )
    })

    // Page break after TOC
    children.push(new Paragraph({ children: [new PageBreak()] }))

    // ── ARTICLES ─────────────────────────────────────────
    for (const article of articles) {
      // Article title
      children.push(
        new Paragraph({
          spacing: { before: 400, after: 200 },
          children: [
            new TextRun({
              text: article.Title || 'Untitled',
              bold: articlePage.titleBold !== false,
              size: Math.min((articlePage.titleSize || 32) * 2, 64),
              color: (articlePage.titleColor || '#0f172a').replace('#', ''),
              font: articlePage.titleFontFamily || 'Arial',
            })
          ]
        })
      )

      // Source + Author + Date metadata
      children.push(
        new Paragraph({
          spacing: { after: 200 },
          children: [
            new TextRun({
              text: `Source: ${article.SiteName || 'Unknown'} | Author: ${article.Author || 'Unknown'} | Date: ${article.Date || ''}`,
              size: 20,
              color: (articlePage.sourceColor || '64748b').replace('#', ''),
              font: articlePage.sourceFontFamily || 'Arial',
              italics: articlePage.sourceItalic || false,
            })
          ]
        })
      )

      // Article image
      if (article.Image) {
        const imageBuffer = await downloadImage(article.Image)
        if (imageBuffer) {
          try {
            const dims = getImageDimensions(imageBuffer)
            children.push(
              new Paragraph({
                spacing: { after: 200 },
                children: [
                  new ImageRun({
                    data: imageBuffer,
                    transformation: {
                      width: dims.width,
                      height: dims.height,
                    }
                  })
                ]
              })
            )
          } catch {
            // Skip image if it fails
          }
        }
      }

      // Article body text with keyword highlighting
      if (article.ArticleText) {
        const paragraphs = article.ArticleText.split('\n').filter(p => p.trim())
        const keywordList = keywords
          ? keywords.split(',').map(k => k.trim().toLowerCase()).filter(k => k)
          : []

        for (const para of paragraphs.slice(0, 10)) {
          const trimmed = para.trim()

          if (keywordList.length === 0) {
            // No keywords — render plain text
            children.push(
              new Paragraph({
                spacing: { after: 160 },
                children: [
                  new TextRun({
                    text: trimmed,
                    size: (articlePage.bodySize || 14) * 2,
                    color: (articlePage.bodyColor || '374151').replace('#', ''),
                    font: articlePage.bodyFontFamily || 'Arial',
                  })
                ]
              })
            )
          } else {
            /*
              Split the paragraph into segments — highlighted and normal.
              We find keyword positions and split around them.
            */
            const runs = []
            let remaining = trimmed
            let lastIndex = 0

            // Build a regex that matches any of the keywords
            const regex = new RegExp(`(${keywordList.map(k => k.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')})`, 'gi')
            const parts = trimmed.split(regex)

            for (const part of parts) {
              const isKeyword = keywordList.includes(part.toLowerCase())
              runs.push(
                new TextRun({
                  text: part,
                  size: (articlePage.bodySize || 14) * 2,
                  color: (articlePage.bodyColor || '374151').replace('#', ''),
                  font: articlePage.bodyFontFamily || 'Arial',
                  // Yellow highlight for matching keywords
                  highlight: isKeyword ? 'yellow' : undefined,
                  bold: isKeyword ? true : (articlePage.bodyBold || false),
                })
              )
            }

            children.push(
              new Paragraph({
                spacing: { after: 160 },
                children: runs
              })
            )
          }
        }
      }

      // Page break between articles
      children.push(new Paragraph({ children: [new PageBreak()] }))
    }

    // ── GENERATE WORD DOCUMENT ────────────────────────────
    const doc = new Document({
      sections: [{
        properties: {},
        children,
      }]
    })

    const buffer = await Packer.toBuffer(doc)

    // Save report to database
    const reportResult = await pool.query(
      `INSERT INTO reports (name, template_id, created_by, keywords, links_count, format, report_shape, report_date, status)
       VALUES ($1, $2, $3, $4, $5, $6, $7, $8, 'success') RETURNING id`,
      [title || 'Report', templateId || null, req.user.userId, keywords || '', articles.length, format || 'word', reportShape || 'combined', reportDate || new Date()]
    )

    // Clean up uploaded Excel file
    fs.unlinkSync(req.file.path)

    // Send the Word file back to the client
    const fileName = `${(title || 'report').replace(/\s+/g, '_')}_${Date.now()}.docx`
    res.setHeader('Content-Disposition', `attachment; filename="${fileName}"`)
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document')
    res.send(buffer)

  } catch (error) {
    console.error('Generate report error:', error)
    res.status(500).json({ message: 'Failed to generate report: ' + error.message })
  }
}

/*
  getReports: returns all reports for the history page.
*/
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

/*
  deleteReport: deletes a report by id.
*/
const deleteReport = async (req, res) => {
  try {
    const { id } = req.params
    const result = await pool.query('DELETE FROM reports WHERE id = $1 RETURNING id', [id])
    if (result.rows.length === 0) return res.status(404).json({ message: 'Report not found' })
    res.json({ message: 'Report deleted successfully' })
  } catch (error) {
    console.error('Delete report error:', error)
    res.status(500).json({ message: 'Internal server error' })
  }
}

module.exports = { generateReport, getReports, deleteReport }