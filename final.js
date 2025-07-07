npm install express cors body-parser exceljs cheerio

const express = require('express');
const cors = require('cors');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const cheerio = require('cheerio');
const fs = require('fs');
const path = require('path');

const app = express();
app.use(cors());
app.use(bodyParser.json({ limit: '100mb' }));

// Health check
app.get('/health', (req, res) => {
  res.json({
    status: 'healthy',
    timestamp: new Date().toISOString(),
    version: '1.0.0'
  });
});

// HTML to Excel conversion helper
async function convertHtmlToExcel(html, outputPath) {
  const $ = cheerio.load(html);
  const table = $('table').first();
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet('Sheet 1');

  if (!table.length) throw new Error('No <table> found in HTML');

  table.find('tr').each((i, row) => {
    const rowData = [];
    $(row).find('th, td').each((j, cell) => {
      rowData.push($(cell).text().trim());
      // You can enhance here to transfer styles, colspan, etc.
    });
    worksheet.addRow(rowData);
  });

  await workbook.xlsx.writeFile(outputPath);
}

// API endpoint: convert HTML to Excel
app.post('/api/convert', async (req, res) => {
  try {
    const { html_content } = req.body;
    if (!html_content) return res.status(400).json({ error: 'Missing html_content in request body' });

    // Decode base64 HTML content
    let html;
    try {
      html = Buffer.from(html_content, 'base64').toString('utf-8');
    } catch (e) {
      return res.status(400).json({ error: 'Invalid base64 format' });
    }

    // Save and convert
    const tempDir = fs.mkdtempSync(path.join(require('os').tmpdir(), 'excelconv-'));
    const excelPath = path.join(tempDir, 'converted.xlsx');
    try {
      await convertHtmlToExcel(html, excelPath);
    } catch (err) {
      return res.status(500).json({ error: 'Error converting HTML to Excel', details: err.message });
    }

    // Read Excel as base64
    const excelBuffer = fs.readFileSync(excelPath);
    const excelBase64 = excelBuffer.toString('base64');

    res.json({
      success: true,
      excel_content: excelBase64,
      filename: 'converted.xlsx',
      timestamp: new Date().toISOString()
    });
  } catch (err) {
    res.status(500).json({ error: 'Internal server error', details: err.message });
  }
});

// Start server
const PORT = process.env.PORT || 8080;
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
