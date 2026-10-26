/**
 * @fileoverview Consolidated Demonstrations for node-pptx-templater
 *
 * This single, self-contained file demonstrates all key capabilities of
 * node-pptx-templater v3.0.0. All required JSON data is embedded directly
 * below, and all required templates are referenced relative to this file.
 *
 * Requirements:
 *   Node.js >= 20.12.0
 *
 * How to run:
 *   node showcase/examples.js
 *   or:
 *   npm run examples
 *
 * Generated presentations are saved to: showcase/output/
 */

'use strict'

const path = require('path')
const fs = require('fs-extra')
const { PPTXTemplater } = require('../src/index.js')

// ============================================================================
// EMBEDDED JSON INPUT DATA
// ============================================================================

const SAMPLE_DATA = {
  metadata: {
    title: 'Executive Quarterly Business Review',
    subtitle: 'Strategic Performance & Operations Update',
    author: 'Operations & Engineering Team',
    organization: 'Acme Technologies Global Inc.',
    quarter: 'Q3 2026',
    date: 'October 11, 2026',
    confidentiality: 'CONFIDENTIAL — FOR INTERNAL USE ONLY',
    currency: 'USD ($M)',
  },

  kpis: [
    {
      metric: 'Annual Recurring Revenue (ARR)',
      value: '$124.8M',
      change: '+38.2% YoY',
      status: 'Exceeded',
    },
    {
      metric: 'Net Revenue Retention (NRR)',
      value: '118.4%',
      change: '+2.1% QoQ',
      status: 'On Target',
    },
    {
      metric: 'Customer Acquisition Cost (CAC)',
      value: '$12,400',
      change: '-14.5% YoY',
      status: 'Optimized',
    },
    { metric: 'Gross Operating Margin', value: '78.6%', change: '+3.4% YoY', status: 'Healthy' },
  ],

  financials: {
    categories: ['Q1 2026', 'Q2 2026', 'Q3 2026', 'Q4 2026 (Proj)'],
    series: [
      { name: 'Enterprise Cloud', values: [18.4, 22.1, 26.8, 31.5] },
      { name: 'AI Solutions', values: [8.2, 12.6, 17.4, 23.0] },
      { name: 'Professional Services', values: [5.1, 5.4, 5.8, 6.2] },
    ],
  },

  projectStatusTable: {
    headers: ['Workstream', 'Initiative', 'Lead', 'Timeline', 'Budget'],
    rows: [
      ['Core Platform', 'Zero-Downtime Multi-Region Migration', 'Sarah Jenkins', 'Q1-Q3', '$1.8M'],
      ['Core Platform', 'Automated Failover Verification', 'Marcus Vance', 'Q2-Q3', '$450K'],
      ['AI & Analytics', 'Real-time Predictive Analytics API', 'Elena Rostova', 'Q3-Q4', '$2.4M'],
      ['AI & Analytics', 'Enterprise LLM Fine-Tuning Pipeline', 'Devon Patel', 'Q3-Q4', '$1.2M'],
      ['Security', 'FedRAMP High Authorization Preparation', 'Alicia Gomez', 'Q2-Q4', '$950K'],
      ['Security', 'Zero-Trust IAM Microsegmentation', 'Kenji Sato', 'Q3-Q4', '$620K'],
    ],
  },

  nestedDepartmentTable: [
    {
      department: 'Engineering',
      teams: [
        { team: 'Platform Architecture', headcount: 24, lead: 'Alex Mercer', openRoles: 3 },
        { team: 'Data Infrastructure', headcount: 18, lead: 'Priya Sharma', openRoles: 2 },
        { team: 'Frontend Experience', headcount: 16, lead: 'Liam O’Connor', openRoles: 1 },
      ],
    },
    {
      department: 'Product & Design',
      teams: [
        { team: 'Enterprise Core UI', headcount: 12, lead: 'Hanna Lindqvist', openRoles: 2 },
        { team: 'Design Systems & Tokens', headcount: 8, lead: 'Carlos Ortiz', openRoles: 1 },
      ],
    },
  ],

  teamRoster: [
    { name: 'Dr. Evelyn Reed', title: 'VP of Platform Engineering', location: 'San Francisco, CA' },
    { name: 'David Zhao', title: 'Principal Systems Architect', location: 'Seattle, WA' },
    { name: 'Amara Okafor', title: 'Head of Infrastructure Security', location: 'Austin, TX' },
    { name: 'Nadia Benali', title: 'Senior Director of Product Design', location: 'New York, NY' },
  ],
}

// ============================================================================
// PATH RESOLUTION & SETUP
// ============================================================================

const SHOWCASE_DIR = __dirname
const TEMPLATES_DIR = path.join(SHOWCASE_DIR, 'templates')
const OUTPUT_DIR = path.join(SHOWCASE_DIR, 'output')

const TEMPLATE_SAMPLE = path.join(TEMPLATES_DIR, 'sample.pptx')
const TEMPLATE_OFFICE = path.join(TEMPLATES_DIR, 'officePPT.pptx')
const IMAGE_CHESS = path.join(TEMPLATES_DIR, 'chess.png')
const IMAGE_ROOK = path.join(TEMPLATES_DIR, 'rook.png')

// 1x1 transparent WebP data URI for multi-format demonstrations
const WEBP_BASE64_DATA =
  'data:image/webp;base64,UklGRkAAAABXRUJQVlA4IDQAAADwAQCdASoBAAEAAQAcJaACdLoAAP7/1/4AAAD//v/9AAAA'

// ============================================================================
// DEMONSTRATION 1: BASIC REPORT GENERATION & TEXT REPLACEMENT
// ============================================================================

async function demo01_BasicTextReplacement() {
  console.log('\n--- [Demo 01] Basic Report Generation & Text Replacement ---')
  const ppt = await PPTXTemplater.load(TEMPLATE_SAMPLE, { logLevel: 'warn' })
  ppt.useAllSlides()

  // 1. Inspect slide dimensions
  const dims = ppt.getSlideDimensions()
  console.log(`✓ Slide dimensions: ${dims.width} x ${dims.height} EMUs`)

  // 2. Discover placeholders across slides
  const placeholders = ppt.findText('{{')
  console.log(`✓ Discovered ${placeholders.length} placeholder tag(s) across presentation.`)

  // 3. Perform text tag replacements across all slides
  const replacements = {
    '{{title}}': SAMPLE_DATA.metadata.title,
    '{{company}}': SAMPLE_DATA.metadata.organization,
    '{{quarter}}': SAMPLE_DATA.metadata.quarter,
    '{{date}}': SAMPLE_DATA.metadata.date,
    '{{author}}': SAMPLE_DATA.metadata.author,
    '{{confidentiality}}': SAMPLE_DATA.metadata.confidentiality,
  }

  ppt.replaceText(replacements)
  console.log(`✓ Replaced placeholder tags across all slides.`)

  const outPath = path.join(OUTPUT_DIR, '01-basic-report.pptx')
  await ppt.saveToFile(outPath)
  console.log(`✓ Saved: ${outPath}`)
  return outPath
}

// ============================================================================
// DEMONSTRATION 2: CHART UPDATES & DATA LABELS WITH EXCEL CACHE
// ============================================================================

async function demo02_ChartUpdate() {
  console.log('\n--- [Demo 02] Chart Updates & Data Labels with Excel Workbook Sync ---')
  const ppt = await PPTXTemplater.load(TEMPLATE_SAMPLE, { logLevel: 'warn' })
  ppt.useSlide(1)

  // Update chart data on Slide 1 using embedded JSON financial series
  const chartData = {
    categories: SAMPLE_DATA.financials.categories,
    series: SAMPLE_DATA.financials.series,
  }

  ppt.updateChart('chart1', chartData)
  console.log('✓ Updated chart "chart1" series data and synchronized internal Excel cache.')

  // Update data labels visibility
  try {
    ppt.updateDataLabels('chart1', {
      showVal: true,
      showCatName: false,
      showLegendKey: false,
    })
    console.log('✓ Configured chart data labels (showVal=true, showCatName=false).')
  } catch (err) {
    console.log(`  (Note: Data labels update skipped: ${err.message})`)
  }

  const outPath = path.join(OUTPUT_DIR, '02-chart-update.pptx')
  await ppt.saveToFile(outPath)
  console.log(`✓ Saved: ${outPath}`)
  return outPath
}

// ============================================================================
// DEMONSTRATION 3: DYNAMIC TABLES, FORMATTING, STRIPING & BORDERS
// ============================================================================

async function demo03_TableOperations() {
  console.log('\n--- [Demo 03] Table Operations: Creation, Columns, Row Striping & Borders ---')
  const ppt = await PPTXTemplater.load(TEMPLATE_SAMPLE, { logLevel: 'warn' })
  // Slide 3 contains the template Table
  ppt.useSlide(3)

  const tables = ppt.getTables()
  if (tables.length === 0) {
    throw new Error('No table found on Slide 3 in template')
  }
  const tableId = tables[0].name

  // 1. Update existing table with financial/initiative data
  const tableData = [SAMPLE_DATA.projectStatusTable.headers, ...SAMPLE_DATA.projectStatusTable.rows]

  ppt.updateTable(tableId, tableData)
  console.log(`✓ Updated table "${tableId}" with ${tableData.length} rows.`)

  // 2. Append an additional summary row
  ppt.addTableRow(tableId, ['Totals', '6 Active Initiatives', 'All Leads', 'FY26', '$7.42M'])
  console.log('✓ Appended summary row via addTableRow().')

  // 3. Apply table styling (alternating row background fills)
  ppt.styleTable(tableId, {
    bandRow: true,
    oddFill: 'F8FAFC',
    evenFill: 'EDF2F7',
  })
  console.log('✓ Applied alternating row striping (#F8FAFC / #EDF2F7).')

  // 4. Set custom cell border on header cell via updateTableCell
  ppt.updateTableCell(tableId, 0, 0, {
    text: 'Workstream',
    border: { bottom: { color: '3B82F6', width: 2 } },
  })
  console.log('✓ Applied custom bottom border to table header.')

  const outPath = path.join(OUTPUT_DIR, '03-table-operations.pptx')
  await ppt.saveToFile(outPath)
  console.log(`✓ Saved: ${outPath}`)
  return outPath
}

// ============================================================================
// DEMONSTRATION 4: NESTED TABLES, CELL MERGING & STRUCTURED EXTRACTION
// ============================================================================

async function demo04_NestedTablesAndExtraction() {
  console.log('\n--- [Demo 04] Nested Tables, Vertical Cell Merging & Structured Extraction ---')
  const ppt = await PPTXTemplater.load(TEMPLATE_SAMPLE, { logLevel: 'warn' })
  ppt.useSlide(3)

  const tables = ppt.getTables()
  const tableId = tables[0].name

  // 1. Add rows
  ppt.addTableRow(tableId, ['East Region', 'Q1 Sales', '$5,200', '$12,000', 'Active'])
  ppt.addTableRow(tableId, ['East Region', 'Q2 Sales', '$6,800', '$12,000', 'Active'])
  console.log('✓ Added regional rows via addTableRow().')

  // 2. Merge region cells
  ppt.mergeCells(tableId, 1, 0, 2, 0)
  console.log('✓ Merged adjacent cells via mergeCells().')

  // 3. Extract table data matrix
  const extracted = ppt.getTableData(tableId)
  console.log(`✓ Extracted table matrix from "${tableId}": ${extracted.length} rows.`)

  const outPath = path.join(OUTPUT_DIR, '04-nested-table-merges.pptx')
  await ppt.saveToFile(outPath)
  console.log(`✓ Saved: ${outPath}`)
  return outPath
}

// ============================================================================
// DEMONSTRATION 5: SHAPES, Z-ORDER, ALIGNMENT & DISTRIBUTION
// ============================================================================

async function demo05_ShapesAndZOrder() {
  console.log('\n--- [Demo 05] Shape Operations, Z-Order, Spatial Alignment & Distribution ---')
  const ppt = await PPTXTemplater.load(TEMPLATE_SAMPLE, { logLevel: 'warn' })
  ppt.useSlide(1)

  // 1. Add background decorative card shape
  await ppt.addShape('rectangle', {
    id: 'CardBackground',
    name: 'CardBackground',
    x: 50,
    y: 100,
    width: 400,
    height: 220,
    fill: '#1E293B',
  })
  console.log('✓ Added background card shape.')

  // 2. Add badge pill shape
  await ppt.addShape('roundedRectangle', {
    id: 'StatusBadge',
    name: 'StatusBadge',
    x: 70,
    y: 120,
    width: 120,
    height: 32,
    fill: '#3B82F6',
    text: 'ACTIVE KPI',
  })
  console.log('✓ Added status badge pill shape.')

  // 3. Add secondary shape
  await ppt.addShape('rectangle', {
    id: 'KpiStatBox',
    name: 'KpiStatBox',
    x: 220,
    y: 120,
    width: 200,
    height: 160,
    fill: '#0F172A',
    text: '+38.2% ARR GROWTH',
  })
  console.log('✓ Added stat highlight box.')

  // 4. Update shape coordinates via updateShape
  ppt.updateShape('StatusBadge', { y: 120 })
  console.log('✓ Aligned shape coordinates via updateShape.')

  // 5. Demonstrate Z-Ordering
  ppt.sendToBack('CardBackground')
  ppt.bringToFront('StatusBadge')
  console.log(
    '✓ Reordered visual layering: CardBackground sent to back, StatusBadge brought to front.'
  )

  const outPath = path.join(OUTPUT_DIR, '05-shape-operations-and-zorder.pptx')
  await ppt.saveToFile(outPath)
  console.log(`✓ Saved: ${outPath}`)
  return outPath
}

// ============================================================================
// DEMONSTRATION 6: CELL SHAPES & CELL BOUNDS INSPECTION
// ============================================================================

async function demo06_CellShapes() {
  console.log('\n--- [Demo 06] Table Cell Formatting & Status Badging ---')
  const ppt = await PPTXTemplater.load(TEMPLATE_SAMPLE, { logLevel: 'warn' })
  ppt.useSlide(3)

  const tables = ppt.getTables()
  const tableId = tables[0].name

  // Add row with formatted cell values
  ppt.addTableRow(tableId, ['● Active', 'Production Cluster', 'US-East-1', 'Operational', '100%'])
  ppt.styleCell(tableId, 1, 0, { fill: 'DCFCE7', color: '166534', bold: true })
  console.log('✓ Added table row with styled status cell.')

  const outPath = path.join(OUTPUT_DIR, '06-cell-shapes.pptx')
  await ppt.saveToFile(outPath)
  console.log(`✓ Saved: ${outPath}`)
  return outPath
}

// ============================================================================
// DEMONSTRATION 7: IMAGE OPERATIONS & MULTI-FORMAT MEDIA INSERTION
// ============================================================================

async function demo07_ImageOperations() {
  console.log('\n--- [Demo 07] Image Operations: PNG, WebP & Base64 Insertion ---')
  const ppt = await PPTXTemplater.load(TEMPLATE_SAMPLE, { logLevel: 'warn' })
  ppt.useSlide(1)

  // 1. Add a new PNG image with explicit position and size
  if (fs.existsSync(IMAGE_ROOK)) {
    await ppt.AddImage(1, IMAGE_ROOK, {
      x: 914400,
      y: 3500000,
      width: 1200000,
      height: 1200000,
      id: 'demo-logo',
    })
    console.log(`✓ Added new PNG image from file with dimensions.`)

    // 2. Replace that image with another image using UpdateImage
    if (fs.existsSync(IMAGE_CHESS)) {
      await ppt.UpdateImage('demo-logo', { source: IMAGE_CHESS })
      console.log(`✓ Replaced image "demo-logo" with chess image.`)
    }
  }

  // 3. Add WebP / Base64 image
  try {
    await ppt.AddImage(1, WEBP_BASE64_DATA, {
      x: 2500000,
      y: 3500000,
      width: 800000,
      height: 800000,
      id: 'webp-icon',
    })
    console.log('✓ Added WebP image from Base64 Data URI.')
  } catch (err) {
    console.log(`  (Base64 WebP insert: ${err.message})`)
  }

  const outPath = path.join(OUTPUT_DIR, '07-images-and-media.pptx')
  await ppt.saveToFile(outPath)
  console.log(`✓ Saved: ${outPath}`)
  return outPath
}

// ============================================================================
// DEMONSTRATION 8: SLIDE ARCHITECTURE, DUPLICATION & CROSS-PRESENTATION IMPORT
// ============================================================================

async function demo08_SlideOperations() {
  console.log('\n--- [Demo 08] Slide Management: Duplication, Reordering & Cross-Deck Import ---')
  const ppt = await PPTXTemplater.load(TEMPLATE_SAMPLE, { logLevel: 'warn' })

  console.log(`✓ Initial slide count: ${ppt.slideCount}`)

  // 1. Duplicate Slide 1
  ppt.duplicateSlide(1)
  console.log(`✓ Duplicated Slide 1. New slide count: ${ppt.slideCount}`)

  // 2. Reorder slides: shift first slide to the end
  const currentCount = ppt.slideCount
  const newOrder = Array.from({ length: currentCount }, (_, i) => i + 1)
  newOrder.push(newOrder.shift()) // shift first slide to last
  ppt.reorderSlides(newOrder)
  console.log(`✓ Reordered ${currentCount} slides array.`)

  // 3. Cross-presentation slide import from corporate template
  if (fs.existsSync(TEMPLATE_OFFICE)) {
    try {
      const sourceDeck = await PPTXTemplater.load(TEMPLATE_OFFICE, { logLevel: 'warn' })
      await ppt.importSlideFrom(sourceDeck, 1)
      console.log(`✓ Imported slide 1 from corporate template. New slide count: ${ppt.slideCount}`)
    } catch (err) {
      console.log(`  (Cross-deck slide import: ${err.message})`)
    }
  }

  const outPath = path.join(OUTPUT_DIR, '08-slide-operations.pptx')
  await ppt.saveToFile(outPath)
  console.log(`✓ Saved: ${outPath}`)
  return outPath
}

// ============================================================================
// DEMONSTRATION 9: HYPERLINKS & INTERACTIVE PRESENTATION NAVIGATION
// ============================================================================

async function demo09_HyperlinksAndNavigation() {
  console.log('\n--- [Demo 09] Hyperlinks & Interactive Navigation Actions ---')
  const ppt = await PPTXTemplater.load(TEMPLATE_SAMPLE, { logLevel: 'warn' })
  ppt.useSlide(1)

  // 1. Add shape hyperlink on Slide 1
  try {
    ppt.addHyperlink('Title', 2, { tooltip: 'Jump to Slide 2' })
    console.log('✓ Added shape slide hyperlink.')
  } catch (err) {
    console.log(`  (addHyperlink: ${err.message})`)
  }

  // 2. Add slide link
  try {
    ppt.addHyperlink('first', 3, { tooltip: 'Jump to Slide 3' })
    console.log('✓ Added internal slide jump link to Slide 3.')
  } catch (err) {
    console.log(`  (addHyperlink: ${err.message})`)
  }

  // 3. Inspect existing hyperlinks on slide
  const links = ppt.getHyperlinks(1)
  console.log(`✓ Discovered ${links.length} hyperlink relationship(s) on Slide 1.`)

  const outPath = path.join(OUTPUT_DIR, '09-hyperlinks-and-navigation.pptx')
  await ppt.saveToFile(outPath)
  console.log(`✓ Saved: ${outPath}`)
  return outPath
}

// ============================================================================
// DEMONSTRATION 10: RICH TYPOGRAPHY & MULTI-STYLE TEXT RUNS
// ============================================================================

async function demo10_TypographyAndTextRuns() {
  console.log('\n--- [Demo 10] Rich Typography & Text Runs (Bold, Color, Fonts) ---')
  const ppt = await PPTXTemplater.load(TEMPLATE_SAMPLE, { logLevel: 'warn' })
  ppt.useSlide(1)

  // 1. Add shape with advanced typography and text styles
  ppt.addShape('rectangle', {
    id: 'TypographyBox',
    x: 60,
    y: 200,
    width: 450,
    height: 120,
    fill: '#F1F5F9',
    text: 'SYSTEM STATUS: PRODUCTION OPERATIONAL',
    textStyle: {
      fontSize: 16,
      bold: true,
      color: '#059669',
      fontFamily: 'Helvetica',
      align: 'center',
    },
  })
  console.log('✓ Added container shape with styled typography.')

  const outPath = path.join(OUTPUT_DIR, '10-typography-text-runs.pptx')
  await ppt.saveToFile(outPath)
  console.log(`✓ Saved: ${outPath}`)
  return outPath
}

// ============================================================================
// DEMONSTRATION 11: XML FOLDER WORKFLOW (EXTRACT, EDIT & REBUILD)
// ============================================================================

async function demo11_XmlFolderWorkflow() {
  console.log('\n--- [Demo 11] OpenXML Folder Workflow: Extract, Inspect, Edit & Rebuild ---')
  const workFolder = path.join(OUTPUT_DIR, 'xml-extracted-demo')
  const outPath = path.join(OUTPUT_DIR, '11-xml-folder-rebuilt.pptx')

  // 1. Extract PPTX to folder
  await PPTXTemplater.extractPptx(TEMPLATE_SAMPLE, workFolder, { overwrite: true })
  console.log(`✓ Extracted presentation OpenXML to folder: ${workFolder}`)

  // 2. Load directly from folder
  const ppt = await PPTXTemplater.load(workFolder, { logLevel: 'warn' })
  ppt.useAllSlides()
  ppt.replaceTextByTag('Title', 'Acme OpenXML Unpacked Systems')
  console.log('✓ Loaded directly from extracted XML folder and replaced placeholders.')

  // 3. Save modified XML in-place
  await ppt.extractToXml(workFolder)
  console.log('✓ Saved modifications back to XML folder.')

  // 4. Rebuild folder back into binary PPTX
  await PPTXTemplater.buildPptx(workFolder, outPath)
  console.log(`✓ Rebuilt binary presentation from folder: ${outPath}`)

  // 5. Clean up temporary extraction folder
  await fs.remove(workFolder)
  console.log('✓ Cleaned up temporary XML extraction directory.')

  return outPath
}

// ============================================================================
// DEMONSTRATION 12: SERVERLESS / BUFFER EXPORT FOR CLOUD & STREAMING
// ============================================================================

async function demo12_BufferExport() {
  console.log('\n--- [Demo 12] In-Memory Buffer Export (AWS Lambda / Serverless / Streaming) ---')
  const ppt = await PPTXTemplater.load(TEMPLATE_SAMPLE, { logLevel: 'warn' })
  ppt.useSlide(1)
  ppt.replaceTextByTag('Title', 'Serverless Generated Report')

  // Export directly to Buffer in memory without disk I/O
  const buffer = await ppt.toBuffer()
  console.log(`✓ Exported presentation to in-memory Buffer (${buffer.length} bytes).`)

  // Verify ZIP magic bytes (0x50 0x4B 0x03 0x04)
  const isZip = buffer[0] === 0x50 && buffer[1] === 0x4b && buffer[2] === 0x03 && buffer[3] === 0x04
  console.log(`✓ Verified valid OpenXML ZIP binary signature: ${isZip ? 'VALID' : 'INVALID'}`)

  const outPath = path.join(OUTPUT_DIR, '12-serverless-buffer.pptx')
  await fs.writeFile(outPath, buffer)
  console.log(`✓ Written Buffer to test output file: ${outPath}`)
  return outPath
}

// ============================================================================
// MAIN RUNNER
// ============================================================================

async function runAllDemos() {
  const startTime = Date.now()
  console.log('================================================================')
  console.log(' node-pptx-templater v3.0.0 — Showcase & Feature Demonstrations')
  console.log('================================================================')
  console.log(`Node.js version : ${process.version}`)
  console.log(`Working dir     : ${process.cwd()}`)
  console.log(`Templates dir   : ${TEMPLATES_DIR}`)
  console.log(`Output dir      : ${OUTPUT_DIR}`)

  // Ensure output directory exists
  await fs.ensureDir(OUTPUT_DIR)

  const demos = [
    { name: '01-BasicTextReplacement', fn: demo01_BasicTextReplacement },
    { name: '02-ChartUpdate', fn: demo02_ChartUpdate },
    { name: '03-TableOperations', fn: demo03_TableOperations },
    { name: '04-NestedTablesAndExtraction', fn: demo04_NestedTablesAndExtraction },
    { name: '05-ShapesAndZOrder', fn: demo05_ShapesAndZOrder },
    { name: '06-CellShapes', fn: demo06_CellShapes },
    { name: '07-ImageOperations', fn: demo07_ImageOperations },
    { name: '08-SlideOperations', fn: demo08_SlideOperations },
    { name: '09-HyperlinksAndNavigation', fn: demo09_HyperlinksAndNavigation },
    { name: '10-TypographyAndTextRuns', fn: demo10_TypographyAndTextRuns },
    { name: '11-XmlFolderWorkflow', fn: demo11_XmlFolderWorkflow },
    { name: '12-BufferExport', fn: demo12_BufferExport },
  ]

  const results = []
  let failed = 0

  for (const demo of demos) {
    const demoStart = Date.now()
    try {
      const out = await demo.fn()
      results.push({
        name: demo.name,
        status: 'PASSED',
        durationMs: Date.now() - demoStart,
        output: out,
      })
    } catch (err) {
      console.error(`\n❌ Error in demo "${demo.name}":`, err.stack || err.message)
      results.push({
        name: demo.name,
        status: 'FAILED',
        durationMs: Date.now() - demoStart,
        error: err.message,
      })
      failed++
    }
  }

  const totalTime = ((Date.now() - startTime) / 1000).toFixed(2)
  console.log('\n================================================================')
  console.log(' SHOWCASE EXECUTION SUMMARY')
  console.log('================================================================')
  for (const res of results) {
    const icon = res.status === 'PASSED' ? '✅' : '❌'
    console.log(`${icon} ${res.name.padEnd(30)} ${res.status.padEnd(8)} (${res.durationMs}ms)`)
  }
  console.log('----------------------------------------------------------------')
  console.log(
    `Total: ${demos.length} demos | Passed: ${demos.length - failed} | Failed: ${failed} | Duration: ${totalTime}s`
  )
  console.log(`Generated presentations stored in: ${OUTPUT_DIR}`)
  console.log('================================================================\n')

  if (failed > 0) {
    process.exit(1)
  }
}

// Execute runner if invoked directly
if (require.main === module) {
  runAllDemos().catch(err => {
    console.error('Fatal execution error:', err)
    process.exit(1)
  })
}

module.exports = {
  SAMPLE_DATA,
  runAllDemos,
}
