/**
 * Nested Table Rows Example
 *
 * Demonstrates addTableRow with nested arrays and mergeStrategy options.
 * Nested arrays indicate columns that should span multiple rows (rowspan).
 *
 * mergeStrategy options:
 *   'rowspan' — wrap nested arrays into vertically merged cells
 *   'auto'    — automatically merge cells with identical adjacent values
 *   'none'    — expand nested arrays into multiple flat rows, no merging
 *
 * Run: node examples/nested-table-rows.js
 */

'use strict'

const path = require('path')
const { PPTXTemplater } = require('../src/index.js')

const TEMPLATE = path.join(__dirname, '../templates/sample.pptx')
const OUTPUT = path.join(__dirname, 'output/nested-table-rows.pptx')

async function main() {
  const ppt = await PPTXTemplater.load(TEMPLATE, { logLevel: 'warn' })
  ppt.useSlide(3)

  const tables = ppt.getTables()
  if (tables.length === 0) {
    console.error('No table found on slide 3. Please adjust the slide index.')
    process.exit(1)
  }
  const tableId = tables[0].name
  console.log(`Working with table: "${tableId}"`)

  // 1. Simple flat row
  ppt.addTableRow(tableId, ['North', '1200', '15%', '✓', 'Active'])
  console.log('✓ Added flat row: ["North", "1200", "15%", "✓", "Active"]')

  // 2. Nested row with rowspan merging
  //    ['Region', ['Q1', 'Q2'], ...] means "Region" spans 2 rows;
  //    'Q1' and 'Q2' go into the merged region
  ppt.addTableRow(
    tableId,
    ['East Region', ['Q1 Sales', 'Q2 Sales'], ['$5,200', '$6,800'], '$12,000', ''],
    { mergeStrategy: 'rowspan' }
  )
  console.log('✓ Added nested row with rowspan strategy')

  // 3. Auto-merge: duplicate adjacent values get merged automatically
  ppt.addTableRow(tableId, ['West', 'West', '$9,100', '$9,100', 'Active'], {
    mergeStrategy: 'auto',
  })
  console.log('✓ Added row with auto-merge (duplicates collapsed)')

  // 4. No-merge: nested arrays expanded into flat rows
  ppt.addTableRow(tableId, ['South', ['$3,100', '$4,200'], 'Mixed', '', ''], {
    mergeStrategy: 'none',
  })
  console.log('✓ Added row with no-merge (nested arrays flattened)')

  await ppt.saveToFile(OUTPUT)
  console.log('✓ Done:', OUTPUT)
}

main().catch(err => {
  console.error('Error:', err.message)
  process.exit(1)
})
