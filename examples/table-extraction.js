/**
 * Table Data Extraction Example
 *
 * Demonstrates: getTables, getTableRows (all 3 modes),
 *               getMergedCells, isMergedCell
 *
 * Run: node examples/table-extraction.js
 */

'use strict'

const path = require('path')
const { PPTXTemplater } = require('../src/index.js')

const TEMPLATE = path.join(__dirname, '../templates/sample.pptx')

async function main() {
  const ppt = await PPTXTemplater.load(TEMPLATE, { logLevel: 'warn' })
  ppt.useSlide(3)

  // 1. Discover all tables on slide 3
  const tables = ppt.getTables()
  console.log(`Found ${tables.length} table(s) on slide 3:`)
  tables.forEach(t => console.log(`  • "${t.name}" (${t.rows} rows × ${t.cols} cols)`))

  if (tables.length === 0) {
    console.log('No tables found — try adjusting the slide index.')
    return
  }

  const tableId = tables[0].name

  // 2. Default mode: object array (first row = headers)
  const objRows = await ppt.getTableRows(tableId)
  console.log(`\n--- Object mode (${objRows.length} rows) ---`)
  console.log(JSON.stringify(objRows.slice(0, 3), null, 2))

  // 3. Raw mode: array of string arrays
  const rawRows = await ppt.getTableRows(tableId, { raw: true })
  console.log(`\n--- Raw mode (${rawRows.length} rows) ---`)
  rawRows.slice(0, 3).forEach((row, i) => console.log(`Row ${i}:`, row))

  // 4. Metadata mode: includes row/column count and merge info
  const meta = await ppt.getTableRows(tableId, { includeMetadata: true })
  console.log(`\n--- Metadata mode ---`)
  console.log(`Rows: ${meta.rowCount}, Columns: ${meta.columnCount}`)
  console.log(`Merged cells: ${meta.mergedCells?.length ?? 0}`)

  // 5. Check for merged cells
  const merges = ppt.getMergedCells(tableId)
  console.log(`\n--- Merged cells in "${tableId}" ---`)
  if (merges.length === 0) {
    console.log('  No merged cells.')
  } else {
    merges.forEach(m =>
      console.log(`  • Row ${m.startRow}:${m.startCol} → ${m.endRow}:${m.endCol}`)
    )
  }

  console.log('\n✓ Done (read-only — no output file saved)')
}

main().catch(err => {
  console.error('Error:', err.message)
  process.exit(1)
})
