/**
 * @fileoverview showcase/table_advanced.js — Master Demonstration for Strict 30-Method Tables Subsystem.
 *
 * Demonstrates all 30 approved DrawingML table methods for node-pptx-templater:
 *
 * 1.  addTable(slideId, rows, columns, data, options)
 * 2.  updateTable(tableId, options)
 * 3.  updateTableData(tableId, data)
 * 4.  removeTable(tableId)
 * 5.  getTables(slideId)
 * 6.  getTable(tableId)
 * 7.  cloneTable(tableId, clonedTableId, clonedOptions)
 * 8.  addTableRow(tableId, rowIndex, data)
 * 9.  addTableColumn(tableId, columnIndex, data)
 * 10. removeTableRow(tableId, rowIndex)
 * 11. removeTableColumn(tableId, columnIndex)
 * 12. updateRowData(tableId, rowIndex, data)
 * 13. updateColumnData(tableId, columnIndex, data)
 * 14. updateTableCell(tableId, rowIndex, columnIndex, options)
 * 15. clearRowData(tableId, rowIndex)
 * 16. clearColumnData(tableId, columnIndex)
 * 17. clearCellData(tableId, rowIndex, columnIndex)
 * 18. getTableData(tableId)
 * 19. getColumnData(tableId, columnIndex)
 * 20. getRowData(tableId, rowIndex)
 * 21. getColumn(tableId, columnIndex)
 * 22. getRow(tableId, rowIndex)
 * 23. getCell(tableId, rowIndex, columnIndex)
 * 24. getCellData(tableId, rowIndex, columnIndex)
 * 25. mergeCells(tableId, startRow, startColumn, endRow, endColumn)
 * 26. splitCell(tableId, rowIndex, columnIndex)
 * 27. styleTable(tableId, options)
 * 28. styleColumn(tableId, columnIndex, options)
 * 29. styleRow(tableId, rowIndex, options)
 * 30. styleCell(tableId, rowIndex, columnIndex, options)
 *
 * How to run:
 *   node showcase/table_advanced.js
 */

'use strict'

const path = require('path')
const fs = require('fs-extra')
const { PPTXTemplater } = require('../src/index.js')

const OUTPUT_DIR = path.resolve(__dirname, 'output')
const OUTPUT_FILE = path.join(OUTPUT_DIR, 'table_advanced_output.pptx')

async function run() {
  console.log('='.repeat(75))
  console.log('node-pptx-templater: Strict 30-Method Tables Subsystem Master Showcase')
  console.log('='.repeat(75))

  await fs.ensureDir(OUTPUT_DIR)
  const startTime = Date.now()

  // Initialize a fresh presentation
  const ppt = await PPTXTemplater.createBlank()

  // Add 4 additional slides for distinct topic demonstrations
  ppt.addSlide() // Slide 2
  ppt.addSlide() // Slide 3
  ppt.addSlide() // Slide 4
  ppt.addSlide() // Slide 5

  // --------------------------------------------------------------------------
  // PART 1: Table Lifecycle & Cloning (Slide 1)
  // --------------------------------------------------------------------------
  console.log('\n[Part 1] Table Lifecycle & Cloning...')

  // 1. addTable
  ppt.addTable(
    1,
    3,
    3,
    [
      ['Department', 'Q1 Target', 'Q1 Actual'],
      ['Engineering', '$450,000', '$475,000'],
      ['Marketing', '$200,000', '$210,000'],
    ],
    {
      id: 'PerformanceTable',
      name: 'PerformanceTable',
      x: '10%',
      y: '10%',
      width: '80%',
      height: '35%',
      firstRow: true,
      bandRow: true,
    }
  )
  console.log('  ✓ 1. addTable: Created "PerformanceTable" with percentage coordinates.')

  // 2. updateTable
  ppt.updateTable('PerformanceTable', {
    y: '12%',
    bandCol: false,
  })
  console.log('  ✓ 2. updateTable: Updated position and properties.')

  // 3. updateTableData
  ppt.updateTableData('PerformanceTable', [
    ['Department', 'Q1 Target', 'Q1 Actual'],
    ['Engineering & Product', '$500,000', '$520,000'],
    ['Global Marketing', '$250,000', '$245,000'],
  ])
  console.log('  ✓ 3. updateTableData: Replaced matrix data seamlessly.')

  // 4. cloneTable
  ppt.cloneTable('PerformanceTable', 'ClonedPerformanceTable', {
    y: '52%',
  })
  console.log('  ✓ 4. cloneTable: Cloned "PerformanceTable" -> "ClonedPerformanceTable".')

  // 5. getTables
  const slide1Tables = ppt.getTables(1)
  console.log(`  ✓ 5. getTables: Discovered ${slide1Tables.length} tables on Slide 1.`)

  // 6. getTable
  const tableMeta = ppt.getTable('PerformanceTable')
  console.log(
    `  ✓ 6. getTable: Fetched metadata for "${tableMeta.name}" (${tableMeta.rows}x${tableMeta.columns}).`
  )

  // Temporary table to test removeTable
  ppt.addTable(
    1,
    2,
    2,
    [
      ['X', 'Y'],
      ['1', '2'],
    ],
    { id: 'TempTable' }
  )
  // 7. removeTable
  ppt.removeTable('TempTable')
  console.log('  ✓ 7. removeTable: Cleanly purged temporary table.')

  // --------------------------------------------------------------------------
  // PART 2: Row & Column Structural Manipulation & Data I/O (Slide 2)
  // --------------------------------------------------------------------------
  console.log('\n[Part 2] Row & Column Operations (Slide 2)...')

  ppt.addTable(
    2,
    3,
    3,
    [
      ['Product', 'Category', 'Price'],
      ['Widget A', 'Hardware', '$29.99'],
      ['Service B', 'Cloud', '$99.00'],
    ],
    { id: 'ProductTable', name: 'ProductTable', x: '10%', y: '15%', width: '80%', height: '50%' }
  )

  // 8. addTableRow
  ppt.addTableRow('ProductTable', ['Widget C', 'Hardware', '$49.99'])
  ppt.addTableRow('ProductTable', 1, ['Widget Premium', 'Hardware', '$199.99'])
  console.log('  ✓ 8. addTableRow: Appended and inserted rows.')

  // 9. addTableColumn
  ppt.addTableColumn('ProductTable', ['Stock', '450', '1,200', '800', '150'])
  console.log('  ✓ 9. addTableColumn: Added "Stock" column across all rows.')

  // 10. removeTableRow
  ppt.removeTableRow('ProductTable', 2)
  console.log('  ✓ 10. removeTableRow: Removed row at index 2.')

  // 11. removeTableColumn
  ppt.removeTableColumn('ProductTable', 1)
  console.log('  ✓ 11. removeTableColumn: Removed column at index 1 ("Category").')

  // 12. updateRowData
  ppt.updateRowData('ProductTable', 1, ['Widget Ultra', '$249.99', '500'])
  console.log('  ✓ 12. updateRowData: Updated row 1 values.')

  // 13. updateColumnData
  ppt.updateColumnData('ProductTable', 2, ['Units In Stock', '600', '900', '250'])
  console.log('  ✓ 13. updateColumnData: Updated column 2 values.')

  // 14. clearRowData
  ppt.addTableRow('ProductTable', ['Temporary', 'To Clear', '0'])
  ppt.clearRowData('ProductTable', 4)
  console.log('  ✓ 14. clearRowData: Cleared row text while keeping cell structure.')

  // 15. clearColumnData
  ppt.addTableColumn('ProductTable', ['DraftCol', 'D1', 'D2', 'D3', 'D4'])
  ppt.clearColumnData('ProductTable', 3)
  console.log('  ✓ 15. clearColumnData: Cleared column text while keeping structure.')

  // Remove temporary row and column
  ppt.removeTableRow('ProductTable', 4)
  ppt.removeTableColumn('ProductTable', 3)

  // 16. getRowData
  const rowData = ppt.getRowData('ProductTable', 0)
  console.log('  ✓ 16. getRowData: Header row is', JSON.stringify(rowData))

  // 17. getColumnData
  const colData = ppt.getColumnData('ProductTable', 0)
  console.log('  ✓ 17. getColumnData: Product names column is', JSON.stringify(colData))

  // 18. getRow
  const rowObj = ppt.getRow('ProductTable', 1)
  console.log(
    `  ✓ 18. getRow: Row 1 height=${rowObj.height} EMU, data=${JSON.stringify(rowObj.data)}`
  )

  // 19. getColumn
  const colObj = ppt.getColumn('ProductTable', 1)
  console.log(
    `  ✓ 19. getColumn: Col 1 width=${colObj.width} EMU, data=${JSON.stringify(colObj.data)}`
  )

  // --------------------------------------------------------------------------
  // PART 3: Cell Operations & Data Matrix (Slide 3)
  // --------------------------------------------------------------------------
  console.log('\n[Part 3] Cell Operations & Accessors (Slide 3)...')

  ppt.addTable(
    3,
    3,
    3,
    [
      ['Metric', 'Target', 'Status'],
      ['Uptime', '99.99%', 'Operational'],
      ['Latency', '< 20ms', 'Nominal'],
    ],
    { id: 'StatusTable', name: 'StatusTable', x: '10%', y: '15%', width: '80%', height: '40%' }
  )

  // 20. updateTableCell
  ppt.updateTableCell('StatusTable', 1, 2, {
    text: 'Optimal',
    bold: true,
    fill: 'DCFCE7',
    color: '166534',
    align: 'ctr',
    verticalAlign: 'ctr',
  })
  console.log('  ✓ 20. updateTableCell: Formatted cell (1, 2) with rich fill and typography.')

  // 21. clearCellData
  ppt.updateTableCell('StatusTable', 2, 2, { text: 'To Clear', fill: 'FEF3C7' })
  ppt.clearCellData('StatusTable', 2, 2)
  console.log('  ✓ 21. clearCellData: Emptied cell (2, 2) while keeping fill.')

  // 22. getTableData
  const matrix = ppt.getTableData('StatusTable')
  console.log('  ✓ 22. getTableData: 2D Matrix has', matrix.length, 'rows.')

  // 23. getCell
  const cellMeta = ppt.getCell('StatusTable', 1, 2)
  console.log(`  ✓ 23. getCell: Cell (1,2) text="${cellMeta.text}", fill="${cellMeta.fill}"`)

  // 24. getCellData
  const rawText = ppt.getCellData('StatusTable', 1, 1)
  console.log(`  ✓ 24. getCellData: Raw cell string at (1,1) is "${rawText}"`)

  // --------------------------------------------------------------------------
  // PART 4: Cell Merging & Splitting Subsystem (Slide 4)
  // --------------------------------------------------------------------------
  console.log('\n[Part 4] Merging & Splitting (Slide 4)...')

  ppt.addTable(
    4,
    4,
    4,
    [
      ['Q1 Total', '', '', ''],
      ['East', '$120K', 'West', '$180K'],
      ['North', '$90K', 'South', '$140K'],
      ['Summary Notes', '', '', ''],
    ],
    { id: 'SalesGrid', name: 'SalesGrid', x: '10%', y: '15%', width: '80%', height: '50%' }
  )

  // 25. mergeCells
  ppt.mergeCells('SalesGrid', 0, 0, 0, 3) // Merge entire top row
  ppt.mergeCells('SalesGrid', 3, 0, 3, 3) // Merge summary row
  console.log('  ✓ 25. mergeCells: Merged rows 0 and 3 across all 4 columns.')

  // 26. splitCell
  ppt.splitCell('SalesGrid', 3, 0) // Restore summary row to 4 independent cells
  console.log('  ✓ 26. splitCell: Restored row 3 to independent matrix cells.')

  // --------------------------------------------------------------------------
  // PART 5: Multi-Tier Styling Engine (Slide 5)
  // --------------------------------------------------------------------------
  console.log('\n[Part 5] Multi-Tier Styling Engine (Slide 5)...')

  ppt.addTable(
    5,
    5,
    4,
    [
      ['Region', 'Q1 Sales', 'Q2 Sales', 'YoY Growth'],
      ['North America', '$1,200,000', '$1,350,000', '+12.5%'],
      ['Europe', '$850,000', '$920,000', '+8.2%'],
      ['Asia-Pacific', '$1,450,000', '$1,780,000', '+22.8%'],
      ['Latin America', '$420,000', '$460,000', '+9.5%'],
    ],
    {
      id: 'StyledReportTable',
      name: 'StyledReportTable',
      x: '10%',
      y: '10%',
      width: '80%',
      height: '60%',
    }
  )

  // 27. styleTable
  ppt.styleTable('StyledReportTable', {
    firstRow: true,
    bandRow: true,
    headerFill: '1E3A8A',
    oddFill: 'F8FAFC',
    evenFill: 'FFFFFF',
    border: { color: 'CBD5E1', width: 1 },
    fontSize: 12,
  })
  console.log('  ✓ 27. styleTable: Applied alternating bands and border theme.')

  // 28. styleRow
  ppt.styleRow('StyledReportTable', 0, {
    fill: '172554',
    color: 'FFFFFF',
    bold: true,
    height: 35,
    fontSize: 13,
  })
  console.log('  ✓ 28. styleRow: Applied corporate dark header styling with white bold font.')

  // 29. styleColumn
  ppt.styleColumn('StyledReportTable', 3, {
    bold: true,
    align: 'r',
    color: '15803D',
  })
  console.log('  ✓ 29. styleColumn: Styled Growth column in bold green with right alignment.')

  // 30. styleCell
  ppt.styleCell('StyledReportTable', 3, 3, {
    fill: 'DCFCE7',
    color: '166534',
    border: { color: '22C55E', width: 2 },
  })
  console.log('  ✓ 30. styleCell: Highlighted top performing cell (3, 3).')

  // --------------------------------------------------------------------------
  // Save & Verify Output Deck
  // --------------------------------------------------------------------------
  console.log('\n[Export & Verification] Writing output deck...')
  await ppt.saveToFile(OUTPUT_FILE)
  const stats = await fs.stat(OUTPUT_FILE)
  console.log(`Saved: ${OUTPUT_FILE} (${stats.size.toLocaleString()} bytes)`)

  // Reload and validate
  const reloaded = await PPTXTemplater.load(OUTPUT_FILE)
  const validation = reloaded.validate()
  if (!validation.valid || validation.errors.length > 0) {
    console.error('Validation errors:', validation.errors)
    throw new Error('Exported presentation failed OpenXML validation!')
  }

  const duration = Date.now() - startTime
  console.log(`\n🎉 Success! All 30 table methods demonstrated and verified in ${duration}ms.`)
}

run().catch(err => {
  console.error('\n❌ Showcase failed:', err)
  process.exit(1)
})
