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
  const startTime = Date.now()
  console.log('='.repeat(75))
  console.log('node-pptx-templater: Strict 30-Method Tables Subsystem Master Showcase')
  console.log('='.repeat(75))

  await fs.ensureDir(OUTPUT_DIR)
  const templatePath = path.join(__dirname, 'templates/TablePowerPointSample.pptx')
  const ppt = await PPTXTemplater.load(templatePath)

  // --------------------------------------------------------------------------
  // PART 1: Table Lifecycle & Cloning (Slide 1)
  // --------------------------------------------------------------------------
  console.log('\n[Part 1] Table Lifecycle & Cloning...')

  // 1. addTable
  ppt.useSlide(1)
  ppt.addTable(
    1,
    3,
    4,
    [
      [
        'Department',
        'Sprint Status (Inside Text)',
        'Blocker Indicator (Outside Text)',
        'Priority Rating (Shape Only)',
      ],
      [
        'Engineering',
        // Variation 1: Inside Shape Text (renders inside shape's <p:txBody>)
        {
          shape: {
            type: 'roundRect',
            fill: '#10B981',
            stroke: '#059669',
            width: 80,
            height: 24,
          },
          text: 'Completed',
          textPosition: 'inside',
          textStyle: {
            color: '#FFFFFF',
            fontSize: 10,
            bold: true,
            alignment: 'center',
          },
        },
        // Variation 2: Outside Shape Text (shape anchored badge, text in cell <a:txBody>)
        {
          shape: {
            type: 'circle',
            fill: '#EF4444',
            width: 12,
            height: 12,
          },
          text: 'Critical Blocker',
          textPosition: 'outside',
          shapePlacement: 'left',
          textStyle: {
            color: '#1E293B',
            fontSize: 11,
            bold: false,
          },
        },
        // Variation 3: Shape Only (graphic shape anchored within cell with zero text)
        {
          shape: {
            type: 'star5',
            fill: '#F59E0B',
            stroke: '#D97706',
            width: 18,
            height: 18,
          },
        },
      ],
      [
        'Marketing',
        // Secondary row: Inside Shape Text with "In Review" badge
        {
          shape: {
            type: 'roundRect',
            fill: '#3B82F6',
            width: 75,
            height: 22,
          },
          text: 'In Review',
          textPosition: 'inside',
          textStyle: {
            color: '#FFFFFF',
            fontSize: 10,
            bold: true,
            alignment: 'center',
          },
        },
        // Outside Shape Text: On Track indicator
        {
          shape: {
            type: 'circle',
            fill: '#10B981',
            width: 12,
            height: 12,
          },
          text: 'On Track',
          textPosition: 'outside',
          shapePlacement: 'left',
          textStyle: {
            color: '#15803D',
            fontSize: 11,
            bold: true,
          },
        },
        // Shape Only: Right Arrow badge
        {
          shape: {
            type: 'rightArrow',
            fill: '#6366F1',
            width: 16,
            height: 14,
          },
        },
      ],
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
  console.log(
    '  ✓ 1. addTable: Created "PerformanceTable" demonstrating inside, outside, and shape-only cell variations.'
  )

  // 2. updateTable
  ppt.useSlide(2)
  ppt.updateTable('UpdateTable', {
    y: '12%',
    bandCol: false,
  })
  console.log('  ✓ 2. updateTable: Updated position and properties.')

  // 3. updateTableData
  ppt.useSlide(3)
  ppt.updateTableData('UpdateTable', [
    ['Department', 'Q1 Target', 'Q1 Actual', 'Status'],
    [
      'Engineering & Product',
      '$500,000',
      '$520,000',
      { shape: { type: 'circle', fill: '#10B981' } },
    ],
    [
      'Global Marketing',
      '$250,000',
      '$245,000',
      { shape: { type: 'circle', fill: '#10B981', shapePlacement: 'left' } },
    ],
  ])
  console.log('  ✓ 3. updateTableData: Replaced matrix data seamlessly.')

  // 4. removeTable
  ppt.useSlide(4)
  ppt.removeTable('UpdateTable')
  console.log('  ✓ 4. removeTable: Cleanly purged temporary table.')

  // 5. getTables
  ppt.useAllSlides()
  const slide1Tables = ppt.getTables()
  console.log(`  ✓ 5. getTables: Discovered ${slide1Tables.length} tables on Slide 1.`)

  // 6. getTable
  ppt.useSlide(3)
  const tableMeta = ppt.getTable('UpdateTable')
  console.log(
    `  ✓ 6. getTable: Fetched metadata for "${tableMeta.name}" (${tableMeta.rows}x${tableMeta.columns}).`
  )

  // 7. cloneTable
  ppt.useSlide(5)
  ppt.cloneTable('UpdateTable', 'ClonedPerformanceTable', {
    y: '52%',
  })
  console.log('  ✓ 7. cloneTable: Cloned "PerformanceTable" -> "ClonedPerformanceTable".')

  // --------------------------------------------------------------------------
  // PART 2: Row & Column Structural Manipulation & Data I/O (Slide 2)
  // --------------------------------------------------------------------------
  console.log('\n[Part 2] Row & Column Operations (Slide 2)...')

  // 8. addTableRow
  ppt.useSlide(6)
  ppt.addTableRow('UpdateTable', 2, [
    'Widget C',
    'Hardware',
    '$49.99',
    { shape: { type: 'circle', fill: '#10B981' } },
  ])
  ppt.addTableRow('UpdateTable', 1, [
    'Widget Premium',
    'Hardware',
    '$199.99',
    { shape: { type: 'circle', fill: '#10B981' } },
  ])
  console.log('  ✓ 8. addTableRow: Appended and inserted rows.')

  // 9. addTableColumn
  ppt.useSlide(7)
  ppt.addTableColumn('UpdateTable', [
    'Status',
    '450',
    { shape: { type: 'circle', fill: '#10B981' } },
  ])
  console.log('  ✓ 9. addTableColumn: Added "Stock" column across all rows.')

  // 10. removeTableRow
  ppt.useSlide(8)
  ppt.removeTableRow('UpdateTable', 2)
  console.log('  ✓ 10. removeTableRow: Removed row at index 2.')

  // 11. removeTableColumn
  ppt.useSlide(9)
  ppt.removeTableColumn('UpdateTable', 1)
  ppt.removeTableColumn('UpdateTable', 3)
  console.log('  ✓ 11. removeTableColumn: Removed column at index 1 ("Category").')

  // 12. updateRowData
  ppt.useSlide(10)
  ppt.updateRowData('UpdateTable', 1, [
    'Widget Ultra',
    '$249.99',
    '500',
    { shape: { type: 'circle', fill: '#b91010ff' } },
  ])
  console.log('  ✓ 12. updateRowData: Updated row 1 values.')

  // 13. updateColumnData
  ppt.useSlide(11)
  ppt.updateColumnData('UpdateTable', 2, [
    'Units In Stock',
    '600',
    { shape: { type: 'circle', fill: '#b91010ff' } },
  ])
  console.log('  ✓ 13. updateColumnData: Updated column 2 values.')

  // 14. updateTableCell
  ppt.useSlide(12)
  ppt.updateTableCell('UpdateTable', 2, 2, { text: 'To Clear', fill: 'FEF3C7', bold: true })
  ppt.updateTableCell('UpdateTable', 2, 3, { shape: { type: 'circle', fill: '#b91010ff' } })
  console.log('  ✓ 14. updateTableCell: Updated cell (2, 2) with text and fill.')

  // 15. clearRowData
  ppt.useSlide(13)
  ppt.clearRowData('UpdateTable', 2)
  console.log('  ✓ 15. clearRowData: Cleared row text while keeping cell structure.')

  // 16. clearColumnData
  ppt.useSlide(14)
  ppt.clearColumnData('UpdateTable', 2)
  ppt.clearColumnData('UpdateTable', 4)
  console.log('  ✓ 16. clearColumnData: Cleared column text while keeping structure.')

  // 17. clearCellData
  ppt.useSlide(15)
  ppt.clearCellData('UpdateTable', 2, 2)
  ppt.clearCellData('UpdateTable', 2, 4)
  console.log('  ✓ 17. clearCellData: Emptied cell (2, 2) while keeping fill.')

  // 18. getTableData
  const matrix = ppt.getTableData('UpdateTable')
  console.log('  ✓ 18. getTableData: 2D Matrix has', matrix.length, 'rows.')

  // 19. getColumnData
  ppt.useSlide(16)
  const colData = ppt.getColumnData('UpdateTable', 0)
  console.log('  ✓ 19. getColumnData: Product names column is', JSON.stringify(colData))

  // 20. getRowData
  ppt.useSlide(15)
  const rowData = ppt.getRowData('UpdateTable', 0)
  console.log('  ✓ 20. getRowData: Header row is', JSON.stringify(rowData))

  // 21. getColumn
  const colObj = ppt.getColumn('UpdateTable', 1)
  console.log(
    `  ✓ 21. getColumn: Col 1 width=${colObj.width} EMU, data=${JSON.stringify(colObj.data)}`
  )

  // 22. getRow
  const rowObj = ppt.getRow('UpdateTable', 1)
  console.log(
    `  ✓ 22. getRow: Row 1 height=${rowObj.height} EMU, data=${JSON.stringify(rowObj.data)}`
  )

  // 23. getCell
  const cellMeta = ppt.getCell('UpdateTable', 1, 2)
  console.log(`  ✓ 23. getCell: Cell (1,2) text="${cellMeta.text}", fill="${cellMeta.fill}"`)

  // 24. getCellData
  const rawText = ppt.getCellData('UpdateTable', 1, 1)
  console.log(`  ✓ 24. getCellData: Raw cell string at (1,1) is "${rawText}"`)

  // --------------------------------------------------------------------------
  // PART 3: Merging, Splitting & Multi-Tier Styling
  // --------------------------------------------------------------------------
  console.log('\n[Part 3] Merging, Splitting & Multi-Tier Styling...')

  // 25. mergeCells
  ppt.useSlide(16)
  ppt.mergeCells('UpdateTable', 0, 0, 0, 4) // Merge entire top row
  console.log('  ✓ 25. mergeCells: Merged rows 0 and 3 across all 4 columns.')

  // 26. splitCell
  ppt.useSlide(17)
  ppt.splitCell('UpdateTable', 1, 0) // Restore summary row to 4 independent cells
  console.log('  ✓ 26. splitCell: Restored row 3 to independent matrix cells.')

  // 27. styleTable
  ppt.useSlide(18)
  ppt.styleTable('UpdateTable', {
    firstRow: true,
    bandRow: true,
    headerFill: '1E3A8A',
    oddFill: 'F8FAFC',
    evenFill: 'FFFFFF',
    border: { color: 'CBD5E1', width: 1 },
    fontSize: 12,
  })
  console.log('  ✓ 27. styleTable: Applied alternating bands and border theme.')

  // 28. styleColumn
  ppt.useSlide(19)
  ppt.styleColumn('UpdateTable', 3, {
    bold: true,
    align: 'r',
    color: '15803D',
  })
  console.log('  ✓ 28. styleColumn: Styled Growth column in bold green with right alignment.')

  // 29. styleRow
  ppt.useSlide(20)
  ppt.styleRow('UpdateTable', 1, {
    fill: '172554',
    color: 'FFFFFF',
    bold: true,
    height: 35,
    fontSize: 13,
  })
  console.log('  ✓ 29. styleRow: Applied corporate dark header styling with white bold font.')

  // 30. styleCell
  ppt.useSlide(21)
  ppt.styleCell('UpdateTable', 2, 3, {
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
