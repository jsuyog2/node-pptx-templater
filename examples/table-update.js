/**
 * @fileoverview Table update example.
 *
 * Demonstrates table manipulation APIs: adding, inserting, cloning,
 * removing rows, merging cells, cell formatting, resizing and auto-fitting.
 *
 * To run: node examples/table-update.js
 */

const { PPTXTemplater } = require('../src/index.js')
const { existsSync } = require('fs')
const { resolve } = require('path')
const fsExtra = require('fs-extra')

const TEMPLATE = resolve(__dirname, '../templates/sample.pptx')
const OUTPUT = resolve(__dirname, './output/table-output.pptx')

async function main() {
  if (!existsSync(TEMPLATE)) {
    console.log('ℹ Template file not found at:', TEMPLATE)
    return
  }

  try {
    console.log('📂 Loading template:', TEMPLATE)
    const ppt = await PPTXTemplater.load(TEMPLATE)
    console.log(`📊 Loaded ${ppt.slideCount} slides`)

    // Slide 3 has our table element
    ppt.useSlide(3)
    const tableId = 'Table'

    console.log('Updating cell value and formatting at row 0, col 0...')
    ppt.updateCell(tableId, 0, 0, 'Client Name', { fill: '00FF00', fontSize: 14, align: 'ctr' })

    console.log('Adding table row to the end...')
    ppt.addTableRow(tableId, ['Acme Client', 'Consulting', 'Q1 2026', 'Completed', '$50,000'])

    console.log('Inserting table row at index 1...')
    ppt.insertTableRow(tableId, 1, ['Beta Inc', 'Development', 'Q2 2026', 'In Progress', '$75,000'])

    console.log('Cloning row 2 to index 3...')
    ppt.cloneTableRow(tableId, 2, 3)

    console.log('Merging cells at row 1, cols 0 and 1...')
    ppt.mergeCells(tableId, 1, 0, 1, 1)

    console.log('Resizing table...')
    ppt.resizeTable(tableId, 9.0, 3.5)

    console.log('Autofitting table columns...')
    ppt.autoFitTable(tableId)

    console.log('Align shapes with table cells...')
    ppt.alignShapeToCell('Shape01', 'Table', 1, 5, { horizontal: 'left', vertical: 'top' })
    ppt.alignShapeToCell('Shape02', 'Table', 2, 5, { horizontal: 'center', vertical: 'middle' })
    ppt.alignShapeToCell('Shape03', 'Table', 3, 5, { horizontal: 'right', vertical: 'bottom' })
    ppt.alignShapeToCell('Shape04', 'Table', 4, 5, { horizontal: 'center', vertical: 'bottom' })
    ppt.alignShapeToCell('Shape05', 'Table', 5, 5, { horizontal: 'left', vertical: 'bottom' })
    ppt.alignShapeToCell('Shape06', 'Table', 6, 5, { horizontal: 'center', vertical: 'top' })

    // Ensure output directory exists
    await fsExtra.ensureDir(resolve(__dirname, './output'))

    await ppt.saveToFile(OUTPUT)
    console.log('💾 Saved table output to:', OUTPUT)

    const validation = ppt.validate()
    if (validation.valid) {
      console.log('✅ Validation passed')
    } else {
      console.warn('⚠ Validation warnings:', validation.errors)
    }

    console.log('\n🎉 Table updates example completed successfully!')
  } catch (err) {
    console.error('❌ Error:', err.message)
    if (process.env.DEBUG) console.error(err.stack)
    process.exit(1)
  }
}

main()
