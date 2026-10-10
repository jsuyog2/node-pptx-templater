import { describe, it, expect } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')

describe('Table Creation, Columns, Striping, and Summary Rows Integration Tests', () => {
  const runTests = existsSync(FIXTURE_FILE)

  if (!runTests) {
    it.skip('Skipping: sample.pptx fixture not found', () => {})
    return
  }

  it('1. should add a new table with 2D data and formatting to a slide', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1)

    const tableId = ppt.addTable({
      id: 'NewSalesTable',
      x: 100,
      y: 150,
      width: 500,
      height: 120,
      data: [
        ['Product', 'Q1', 'Q2'],
        ['Alpha', '100', '150'],
        ['Beta', '200', '250'],
      ],
      firstRow: true,
      bandRow: true,
    })

    expect(tableId).toBe('NewSalesTable')

    const tables = ppt.getTables()
    const found = tables.find(t => t.name === 'NewSalesTable' || t.id === 'NewSalesTable')
    expect(found).toBeDefined()
    expect(found.rows).toBe(3)
    expect(found.cols).toBe(3)

    const meta = ppt.getTableRows('NewSalesTable', { includeMetadata: true })
    expect(meta.rowCount).toBe(3)
    expect(meta.columnCount).toBe(3)

    const rows = ppt.getTableRows('NewSalesTable')
    expect(rows.length).toBe(2)
    expect(rows[0].Product).toBe('Alpha')
    expect(rows[0].Q1).toBe('100')
    expect(rows[1].Q2).toBe('250')

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeDefined()
    expect(buffer.length).toBeGreaterThan(0)
  })

  it('2. should add a new table with rich cell options (borders, bold, colors)', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1)

    ppt.addTable({
      id: 'StyledTable',
      x: 50,
      y: 50,
      data: [
        [
          { value: 'Metric', bold: true, fill: '003366', color: 'FFFFFF' },
          { value: 'Score', bold: true, fill: '003366', color: 'FFFFFF' },
        ],
        [
          { value: 'Quality', border: { bottom: { color: 'FF0000', width: 2 } } },
          { value: '98%', bold: true, color: '008000' },
        ],
      ],
    })

    const rows = ppt.getTableRows('StyledTable')
    expect(rows.length).toBe(1)
    expect(rows[0].Metric).toBe('Quality')
    expect(rows[0].Score).toBe('98%')

    const buffer = await ppt.toBuffer()
    expect(buffer.length).toBeGreaterThan(0)
  })

  it('3. should append and insert columns into an existing table', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1)

    ppt.addTable({
      id: 'ColTestTable',
      data: [
        ['A', 'B'],
        ['1', '2'],
      ],
    })

    // Append column at the end
    ppt.addColumn('ColTestTable', ['C', '3'], { width: 100 })
    let rows = ppt.getTableRows('ColTestTable')
    expect(rows[0].C).toBe('3')

    // Insert column at index 1
    ppt.addColumn('ColTestTable', 1, ['Mid', '1.5'], { width: 80 })
    rows = ppt.getTableRows('ColTestTable')
    expect(rows[0].Mid).toBe('1.5')
    expect(rows[0].B).toBe('2')

    // Delete column at index 1
    ppt.deleteColumn('ColTestTable', 1)
    rows = ppt.getTableRows('ColTestTable')
    expect(rows[0].Mid).toBeUndefined()
    expect(rows[0].B).toBe('2')
    expect(rows[0].C).toBe('3')

    const buffer = await ppt.toBuffer()
    expect(buffer.length).toBeGreaterThan(0)
  })

  it('4. should apply alternating row striping to a table', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1)

    ppt.addTable({
      id: 'StripedTable',
      data: [
        ['Header 1', 'Header 2'],
        ['Row 1', 'Val 1'],
        ['Row 2', 'Val 2'],
        ['Row 3', 'Val 3'],
        ['Row 4', 'Val 4'],
      ],
    })

    ppt.stripeRows('StripedTable', {
      headerFill: '1F4E79',
      evenFill: 'F2F2F2',
      oddFill: 'FFFFFF',
      startRow: 1,
    })

    const buffer = await ppt.toBuffer()
    expect(buffer.length).toBeGreaterThan(0)
  })

  it('5. should preserve summary row on final slide when table overflows', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3) // Slide 3 contains 'Table'

    // Add 25 rows
    for (let i = 0; i < 25; i++) {
      ppt.addTableRow('Table', [`Data ${i}`, 'B', 'C', 'D', 'E'])
    }
    // Add summary row as the last row
    ppt.addTableRow('Table', ['TOTAL SUMMARY', '999', '888', '777', '666'])

    const initialSlides = ppt.slideCount

    // Split on overflow preserving 1 summary row on final slide
    await ppt.splitTableOnOverflow('Table', {
      bottomMargin: 1,
      repeatHeaders: 1,
      summaryRows: 1,
    })

    const finalSlides = ppt.slideCount
    expect(finalSlides).toBeGreaterThan(initialSlides)

    // Check slide 3 (first/intermediate slide): should NOT have 'TOTAL SUMMARY'
    ppt.useSlide(3)
    const slide3Rows = ppt.getTableRows('Table', { raw: true })
    const hasTotalOnSlide3 = slide3Rows.some(row =>
      row.some(cellText => cellText === 'TOTAL SUMMARY')
    )
    expect(hasTotalOnSlide3).toBe(false)

    // Check last generated slide for this table (slide 3 + added slides): SHOULD have 'TOTAL SUMMARY'
    const lastTableSlideIndex = 3 + (finalSlides - initialSlides)
    ppt.useSlide(lastTableSlideIndex)
    const lastSlideRows = ppt.getTableRows('Table', { raw: true })
    const hasTotalOnLastSlide = lastSlideRows.some(row =>
      row.some(cellText => cellText === 'TOTAL SUMMARY')
    )
    expect(hasTotalOnLastSlide).toBe(true)
  })
})
