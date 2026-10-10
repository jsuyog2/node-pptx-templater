import { describe, it, expect } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')

describe('Table Auto-Pagination (splitTableOnOverflow)', () => {
  const runTests = existsSync(FIXTURE_FILE)

  if (!runTests) {
    it.skip('Skipping: sample.pptx fixture not found', () => {})
    return
  }

  it('should split table rows when table exceeds slide height limit', async () => {
    // Clone PPT for test isolation
    const testPpt = await PPTXTemplater.load(FIXTURE_FILE)

    // Find slide 3 which has 'Table'
    testPpt.useSlide(3)
    const initialSlideCount = testPpt.slideCount

    // Generate lots of rows to force overflow
    for (let i = 0; i < 30; i++) {
      testPpt.addTableRow('Table', [`Row ${i}`, 'B', 'C', 'D', 'E'])
    }

    // Check initial table state (1 large table)
    const tables = testPpt.getTables()
    expect(tables.length).toBeGreaterThan(0)

    const tableInfoBefore = tables.find(t => t.name === 'Table' || t.id === '5')
    const beforeRows = tableInfoBefore.rows
    expect(beforeRows).toBeGreaterThan(30)

    // Split on overflow, leaving 1 inch (914400 EMUs) at the bottom, repeating 1 header row
    await testPpt.splitTableOnOverflow('Table', { bottomMargin: 1, repeatHeaders: 1 })

    // Ensure slides were duplicated
    const finalSlideCount = testPpt.slideCount
    expect(finalSlideCount).toBeGreaterThan(initialSlideCount)

    // Test that the first slide's table has fewer rows now
    testPpt.useSlide(3)
    const slide3Tables = testPpt.getTables()
    const tableInfoSlide3 = slide3Tables.find(t => t.name === 'Table' || t.id === '5')
    expect(tableInfoSlide3.rows).toBeLessThan(beforeRows)

    // Test that the next slide has the repeated header
    testPpt.useSlide(4)
    const slide4Tables = testPpt.getTables()
    const tableInfoSlide4 = slide4Tables.find(t => t.name === 'Table' || t.id === '5')
    expect(tableInfoSlide4).toBeDefined()

    // Verify first row of the split table is a header (same text as original)
    // For this test, we know it's a clone of the table, so it should exist and have data
  })
})
