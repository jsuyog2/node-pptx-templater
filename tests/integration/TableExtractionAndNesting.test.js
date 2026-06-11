import { describe, it, expect } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'
import fsExtra from 'fs-extra'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')
const TEMP_DIR = resolve(__dirname, '../temp-nesting-tests')

describe('Table JSON Extraction and Nested Rows', () => {
  const runTests = existsSync(FIXTURE_FILE)

  if (!runTests) {
    it.skip('Skipping: sample.pptx fixture not found', () => {})
    return
  }

  it('should support getTableRows JSON extraction with options', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3) // Slide 3 contains the 'Table'

    // 1. Extract raw array data
    const rawRows = await ppt.getTableRows('Table', { raw: true })
    expect(Array.isArray(rawRows)).toBe(true)
    expect(rawRows.length).toBeGreaterThan(0)
    expect(Array.isArray(rawRows[0])).toBe(true)

    // 2. Extract mapped objects (default)
    const objRows = await ppt.getTableRows('Table')
    expect(Array.isArray(objRows)).toBe(true)
    expect(objRows.length).toBeGreaterThan(0)
    expect(typeof objRows[0]).toBe('object')
    expect(objRows[0].A).toBeDefined()

    // 3. Extract with metadata
    const metaData = await ppt.getTableRows('Table', { includeMetadata: true })
    expect(metaData.rows).toBeDefined()
    expect(metaData.rowCount).toBeDefined()
    expect(metaData.columnCount).toBeDefined()
    expect(metaData.mergedCells).toBeDefined()
    expect(metaData.rowCount).toBe(metaData.rows.length + 1) // header row + data rows
  })

  it('should support addTableRow with nested rows and merges', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    // Append a nested row
    // Col 0: 'South' (1 item)
    // Col 1: ['Sales', '1200'] (2 items)
    // Col 2: ['Growth', '15%'] (2 items)
    // Col 3: 'Product' (1 item)
    // Col 4: '' (1 item)
    await ppt.addTableRow('Table', ['South', ['Sales', '1200'], ['Growth', '15%'], 'Product', ''], {
      mergeStrategy: 'rowspan',
    })

    const meta = await ppt.getTableRows('Table', { includeMetadata: true })

    // The table should have expanded by 2 rows
    // Let's verify the last two rows in raw extraction
    const rawRows = await ppt.getTableRows('Table', { raw: true })
    const len = rawRows.length
    expect(rawRows[len - 2][0]).toBe('South')
    expect(rawRows[len - 1][0]).toBe('South') // Resolves to parent's text since it's merged

    expect(rawRows[len - 2][1]).toBe('Sales')
    expect(rawRows[len - 1][1]).toBe('1200')

    expect(rawRows[len - 2][2]).toBe('Growth')
    expect(rawRows[len - 1][2]).toBe('15%')

    // Verify merge mapping metadata contains a vertical merge spanning the last two rows for column 0 and column 3
    const lastRowIndex = meta.rowCount - 1
    const secondLastRowIndex = lastRowIndex - 1
    const mergeCol0 = meta.mergedCells.find(
      m => m.startRow === secondLastRowIndex && m.endRow === lastRowIndex && m.startCol === 0
    )
    expect(mergeCol0).toBeDefined()
  })

  it('should support none mergeStrategy (no rowspan/merging)', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    await ppt.addTableRow('Table', ['North', ['Sales', '500'], 'Text'], { mergeStrategy: 'none' })

    const meta = await ppt.getTableRows('Table', { includeMetadata: true })
    // With mergeStrategy none, there should be no vertical merge registered for the new rows
    const lastRowIndex = meta.rowCount - 1
    const secondLastRowIndex = lastRowIndex - 1

    const mergeCol0 = meta.mergedCells.find(
      m => m.startRow === secondLastRowIndex && m.endRow === lastRowIndex && m.startCol === 0
    )
    expect(mergeCol0).toBeUndefined()
  })

  it('should support auto mergeStrategy (adjacent identical value merging)', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    // We add a single nested row with duplicate values in the first column
    await ppt.addTableRow(
      'Table',
      [
        ['Duplicate', 'Duplicate'],
        ['Test 1', 'Test 2'],
        ['Val 1', 'Val 2'],
      ],
      { mergeStrategy: 'auto' }
    )

    const meta = await ppt.getTableRows('Table', { includeMetadata: true })
    const lastRowIndex = meta.rowCount - 1
    const secondLastRowIndex = lastRowIndex - 1

    const mergeDuplicate = meta.mergedCells.find(
      m => m.startRow === secondLastRowIndex && m.endRow === lastRowIndex && m.startCol === 0
    )
    expect(mergeDuplicate).toBeDefined()
  })

  it('should support extractPptx and buildPptx template folder utilities', async () => {
    const outFolder = resolve(TEMP_DIR, 'unzipped-template')
    const rebuiltPptx = resolve(TEMP_DIR, 'rebuilt.pptx')

    // Clean up
    await fsExtra.remove(TEMP_DIR)
    await fsExtra.ensureDir(TEMP_DIR)

    // 1. Extract
    await PPTXTemplater.extractPptx(FIXTURE_FILE, outFolder)

    // Check critical files exist
    expect(existsSync(resolve(outFolder, '[Content_Types].xml'))).toBe(true)
    expect(existsSync(resolve(outFolder, 'ppt/presentation.xml'))).toBe(true)
    expect(existsSync(resolve(outFolder, 'ppt/slides'))).toBe(true)
    expect(existsSync(resolve(outFolder, 'ppt/_rels'))).toBe(true)

    // Check overwrite error protection
    await expect(PPTXTemplater.extractPptx(FIXTURE_FILE, outFolder)).rejects.toThrow()

    // 2. Build
    await PPTXTemplater.buildPptx(outFolder, rebuiltPptx)
    expect(existsSync(rebuiltPptx)).toBe(true)

    // 3. Load rebuilt
    const pptRebuilt = await PPTXTemplater.load(rebuiltPptx)
    expect(pptRebuilt.slideCount).toBeGreaterThan(0)

    // Clean up
    await fsExtra.remove(TEMP_DIR)
  })
})
