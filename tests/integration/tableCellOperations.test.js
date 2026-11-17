/**
 * @fileoverview Integration tests for TASK-TABLE-004: Cell-Level Operations, Clear & Data Accessors.
 * Tests updateTableCell, clearCellData, getTableData, getCell, and getCellData.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('TASK-TABLE-004: Cell-Level Operations, Clear & Data Accessors', () => {
  let ppt

  beforeEach(async () => {
    ppt = await PPTXTemplater.createBlank()
    ppt.addTable(
      1,
      3,
      3,
      [
        ['H1', 'H2', 'H3'],
        ['Val A', 'Val B', 'Val C'],
        ['Val D', 'Val E', 'Val F'],
      ],
      { id: 'CellTable', name: 'CellTable' }
    )
  })

  it('should update cell text with string and return this', () => {
    const res = ppt.updateTableCell('CellTable', 1, 1, 'Updated Val B')
    expect(res).toBe(ppt) // Fluent chaining

    expect(ppt.getCellData('CellTable', 1, 1)).toBe('Updated Val B')
    const matrix = ppt.getTableData('CellTable')
    expect(matrix[1][1]).toBe('Updated Val B')
  })

  it('should update cell formatting, typography, fill, and multiline text with updateTableCell', () => {
    const res = ppt.updateTableCell('CellTable', 0, 0, {
      text: 'Title Line 1\nTitle Line 2',
      bold: true,
      italic: false,
      fontSize: 14,
      color: 'FFFFFF',
      fill: '1E3A8A',
      align: 'ctr',
      verticalAlign: 'ctr',
    })
    expect(res).toBe(ppt)

    const cellData = ppt.getCellData('CellTable', 0, 0)
    expect(cellData).toBe('Title Line 1\nTitle Line 2')

    const cell = ppt.getCell('CellTable', 0, 0)
    expect(cell).not.toBeNull()
    expect(cell.text).toBe('Title Line 1\nTitle Line 2')
    expect(cell.fill).toBe('1E3A8A')
    expect(cell.row).toBe(0)
    expect(cell.column).toBe(0)
    expect(cell.bounds).toBeDefined()
  })

  it('should clear cell text with clearCellData while preserving cell properties', () => {
    // Style cell first
    ppt.updateTableCell('CellTable', 2, 2, {
      text: 'To be cleared',
      fill: 'FF0000',
    })

    const preClear = ppt.getCell('CellTable', 2, 2)
    expect(preClear.text).toBe('To be cleared')
    expect(preClear.fill).toBe('FF0000')

    // Clear cell data
    const res = ppt.clearCellData('CellTable', 2, 2)
    expect(res).toBe(ppt)

    const postClear = ppt.getCell('CellTable', 2, 2)
    expect(postClear.text).toBe('')
    expect(postClear.fill).toBe('FF0000') // Fill preserved
    expect(ppt.getCellData('CellTable', 2, 2)).toBe('')
  })

  it('should return complete 2D string matrix via getTableData', () => {
    const data = ppt.getTableData('CellTable')
    expect(data).toEqual([
      ['H1', 'H2', 'H3'],
      ['Val A', 'Val B', 'Val C'],
      ['Val D', 'Val E', 'Val F'],
    ])
  })

  it('should export valid PPTX package and preserve cell mutations on reload', async () => {
    ppt.updateTableCell('CellTable', 1, 0, 'New Value A')
    ppt.updateTableCell('CellTable', 2, 1, {
      text: 'Styled Cell',
      fill: '00AA00',
      bold: true,
    })

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeInstanceOf(Buffer)

    const loaded = await PPTXTemplater.load(buffer)
    expect(loaded.getCellData('CellTable', 1, 0)).toBe('New Value A')
    expect(loaded.getCellData('CellTable', 2, 1)).toBe('Styled Cell')

    const cell = loaded.getCell('CellTable', 2, 1)
    expect(cell.fill).toBe('00AA00')

    const validation = loaded.validate()
    expect(validation.valid).toBe(true)
    expect(validation.errors).toHaveLength(0)
  })
})
