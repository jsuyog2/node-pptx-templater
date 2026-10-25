/**
 * @fileoverview Integration tests for TASK-TABLE-003: Row & Column Structural Manipulation & Data I/O.
 * Tests addTableRow, addTableColumn, removeTableRow, removeTableColumn,
 * updateRowData, updateColumnData, clearRowData, clearColumnData,
 * getRowData, getColumnData, getRow, and getColumn.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('TASK-TABLE-003: Row & Column Structural Manipulation & Data I/O', () => {
  let ppt

  beforeEach(async () => {
    ppt = await PPTXTemplater.createBlank()
    // Create base 3x3 table on slide 1
    ppt.addTable(
      1,
      3,
      3,
      [
        ['H1', 'H2', 'H3'],
        ['R1C1', 'R1C2', 'R1C3'],
        ['R2C1', 'R2C2', 'R2C3'],
      ],
      { id: 'GridTable', name: 'GridTable' }
    )
  })

  it('should append and insert rows with addTableRow and return this', () => {
    // Append row
    const res1 = ppt.addTableRow('GridTable', ['R3C1', 'R3C2', 'R3C3'])
    expect(res1).toBe(ppt) // Fluent chaining

    let data = ppt.getTableData('GridTable')
    expect(data).toHaveLength(4)
    expect(data[3]).toEqual(['R3C1', 'R3C2', 'R3C3'])

    // Insert row at index 1
    const res2 = ppt.addTableRow('GridTable', 1, ['InsertC1', 'InsertC2', 'InsertC3'])
    expect(res2).toBe(ppt)

    data = ppt.getTableData('GridTable')
    expect(data).toHaveLength(5)
    expect(data[1]).toEqual(['InsertC1', 'InsertC2', 'InsertC3'])
    expect(data[2]).toEqual(['R1C1', 'R1C2', 'R1C3'])
  })

  it('should append and insert columns with addTableColumn and return this', () => {
    // Append column
    const res1 = ppt.addTableColumn('GridTable', ['H4', 'R1C4', 'R2C4'])
    expect(res1).toBe(ppt)

    let data = ppt.getTableData('GridTable')
    expect(data[0]).toHaveLength(4)
    expect(ppt.getColumnData('GridTable', 3)).toEqual(['H4', 'R1C4', 'R2C4'])

    // Insert column at index 1
    const res2 = ppt.addTableColumn('GridTable', 1, ['InsertH', 'InsertR1', 'InsertR2'])
    expect(res2).toBe(ppt)

    data = ppt.getTableData('GridTable')
    expect(data[0]).toHaveLength(5)
    expect(ppt.getColumnData('GridTable', 1)).toEqual(['InsertH', 'InsertR1', 'InsertR2'])
    expect(ppt.getColumnData('GridTable', 2)).toEqual(['H2', 'R1C2', 'R2C2'])
  })

  it('should remove row with removeTableRow and column with removeTableColumn', () => {
    // Remove row 1
    const res1 = ppt.removeTableRow('GridTable', 1)
    expect(res1).toBe(ppt)

    let data = ppt.getTableData('GridTable')
    expect(data).toHaveLength(2)
    expect(data[0]).toEqual(['H1', 'H2', 'H3'])
    expect(data[1]).toEqual(['R2C1', 'R2C2', 'R2C3'])

    // Remove column 1
    const res2 = ppt.removeTableColumn('GridTable', 1)
    expect(res2).toBe(ppt)

    data = ppt.getTableData('GridTable')
    expect(data[0]).toEqual(['H1', 'H3'])
    expect(data[1]).toEqual(['R2C1', 'R2C3'])

    // Out of bounds checking
    expect(() => ppt.removeTableRow('GridTable', 99)).toThrow(/out of bounds/i)
    expect(() => ppt.removeTableColumn('GridTable', 99)).toThrow(/out of bounds/i)
  })

  it('should update row data with updateRowData and column data with updateColumnData', () => {
    const res1 = ppt.updateRowData('GridTable', 1, ['Updated1', 'Updated2', 'Updated3'])
    expect(res1).toBe(ppt)
    expect(ppt.getRowData('GridTable', 1)).toEqual(['Updated1', 'Updated2', 'Updated3'])

    const res2 = ppt.updateColumnData('GridTable', 0, ['NewH1', 'NewR1', 'NewR2'])
    expect(res2).toBe(ppt)
    expect(ppt.getColumnData('GridTable', 0)).toEqual(['NewH1', 'NewR1', 'NewR2'])
  })

  it('should clear row data with clearRowData and column data with clearColumnData', () => {
    const res1 = ppt.clearRowData('GridTable', 2)
    expect(res1).toBe(ppt)
    expect(ppt.getRowData('GridTable', 2)).toEqual(['', '', ''])

    const res2 = ppt.clearColumnData('GridTable', 1)
    expect(res2).toBe(ppt)
    expect(ppt.getColumnData('GridTable', 1)).toEqual(['', '', ''])
  })

  it('should query getRow and getColumn returning comprehensive descriptors', () => {
    const rowInfo = ppt.getRow('GridTable', 0)
    expect(rowInfo).not.toBeNull()
    expect(rowInfo.index).toBe(0)
    expect(rowInfo.height).toBeGreaterThan(0)
    expect(rowInfo.data).toEqual(['H1', 'H2', 'H3'])

    const colInfo = ppt.getColumn('GridTable', 2)
    expect(colInfo).not.toBeNull()
    expect(colInfo.index).toBe(2)
    expect(colInfo.width).toBeGreaterThan(0)
    expect(colInfo.data).toEqual(['H3', 'R1C3', 'R2C3'])
  })

  it('should export valid, repair-free PPTX package after row and column operations', async () => {
    ppt.addTableRow('GridTable', ['NewRow1', 'NewRow2', 'NewRow3'])
    ppt.addTableColumn('GridTable', ['Col4H', 'Col4R1', 'Col4R2', 'Col4R3'])
    ppt.updateRowData('GridTable', 0, ['HeadA', 'HeadB', 'HeadC', 'HeadD'])

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeInstanceOf(Buffer)

    const reloaded = await PPTXTemplater.load(buffer)
    const tableData = reloaded.getTableData('GridTable')
    expect(tableData).toHaveLength(4)
    expect(tableData[0]).toEqual(['HeadA', 'HeadB', 'HeadC', 'HeadD'])

    const validation = reloaded.validate()
    expect(validation.valid).toBe(true)
    expect(validation.errors).toHaveLength(0)
  })
})
