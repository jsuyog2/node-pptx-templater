/**
 * @fileoverview Integration tests for polymorphic table cell data types (primitives, numbers, booleans, styled objects).
 * Pure JavaScript, zero native dependencies, ECMA-376 PresentationML & DrawingML compliant.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('TASK-TABLE-CELL-ENHANCE-002: Table Cell Polymorphic Data Types', () => {
  let ppt

  beforeEach(async () => {
    ppt = await PPTXTemplater.createBlank()
  })

  it('should ingest primitives (numbers, booleans, strings) in addTable', () => {
    ppt.addTable(
      1,
      3,
      3,
      [
        ['Metric', 'Target', 'Met?'],
        ['Q1 Sales', 125000, true],
        ['Q2 Growth', 14.5, false],
      ],
      { id: 'PrimitiveTable', name: 'PrimitiveTable' }
    )

    const data = ppt.getTableData('PrimitiveTable')
    expect(data[0]).toEqual(['Metric', 'Target', 'Met?'])
    expect(data[1]).toEqual(['Q1 Sales', '125000', 'true'])
    expect(data[2]).toEqual(['Q2 Growth', '14.5', 'false'])

    expect(ppt.getCellData('PrimitiveTable', 1, 1)).toBe('125000')
    expect(ppt.getCellData('PrimitiveTable', 1, 2)).toBe('true')
  })

  it('should ingest styled values in addTable and apply font, color, alignment, and fill', () => {
    ppt.addTable(
      1,
      2,
      2,
      [
        [
          {
            value: 99.9,
            style: {
              color: '#0F172A',
              fontSize: 14,
              bold: true,
              italic: false,
              alignment: 'center',
              verticalAlignment: 'middle',
              fill: '#F1F5F9',
            },
          },
          {
            value: 'Status OK',
            style: {
              color: '#10B981',
              bold: true,
              fill: '#ECFDF5',
            },
          },
        ],
        ['Bottom Left', 42],
      ],
      { id: 'StyledTable' }
    )

    expect(ppt.getCellData('StyledTable', 0, 0)).toBe('99.9')
    expect(ppt.getCellData('StyledTable', 0, 1)).toBe('Status OK')

    const cell00 = ppt.getCell('StyledTable', 0, 0)
    expect(cell00).not.toBeNull()
    expect(cell00.text).toBe('99.9')
    expect(cell00.fill).toBe('F1F5F9')
    expect(cell00.anchor).toBe('ctr')

    const cell01 = ppt.getCell('StyledTable', 0, 1)
    expect(cell01.fill).toBe('ECFDF5')
  })

  it('should update cell with primitives and styled values via updateTableCell', () => {
    ppt.addTable(
      1,
      2,
      2,
      [
        ['A', 'B'],
        ['C', 'D'],
      ],
      { id: 'UpdateCellTable' }
    )

    // Update with number
    ppt.updateTableCell('UpdateCellTable', 0, 0, 9999)
    expect(ppt.getCellData('UpdateCellTable', 0, 0)).toBe('9999')

    // Update with styled value object
    ppt.updateTableCell('UpdateCellTable', 0, 1, {
      value: 125000,
      style: {
        color: '#FFFFFF',
        fill: '#3B82F6',
        bold: true,
        fontSize: 12,
        alignment: 'center',
      },
    })
    expect(ppt.getCellData('UpdateCellTable', 0, 1)).toBe('125000')

    const cell01 = ppt.getCell('UpdateCellTable', 0, 1)
    expect(cell01.fill).toBe('3B82F6')
  })

  it('should update rows and columns with polymorphic values (updateRowData, updateColumnData)', () => {
    ppt.addTable(
      1,
      3,
      3,
      [
        ['A', 'B', 'C'],
        ['D', 'E', 'F'],
        ['G', 'H', 'I'],
      ],
      { id: 'VectorTable' }
    )

    // Update entire row with numbers and styled values
    ppt.updateRowData('VectorTable', 1, [
      100,
      { value: 200, style: { fill: '#FDE047', bold: true } },
      300,
    ])

    const row1 = ppt.getRowData('VectorTable', 1)
    expect(row1).toEqual(['100', '200', '300'])
    expect(ppt.getCell('VectorTable', 1, 1).fill).toBe('FDE047')

    // Update entire column with mixed primitives and styled values
    ppt.updateColumnData('VectorTable', 2, [
      'Top',
      { value: 'Mid', style: { fill: '#6EE7B7' } },
      'Bottom',
    ])

    const col2 = ppt.getColumnData('VectorTable', 2)
    expect(col2).toEqual(['Top', 'Mid', 'Bottom'])
    expect(ppt.getCell('VectorTable', 1, 2).fill).toBe('6EE7B7')
  })

  it('should replace entire matrix using updateTableData with polymorphic elements', () => {
    ppt.addTable(
      1,
      2,
      2,
      [
        ['Old1', 'Old2'],
        ['Old3', 'Old4'],
      ],
      { id: 'MatrixTable' }
    )

    ppt.updateTableData('MatrixTable', [
      [
        { value: 'Q1', style: { bold: true, fill: '#E2E8F0' } },
        { value: 150000, style: { fill: '#FEF08A' } },
      ],
      [
        { value: 'Q2', style: { bold: true, fill: '#E2E8F0' } },
        { value: 185000, style: { fill: '#BBF7D0' } },
      ],
    ])

    const matrix = ppt.getTableData('MatrixTable')
    expect(matrix).toEqual([
      ['Q1', '150000'],
      ['Q2', '185000'],
    ])

    expect(ppt.getCell('MatrixTable', 0, 0).fill).toBe('E2E8F0')
    expect(ppt.getCell('MatrixTable', 0, 1).fill).toBe('FEF08A')
    expect(ppt.getCell('MatrixTable', 1, 1).fill).toBe('BBF7D0')
  })

  it('should export valid PPTX package containing polymorphic cell types and reload cleanly', async () => {
    ppt.addTable(
      1,
      2,
      2,
      [
        [
          { value: 125000, style: { color: '#0F172A', fontSize: 12, bold: true, fill: '#F1F5F9' } },
          'Standard Text',
        ],
        [true, 42.5],
      ],
      { id: 'ExportPolyTable' }
    )

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeInstanceOf(Buffer)
    expect(buffer.length).toBeGreaterThan(1000)

    const reloaded = await PPTXTemplater.load(buffer)
    const tableData = reloaded.getTableData('ExportPolyTable')
    expect(tableData).toEqual([
      ['125000', 'Standard Text'],
      ['true', '42.5'],
    ])

    const cell = reloaded.getCell('ExportPolyTable', 0, 0)
    expect(cell.text).toBe('125000')
    expect(cell.fill).toBe('F1F5F9')
  })
})
