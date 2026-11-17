/**
 * @fileoverview Integration tests for table cell shape overlays (inside/outside text, pill/circle/roundRect badges).
 * ECMA-376 PresentationML & DrawingML compliant.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('TASK-TABLE-CELL-ENHANCE-002: Table Cell Shapes & Overlay Engine', () => {
  let ppt

  beforeEach(async () => {
    ppt = await PPTXTemplater.createBlank()
  })

  it('should render shape overlay with text inside the shape', () => {
    ppt.addTable(
      1,
      2,
      2,
      [
        [
          'Task',
          {
            shape: {
              type: 'roundRect',
              fill: '#10B981',
              stroke: '#059669',
              width: 60,
              height: 20,
            },
            text: 'Completed',
            textPosition: 'inside',
            textStyle: { color: '#FFFFFF', fontSize: 10, bold: true, alignment: 'center' },
          },
        ],
        ['Build CI/CD', 'Done'],
      ],
      { id: 'InsideShapeTable' }
    )

    // Verify cell query returns inside-shape text
    expect(ppt.getCellData('InsideShapeTable', 0, 1)).toBe('Completed')
    const matrix = ppt.getTableData('InsideShapeTable')
    expect(matrix[0][1]).toBe('Completed')

    const cell = ppt.getCell('InsideShapeTable', 0, 1)
    expect(cell).not.toBeNull()
    expect(cell.text).toBe('Completed')
    expect(cell.shape).not.toBeNull()
    expect(cell.shape.text).toBe('Completed')
  })

  it('should render shape overlay with text outside the shape and adjust margins', () => {
    ppt.addTable(
      1,
      2,
      2,
      [
        [
          'Alert',
          {
            shape: {
              type: 'circle',
              fill: '#EF4444',
              width: 10,
              height: 10,
            },
            text: 'Critical Block',
            textPosition: 'outside',
            textStyle: { color: '#1E293B', fontSize: 11 },
            shapePlacement: 'left',
          },
        ],
        ['Warning', 'Notice'],
      ],
      { id: 'OutsideShapeTable' }
    )

    expect(ppt.getCellData('OutsideShapeTable', 0, 1)).toBe('Critical Block')
    const cell = ppt.getCell('OutsideShapeTable', 0, 1)
    expect(cell.text).toBe('Critical Block')
  })

  it('should update cell with a shape overlay using updateTableCell', () => {
    ppt.addTable(
      1,
      2,
      2,
      [
        ['A', 'B'],
        ['C', 'D'],
      ],
      { id: 'DynamicShapeTable' }
    )

    ppt.updateTableCell('DynamicShapeTable', 1, 1, {
      shape: {
        type: 'roundRect',
        fill: '#3B82F6',
        stroke: '#1D4ED8',
        width: 50,
        height: 18,
      },
      text: 'In Progress',
      textPosition: 'inside',
      textStyle: { color: '#FFFFFF', fontSize: 9, bold: true },
    })

    expect(ppt.getCellData('DynamicShapeTable', 1, 1)).toBe('In Progress')
    const cell = ppt.getCell('DynamicShapeTable', 1, 1)
    expect(cell.shape).not.toBeNull()

    // Overwriting cell with primitive removes the shape overlay
    ppt.updateTableCell('DynamicShapeTable', 1, 1, 'Plain Text')
    expect(ppt.getCellData('DynamicShapeTable', 1, 1)).toBe('Plain Text')
    const updatedCell = ppt.getCell('DynamicShapeTable', 1, 1)
    expect(updatedCell.shape).toBeNull()
  })

  it('should remove cell shape on clearCellData, removeTableRow, and removeTable', () => {
    ppt.addTable(
      1,
      3,
      2,
      [
        [
          'Row 0',
          {
            shape: { type: 'circle', fill: '#10B981' },
            text: 'Badge 0',
            textPosition: 'inside',
          },
        ],
        [
          'Row 1',
          {
            shape: { type: 'circle', fill: '#F59E0B' },
            text: 'Badge 1',
            textPosition: 'inside',
          },
        ],
        ['Row 2', 'Standard'],
      ],
      { id: 'LifecycleShapeTable' }
    )

    // Clear cell at (0, 1)
    ppt.clearCellData('LifecycleShapeTable', 0, 1)
    expect(ppt.getCellData('LifecycleShapeTable', 0, 1)).toBe('')
    expect(ppt.getCell('LifecycleShapeTable', 0, 1).shape).toBeNull()

    // Remove row 1
    ppt.removeTableRow('LifecycleShapeTable', 1)
    expect(ppt.getTableData('LifecycleShapeTable').length).toBe(2)

    // Remove entire table
    ppt.removeTable('LifecycleShapeTable')
    expect(ppt.getTables(1).find(t => t.id === 'LifecycleShapeTable')).toBeUndefined()
  })

  it('should clone cell shapes when table is cloned via cloneTable', () => {
    ppt.addTable(
      1,
      2,
      2,
      [
        [
          'Task',
          {
            shape: { type: 'roundRect', fill: '#6366F1' },
            text: 'Cloned Badge',
            textPosition: 'inside',
          },
        ],
        ['Item', 'Value'],
      ],
      { id: 'SourceTable', x: 50, y: 50 }
    )

    ppt.cloneTable('SourceTable', 'ClonedTable', { x: 50, y: 250 })

    const cloned = ppt.getTable('ClonedTable')
    expect(cloned).not.toBeNull()
    expect(cloned.name).toBe('ClonedTable')

    const clonedData = ppt.getCellData('ClonedTable', 0, 1)
    expect(clonedData).toBe('Cloned Badge')
  })

  it('should export presentation with cell shapes and reload cleanly', async () => {
    ppt.addTable(
      1,
      2,
      2,
      [
        [
          'Status',
          {
            shape: { type: 'roundRect', fill: '#10B981', width: 60, height: 20 },
            text: 'Success',
            textPosition: 'inside',
            textStyle: { color: '#FFFFFF', bold: true },
          },
        ],
        [
          'Priority',
          {
            shape: { type: 'circle', fill: '#EF4444', width: 12, height: 12 },
            text: 'High',
            textPosition: 'outside',
            shapePlacement: 'left',
          },
        ],
      ],
      { id: 'ExportShapeTable' }
    )

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeInstanceOf(Buffer)
    expect(buffer.length).toBeGreaterThan(1000)

    const reloaded = await PPTXTemplater.load(buffer)
    expect(reloaded.getCellData('ExportShapeTable', 0, 1)).toBe('Success')
    expect(reloaded.getCellData('ExportShapeTable', 1, 1)).toBe('High')
  })
})
