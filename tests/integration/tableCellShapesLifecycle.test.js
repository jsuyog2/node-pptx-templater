/**
 * @fileoverview Integration tests verifying all 11 table cell shape lifecycle,
 * indexing, positioning, coexistence, cloning, and cascading deletion bugs.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('Table Cell Shapes Lifecycle & Indexing Subsystem (11 Bug Fixes)', () => {
  let pptx

  beforeEach(async () => {
    pptx = await PPTXTemplater.createBlank()
  })

  // --------------------------------------------------------------------------
  // Bug #1: addTable Initial Cell Shape Rendering
  // --------------------------------------------------------------------------
  it('Bug #1: addTable correctly parses polymorphic cell shapes (inside, outside, shape-only) during initial creation', async () => {
    const tableData = [
      ['Metric', 'Target', 'Status', 'Badge'],
      [
        'Revenue',
        '$1,000,000',
        {
          shape: { type: 'roundRect', fill: '#10B981', width: 70, height: 22 },
          text: 'Exceeded',
          textPosition: 'inside',
          textStyle: { color: '#FFFFFF', fontSize: 10, bold: true },
        },
        {
          shape: { type: 'star5', fill: '#F59E0B', width: 18, height: 18 },
        },
      ],
      [
        'Churn',
        '2.5%',
        {
          shape: { type: 'circle', fill: '#EF4444', width: 12, height: 12 },
          text: 'Warning Level',
          textPosition: 'outside',
          shapePlacement: 'left',
          textStyle: { color: '#1E293B', fontSize: 10 },
        },
        'Standard',
      ],
    ]

    pptx.addTable(1, 3, 4, tableData, {
      id: 'InitShapeTable',
      x: 100,
      y: 100,
      width: 600,
      height: 250,
    })

    const shapes = pptx.getShapes(1)
    const cellShapes = shapes.filter(s => s.name?.startsWith('CellShape_InitShapeTable_'))

    // Should create 3 shape overlays: row 1 col 2 (inside), row 1 col 3 (star5), row 2 col 2 (outside)
    expect(cellShapes.length).toBe(3)

    // Verify inside text shape
    const insideShape = cellShapes.find(s => s.name.includes('_r1_c2'))
    expect(insideShape).toBeDefined()
    expect(insideShape.text).toBe('Exceeded')

    // Verify shape-only overlay
    const starShape = cellShapes.find(s => s.name.includes('_r1_c3'))
    expect(starShape).toBeDefined()
    expect(starShape.text || '').toBe('')

    // Verify outside shape
    const outsideShape = cellShapes.find(s => s.name.includes('_r2_c2'))
    expect(outsideShape).toBeDefined()

    // Table cell itself should have the text "Warning Level"
    const cellData = pptx.getCellData('InitShapeTable', 2, 2, 1)
    expect(cellData).toBe('Warning Level')

    // Inside text cell should return inside text from getCellData
    const insideCellData = pptx.getCellData('InitShapeTable', 1, 2, 1)
    expect(insideCellData).toBe('Exceeded')
  })

  // --------------------------------------------------------------------------
  // Bug #2: updateTableData Placement, Alignment & Dimensions
  // --------------------------------------------------------------------------
  it('Bug #2: updateTableData accurately applies width, height, offsets, and placement without dropping metadata', async () => {
    pptx.addTable(
      1,
      2,
      2,
      [
        ['A', 'B'],
        ['C', 'D'],
      ],
      {
        id: 'UpdatePlacementTable',
        x: 100,
        y: 100,
        width: 400,
        height: 200,
      }
    )

    // Update with customized shapes
    pptx.updateTableData(
      'UpdatePlacementTable',
      [
        [
          'A',
          {
            shape: {
              type: 'roundRect',
              fill: '#6366F1',
              width: 80,
              height: 25,
              offsetX: 10,
              offsetY: 5,
              alignment: 'center',
            },
            text: 'Custom Badge',
            textPosition: 'inside',
          },
        ],
        [
          {
            shape: {
              type: 'circle',
              fill: '#10B981',
              width: 15,
              height: 15,
              shapePlacement: 'right',
            },
            text: 'Active',
            textPosition: 'outside',
          },
          'Standard Text',
        ],
      ],
      1
    )

    const shapes = pptx.getShapes(1)
    const customShape = shapes.find(s => s.name?.includes('CellShape_UpdatePlacementTable_r0_c1'))
    expect(customShape).toBeDefined()
    expect(customShape.text).toBe('Custom Badge')

    const circleShape = shapes.find(s => s.name?.includes('CellShape_UpdatePlacementTable_r1_c0'))
    expect(circleShape).toBeDefined()

    // Cell data for outside text
    expect(pptx.getCellData('UpdatePlacementTable', 1, 0, 1)).toBe('Active')

    // Replace shape with primitive string: shape must be purged cleanly
    pptx.updateTableData(
      'UpdatePlacementTable',
      [
        ['A', 'Plain Text B'],
        ['Plain Text C', 'Plain Text D'],
      ],
      1
    )

    const remainingShapes = pptx
      .getShapes(1)
      .filter(s => s.name?.startsWith('CellShape_UpdatePlacementTable_'))
    expect(remainingShapes.length).toBe(0)
    expect(pptx.getCellData('UpdatePlacementTable', 0, 1, 1)).toBe('Plain Text B')
  })

  // --------------------------------------------------------------------------
  // Bug #3: cloneTable Deep Shape Cloning with Offsets & Unique IDs
  // --------------------------------------------------------------------------
  it('Bug #3: cloneTable duplicates all cell shapes, shifts coordinates, and allocates unique shape IDs', async () => {
    pptx.addTable(
      1,
      2,
      2,
      [
        [
          'OrigHeader',
          {
            shape: { type: 'roundRect', fill: '#0EA5E9', width: 60, height: 20 },
            text: 'Original',
            textPosition: 'inside',
          },
        ],
        ['Row2Col1', 'Row2Col2'],
      ],
      {
        id: 'SourceTable',
        x: 100,
        y: 100,
        width: 300,
        height: 150,
      }
    )

    const origShape = pptx.getShapes(1).find(s => s.name?.includes('CellShape_SourceTable_r0_c1'))
    expect(origShape).toBeDefined()

    // Clone table with +200 x offset and +150 y offset
    pptx.cloneTable('SourceTable', 'ClonedTable', { x: 300, y: 250 }, 1)

    const clonedTable = pptx.getTable('ClonedTable', 1)
    expect(clonedTable).toBeDefined()
    expect(clonedTable.name).toBe('ClonedTable')

    const allShapes = pptx.getShapes(1)
    const clonedCellShape = allShapes.find(s => s.name?.includes('CellShape_ClonedTable_r0_c1'))
    expect(clonedCellShape).toBeDefined()
    expect(clonedCellShape.text).toBe('Original')
    expect(clonedCellShape.id).not.toBe(origShape.id)

    // Cloned shape X and Y must be shifted
    expect(clonedCellShape.x).toBeGreaterThan(origShape.x)
    expect(clonedCellShape.y).toBeGreaterThan(origShape.y)
  })

  // --------------------------------------------------------------------------
  // Bug #4: addTableRow Dynamic Indexing & Downward Shape Shifting
  // --------------------------------------------------------------------------
  it('Bug #4: addTableRow dynamically shifts existing shape row indices and EMU Y positions downward', async () => {
    pptx.addTable(
      1,
      3,
      2,
      [
        ['Row0', 'Row0Col1'],
        [
          'Row1',
          {
            shape: { type: 'roundRect', fill: '#10B981', width: 50, height: 18 },
            text: 'R1-Shape',
            textPosition: 'inside',
          },
        ],
        [
          'Row2',
          {
            shape: { type: 'circle', fill: '#EF4444', width: 14, height: 14 },
            text: 'R2-Shape',
            textPosition: 'inside',
          },
        ],
      ],
      {
        id: 'ShiftRowTable',
        x: 100,
        y: 100,
        width: 300,
        height: 200,
      }
    )

    const initialR1Shape = pptx
      .getShapes(1)
      .find(s => s.name?.includes('CellShape_ShiftRowTable_r1_c1'))
    const initialR2Shape = pptx
      .getShapes(1)
      .find(s => s.name?.includes('CellShape_ShiftRowTable_r2_c1'))
    expect(initialR1Shape).toBeDefined()
    expect(initialR2Shape).toBeDefined()

    const origR1Y = initialR1Shape.y
    const origR2Y = initialR2Shape.y

    // Insert row at index 2 (between row 1 and row 2)
    pptx.addTableRow('ShiftRowTable', 2, ['InsertedRow2', 'InsertedVal'], 1)

    // Row 2 shape should now be shifted to Row 3!
    const updatedShapesAfterFirst = pptx.getShapes(1)
    const shiftedToR3 = updatedShapesAfterFirst.find(s =>
      s.name?.includes('CellShape_ShiftRowTable_r3_c1')
    )
    expect(shiftedToR3).toBeDefined()
    expect(shiftedToR3.text).toBe('R2-Shape')
    expect(shiftedToR3.y).toBeGreaterThan(origR2Y)

    // Row 1 shape should remain at Row 1 (since 1 < 2)
    const unshiftedR1 = updatedShapesAfterFirst.find(s =>
      s.name?.includes('CellShape_ShiftRowTable_r1_c1')
    )
    expect(unshiftedR1).toBeDefined()
    expect(unshiftedR1.y).toBe(origR1Y)

    // Now insert a row at index 1 (between row 0 and row 1)
    pptx.addTableRow('ShiftRowTable', 1, ['InsertedRow1', 'InsertedVal1'], 1)

    // Row 1 shape should now shift to Row 2, and Row 3 shape should shift to Row 4!
    const updatedShapesAfterSecond = pptx.getShapes(1)
    const r1NowR2 = updatedShapesAfterSecond.find(s =>
      s.name?.includes('CellShape_ShiftRowTable_r2_c1')
    )
    const r3NowR4 = updatedShapesAfterSecond.find(s =>
      s.name?.includes('CellShape_ShiftRowTable_r4_c1')
    )

    expect(r1NowR2).toBeDefined()
    expect(r1NowR2.text).toBe('R1-Shape')
    expect(r1NowR2.y).toBeGreaterThan(origR1Y)

    expect(r3NowR4).toBeDefined()
    expect(r3NowR4.text).toBe('R2-Shape')
    expect(r3NowR4.y).toBeGreaterThan(shiftedToR3.y)
  })

  // --------------------------------------------------------------------------
  // Bug #5: addTableColumn Text and Shape Coexistence & Rightward Shifting
  // --------------------------------------------------------------------------
  it('Bug #5: addTableColumn renders both DrawingML text and shape overlay without dropping text runs', async () => {
    pptx.addTable(
      1,
      2,
      2,
      [
        [
          'H1',
          {
            shape: { type: 'star5', fill: '#F59E0B', width: 16, height: 16 },
            text: 'StarCol1',
            textPosition: 'inside',
          },
        ],
        ['R1C0', 'R1C1'],
      ],
      {
        id: 'ColCoexistTable',
        x: 100,
        y: 100,
        width: 300,
        height: 150,
      }
    )

    const initialShape = pptx
      .getShapes(1)
      .find(s => s.name?.includes('CellShape_ColCoexistTable_r0_c1'))
    expect(initialShape).toBeDefined()
    const origShapeX = initialShape.x

    // Insert a new column at index 1 with an outside-text shape
    pptx.addTableColumn(
      'ColCoexistTable',
      1,
      [
        {
          shape: { type: 'circle', fill: '#10B981', width: 12, height: 12 },
          text: 'Coexisting Text',
          textPosition: 'outside',
          shapePlacement: 'left',
          textStyle: { color: '#0F172A', bold: true },
        },
        'Standard New Col',
      ],
      1
    )

    // The newly added column at index 1 must have BOTH text in the cell AND shape overlay
    const newCellData = pptx.getCellData('ColCoexistTable', 0, 1, 1)
    expect(newCellData).toBe('Coexisting Text')

    const allShapes = pptx.getShapes(1)
    const newColShape = allShapes.find(s => s.name?.includes('CellShape_ColCoexistTable_r0_c1'))
    expect(newColShape).toBeDefined()

    // The previous column 1 shape should have shifted to column 2!
    const shiftedCol2Shape = allShapes.find(s =>
      s.name?.includes('CellShape_ColCoexistTable_r0_c2')
    )
    expect(shiftedCol2Shape).toBeDefined()
    expect(shiftedCol2Shape.text).toBe('StarCol1')
    expect(shiftedCol2Shape.x).toBeGreaterThan(origShapeX)
  })

  // --------------------------------------------------------------------------
  // Bug #6: removeTable Cascading Deletion
  // --------------------------------------------------------------------------
  it('Bug #6: removeTable cleanly purges all anchored cell shapes and removes table frame', async () => {
    pptx.addTable(
      1,
      2,
      2,
      [
        [
          {
            shape: { type: 'roundRect', fill: '#10B981', width: 50, height: 20 },
            text: 'Badge1',
            textPosition: 'inside',
          },
          'B',
        ],
        [
          'C',
          {
            shape: { type: 'circle', fill: '#EF4444', width: 12, height: 12 },
            text: 'Badge2',
            textPosition: 'inside',
          },
        ],
      ],
      { id: 'TableToDelete' }
    )

    expect(
      pptx.getShapes(1).filter(s => s.name?.startsWith('CellShape_TableToDelete_')).length
    ).toBe(2)

    pptx.removeTable('TableToDelete', 1)

    // Both table and shapes must be removed completely
    expect(pptx.getTable('TableToDelete', 1)).toBeNull()
    expect(
      pptx.getShapes(1).filter(s => s.name?.startsWith('CellShape_TableToDelete_')).length
    ).toBe(0)
  })

  // --------------------------------------------------------------------------
  // Bug #7: removeTableRow Purging and Upward Shape Shifting
  // --------------------------------------------------------------------------
  it('Bug #7: removeTableRow purges row shapes and shifts lower row shapes upward', async () => {
    pptx.addTable(
      1,
      4,
      2,
      [
        ['Row0', 'Col1'],
        [
          'Row1',
          {
            shape: { type: 'roundRect', fill: '#3B82F6', width: 50, height: 20 },
            text: 'R1-Badge',
            textPosition: 'inside',
          },
        ],
        [
          'Row2',
          {
            shape: { type: 'circle', fill: '#EF4444', width: 14, height: 14 },
            text: 'R2-Badge-Delete',
            textPosition: 'inside',
          },
        ],
        [
          'Row3',
          {
            shape: { type: 'diamond', fill: '#10B981', width: 16, height: 16 },
            text: 'R3-Badge-Shift',
            textPosition: 'inside',
          },
        ],
      ],
      { id: 'RemoveRowTable', y: 100, height: 300 }
    )

    const initialR3 = pptx
      .getShapes(1)
      .find(s => s.name?.includes('CellShape_RemoveRowTable_r3_c1'))
    expect(initialR3).toBeDefined()
    const origR3Y = initialR3.y

    // Remove row 2 (which contains R2-Badge-Delete)
    pptx.removeTableRow('RemoveRowTable', 2, 1)

    const shapesAfter = pptx.getShapes(1)

    // Row 2 shape must be gone
    const deletedShape = shapesAfter.find(s => s.text === 'R2-Badge-Delete')
    expect(deletedShape).toBeUndefined()

    // Row 3 shape must now be at Row 2, and its Y coordinate shifted upward!
    const shiftedR3NowR2 = shapesAfter.find(s => s.name?.includes('CellShape_RemoveRowTable_r2_c1'))
    expect(shiftedR3NowR2).toBeDefined()
    expect(shiftedR3NowR2.text).toBe('R3-Badge-Shift')
    expect(shiftedR3NowR2.y).toBeLessThan(origR3Y)

    // Row 1 shape remains unaffected
    const r1Shape = shapesAfter.find(s => s.name?.includes('CellShape_RemoveRowTable_r1_c1'))
    expect(r1Shape).toBeDefined()
    expect(r1Shape.text).toBe('R1-Badge')
  })

  // --------------------------------------------------------------------------
  // Bug #8: removeTableColumn Purging and Leftward Shape Shifting
  // --------------------------------------------------------------------------
  it('Bug #8: removeTableColumn purges column shapes and shifts right-side shapes leftward', async () => {
    pptx.addTable(
      1,
      2,
      4,
      [
        [
          'Col0',
          {
            shape: { type: 'roundRect', fill: '#3B82F6', width: 50, height: 20 },
            text: 'C1-Badge',
            textPosition: 'inside',
          },
          {
            shape: { type: 'circle', fill: '#EF4444', width: 14, height: 14 },
            text: 'C2-Badge-Delete',
            textPosition: 'inside',
          },
          {
            shape: { type: 'diamond', fill: '#10B981', width: 16, height: 16 },
            text: 'C3-Badge-Shift',
            textPosition: 'inside',
          },
        ],
        ['R1C0', 'R1C1', 'R1C2', 'R1C3'],
      ],
      { id: 'RemoveColTable', x: 100, width: 400 }
    )

    const initialC3 = pptx
      .getShapes(1)
      .find(s => s.name?.includes('CellShape_RemoveColTable_r0_c3'))
    expect(initialC3).toBeDefined()
    const origC3X = initialC3.x

    // Remove column 2 (which contains C2-Badge-Delete)
    pptx.removeTableColumn('RemoveColTable', 2, 1)

    const shapesAfter = pptx.getShapes(1)

    // Column 2 shape must be deleted
    expect(shapesAfter.find(s => s.text === 'C2-Badge-Delete')).toBeUndefined()

    // Column 3 shape must now be at Column 2, and its X coordinate shifted leftward!
    const shiftedC3NowC2 = shapesAfter.find(s => s.name?.includes('CellShape_RemoveColTable_r0_c2'))
    expect(shiftedC3NowC2).toBeDefined()
    expect(shiftedC3NowC2.text).toBe('C3-Badge-Shift')
    expect(shiftedC3NowC2.x).toBeLessThan(origC3X)

    // Column 1 shape remains unaffected
    const c1Shape = shapesAfter.find(s => s.name?.includes('CellShape_RemoveColTable_r0_c1'))
    expect(c1Shape).toBeDefined()
    expect(c1Shape.text).toBe('C1-Badge')
  })

  // --------------------------------------------------------------------------
  // Bug #9, #10, #11: Cascading Purging in clearRowData, clearColumnData, clearCellData
  // --------------------------------------------------------------------------
  it('Bugs #9, #10, #11: clearRowData, clearColumnData, and clearCellData cleanly purge associated shapes', async () => {
    pptx.addTable(
      1,
      3,
      3,
      [
        [
          {
            shape: { type: 'roundRect', fill: '#10B981', width: 40, height: 18 },
            text: 'R0C0',
            textPosition: 'inside',
          },
          {
            shape: { type: 'roundRect', fill: '#10B981', width: 40, height: 18 },
            text: 'R0C1',
            textPosition: 'inside',
          },
          {
            shape: { type: 'roundRect', fill: '#10B981', width: 40, height: 18 },
            text: 'R0C2',
            textPosition: 'inside',
          },
        ],
        [
          {
            shape: { type: 'circle', fill: '#3B82F6', width: 14, height: 14 },
            text: 'R1C0',
            textPosition: 'inside',
          },
          {
            shape: { type: 'circle', fill: '#3B82F6', width: 14, height: 14 },
            text: 'R1C1',
            textPosition: 'inside',
          },
          {
            shape: { type: 'circle', fill: '#3B82F6', width: 14, height: 14 },
            text: 'R1C2',
            textPosition: 'inside',
          },
        ],
        [
          {
            shape: { type: 'diamond', fill: '#F59E0B', width: 14, height: 14 },
            text: 'R2C0',
            textPosition: 'inside',
          },
          {
            shape: { type: 'diamond', fill: '#F59E0B', width: 14, height: 14 },
            text: 'R2C1',
            textPosition: 'inside',
          },
          {
            shape: { type: 'diamond', fill: '#F59E0B', width: 14, height: 14 },
            text: 'R2C2',
            textPosition: 'inside',
          },
        ],
      ],
      { id: 'ClearShapesTable' }
    )

    expect(
      pptx.getShapes(1).filter(s => s.name?.startsWith('CellShape_ClearShapesTable_')).length
    ).toBe(9)

    // 1. Clear cell (1, 1)
    pptx.clearCellData('ClearShapesTable', 1, 1, 1)
    expect(
      pptx.getShapes(1).find(s => s.name?.includes('CellShape_ClearShapesTable_r1_c1'))
    ).toBeUndefined()
    expect(
      pptx.getShapes(1).filter(s => s.name?.startsWith('CellShape_ClearShapesTable_')).length
    ).toBe(8)

    // 2. Clear entire row 0
    pptx.clearRowData('ClearShapesTable', 0, 1)
    const shapesAfterRowClear = pptx
      .getShapes(1)
      .filter(s => s.name?.startsWith('CellShape_ClearShapesTable_'))
    expect(shapesAfterRowClear.some(s => s.name?.includes('_r0_'))).toBe(false)
    expect(shapesAfterRowClear.length).toBe(5)

    // 3. Clear entire column 0
    pptx.clearColumnData('ClearShapesTable', 0, 1)
    const shapesAfterColClear = pptx
      .getShapes(1)
      .filter(s => s.name?.startsWith('CellShape_ClearShapesTable_'))
    expect(shapesAfterColClear.some(s => s.name?.includes('_c0'))).toBe(false)
    expect(shapesAfterColClear.length).toBe(3)
  })

  // --------------------------------------------------------------------------
  // Full Export & Round-Trip Reload Verification
  // --------------------------------------------------------------------------
  it('Round-Trip: Presentation with complex table lifecycle exports and reloads cleanly', async () => {
    pptx.addTable(
      1,
      3,
      3,
      [
        [
          'KPI',
          'Value',
          {
            shape: { type: 'roundRect', fill: '#10B981', width: 60, height: 20 },
            text: 'Target',
            textPosition: 'inside',
          },
        ],
        ['Sales', '$2M', 'Normal'],
        ['Retention', '94%', 'Normal'],
      ],
      { id: 'ExportTable', x: 50, y: 50, width: 500, height: 200 }
    )

    pptx.addTableRow('ExportTable', 1, ['Profit', '$500K', 'On Track'], 1)
    pptx.updateTableCell('ExportTable', 2, 2, {
      shape: { type: 'circle', fill: '#EF4444', width: 12, height: 12 },
      text: 'Critical',
      textPosition: 'outside',
      shapePlacement: 'left',
    })

    const buffer = await pptx.toBuffer()
    expect(buffer).toBeInstanceOf(Buffer)
    expect(buffer.length).toBeGreaterThan(1000)

    const reloaded = await PPTXTemplater.load(buffer)

    const reloadedTable = reloaded.getTable('ExportTable', 1)
    expect(reloadedTable).toBeDefined()
    expect(reloadedTable.rows).toBe(4)
    expect(reloadedTable.columns).toBe(3)

    const reloadedShapes = reloaded.getShapes(1)
    const tableShapes = reloadedShapes.filter(s => s.name?.startsWith('CellShape_ExportTable_'))
    expect(tableShapes.length).toBe(2)
  })
})
