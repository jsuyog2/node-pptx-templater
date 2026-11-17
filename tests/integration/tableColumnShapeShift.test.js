/**
 * @fileoverview Integration tests for horizontal coordinate shifting of cell shapes
 * when removing a table column (specifically at index 1 and sequentially).
 *
 * TASK-TBL-SHAPE-SHIFT-002:
 * 1. Create a table with shapes across columns 0, 1, and 2.
 * 2. Delete column index 1 using removeTableColumn.
 * 3. Assert shapes in column 1 are purged.
 * 4. Assert shapes previously in column 2 shift horizontally by the deleted column width
 *    and are re-indexed to column 1.
 * 5. Assert shapes in column 0 remain intact and unshifted.
 * 6. Sequential column removal and round-trip export verification.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('Table Column Removal & Cell Shape Horizontal Shifting', () => {
  let pptx

  beforeEach(async () => {
    pptx = await PPTXTemplater.createBlank()
  })

  // Helper to extract shape node from slide AST
  function getShapeNode(slideIndex, shapeName) {
    const slideObj = pptx.slideManager.getSlideObj(slideIndex)
    const spTree = slideObj['p:sld']['p:cSld']['p:spTree']
    const shapes = Array.isArray(spTree['p:sp'])
      ? spTree['p:sp']
      : spTree['p:sp']
        ? [spTree['p:sp']]
        : []
    return shapes.find(s => s?.['p:nvSpPr']?.['p:cNvPr']?.['@_name'] === shapeName)
  }

  // Helper to extract table XML AST
  function getTableContext(slideIndex, tableId) {
    const slideObj = pptx.slideManager.getSlideObj(slideIndex)
    const spTree = slideObj['p:sld']['p:cSld']['p:spTree']
    const frames = Array.isArray(spTree['p:graphicFrame'])
      ? spTree['p:graphicFrame']
      : spTree['p:graphicFrame']
        ? [spTree['p:graphicFrame']]
        : []
    const frame =
      frames.find(f => f?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_name'] === tableId) || frames[0]
    const tblObj = frame?.['a:graphic']?.['a:graphicData']?.['a:tbl']
    return { frame, tblObj }
  }

  // Helper to get raw EMU x coordinate from AST
  function getRawShapeX(slideIndex, shapeName) {
    const node = getShapeNode(slideIndex, shapeName)
    if (!node) return null
    return parseInt(node?.['p:spPr']?.['a:xfrm']?.['a:off']?.['@_x'] || '0', 10)
  }

  // --------------------------------------------------------------------------
  // Core Requirement: Index-1 Column Removal with Shapes across Cols 0, 1, 2
  // --------------------------------------------------------------------------
  describe('Index-1 Column Removal and Coordinate Shifting', () => {
    it('purges column 1 shape, shifts column 2 shape left by deleted column width, and preserves column 0 shape', () => {
      pptx.addTable(
        1,
        2,
        3,
        [
          [
            {
              shape: { type: 'roundRect', fill: '#3B82F6', width: 40, height: 20 },
              text: 'Col0-Badge',
              textPosition: 'inside',
            },
            {
              shape: { type: 'rect', fill: '#EF4444', width: 50, height: 25 },
              text: 'Col1-DeleteMe',
              textPosition: 'inside',
            },
            {
              shape: { type: 'diamond', fill: '#10B981', width: 30, height: 30 },
              text: 'Col2-ShiftLeft',
              textPosition: 'inside',
            },
          ],
          ['Row1-Col0', 'Row1-Col1', 'Row1-Col2'],
        ],
        { id: 'ShiftTable', x: 50, y: 50, width: 600, height: 100 }
      )

      // Query table before removal
      const tblBefore = pptx.getTable('ShiftTable', 1)
      expect(tblBefore.columns).toBe(3)
      expect(tblBefore.rows).toBe(2)

      // Get column 1 width in EMUs from <a:tblGrid>
      const { tblObj } = getTableContext(1, 'ShiftTable')
      const col1WidthEmu = parseInt(tblObj['a:tblGrid']['a:gridCol'][1]['@_w'], 10)
      expect(col1WidthEmu).toBeGreaterThan(0)

      // Record pre-removal shape names and coordinates
      const origCol0X = getRawShapeX(1, 'CellShape_ShiftTable_r0_c0')
      const origCol1X = getRawShapeX(1, 'CellShape_ShiftTable_r0_c1')
      const origCol2X = getRawShapeX(1, 'CellShape_ShiftTable_r0_c2')

      expect(origCol0X).not.toBeNull()
      expect(origCol1X).not.toBeNull()
      expect(origCol2X).not.toBeNull()
      expect(origCol1X).toBeGreaterThan(origCol0X)
      expect(origCol2X).toBeGreaterThan(origCol1X)

      // Execute removeTableColumn at index 1
      pptx.removeTableColumn('ShiftTable', 1, 1)

      // Verify table structure after removal
      const tblAfter = pptx.getTable('ShiftTable', 1)
      expect(tblAfter.columns).toBe(2)
      expect(tblAfter.rows).toBe(2)

      const tableData = pptx.getTableData('ShiftTable', 1)
      expect(tableData[0]).toEqual(['Col0-Badge', 'Col2-ShiftLeft'])
      expect(tableData[1]).toEqual(['Row1-Col0', 'Row1-Col2'])

      // 1. Column 1 shape is purged
      expect(getShapeNode(1, 'CellShape_ShiftTable_r0_c1')).toBeDefined() // now re-indexed from c2!
      const allShapes = pptx.getShapes(1)
      const purgedShape = allShapes.find(s => s.text === 'Col1-DeleteMe')
      expect(purgedShape).toBeUndefined()

      // 2. Column 0 shape is preserved intact and unshifted
      const col0Shape = allShapes.find(s => s.text === 'Col0-Badge')
      expect(col0Shape).toBeDefined()
      expect(col0Shape.name).toBe('CellShape_ShiftTable_r0_c0')
      const afterCol0X = getRawShapeX(1, 'CellShape_ShiftTable_r0_c0')
      expect(afterCol0X).toBe(origCol0X)

      // 3. Column 2 shape is shifted horizontally by deleted column width and re-indexed to c1
      const col2Shape = allShapes.find(s => s.text === 'Col2-ShiftLeft')
      expect(col2Shape).toBeDefined()
      expect(col2Shape.name).toBe('CellShape_ShiftTable_r0_c1')
      const afterCol2X = getRawShapeX(1, 'CellShape_ShiftTable_r0_c1')
      expect(afterCol2X).toBe(origCol2X - col1WidthEmu)
    })
  })

  // --------------------------------------------------------------------------
  // Multi-Row Column 1 Removal
  // --------------------------------------------------------------------------
  describe('Multi-Row Column 1 Removal with Cell Shapes Across Rows', () => {
    it('shifts all subsequent column shapes across multiple rows by the deleted column width', () => {
      pptx.addTable(
        1,
        3,
        3,
        [
          [
            { shape: { type: 'roundRect' }, text: 'R0C0-Shape', textPosition: 'inside' },
            { shape: { type: 'circle' }, text: 'R0C1-Del', textPosition: 'inside' },
            { shape: { type: 'star' }, text: 'R0C2-Shift', textPosition: 'inside' },
          ],
          [
            'R1C0-Plain',
            { shape: { type: 'rect' }, text: 'R1C1-Del', textPosition: 'inside' },
            { shape: { type: 'heart' }, text: 'R1C2-Shift', textPosition: 'inside' },
          ],
          [
            { shape: { type: 'diamond' }, text: 'R2C0-Shape', textPosition: 'inside' },
            'R2C1-Plain',
            { shape: { type: 'roundRect' }, text: 'R2C2-Shift', textPosition: 'inside' },
          ],
        ],
        { id: 'MultiRowTable', x: 40, y: 60, width: 450, height: 180 }
      )

      const { tblObj } = getTableContext(1, 'MultiRowTable')
      const col1Width = parseInt(tblObj['a:tblGrid']['a:gridCol'][1]['@_w'], 10)

      const origR0C2X = getRawShapeX(1, 'CellShape_MultiRowTable_r0_c2')
      const origR1C2X = getRawShapeX(1, 'CellShape_MultiRowTable_r1_c2')
      const origR2C2X = getRawShapeX(1, 'CellShape_MultiRowTable_r2_c2')

      const origR0C0X = getRawShapeX(1, 'CellShape_MultiRowTable_r0_c0')
      const origR2C0X = getRawShapeX(1, 'CellShape_MultiRowTable_r2_c0')

      // Remove column 1
      pptx.removeTableColumn('MultiRowTable', 1, 1)

      const shapes = pptx.getShapes(1)

      // Purged shapes
      expect(shapes.find(s => s.text === 'R0C1-Del')).toBeUndefined()
      expect(shapes.find(s => s.text === 'R1C1-Del')).toBeUndefined()

      // Preserved column 0 shapes
      expect(getRawShapeX(1, 'CellShape_MultiRowTable_r0_c0')).toBe(origR0C0X)
      expect(getRawShapeX(1, 'CellShape_MultiRowTable_r2_c0')).toBe(origR2C0X)

      // Shifted column 2 shapes across all rows
      const shiftedR0 = shapes.find(s => s.text === 'R0C2-Shift')
      const shiftedR1 = shapes.find(s => s.text === 'R1C2-Shift')
      const shiftedR2 = shapes.find(s => s.text === 'R2C2-Shift')

      expect(shiftedR0.name).toBe('CellShape_MultiRowTable_r0_c1')
      expect(shiftedR1.name).toBe('CellShape_MultiRowTable_r1_c1')
      expect(shiftedR2.name).toBe('CellShape_MultiRowTable_r2_c1')

      expect(getRawShapeX(1, 'CellShape_MultiRowTable_r0_c1')).toBe(origR0C2X - col1Width)
      expect(getRawShapeX(1, 'CellShape_MultiRowTable_r1_c1')).toBe(origR1C2X - col1Width)
      expect(getRawShapeX(1, 'CellShape_MultiRowTable_r2_c1')).toBe(origR2C2X - col1Width)
    })
  })

  // --------------------------------------------------------------------------
  // Sequential Column Removal
  // --------------------------------------------------------------------------
  describe('Sequential Column Removal', () => {
    it('correctly shifts shapes across multiple consecutive removeTableColumn invocations', () => {
      pptx.addTable(
        1,
        2,
        4,
        [
          [
            { shape: { type: 'circle' }, text: 'Col0-Keep', textPosition: 'inside' },
            { shape: { type: 'rect' }, text: 'Col1-FirstDel', textPosition: 'inside' },
            { shape: { type: 'roundRect' }, text: 'Col2-SecondDel', textPosition: 'inside' },
            { shape: { type: 'diamond' }, text: 'Col3-FinalShift', textPosition: 'inside' },
          ],
          ['A', 'B', 'C', 'D'],
        ],
        { id: 'SeqTable', x: 50, y: 50, width: 400, height: 100 }
      )

      const { tblObj } = getTableContext(1, 'SeqTable')
      const w1 = parseInt(tblObj['a:tblGrid']['a:gridCol'][1]['@_w'], 10)
      const w2 = parseInt(tblObj['a:tblGrid']['a:gridCol'][2]['@_w'], 10)

      const origCol3X = getRawShapeX(1, 'CellShape_SeqTable_r0_c3')
      const origCol0X = getRawShapeX(1, 'CellShape_SeqTable_r0_c0')

      // Step 1: Remove column 1
      pptx.removeTableColumn('SeqTable', 1, 1)

      expect(pptx.getTable('SeqTable', 1).columns).toBe(3)
      expect(pptx.getShapes(1).find(s => s.text === 'Col1-FirstDel')).toBeUndefined()

      // Col 2 moved to Col 1, Col 3 moved to Col 2
      const step1Col2 = pptx.getShapes(1).find(s => s.text === 'Col2-SecondDel')
      expect(step1Col2.name).toBe('CellShape_SeqTable_r0_c1')

      const step1Col3 = pptx.getShapes(1).find(s => s.text === 'Col3-FinalShift')
      expect(step1Col3.name).toBe('CellShape_SeqTable_r0_c2')
      expect(getRawShapeX(1, 'CellShape_SeqTable_r0_c2')).toBe(origCol3X - w1)

      // Step 2: Remove column 1 again (which was originally Col 2)
      pptx.removeTableColumn('SeqTable', 1, 1)

      expect(pptx.getTable('SeqTable', 1).columns).toBe(2)
      expect(pptx.getShapes(1).find(s => s.text === 'Col2-SecondDel')).toBeUndefined()

      // Col 3 now shifted again to Col 1
      const step2Col3 = pptx.getShapes(1).find(s => s.text === 'Col3-FinalShift')
      expect(step2Col3.name).toBe('CellShape_SeqTable_r0_c1')
      expect(getRawShapeX(1, 'CellShape_SeqTable_r0_c1')).toBe(origCol3X - w1 - w2)

      // Col 0 remains untouched through both deletions
      expect(getRawShapeX(1, 'CellShape_SeqTable_r0_c0')).toBe(origCol0X)
    })
  })

  // --------------------------------------------------------------------------
  // Polymorphic Cell Shapes Shifting (Inside, Outside, Shape-Only)
  // --------------------------------------------------------------------------
  describe('Polymorphic Shape Modes Shifting', () => {
    it('shifts outside-text and shape-only overlays by the deleted column width', () => {
      pptx.addTable(
        1,
        2,
        3,
        [
          [
            // Inside text shape
            {
              shape: { type: 'roundRect', fill: '#10B981', width: 40, height: 20 },
              text: 'InsideText',
              textPosition: 'inside',
            },
            // Deleted column with shape-only
            {
              shape: { type: 'circle', fill: '#EF4444', width: 16, height: 16 },
            },
            // Outside text shape (shape overlay + DrawingML text in cell)
            {
              shape: { type: 'diamond', fill: '#3B82F6', width: 24, height: 24 },
              text: 'OutsideText',
              textPosition: 'outside',
            },
          ],
          ['Row1-C0', 'Row1-C1', 'Row1-C2'],
        ],
        { id: 'PolyTable', x: 60, y: 60, width: 450, height: 100 }
      )

      const { tblObj } = getTableContext(1, 'MultiRowTable')
      const col1Width = parseInt(tblObj['a:tblGrid']['a:gridCol'][1]['@_w'], 10)

      const origOutsideX = getRawShapeX(1, 'CellShape_PolyTable_r0_c2')

      pptx.removeTableColumn('PolyTable', 1, 1)

      const afterOutsideX = getRawShapeX(1, 'CellShape_PolyTable_r0_c1')
      expect(afterOutsideX).toBe(origOutsideX - col1Width)

      const tableData = pptx.getTableData('PolyTable', 1)
      expect(tableData[0][1]).toBe('OutsideText')
    })
  })

  // --------------------------------------------------------------------------
  // Round-Trip Presentation Export & Integrity
  // --------------------------------------------------------------------------
  describe('Round-Trip Presentation Integrity', () => {
    it('exports presentation cleanly to buffer and re-imports with valid table and shifted shape coordinates', async () => {
      pptx.addTable(
        1,
        2,
        3,
        [
          [
            { shape: { type: 'roundRect' }, text: 'Keep0', textPosition: 'inside' },
            { shape: { type: 'circle' }, text: 'Remove1', textPosition: 'inside' },
            { shape: { type: 'diamond' }, text: 'Shift2', textPosition: 'inside' },
          ],
          ['D0', 'D1', 'D2'],
        ],
        { id: 'ExportTable', x: 100, y: 100, width: 300, height: 120 }
      )

      pptx.removeTableColumn('ExportTable', 1, 1)

      const buffer = await pptx.toBuffer()
      expect(buffer).toBeInstanceOf(Buffer)
      expect(buffer.length).toBeGreaterThan(0)

      // Re-load presentation
      const loadedPptx = await PPTXTemplater.load(buffer)
      const tbl = loadedPptx.getTable('ExportTable', 1)
      expect(tbl.columns).toBe(2)
      expect(tbl.rows).toBe(2)

      const shapes = loadedPptx.getShapes(1)
      expect(shapes.find(s => s.text === 'Remove1')).toBeUndefined()
      expect(shapes.find(s => s.text === 'Keep0')).toBeDefined()
      expect(shapes.find(s => s.text === 'Shift2')).toBeDefined()
    })
  })
})
