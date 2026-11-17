/**
 * @fileoverview Integration tests for Strict Geometric Bounding-Box Shape Removal
 * in Table Operations (removeTable, removeTableRow, removeTableColumn, clearRowData,
 * clearColumnData, and clearCellData).
 *
 * Verifies the strict inequality containment rule:
 * shape.xMin > target.xMin AND shape.yMin > target.yMin AND
 * shape.xMax < target.xMax AND shape.yMax < target.yMax
 *
 * Shapes touching edges or crossing borders MUST be preserved.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'
import { isStrictlyContainedWithin } from '../../src/elements/CellShapeManager.js'

describe('Strict Geometric Bounding-Box Shape Removal Subsystem', () => {
  let pptx

  beforeEach(async () => {
    pptx = await PPTXTemplater.createBlank()
  })

  // --------------------------------------------------------------------------
  // 1. Mathematical Containment Filter Unit Invariants
  // --------------------------------------------------------------------------
  describe('isStrictlyContainedWithin mathematical inequality contract', () => {
    const target = { xMin: 1000, yMin: 1000, xMax: 5000, yMax: 5000 }

    it('returns true when shape is strictly interior', () => {
      const shape = { xMin: 1500, yMin: 1500, xMax: 4500, yMax: 4500 }
      expect(isStrictlyContainedWithin(shape, target)).toBe(true)
    })

    it('returns false when shape touches the left edge (xMin === target.xMin)', () => {
      const shape = { xMin: 1000, yMin: 1500, xMax: 4500, yMax: 4500 }
      expect(isStrictlyContainedWithin(shape, target)).toBe(false)
    })

    it('returns false when shape touches the right edge (xMax === target.xMax)', () => {
      const shape = { xMin: 1500, yMin: 1500, xMax: 5000, yMax: 4500 }
      expect(isStrictlyContainedWithin(shape, target)).toBe(false)
    })

    it('returns false when shape touches the top edge (yMin === target.yMin)', () => {
      const shape = { xMin: 1500, yMin: 1000, xMax: 4500, yMax: 4500 }
      expect(isStrictlyContainedWithin(shape, target)).toBe(false)
    })

    it('returns false when shape touches the bottom edge (yMax === target.yMax)', () => {
      const shape = { xMin: 1500, yMin: 1500, xMax: 4500, yMax: 5000 }
      expect(isStrictlyContainedWithin(shape, target)).toBe(false)
    })

    it('returns false when shape completely aligns with target (all 4 edges touch)', () => {
      const shape = { xMin: 1000, yMin: 1000, xMax: 5000, yMax: 5000 }
      expect(isStrictlyContainedWithin(shape, target)).toBe(false)
    })

    it('returns false when shape crosses an edge / overlaps outside', () => {
      const shape = { xMin: 800, yMin: 1500, xMax: 4500, yMax: 4500 }
      expect(isStrictlyContainedWithin(shape, target)).toBe(false)
    })

    it('returns false when shape is completely external', () => {
      const shape = { xMin: 6000, yMin: 6000, xMax: 7000, yMax: 7000 }
      expect(isStrictlyContainedWithin(shape, target)).toBe(false)
    })

    it('handles width/height bounding format { x, y, width, height }', () => {
      const t = { x: 1000, y: 1000, width: 4000, height: 4000 }
      const sInterior = { x: 1200, y: 1200, width: 2000, height: 2000 }
      const sTouch = { x: 1000, y: 1200, width: 2000, height: 2000 }
      expect(isStrictlyContainedWithin(sInterior, t)).toBe(true)
      expect(isStrictlyContainedWithin(sTouch, t)).toBe(false)
    })

    it('returns false when shape or target is null/undefined', () => {
      expect(isStrictlyContainedWithin(null, target)).toBe(false)
      expect(isStrictlyContainedWithin(target, null)).toBe(false)
    })
  })

  // --------------------------------------------------------------------------
  // 2. Table Bounding Box Calculation Engine
  // --------------------------------------------------------------------------
  describe('Table Bounding Box Calculation Engine', () => {
    it('computes exact EMU and pixel bounds for table, row, column, and cell', () => {
      pptx.addTable(
        1,
        3,
        3,
        [
          ['R0C0', 'R0C1', 'R0C2'],
          ['R1C0', 'R1C1', 'R1C2'],
          ['R2C0', 'R2C1', 'R2C2'],
        ],
        { id: 'BoundsTable', x: 100, y: 100, width: 300, height: 150 }
      )

      const tm = pptx.tableManager
      const sm = pptx.slideManager

      const tblBounds = tm.getTableBounds('BoundsTable', 1, sm)
      expect(tblBounds).toBeDefined()
      expect(tblBounds.x).toBe(952500)
      expect(tblBounds.y).toBe(952500)
      expect(tblBounds.xMin).toBe(tblBounds.x)
      expect(tblBounds.xMax).toBe(tblBounds.x + tblBounds.width)
      expect(tblBounds.yMin).toBe(tblBounds.y)
      expect(tblBounds.yMax).toBe(tblBounds.y + tblBounds.height)

      const rowBounds1 = tm.getTableRowBounds('BoundsTable', 1, 1, sm)
      expect(rowBounds1).toBeDefined()
      expect(rowBounds1.yMin).toBeGreaterThan(tblBounds.yMin)
      expect(rowBounds1.yMax).toBeLessThanOrEqual(tblBounds.yMax)
      expect(rowBounds1.width).toBe(tblBounds.width)

      const colBounds1 = tm.getTableColumnBounds('BoundsTable', 1, 1, sm)
      expect(colBounds1).toBeDefined()
      expect(colBounds1.xMin).toBeGreaterThan(tblBounds.xMin)
      expect(colBounds1.xMax).toBeLessThanOrEqual(tblBounds.xMax)
      expect(colBounds1.height).toBe(tblBounds.height)

      const cellBounds11 = tm.getTableCellBounds('BoundsTable', 1, 1, 1, sm)
      expect(cellBounds11).toBeDefined()
      expect(cellBounds11.xMin).toBe(colBounds1.xMin)
      expect(cellBounds11.yMin).toBe(rowBounds1.yMin)
      expect(cellBounds11.width).toBe(colBounds1.width)
      expect(cellBounds11.height).toBe(rowBounds1.height)

      // Signature flexibility for getCellBounds
      const altCellBounds = tm.getCellBounds('BoundsTable', 1, 1, 1, sm)
      expect(altCellBounds.xMin).toBe(cellBounds11.xMin)
      expect(altCellBounds.yMin).toBe(cellBounds11.yMin)
    })
  })

  // Helper to inject a shape directly into slide <p:spTree> with explicit EMU bounds
  function injectShape(slideIndex, name, x, y, cx, cy) {
    const slideObj = pptx.slideManager.getSlideObj(slideIndex)
    const spTree = slideObj['p:sld']['p:cSld']['p:spTree']
    if (!spTree['p:sp']) spTree['p:sp'] = []
    if (!Array.isArray(spTree['p:sp'])) spTree['p:sp'] = [spTree['p:sp']]

    const sp = {
      'p:nvSpPr': {
        'p:cNvPr': { '@_id': name, '@_name': name },
        'p:cNvSpPr': {},
        'p:nvPr': {},
      },
      'p:spPr': {
        'a:xfrm': {
          'a:off': { '@_x': String(x), '@_y': String(y) },
          'a:ext': { '@_cx': String(cx), '@_cy': String(cy) },
        },
        'a:prstGeom': { '@_prst': 'rect', 'a:avLst': {} },
      },
    }
    spTree['p:sp'].push(sp)
    pptx.slideManager.markSlideObjDirty(slideIndex)
    return sp
  }

  function hasShape(name) {
    const slideObj = pptx.slideManager.getSlideObj(1)
    const spTree = slideObj['p:sld']['p:cSld']['p:spTree']
    const shapes = Array.isArray(spTree['p:sp'])
      ? spTree['p:sp']
      : spTree['p:sp']
        ? [spTree['p:sp']]
        : []
    return shapes.some(s => s?.['p:nvSpPr']?.['p:cNvPr']?.['@_name'] === name)
  }

  // --------------------------------------------------------------------------
  // 3. clearCellData Strict Containment
  // --------------------------------------------------------------------------
  describe('clearCellData strict boundary containment', () => {
    it('purges strictly interior shapes while preserving edge-touching, crossing, and adjacent shapes', () => {
      pptx.addTable(
        1,
        2,
        2,
        [
          ['C00', 'C01'],
          ['C10', 'C11'],
        ],
        { id: 'CellTable', x: 100, y: 100, width: 200, height: 100 }
      )

      const cellBounds = pptx.tableManager.getTableCellBounds(
        'CellTable',
        0,
        0,
        1,
        pptx.slideManager
      )
      const { xMin, yMin, xMax, yMax } = cellBounds

      // 1. Strictly interior shape
      injectShape(1, 'Shape_Interior', xMin + 1000, yMin + 1000, 10000, 10000)

      // 2. Touches left border (x === xMin) -> PRESERVE
      injectShape(1, 'Shape_TouchLeft', xMin, yMin + 1000, 10000, 10000)

      // 3. Touches right border (x + cx === xMax) -> PRESERVE
      injectShape(1, 'Shape_TouchRight', xMax - 10000, yMin + 1000, 10000, 10000)

      // 4. Crosses right boundary (overlaps adjacent cell) -> PRESERVE
      injectShape(1, 'Shape_OverlapsRight', xMin + 1000, yMin + 1000, xMax - xMin + 5000, 10000)

      // 5. Touches bottom border (y + cy === yMax) -> PRESERVE
      injectShape(1, 'Shape_TouchBottom', xMin + 1000, yMax - 10000, 10000, 10000)

      // 6. Shape in adjacent cell (0, 1) -> PRESERVE
      const adjBounds = pptx.tableManager.getTableCellBounds(
        'CellTable',
        0,
        1,
        1,
        pptx.slideManager
      )
      injectShape(
        1,
        'Shape_AdjacentCell',
        adjBounds.xMin + 1000,
        adjBounds.yMin + 1000,
        10000,
        10000
      )

      // Clear cell (0, 0)
      pptx.clearCellData('CellTable', 0, 0, 1)

      expect(hasShape('Shape_Interior')).toBe(false)
      expect(hasShape('Shape_TouchLeft')).toBe(true)
      expect(hasShape('Shape_TouchRight')).toBe(true)
      expect(hasShape('Shape_TouchBottom')).toBe(true)
      expect(hasShape('Shape_OverlapsRight')).toBe(true)
      expect(hasShape('Shape_AdjacentCell')).toBe(true)
    })
  })

  // --------------------------------------------------------------------------
  // 4. clearRowData Strict Containment
  // --------------------------------------------------------------------------
  describe('clearRowData strict boundary containment', () => {
    it('purges strictly interior row shapes while preserving border-touching, overlapping, and other row shapes', () => {
      pptx.addTable(
        1,
        3,
        2,
        [
          ['R0C0', 'R0C1'],
          ['R1C0', 'R1C1'],
          ['R2C0', 'R2C1'],
        ],
        { id: 'RowClearTable', x: 100, y: 100, width: 200, height: 150 }
      )

      const row1Bounds = pptx.tableManager.getTableRowBounds(
        'RowClearTable',
        1,
        1,
        pptx.slideManager
      )
      const { xMin, yMin, xMax, yMax } = row1Bounds

      // 1. Strictly interior shape in Row 1
      injectShape(1, 'Row1_Interior', xMin + 2000, yMin + 2000, 10000, 10000)

      // 2. Touches top edge of Row 1 (y === yMin) -> PRESERVE
      injectShape(1, 'Row1_TouchTop', xMin + 2000, yMin, 10000, 10000)

      // 3. Touches bottom edge of Row 1 (y + cy === yMax) -> PRESERVE
      injectShape(1, 'Row1_TouchBottom', xMin + 2000, yMax - 10000, 10000, 10000)

      // 4. Overlaps across bottom boundary into Row 2 -> PRESERVE
      injectShape(1, 'Row1_OverlapsBottom', xMin + 2000, yMin + 5000, 10000, yMax - yMin + 2000)

      // 5. Touches right edge of Row 1 (x + cx === xMax) -> PRESERVE
      injectShape(1, 'Row1_TouchRight', xMax - 10000, yMin + 2000, 10000, 10000)

      // 6. Shape strictly inside Row 2 -> PRESERVE
      const row2Bounds = pptx.tableManager.getTableRowBounds(
        'RowClearTable',
        2,
        1,
        pptx.slideManager
      )
      injectShape(1, 'Row2_Interior', row2Bounds.xMin + 2000, row2Bounds.yMin + 2000, 10000, 10000)

      // Clear row 1
      pptx.clearRowData('RowClearTable', 1, 1)

      expect(hasShape('Row1_Interior')).toBe(false)
      expect(hasShape('Row1_TouchTop')).toBe(true)
      expect(hasShape('Row1_TouchBottom')).toBe(true)
      expect(hasShape('Row1_TouchRight')).toBe(true)
      expect(hasShape('Row1_OverlapsBottom')).toBe(true)
      expect(hasShape('Row2_Interior')).toBe(true)
    })
  })

  // --------------------------------------------------------------------------
  // 5. clearColumnData Strict Containment
  // --------------------------------------------------------------------------
  describe('clearColumnData strict boundary containment', () => {
    it('purges strictly interior column shapes while preserving border-touching, overlapping, and other column shapes', () => {
      pptx.addTable(
        1,
        2,
        3,
        [
          ['R0C0', 'R0C1', 'R0C2'],
          ['R1C0', 'R1C1', 'R1C2'],
        ],
        { id: 'ColClearTable', x: 100, y: 100, width: 300, height: 100 }
      )

      const col1Bounds = pptx.tableManager.getTableColumnBounds(
        'ColClearTable',
        1,
        1,
        pptx.slideManager
      )
      const { xMin, yMin, xMax, yMax } = col1Bounds

      // 1. Strictly interior shape in Col 1
      injectShape(1, 'Col1_Interior', xMin + 2000, yMin + 2000, 10000, 10000)

      // 2. Touches left edge of Col 1 (x === xMin) -> PRESERVE
      injectShape(1, 'Col1_TouchLeft', xMin, yMin + 2000, 10000, 10000)

      // 3. Touches right edge of Col 1 (x + cx === xMax) -> PRESERVE
      injectShape(1, 'Col1_TouchRight', xMax - 10000, yMin + 2000, 10000, 10000)

      // 4. Overlaps across right boundary into Col 2 -> PRESERVE
      injectShape(1, 'Col1_OverlapsRight', xMin + 2000, yMin + 2000, xMax - xMin + 5000, 10000)

      // 5. Touches bottom edge of Col 1 (y + cy === yMax) -> PRESERVE
      injectShape(1, 'Col1_TouchBottom', xMin + 2000, yMax - 10000, 10000, 10000)

      // 6. Shape strictly inside Col 2 -> PRESERVE
      const col2Bounds = pptx.tableManager.getTableColumnBounds(
        'ColClearTable',
        2,
        1,
        pptx.slideManager
      )
      injectShape(1, 'Col2_Interior', col2Bounds.xMin + 2000, col2Bounds.yMin + 2000, 10000, 10000)

      // Clear column 1
      pptx.clearColumnData('ColClearTable', 1, 1)

      expect(hasShape('Col1_Interior')).toBe(false)
      expect(hasShape('Col1_TouchLeft')).toBe(true)
      expect(hasShape('Col1_TouchRight')).toBe(true)
      expect(hasShape('Col1_TouchBottom')).toBe(true)
      expect(hasShape('Col1_OverlapsRight')).toBe(true)
      expect(hasShape('Col2_Interior')).toBe(true)
    })
  })

  // --------------------------------------------------------------------------
  // 6. removeTableRow Strict Containment & Non-Shifting of Edge/External Shapes
  // --------------------------------------------------------------------------
  describe('removeTableRow strict boundary containment & selective coordinate shifting', () => {
    it('purges strictly interior shapes in target row, preserves touching/overlapping shapes, and shifts only interior shapes', () => {
      pptx.addTable(
        1,
        4,
        2,
        [
          ['Header 0', 'Header 1'],
          ['Row 1 Col 0', 'Row 1 Col 1'],
          [
            'Row 2 Col 0',
            {
              shape: { type: 'circle', width: 20, height: 20 },
              text: 'R2-Interior',
              textPosition: 'inside',
            },
          ],
          [
            'Row 3 Col 0',
            {
              shape: { type: 'roundRect', width: 20, height: 20 },
              text: 'R3-Interior',
              textPosition: 'inside',
            },
          ],
        ],
        { id: 'RemoveRowTable', x: 100, y: 100, width: 200, height: 200 }
      )

      const row2Bounds = pptx.tableManager.getTableRowBounds(
        'RemoveRowTable',
        2,
        1,
        pptx.slideManager
      )
      const tblBounds = pptx.tableManager.getTableBounds('RemoveRowTable', 1, pptx.slideManager)

      // Edge-touching shape on Row 2 (touches top of Row 2) -> PRESERVE
      injectShape(1, 'R2_TouchTop', row2Bounds.xMin + 2000, row2Bounds.yMin, 10000, 10000)

      // External shape outside table -> MUST NOT SHIFT
      const extShape = injectShape(
        1,
        'External_Shape',
        tblBounds.xMax + 50000,
        tblBounds.yMin + 50000,
        20000,
        20000
      )
      const origExtY = parseInt(extShape['p:spPr']['a:xfrm']['a:off']['@_y'], 10)

      const r3Shape = pptx.getShapes(1).find(s => s.text === 'R3-Interior')
      expect(r3Shape).toBeDefined()
      const origR3Y = r3Shape.y

      // Remove row 2
      pptx.removeTableRow('RemoveRowTable', 2, 1)

      // 1. Interior Row 2 shape is purged
      expect(pptx.getShapes(1).find(s => s.text === 'R2-Interior')).toBeUndefined()

      // 2. Edge-touching Row 2 shape is PRESERVED
      expect(hasShape('R2_TouchTop')).toBe(true)

      // 3. Lower interior shape (R3-Interior) shifted upward
      const shiftedR3 = pptx.getShapes(1).find(s => s.text === 'R3-Interior')
      expect(shiftedR3).toBeDefined()
      expect(shiftedR3.y).toBeLessThan(origR3Y)

      // 4. External shape was NOT shifted
      const slide1 = pptx.slideManager.getSlideObj(1)
      const currentExtShape = slide1['p:sld']['p:cSld']['p:spTree']['p:sp'].find(
        s => s?.['p:nvSpPr']?.['p:cNvPr']?.['@_name'] === 'External_Shape'
      )
      const currentExtY = parseInt(currentExtShape['p:spPr']['a:xfrm']['a:off']['@_y'], 10)
      expect(currentExtY).toBe(origExtY)
    })
  })

  // --------------------------------------------------------------------------
  // 7. removeTableColumn Strict Containment & Non-Shifting of Edge/External Shapes
  // --------------------------------------------------------------------------
  describe('removeTableColumn strict boundary containment & selective coordinate shifting', () => {
    it('purges strictly interior shapes in target col, preserves touching/overlapping shapes, and shifts only interior shapes', () => {
      pptx.addTable(
        1,
        2,
        4,
        [
          [
            'Col 0',
            'Col 1',
            {
              shape: { type: 'circle', width: 20, height: 20 },
              text: 'C2-Interior',
              textPosition: 'inside',
            },
            {
              shape: { type: 'diamond', width: 20, height: 20 },
              text: 'C3-Interior',
              textPosition: 'inside',
            },
          ],
          ['R1C0', 'R1C1', 'R1C2', 'R1C3'],
        ],
        { id: 'RemoveColTable', x: 100, y: 100, width: 400, height: 100 }
      )

      const col2Bounds = pptx.tableManager.getTableColumnBounds(
        'RemoveColTable',
        2,
        1,
        pptx.slideManager
      )
      const tblBounds = pptx.tableManager.getTableBounds('RemoveColTable', 1, pptx.slideManager)

      // Edge-touching shape on Col 2 (touches left edge of Col 2) -> PRESERVE
      injectShape(1, 'C2_TouchLeft', col2Bounds.xMin, col2Bounds.yMin + 2000, 10000, 10000)

      // External shape outside table -> MUST NOT SHIFT
      const extShape = injectShape(
        1,
        'External_Col_Shape',
        tblBounds.xMax + 50000,
        tblBounds.yMin + 20000,
        20000,
        20000
      )
      const origExtX = parseInt(extShape['p:spPr']['a:xfrm']['a:off']['@_x'], 10)

      const c3Shape = pptx.getShapes(1).find(s => s.text === 'C3-Interior')
      expect(c3Shape).toBeDefined()
      const origC3X = c3Shape.x

      // Remove column 2
      pptx.removeTableColumn('RemoveColTable', 2, 1)

      // 1. Interior Col 2 shape is purged
      expect(pptx.getShapes(1).find(s => s.text === 'C2-Interior')).toBeUndefined()

      // 2. Edge-touching Col 2 shape is PRESERVED
      expect(hasShape('C2_TouchLeft')).toBe(true)

      // 3. Right interior shape (C3-Interior) shifted leftward
      const shiftedC3 = pptx.getShapes(1).find(s => s.text === 'C3-Interior')
      expect(shiftedC3).toBeDefined()
      expect(shiftedC3.x).toBeLessThan(origC3X)

      // 4. External shape was NOT shifted
      const slide1 = pptx.slideManager.getSlideObj(1)
      const currentExtShape = slide1['p:sld']['p:cSld']['p:spTree']['p:sp'].find(
        s => s?.['p:nvSpPr']?.['p:cNvPr']?.['@_name'] === 'External_Col_Shape'
      )
      const currentExtX = parseInt(currentExtShape['p:spPr']['a:xfrm']['a:off']['@_x'], 10)
      expect(currentExtX).toBe(origExtX)
    })
  })

  // --------------------------------------------------------------------------
  // 8. removeTable Strict Containment
  // --------------------------------------------------------------------------
  describe('removeTable strict boundary containment', () => {
    it('purges table frame and strictly interior shapes, leaving touching and external shapes intact', () => {
      pptx.addTable(
        1,
        2,
        2,
        [
          [
            'A1',
            {
              shape: { type: 'circle', width: 20, height: 20 },
              text: 'T-Interior',
              textPosition: 'inside',
            },
          ],
          ['A2', 'B2'],
        ],
        { id: 'TableToDelete', x: 100, y: 100, width: 200, height: 100 }
      )

      const tblBounds = pptx.tableManager.getTableBounds('TableToDelete', 1, pptx.slideManager)
      const { xMin, yMin, xMax, yMax } = tblBounds

      // 1. Touches table left border -> PRESERVE
      injectShape(1, 'Table_TouchLeft', xMin, yMin + 2000, 10000, 10000)

      // 2. Touches table right border -> PRESERVE
      injectShape(1, 'Table_TouchRight', xMax - 10000, yMin + 2000, 10000, 10000)

      // 3. Crosses table outer border -> PRESERVE
      injectShape(1, 'Table_OverlapsBorder', xMin - 5000, yMin + 2000, 20000, 10000)

      // 4. Completely outside table -> PRESERVE
      injectShape(1, 'Table_External', xMax + 20000, yMax + 20000, 20000, 20000)

      // Remove table
      pptx.removeTable('TableToDelete', 1)

      // Table frame is purged
      expect(pptx.getTables(1).find(t => t.name === 'TableToDelete')).toBeUndefined()

      // Strictly interior shape is purged
      expect(pptx.getShapes(1).find(s => s.text === 'T-Interior')).toBeUndefined()

      // Edge-touching, overlapping, and external shapes remain intact
      expect(hasShape('Table_TouchLeft')).toBe(true)
      expect(hasShape('Table_TouchRight')).toBe(true)
      expect(hasShape('Table_OverlapsBorder')).toBe(true)
      expect(hasShape('Table_External')).toBe(true)
    })
  })

  // --------------------------------------------------------------------------
  // 9. Presentation Export Round-Trip
  // --------------------------------------------------------------------------
  describe('Round-Trip Presentation Export Integrity', () => {
    it('exports cleanly to buffer and re-imports without schema corruption', async () => {
      pptx.addTable(
        1,
        3,
        3,
        [
          ['R0C0', 'R0C1', 'R0C2'],
          ['R1C0', 'R1C1', 'R1C2'],
          ['R2C0', 'R2C1', 'R2C2'],
        ],
        { id: 'RoundTripTable', x: 100, y: 100, width: 300, height: 150 }
      )

      pptx.clearCellData('RoundTripTable', 0, 0, 1)
      pptx.clearRowData('RoundTripTable', 1, 1)
      pptx.clearColumnData('RoundTripTable', 2, 1)
      pptx.removeTableRow('RoundTripTable', 2, 1)

      const buffer = await pptx.toBuffer()
      expect(buffer).toBeInstanceOf(Buffer)
      expect(buffer.length).toBeGreaterThan(0)

      const reloaded = await PPTXTemplater.load(buffer)
      const tables = reloaded.getTables(1)
      expect(tables.length).toBe(1)
      expect(tables[0].name).toBe('RoundTripTable')
      expect(tables[0].rows).toBe(2)
      expect(tables[0].columns).toBe(3)
    })
  })
})
