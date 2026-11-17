/**
 * @fileoverview Integration tests for multi-step / sequential row and column deletions
 * with strict geometric bounding box shape handling and coordinate shifting.
 *
 * TASK-TBL-LIFECYCLE-TEST-004:
 * 1. Multi-step row removal (Slide 8 scenario): 3 data rows + 1 header, shapes in row 1 and row 2.
 *    Removing row 1 and then row 2 must preserve row 2's shifted shape and purge row 1's shape.
 * 2. Multi-step column removal (Slide 9 scenario): 5 columns, shapes in column 4.
 *    Removing column 1 shifts column 4 shapes into column 3; removing column 3 purges them completely.
 * 3. 4-column sequential removal with post-deletion index shifting.
 * 4. Presentation export and round-trip verification with zero repair warnings.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('Sequential Row and Column Shape Removal & Coordinate Shifting', () => {
  let pptx

  beforeEach(async () => {
    pptx = await PPTXTemplater.createBlank()
  })

  // Helper to extract raw EMU coordinates from AST
  function getRawShapeCoords(slideIndex, shapeName) {
    const slideObj = pptx.slideManager.getSlideObj(slideIndex)
    const spTree = slideObj['p:sld']['p:cSld']['p:spTree']
    const shapes = Array.isArray(spTree['p:sp'])
      ? spTree['p:sp']
      : spTree['p:sp']
        ? [spTree['p:sp']]
        : []
    const node = shapes.find(s => s?.['p:nvSpPr']?.['p:cNvPr']?.['@_name'] === shapeName)
    if (!node) return null
    const xfrm = node?.['p:spPr']?.['a:xfrm']
    return {
      x: parseInt(xfrm?.['a:off']?.['@_x'] || '0', 10),
      y: parseInt(xfrm?.['a:off']?.['@_y'] || '0', 10),
      cx: parseInt(xfrm?.['a:ext']?.['@_cx'] || '0', 10),
      cy: parseInt(xfrm?.['a:ext']?.['@_cy'] || '0', 10),
    }
  }

  // --------------------------------------------------------------------------
  // Scenario 1: Slide 8 Replication — Sequential Row Removal
  // --------------------------------------------------------------------------
  describe('TASK-TBL-ROW-SHAPE-FIX-002: Sequential Row Removal (Slide 8 Scenario)', () => {
    it('preserves row 2 shape with shifted Y coordinates when removing row 1 and then row 2', () => {
      // Create table with 4 rows (1 header + 3 data rows), shapes in row 1 and row 2
      pptx.addTable(
        1,
        4,
        5,
        [
          ['Department', '', 'Q1 Target', 'Q1 Actual', ''],
          [
            'Global Marketing',
            '',
            '$500,000',
            '$520,000',
            {
              shape: { type: 'rect', fill: '#3B82F6', width: 36, height: 25 },
              text: 'R1-Shape',
              textPosition: 'inside',
            },
          ],
          [
            'Engineering',
            'Product',
            '$250,000',
            '$245,000',
            {
              shape: { type: 'ellipse', fill: '#10B981', width: 27, height: 27 },
              text: 'R2-Shape',
              textPosition: 'inside',
            },
          ],
          ['Engineering2', 'Product1', '$450,000', '$145,000', ''],
        ],
        { id: 'Slide8Table', x: 88, y: 108, width: 1104, height: 156 }
      )

      const tm = pptx.tableManager
      const sm = pptx.slideManager

      const tblBefore = pptx.getTable('Slide8Table', 1)
      expect(tblBefore.rows).toBe(4)

      const row1BoundsBefore = tm.getTableRowBounds('Slide8Table', 1, 1, sm)
      const row2BoundsBefore = tm.getTableRowBounds('Slide8Table', 2, 1, sm)
      const row1Height = row1BoundsBefore.height

      const r1ShapeBefore = pptx.getShapes(1).find(s => s.text === 'R1-Shape')
      const r2ShapeBefore = pptx.getShapes(1).find(s => s.text === 'R2-Shape')
      expect(r1ShapeBefore).toBeDefined()
      expect(r2ShapeBefore).toBeDefined()

      const r2CoordsBefore = getRawShapeCoords(1, r2ShapeBefore.name)
      expect(r2CoordsBefore.y).toBeGreaterThanOrEqual(row2BoundsBefore.yMin)
      expect(r2CoordsBefore.y + r2CoordsBefore.cy).toBeLessThanOrEqual(row2BoundsBefore.yMax)

      // Step 1: Remove row index 1 (Global Marketing)
      pptx.removeTableRow('Slide8Table', 1, 1)

      const tblAfterStep1 = pptx.getTable('Slide8Table', 1)
      expect(tblAfterStep1.rows).toBe(3)

      const shapesAfterStep1 = pptx.getShapes(1)
      // R1-Shape must be purged
      expect(shapesAfterStep1.find(s => s.text === 'R1-Shape')).toBeUndefined()

      // R2-Shape must be preserved and shifted upward by exactly row1Height
      const r2ShapeAfterStep1 = shapesAfterStep1.find(s => s.text === 'R2-Shape')
      expect(r2ShapeAfterStep1).toBeDefined()
      const r2CoordsAfterStep1 = getRawShapeCoords(1, r2ShapeAfterStep1.name)
      expect(r2CoordsAfterStep1.y).toBe(r2CoordsBefore.y - row1Height)
      expect(r2CoordsAfterStep1.x).toBe(r2CoordsBefore.x)

      // The new Row 1 bounds should now contain R2-Shape
      const row1BoundsAfterStep1 = tm.getTableRowBounds('Slide8Table', 1, 1, sm)
      expect(r2CoordsAfterStep1.y).toBeGreaterThanOrEqual(row1BoundsAfterStep1.yMin)
      expect(r2CoordsAfterStep1.y + r2CoordsAfterStep1.cy).toBeLessThanOrEqual(
        row1BoundsAfterStep1.yMax
      )

      // Step 2: Remove row index 2 (which is now Engineering2, original row 3)
      pptx.removeTableRow('Slide8Table', 2, 1)

      const tblAfterStep2 = pptx.getTable('Slide8Table', 1)
      expect(tblAfterStep2.rows).toBe(2)

      const shapesAfterStep2 = pptx.getShapes(1)
      // R2-Shape MUST STILL BE PRESERVED (NOT accidentally purged by the second deletion)
      const r2ShapeAfterStep2 = shapesAfterStep2.find(s => s.text === 'R2-Shape')
      expect(r2ShapeAfterStep2).toBeDefined()
      const r2CoordsAfterStep2 = getRawShapeCoords(1, r2ShapeAfterStep2.name)
      expect(r2CoordsAfterStep2.y).toBe(r2CoordsBefore.y - row1Height)

      // Data check: remaining rows are Header and Engineering
      const dataAfter = pptx.getTableData('Slide8Table', 1)
      expect(dataAfter.length).toBe(2)
      expect(dataAfter[0][0]).toBe('Department')
      expect(dataAfter[1][0]).toBe('Engineering')
    })

    it('works seamlessly with arbitrary custom shape names in template rows', () => {
      // Create table
      pptx.addTable(
        1,
        3,
        2,
        [
          ['H0', 'H1'],
          ['Row1', 'Val1'],
          ['Row2', 'Val2'],
        ],
        { id: 'CustomShapeTable', x: 100, y: 100, width: 400, height: 120 }
      )

      const tm = pptx.tableManager
      const sm = pptx.slideManager
      const r1Bounds = tm.getTableRowBounds('CustomShapeTable', 1, 1, sm)
      const r2Bounds = tm.getTableRowBounds('CustomShapeTable', 2, 1, sm)

      // Inject custom named shapes directly (like PowerPoint does)
      const slideObj = sm.getSlideObj(1)
      const spTree = slideObj['p:sld']['p:cSld']['p:spTree']
      if (!Array.isArray(spTree['p:sp'])) spTree['p:sp'] = [spTree['p:sp']].filter(Boolean)

      spTree['p:sp'].push({
        'p:nvSpPr': {
          'p:cNvPr': { '@_id': '101', '@_name': 'Rectangle 4' },
          'p:cNvSpPr': {},
          'p:nvPr': {},
        },
        'p:spPr': {
          'a:xfrm': {
            'a:off': { '@_x': String(r1Bounds.xMin + 5000), '@_y': String(r1Bounds.yMin + 5000) },
            'a:ext': { '@_cx': '20000', '@_cy': '20000' },
          },
          'a:prstGeom': { '@_prst': 'rect', 'a:avLst': {} },
        },
      })

      spTree['p:sp'].push({
        'p:nvSpPr': {
          'p:cNvPr': { '@_id': '102', '@_name': 'Oval 6' },
          'p:cNvSpPr': {},
          'p:nvPr': {},
        },
        'p:spPr': {
          'a:xfrm': {
            'a:off': { '@_x': String(r2Bounds.xMin + 5000), '@_y': String(r2Bounds.yMin + 5000) },
            'a:ext': { '@_cx': '20000', '@_cy': '20000' },
          },
          'a:prstGeom': { '@_prst': 'ellipse', 'a:avLst': {} },
        },
      })
      sm.markSlideObjDirty(1)

      const origOvalY = r2Bounds.yMin + 5000

      // Remove row 1
      pptx.removeTableRow('CustomShapeTable', 1, 1)

      // Rectangle 4 must be purged
      expect(pptx.getShapes(1).find(s => s.name === 'Rectangle 4')).toBeUndefined()

      // Oval 6 must be preserved and shifted upward
      const ovalAfter = pptx.getShapes(1).find(s => s.name === 'Oval 6')
      expect(ovalAfter).toBeDefined()
      const ovalCoords = getRawShapeCoords(1, 'Oval 6')
      expect(ovalCoords.y).toBe(origOvalY - r1Bounds.height)
    })
  })

  // --------------------------------------------------------------------------
  // Scenario 2: Slide 9 Replication — Sequential Column Removal
  // --------------------------------------------------------------------------
  describe('TASK-TBL-COL-SHAPE-FIX-003: Sequential Column Removal (Slide 9 Scenario)', () => {
    it('shifts column 4 shapes to column 3 upon removing column 1, then purges them when column 3 is removed', () => {
      // 5 columns (Col 0: Department, Col 1: '', Col 2: Q1 Target, Col 3: Q1 Actual, Col 4: Status)
      pptx.addTable(
        1,
        3,
        5,
        [
          ['Department', '', 'Q1 Target', 'Q1 Actual', 'Status'],
          [
            'Global Marketing',
            '',
            '$500,000',
            '$520,000',
            {
              shape: { type: 'rect', fill: '#3B82F6', width: 36, height: 25 },
              text: 'Col4-TopShape',
              textPosition: 'inside',
            },
          ],
          [
            'Engineering',
            'Product',
            '$250,000',
            '$245,000',
            {
              shape: { type: 'ellipse', fill: '#10B981', width: 27, height: 27 },
              text: 'Col4-BottomShape',
              textPosition: 'inside',
            },
          ],
        ],
        { id: 'Slide9Table', x: 88, y: 108, width: 1104, height: 156 }
      )

      const tm = pptx.tableManager
      const sm = pptx.slideManager

      const col1Bounds = tm.getTableColumnBounds('Slide9Table', 1, 1, sm)
      const col1Width = col1Bounds.width

      const topShapeBefore = pptx.getShapes(1).find(s => s.text === 'Col4-TopShape')
      const bottomShapeBefore = pptx.getShapes(1).find(s => s.text === 'Col4-BottomShape')
      expect(topShapeBefore).toBeDefined()
      expect(bottomShapeBefore).toBeDefined()

      const topCoordsBefore = getRawShapeCoords(1, topShapeBefore.name)
      const bottomCoordsBefore = getRawShapeCoords(1, bottomShapeBefore.name)

      // Step 1: Remove column index 1
      pptx.removeTableColumn('Slide9Table', 1, 1)

      const tblAfterCol1 = pptx.getTable('Slide9Table', 1)
      expect(tblAfterCol1.columns).toBe(4)

      // Both shapes must shift left by exactly col1Width
      const topShapeAfterStep1 = pptx.getShapes(1).find(s => s.text === 'Col4-TopShape')
      const bottomShapeAfterStep1 = pptx.getShapes(1).find(s => s.text === 'Col4-BottomShape')
      expect(topShapeAfterStep1).toBeDefined()
      expect(bottomShapeAfterStep1).toBeDefined()

      const topCoordsAfterStep1 = getRawShapeCoords(1, topShapeAfterStep1.name)
      const bottomCoordsAfterStep1 = getRawShapeCoords(1, bottomShapeAfterStep1.name)
      expect(topCoordsAfterStep1.x).toBe(topCoordsBefore.x - col1Width)
      expect(bottomCoordsAfterStep1.x).toBe(bottomCoordsBefore.x - col1Width)

      // The shifted shapes must now fall strictly inside the new Column 3 bounds
      const col3BoundsAfterStep1 = tm.getTableColumnBounds('Slide9Table', 3, 1, sm)
      expect(topCoordsAfterStep1.x).toBeGreaterThanOrEqual(col3BoundsAfterStep1.xMin)
      expect(topCoordsAfterStep1.x + topCoordsAfterStep1.cx).toBeLessThanOrEqual(
        col3BoundsAfterStep1.xMax
      )

      // Step 2: Remove column index 3 (which now holds the Status column and shapes)
      pptx.removeTableColumn('Slide9Table', 3, 1)

      const tblAfterCol3 = pptx.getTable('Slide9Table', 1)
      expect(tblAfterCol3.columns).toBe(3)

      // BOTH shapes must now be completely purged with zero orphans left behind!
      const shapesAfterStep2 = pptx.getShapes(1)
      expect(shapesAfterStep2.find(s => s.text === 'Col4-TopShape')).toBeUndefined()
      expect(shapesAfterStep2.find(s => s.text === 'Col4-BottomShape')).toBeUndefined()

      // Table data reflects the removals
      const dataAfter = pptx.getTableData('Slide9Table', 1)
      expect(dataAfter[0]).toEqual(['Department', 'Q1 Target', 'Q1 Actual'])
      expect(dataAfter[1]).toEqual(['Global Marketing', '$500,000', '$520,000'])
    })

    it('handles 4-column sequential removal with accurate post-deletion index shifting', () => {
      // 4 columns: Col 0 (Header), Col 1 (Del 1), Col 2 (Keep), Col 3 (Del 2)
      pptx.addTable(
        1,
        2,
        4,
        [
          [
            'Col0',
            { shape: { type: 'rect' }, text: 'Col1-Badge', textPosition: 'inside' },
            'Col2-Keep',
            { shape: { type: 'circle' }, text: 'Col3-Badge', textPosition: 'inside' },
          ],
          ['A0', 'A1', 'A2', 'A3'],
        ],
        { id: 'FourColTable', x: 50, y: 50, width: 400, height: 100 }
      )

      // Step 1: Remove column index 1
      pptx.removeTableColumn('FourColTable', 1, 1)

      expect(pptx.getTable('FourColTable', 1).columns).toBe(3)
      expect(pptx.getShapes(1).find(s => s.text === 'Col1-Badge')).toBeUndefined()

      // Col3 has shifted to index 2
      const shiftedCol3 = pptx.getShapes(1).find(s => s.text === 'Col3-Badge')
      expect(shiftedCol3).toBeDefined()

      // Step 2: Remove the shifted column (now at index 2)
      pptx.removeTableColumn('FourColTable', 2, 1)

      expect(pptx.getTable('FourColTable', 1).columns).toBe(2)
      expect(pptx.getShapes(1).find(s => s.text === 'Col3-Badge')).toBeUndefined()

      const dataFinal = pptx.getTableData('FourColTable', 1)
      expect(dataFinal[0]).toEqual(['Col0', 'Col2-Keep'])
      expect(dataFinal[1]).toEqual(['A0', 'A2'])
    })
  })

  // --------------------------------------------------------------------------
  // Scenario 3: Round-Trip Presentation Export Integrity
  // --------------------------------------------------------------------------
  describe('Round-Trip Presentation Export Integrity', () => {
    it('exports presentation cleanly with zero OpenXML validation errors', async () => {
      pptx.addTable(
        1,
        4,
        4,
        [
          ['H0', 'H1', { shape: { type: 'rect' }, text: 'H2-Shape', textPosition: 'inside' }, 'H3'],
          [
            'R1C0',
            { shape: { type: 'circle' }, text: 'R1C1-Shape', textPosition: 'inside' },
            'R1C2',
            'R1C3',
          ],
          ['R2C0', 'R2C1', 'R2C2', 'R2C3'],
          ['R3C0', 'R3C1', 'R3C2', 'R3C3'],
        ],
        { id: 'ExportTable', x: 60, y: 60, width: 400, height: 160 }
      )

      // Remove row 1 and column 1
      pptx.removeTableRow('ExportTable', 1, 1)
      pptx.removeTableColumn('ExportTable', 1, 1)

      const buffer = await pptx.toBuffer()
      expect(buffer).toBeInstanceOf(Buffer)
      expect(buffer.length).toBeGreaterThan(0)

      const reloaded = await PPTXTemplater.load(buffer)
      const validation = reloaded.validate()
      expect(validation.valid).toBe(true)
      expect(validation.errors).toHaveLength(0)

      const tbl = reloaded.getTable('ExportTable', 1)
      expect(tbl.rows).toBe(3)
      expect(tbl.columns).toBe(3)
    })
  })
})
