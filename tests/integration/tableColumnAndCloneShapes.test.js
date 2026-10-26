/**
 * @fileoverview Integration tests for removeTableColumn shape purging and column deletion,
 * and cloneTable deep shape duplication with unique 32-bit IDs and coordinate translations.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('Table Column Removal & Clone Table Shapes Subsystem', () => {
  let pptx

  beforeEach(async () => {
    pptx = await PPTXTemplater.createBlank()
  })

  // Helper to inject an edge-touching shape directly into <p:spTree>
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
  // TASK-TBL-SHAPE-FIX-001: removeTableColumn Column Deletion & Shape Purging
  // --------------------------------------------------------------------------
  describe('TASK-TBL-SHAPE-FIX-001: removeTableColumn', () => {
    it('deletes the designated column (not rows) and purges strictly interior column shapes', () => {
      pptx.addTable(
        1,
        3,
        4,
        [
          [
            'Col0-Header',
            {
              shape: { type: 'roundRect', fill: '#3B82F6', width: 40, height: 20 },
              text: 'Col1-Badge',
              textPosition: 'inside',
            },
            {
              shape: { type: 'circle', fill: '#EF4444', width: 14, height: 14 },
              text: 'Col2-DeleteMe',
              textPosition: 'inside',
            },
            {
              shape: { type: 'diamond', fill: '#10B981', width: 16, height: 16 },
              text: 'Col3-ShiftLeft',
              textPosition: 'inside',
            },
          ],
          ['R1C0', 'R1C1', 'R1C2', 'R1C3'],
          ['R2C0', 'R2C1', 'R2C2', 'R2C3'],
        ],
        { id: 'ColumnFixTable', x: 100, y: 100, width: 400, height: 150 }
      )

      const tblBefore = pptx.getTable('ColumnFixTable', 1)
      expect(tblBefore.columns).toBe(4)
      expect(tblBefore.rows).toBe(3)

      const col2Bounds = pptx.tableManager.getTableColumnBounds(
        'ColumnFixTable',
        2,
        1,
        pptx.slideManager
      )
      expect(col2Bounds).toBeDefined()

      // Inject edge-touching shape on left border of Col 2 (x === col2Bounds.xMin) -> PRESERVE
      injectShape(1, 'EdgeTouch_Col2_Left', col2Bounds.xMin, col2Bounds.yMin + 2000, 10000, 10000)

      const col3Initial = pptx.getShapes(1).find(s => s.text === 'Col3-ShiftLeft')
      expect(col3Initial).toBeDefined()
      const origCol3X = col3Initial.x

      // Execute removeTableColumn for column index 2
      pptx.removeTableColumn('ColumnFixTable', 2, 1)

      const tblAfter = pptx.getTable('ColumnFixTable', 1)

      // 1. Invariant: Rows MUST NOT be deleted (rows stay 3)
      expect(tblAfter.rows).toBe(3)

      // 2. Column count decremented to 3
      expect(tblAfter.columns).toBe(3)

      // 3. Table data reflects column removal across all rows
      const dataAfter = pptx.getTableData('ColumnFixTable', 1)
      expect(dataAfter.length).toBe(3)
      expect(dataAfter[0]).toEqual(['Col0-Header', 'Col1-Badge', 'Col3-ShiftLeft'])
      expect(dataAfter[1]).toEqual(['R1C0', 'R1C1', 'R1C3'])
      expect(dataAfter[2]).toEqual(['R2C0', 'R2C1', 'R2C3'])

      const shapesAfter = pptx.getShapes(1)

      // 4. Target column shape 'Col2-DeleteMe' is completely purged
      expect(shapesAfter.find(s => s.text === 'Col2-DeleteMe')).toBeUndefined()

      // 5. Edge-touching shape is preserved
      expect(hasShape('EdgeTouch_Col2_Left')).toBe(true)

      // 6. Right column shape 'Col3-ShiftLeft' shifted leftward and renamed
      const shiftedCol3 = shapesAfter.find(s => s.text === 'Col3-ShiftLeft')
      expect(shiftedCol3).toBeDefined()
      expect(shiftedCol3.x).toBeLessThan(origCol3X)
      expect(shiftedCol3.name).toContain('CellShape_ColumnFixTable_r0_c2')

      // 7. Column 1 shape remains unaffected
      const col1Shape = shapesAfter.find(s => s.text === 'Col1-Badge')
      expect(col1Shape).toBeDefined()
      expect(col1Shape.name).toContain('CellShape_ColumnFixTable_r0_c1')
    })

    it('purges multiple shapes positioned across multiple rows in the targeted column', () => {
      pptx.addTable(
        1,
        3,
        3,
        [
          [
            'H0',
            {
              shape: { type: 'roundRect', fill: '#6366F1', width: 30, height: 15 },
              text: 'R0-Col1-Badge',
              textPosition: 'inside',
            },
            'H2',
          ],
          [
            'R1C0',
            {
              shape: { type: 'circle', fill: '#EC4899', width: 12, height: 12 },
              text: 'R1-Col1-Badge',
              textPosition: 'inside',
            },
            'R1C2',
          ],
          ['R2C0', 'R2C1', 'R2C2'],
        ],
        { id: 'MultiRowColTable', x: 50, y: 50, width: 300, height: 150 }
      )

      expect(pptx.getShapes(1).filter(s => s.text?.includes('Col1-Badge')).length).toBe(2)

      pptx.removeTableColumn('MultiRowColTable', 1, 1)

      const tbl = pptx.getTable('MultiRowColTable', 1)
      expect(tbl.rows).toBe(3)
      expect(tbl.columns).toBe(2)

      // Both shapes in Column 1 are purged
      expect(pptx.getShapes(1).filter(s => s.text?.includes('Col1-Badge')).length).toBe(0)
    })
  })

  // --------------------------------------------------------------------------
  // TASK-TBL-SHAPE-FIX-002: cloneTable Deep Shape Duplication
  // --------------------------------------------------------------------------
  describe('TASK-TBL-SHAPE-FIX-002: cloneTable Deep Shape Duplication', () => {
    it('deep-clones all associated cell shapes with unique 32-bit IDs, names, and spatial coordinate offsets', () => {
      pptx.addTable(
        1,
        3,
        3,
        [
          [
            'Header 0',
            {
              shape: { type: 'roundRect', fill: '#3B82F6', width: 50, height: 20 },
              text: 'R0C1-Badge',
              textPosition: 'inside',
            },
            'Header 2',
          ],
          [
            {
              shape: { type: 'circle', fill: '#10B981', width: 14, height: 14 },
              text: 'R1C0-Badge',
              textPosition: 'inside',
            },
            'R1C1',
            'R1C2',
          ],
          [
            'R2C0',
            'R2C1',
            {
              shape: { type: 'diamond', fill: '#F59E0B', width: 16, height: 16 },
              text: 'R2C2-Badge',
              textPosition: 'inside',
            },
          ],
        ],
        { id: 'SourceTableToClone', x: 50, y: 50, width: 300, height: 150 }
      )

      const origShapes = pptx
        .getShapes(1)
        .filter(s => s.name?.startsWith('CellShape_SourceTableToClone_'))
      expect(origShapes.length).toBe(3)

      // Clone table with +250 pt x-offset and +180 pt y-offset
      pptx.cloneTable('SourceTableToClone', 'ClonedTableTarget', { x: 300, y: 230 }, 1)

      const clonedTable = pptx.getTable('ClonedTableTarget', 1)
      expect(clonedTable).toBeDefined()
      expect(clonedTable.name).toBe('ClonedTableTarget')
      expect(clonedTable.rows).toBe(3)
      expect(clonedTable.columns).toBe(3)

      const clonedShapes = pptx
        .getShapes(1)
        .filter(s => s.name?.startsWith('CellShape_ClonedTableTarget_'))
      expect(clonedShapes.length).toBe(3)

      // 1. Unique ID invariant: Cloned shapes MUST NOT collide with source shape IDs
      const slideObj = pptx.slideManager.getSlideObj(1)
      const allAstShapes = slideObj['p:sld']['p:cSld']['p:spTree']['p:sp']
      const astClonedShapes = allAstShapes.filter(s =>
        s['p:nvSpPr']?.['p:cNvPr']?.['@_name']?.startsWith('CellShape_ClonedTableTarget_')
      )
      const astOrigShapes = allAstShapes.filter(s =>
        s['p:nvSpPr']?.['p:cNvPr']?.['@_name']?.startsWith('CellShape_SourceTableToClone_')
      )

      const origNumericIds = new Set(
        astOrigShapes.map(s => Number(s['p:nvSpPr']['p:cNvPr']['@_id']))
      )

      for (const cs of astClonedShapes) {
        const idNum = Number(cs['p:nvSpPr']['p:cNvPr']['@_id'])
        expect(origNumericIds.has(idNum)).toBe(false)
        expect(idNum).toBeGreaterThan(0)
        expect(idNum).toBeLessThan(2147483647)
      }

      // 2. Cloned shapes retain matching cell content
      const r0c1Cloned = clonedShapes.find(s => s.name?.includes('_r0_c1'))
      expect(r0c1Cloned).toBeDefined()
      expect(r0c1Cloned.text).toBe('R0C1-Badge')

      const r1c0Cloned = clonedShapes.find(s => s.name?.includes('_r1_c0'))
      expect(r1c0Cloned).toBeDefined()
      expect(r1c0Cloned.text).toBe('R1C0-Badge')

      const r2c2Cloned = clonedShapes.find(s => s.name?.includes('_r2_c2'))
      expect(r2c2Cloned).toBeDefined()
      expect(r2c2Cloned.text).toBe('R2C2-Badge')

      // 3. Spatial offset invariant: Cloned shape coordinates must be shifted
      const r0c1Orig = origShapes.find(s => s.name?.includes('_r0_c1'))
      expect(r0c1Cloned.x).toBeGreaterThan(r0c1Orig.x)
      expect(r0c1Cloned.y).toBeGreaterThan(r0c1Orig.y)
    })

    it('clones table and shapes without explicit coordinate offsets (in-place cloning)', () => {
      pptx.addTable(
        1,
        2,
        2,
        [
          [
            'A',
            {
              shape: { type: 'pill', fill: '#8B5CF6', width: 40, height: 18 },
              text: 'InPlaceBadge',
              textPosition: 'inside',
            },
          ],
          ['C', 'D'],
        ],
        { id: 'InPlaceSourceTable', x: 80, y: 80, width: 200, height: 100 }
      )

      pptx.cloneTable('InPlaceSourceTable', 'InPlaceClonedTable', {}, 1)

      const clonedShapes = pptx
        .getShapes(1)
        .filter(s => s.name?.startsWith('CellShape_InPlaceClonedTable_'))
      expect(clonedShapes.length).toBe(1)
      expect(clonedShapes[0].text).toBe('InPlaceBadge')
      expect(clonedShapes[0].name).toContain('CellShape_InPlaceClonedTable_r0_c1')
    })
  })

  // --------------------------------------------------------------------------
  // Round-Trip Presentation Export & Reloading Integrity
  // --------------------------------------------------------------------------
  describe('Round-Trip Presentation Integrity', () => {
    it('exports presentation with removed column and cloned table shapes cleanly to buffer', async () => {
      pptx.addTable(
        1,
        3,
        3,
        [
          ['H0', 'H1', 'H2'],
          [
            'A1',
            {
              shape: { type: 'roundRect', fill: '#10B981', width: 40, height: 16 },
              text: 'Active',
              textPosition: 'inside',
            },
            'C1',
          ],
          ['A2', 'B2', 'C2'],
        ],
        { id: 'RoundTripTable', x: 60, y: 60, width: 300, height: 120 }
      )

      pptx.cloneTable('RoundTripTable', 'RoundTripCloned', { x: 60, y: 220 }, 1)
      pptx.removeTableColumn('RoundTripTable', 0, 1)

      const buffer = await pptx.toBuffer()
      expect(buffer).toBeInstanceOf(Buffer)
      expect(buffer.length).toBeGreaterThan(0)

      const reloaded = await PPTXTemplater.load(buffer)
      const tables = reloaded.getTables(1)
      expect(tables.length).toBe(2)

      const origTable = tables.find(t => t.name === 'RoundTripTable')
      expect(origTable.columns).toBe(2)
      expect(origTable.rows).toBe(3)

      const clonedTable = tables.find(t => t.name === 'RoundTripCloned')
      expect(clonedTable.columns).toBe(3)
      expect(clonedTable.rows).toBe(3)
    })
  })
})
