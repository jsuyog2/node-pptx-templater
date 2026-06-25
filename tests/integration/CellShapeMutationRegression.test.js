/**
 * Regression tests for cell shape positioning after table structural mutations.
 *
 * Each test verifies that shapes placed via addTableRow() / addCellShape() remain
 * correctly positioned (x, y within cell bounds) after subsequent operations such as
 * removeTableRow(), mergeCells(), and insertTableRow().
 *
 * These tests guard against the bug where shapes were placed with stale absolute
 * coordinates that were not updated when the table layout changed.
 */
import { describe, it, expect } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')

/**
 * Helper: checks that a shape returned by getCellShape() has its entire body
 * inside the cell bounds returned by getCellBounds().
 */
function shapeIsInsideCell(shape, bounds) {
  if (!shape || !bounds) return false
  // Allow 2px tolerance for rounding
  const tol = 2
  return (
    shape.x >= bounds.x - tol &&
    shape.y >= bounds.y - tol &&
    shape.x + shape.width <= bounds.x + bounds.width + tol &&
    shape.y + shape.height <= bounds.y + bounds.height + tol
  )
}

/**
 * Helper: checks that a shape is centred inside the given bounds.
 * Returns true when the shape centre is within `tol` pixels of the cell centre.
 */
function shapeIsCentred(shape, bounds, tol = 3) {
  if (!shape || !bounds) return false
  const shapeCX = shape.x + shape.width / 2
  const shapeCY = shape.y + shape.height / 2
  const cellCX = bounds.x + bounds.width / 2
  const cellCY = bounds.y + bounds.height / 2
  return Math.abs(shapeCX - cellCX) <= tol && Math.abs(shapeCY - cellCY) <= tol
}

describe('Cell Shape Mutation Regression Tests', () => {
  const runTests = existsSync(FIXTURE_FILE)

  if (!runTests) {
    it.skip('Skipping: sample.pptx fixture not found', () => {})
    return
  }

  // ─── Test 1: addTableRow + mergeCells ────────────────────────────────────────
  it('addTableRow() then mergeCells() – shape stays inside the correct cell', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    // Add two rows, each with a shape in col 4
    ppt.addTableRow('Table', [
      'Alice',
      '',
      'Designer',
      'Value',
      { type: 'circle', color: '#10B981', width: 12, height: 12, position: 'center' },
    ])
    ppt.addTableRow('Table', [
      'Bob',
      '',
      'Engineer',
      'Value',
      { type: 'circle', color: '#3B82F6', width: 12, height: 12, position: 'center' },
    ])

    // Get data rows length to compute XML row indices (including header at 0)
    const rows = ppt.getTableRows('Table', { raw: true })
    const r1 = rows.length - 1 // 2 (first added row index)
    const r2 = rows.length // 3 (second added row index)

    // Merge col 0 across the two new rows
    ppt.mergeCells('Table', r1, 0, r2, 0)

    // Both shapes must still be inside their respective cells
    const bounds1 = ppt.getCellBounds('Table', r1, 4)
    const bounds2 = ppt.getCellBounds('Table', r2, 4)
    const shape1 = ppt.getCellShape('Table', r1, 4, 0)
    const shape2 = ppt.getCellShape('Table', r2, 4, 0)

    expect(shape1).not.toBeNull()
    expect(shape2).not.toBeNull()
    expect(shapeIsInsideCell(shape1, bounds1)).toBe(true)
    expect(shapeIsInsideCell(shape2, bounds2)).toBe(true)

    const validation = await ppt.validatePresentation()
    expect(validation.valid).toBe(true)
  })

  // ─── Test 2: removeTableRow + mergeCells ─────────────────────────────────────
  it('removeTableRow() then mergeCells() – shapes have correct y-coord after row removal', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    // Add three rows, each with a circle in the last column
    ppt.addTableRow('Table', [
      'Bob',
      '',
      'The implementation should allow users to build highly visual dashboards.',
      'Value',
      { type: 'circle', color: '#af4c4c', width: 12, height: 12, position: 'center' },
    ])
    ppt.addTableRow('Table', [
      '',
      '',
      'Designer',
      'Value',
      { type: 'circle', color: '#0021db', width: 12, height: 12, position: 'center' },
    ])
    ppt.addTableRow('Table', [
      'Bob',
      '',
      'Designer',
      'Value',
      { type: 'circle', color: '#4CAF50', width: 12, height: 12, position: 'center' },
    ])

    // Remove the original data row (row index 1) — appended rows shift up
    ppt.removeTableRow('Table', 1)

    const allRows = ppt.getTableRows('Table', { raw: true })
    const r0 = allRows.length - 2 // 1 (first shifted row index)
    const r1 = allRows.length - 1 // 2 (second shifted row index)
    const r2 = allRows.length // 3 (third shifted row index)

    // Merge col 0 and col 1 across the first two appended rows
    ppt.mergeCells('Table', r0, 0, r1, 0)
    ppt.mergeCells('Table', r0, 1, r1, 1)

    // Each circle must still be inside its logical cell in the last column
    for (const rowIdx of [r0, r1, r2]) {
      const bounds = ppt.getCellBounds('Table', rowIdx, 4)
      const shape = ppt.getCellShape('Table', rowIdx, 4, 0)
      expect(shape, `Shape missing at row ${rowIdx}`).not.toBeNull()
      expect(
        shapeIsInsideCell(shape, bounds),
        `Shape at row ${rowIdx} outside cell bounds\n  shape=${JSON.stringify(shape)}\n  bounds=${JSON.stringify(bounds)}`
      ).toBe(true)
    }

    const validation = await ppt.validatePresentation()
    expect(validation.valid).toBe(true)
  })

  // ─── Test 3: Wrapped text causing variable row heights ────────────────────────
  it('shape stays inside its cell when row height grows due to wrapped text', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    ppt.addTableRow('Table', [
      'Alice',
      '',
      'This is a very long description that will wrap across multiple lines in the cell and force the row to grow taller than default.',
      'Value',
      { type: 'circle', color: '#EF4444', width: 12, height: 12, position: 'center' },
    ])

    // Remove the original data row so our row shifts up
    ppt.removeTableRow('Table', 1)

    const newRows = ppt.getTableRows('Table', { raw: true })
    const newLastRow = newRows.length // 1 (the shifted added row)

    const bounds = ppt.getCellBounds('Table', newLastRow, 4)
    const shape = ppt.getCellShape('Table', newLastRow, 4, 0)

    expect(shape).not.toBeNull()
    expect(shapeIsInsideCell(shape, bounds)).toBe(true)

    const validation = await ppt.validatePresentation()
    expect(validation.valid).toBe(true)
  })

  // ─── Test 4: Multiple consecutive merges ─────────────────────────────────────
  it('multiple consecutive mergeCells() calls leave all shapes correctly positioned', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    ppt.addTableRow('Table', [
      'Alice',
      '',
      'Engineer',
      'V1',
      { type: 'circle', color: '#F59E0B', width: 10, height: 10, position: 'center' },
    ])
    ppt.addTableRow('Table', [
      'Alice',
      '',
      'Designer',
      'V2',
      { type: 'circle', color: '#8B5CF6', width: 10, height: 10, position: 'center' },
    ])
    ppt.addTableRow('Table', [
      'Alice',
      '',
      'Manager',
      'V3',
      { type: 'circle', color: '#EC4899', width: 10, height: 10, position: 'center' },
    ])

    const rows = ppt.getTableRows('Table', { raw: true })
    const base = rows.length - 2 // 2 (index of first added row)

    // First merge: col 0 across all three new rows
    ppt.mergeCells('Table', base, 0, base + 2, 0)
    // Second merge: col 1 across first two new rows
    ppt.mergeCells('Table', base, 1, base + 1, 1)

    // All three shapes in col 4 must still be inside their cells
    for (let i = base; i <= base + 2; i++) {
      const bounds = ppt.getCellBounds('Table', i, 4)
      const shape = ppt.getCellShape('Table', i, 4, 0)
      expect(shape, `Shape missing at row ${i}`).not.toBeNull()
      expect(shapeIsInsideCell(shape, bounds)).toBe(true)
    }

    const validation = await ppt.validatePresentation()
    expect(validation.valid).toBe(true)
  })

  // ─── Test 5: Multiple shapes in the same column ───────────────────────────────
  it('multiple shapes in the same column each stay in their own cell after row removal', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    const colors = ['#EF4444', '#10B981', '#3B82F6']
    for (const color of colors) {
      ppt.addTableRow('Table', [
        'X',
        '',
        'Role',
        'V',
        { type: 'circle', color, width: 10, height: 10, position: 'center' },
      ])
    }

    // Remove the original data row — all three appended rows shift
    ppt.removeTableRow('Table', 1)

    const allRows = ppt.getTableRows('Table', { raw: true })
    const base = allRows.length - 2 // 1 (index of first shifted added row)

    for (let i = 0; i < 3; i++) {
      const rowIdx = base + i
      const bounds = ppt.getCellBounds('Table', rowIdx, 4)
      const shape = ppt.getCellShape('Table', rowIdx, 4, 0)
      expect(shape, `Shape missing at row ${rowIdx}`).not.toBeNull()
      expect(shapeIsInsideCell(shape, bounds), `Shape at row ${rowIdx} outside bounds`).toBe(true)
    }

    const validation = await ppt.validatePresentation()
    expect(validation.valid).toBe(true)
  })

  // ─── Test 6: Center alignment in merged cells ─────────────────────────────────
  it('center-aligned shape is centred within the merged cell bounding box', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    ppt.addTableRow('Table', [
      'Alice',
      '',
      'Engineer',
      'V1',
      { type: 'circle', color: '#14B8A6', width: 12, height: 12, position: 'center' },
    ])
    ppt.addTableRow('Table', [
      'Alice',
      '',
      'Designer',
      'V2',
      { type: 'circle', color: '#F97316', width: 12, height: 12, position: 'center' },
    ])

    const rows = ppt.getTableRows('Table', { raw: true })
    const base = rows.length - 1 // 2 (index of first added row)

    // Merge the last column across both rows — the shape must now be centred in the merged area
    ppt.mergeCells('Table', base, 4, base + 1, 4)

    const mergedBounds = ppt.getCellBounds('Table', base, 4)
    const shape = ppt.getCellShape('Table', base, 4, 0)

    expect(shape).not.toBeNull()
    expect(shapeIsInsideCell(shape, mergedBounds)).toBe(true)
    expect(shapeIsCentred(shape, mergedBounds)).toBe(true)

    const validation = await ppt.validatePresentation()
    expect(validation.valid).toBe(true)
  })

  // ─── Test 7: Non-center alignment after row removal ───────────────────────────
  it('left-aligned and right-aligned shapes remain inside cell bounds after removeTableRow()', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    ppt.addTableRow('Table', [
      'A',
      '',
      { type: 'rectangle', color: '#EF4444', width: 20, height: 8, position: 'left' },
      'V',
      { type: 'circle', color: '#10B981', width: 10, height: 10, position: 'right' },
    ])

    // Remove the row above to shift our appended row
    ppt.removeTableRow('Table', 1)

    const newRows = ppt.getTableRows('Table', { raw: true })
    const newLastRow = newRows.length // 1 (index of the shifted added row)

    // Check col 2 (left-aligned rectangle)
    const bounds2 = ppt.getCellBounds('Table', newLastRow, 2)
    const shape2 = ppt.getCellShape('Table', newLastRow, 2, 0)
    expect(shape2).not.toBeNull()
    expect(shapeIsInsideCell(shape2, bounds2)).toBe(true)

    // Check col 4 (right-aligned circle)
    const bounds4 = ppt.getCellBounds('Table', newLastRow, 4)
    const shape4 = ppt.getCellShape('Table', newLastRow, 4, 0)
    expect(shape4).not.toBeNull()
    expect(shapeIsInsideCell(shape4, bounds4)).toBe(true)

    const validation = await ppt.validatePresentation()
    expect(validation.valid).toBe(true)
  })

  // ─── Test 8: Full bug-report sequence ────────────────────────────────────────
  it('exact bug-report sequence: addTableRow×3 + removeTableRow + mergeCells×2 – all circles centred in STATUS cells', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    // Reproduce the exact sequence from the bug report
    ppt.addTableRow('Table', [
      'Bob',
      '',
      'The final implementation should allow users to build highly visual dashboards and reports.',
      'Value',
      { type: 'circle', color: '#af4c4cff', width: 12, height: 12, position: 'center' },
    ])
    ppt.addTableRow('Table', [
      '',
      '',
      'Designer',
      'Value',
      { type: 'circle', color: '#0021dbff', width: 12, height: 12, position: 'center' },
    ])
    ppt.addTableRow('Table', [
      'Bob',
      '',
      'Designer',
      'Value',
      { type: 'circle', color: '#4CAF50', width: 12, height: 12, position: 'center' },
    ])

    ppt.removeTableRow('Table', 1)

    const allRows = ppt.getTableRows('Table', { raw: true })
    const r0 = allRows.length - 2 // 1
    const r1 = allRows.length - 1 // 2
    const r2 = allRows.length // 3

    ppt.mergeCells('Table', r0, 0, r1, 0)
    ppt.mergeCells('Table', r0, 1, r1, 1)

    // Every circle must be inside its STATUS cell (col 4) and be centred
    for (const rowIdx of [r0, r1, r2]) {
      const bounds = ppt.getCellBounds('Table', rowIdx, 4)
      const shape = ppt.getCellShape('Table', rowIdx, 4, 0)

      expect(shape, `Circle missing at row ${rowIdx} after full mutation sequence`).not.toBeNull()

      expect(
        shapeIsInsideCell(shape, bounds),
        `Circle at row ${rowIdx} outside its cell bounds.\n` +
          `  shape: ${JSON.stringify(shape)}\n  bounds: ${JSON.stringify(bounds)}`
      ).toBe(true)

      expect(
        shapeIsCentred(shape, bounds),
        `Circle at row ${rowIdx} is not centred.\n` +
          `  shape: ${JSON.stringify(shape)}\n  bounds: ${JSON.stringify(bounds)}`
      ).toBe(true)
    }

    // No PowerPoint repair dialogs
    const validation = await ppt.validatePresentation()
    expect(validation.valid).toBe(true)
  })
})
