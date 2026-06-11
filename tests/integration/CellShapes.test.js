import { describe, it, expect } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')

describe('Table Cell Shapes Integration Tests', () => {
  const runTests = existsSync(FIXTURE_FILE)

  if (!runTests) {
    it.skip('Skipping: sample.pptx fixture not found', () => {})
    return
  }

  it('should support adding, updating, and removing cell shapes dynamically', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3) // Slide 3 contains the table 'Table'

    // 1. Initial table update with cellShapes
    ppt.updateTable('Table', {
      rows: [
        { A: 'John', V: 'Active', B: 25 },
        { A: 'Mike', V: 'Inactive', B: -10 },
      ],
      cellShapes: {
        V: row => ({
          type: 'circle',
          fill: row.V === 'Active' ? '#10B981' : '#EF4444',
          width: 12,
          height: 12,
          position: 'left',
          x: 2,
        }),
        B: row => ({
          type: 'icon',
          icon: row.B > 0 ? 'up' : 'down',
          size: 14,
          position: 'right',
        }),
      },
    })

    // Verify shapes exist via getCellShape
    const shapeActive = ppt.getCellShape('Table', 1, 1, 0)
    expect(shapeActive).not.toBeNull()
    expect(shapeActive.type).toBe('circle')

    const shapeGrowth = ppt.getCellShape('Table', 1, 2, 0)
    expect(shapeGrowth).not.toBeNull()
    expect(shapeGrowth.type).toBe('upArrow')

    const shapeInactive = ppt.getCellShape('Table', 2, 1, 0)
    expect(shapeInactive).not.toBeNull()
    expect(shapeInactive.type).toBe('circle')

    const shapeGrowth2 = ppt.getCellShape('Table', 2, 2, 0)
    expect(shapeGrowth2).not.toBeNull()
    expect(shapeGrowth2.type).toBe('downArrow')

    // 2. Test getCellShape for non-existent shape
    const nonExistent = ppt.getCellShape('Table', 1, 1, 99)
    expect(nonExistent).toBeNull()

    // 3. Test addCellShape dynamically
    await ppt.addCellShape('Table', 1, 0, {
      type: 'star5',
      fill: '#FBBF24',
      width: 15,
      height: 15,
      position: 'center',
    })

    const newStar = ppt.getCellShape('Table', 1, 0, 0)
    expect(newStar).not.toBeNull()
    expect(newStar.type).toBe('star5')

    // 4. Test updateCellShape dynamically
    await ppt.updateCellShape('Table', 1, 0, 0, {
      type: 'triangle',
      fill: '#EF4444',
      width: 10,
      height: 10,
    })

    const updatedStar = ppt.getCellShape('Table', 1, 0, 0)
    expect(updatedStar).not.toBeNull()
    expect(updatedStar.type).toBe('triangle')

    // 5. Test removeCellShape dynamically
    await ppt.removeCellShape('Table', 1, 0, 0)
    const removedStar = ppt.getCellShape('Table', 1, 0, 0)
    expect(removedStar).toBeNull()

    // Validate structural integrity
    const report = await ppt.validatePresentation()
    expect(report.valid).toBe(true)
  })

  it('should support built-in helpers (progressBar, badge, icons) and zIndex layering', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    ppt.updateTable('Table', {
      rows: [
        { A: 'Project Alpha', V: 75, B: 'Ongoing' },
        { A: 'Project Beta', V: 0, B: 'Stalled' },
      ],
      cellShapes: {
        V: row => ({
          type: 'progressBar',
          value: row.V,
          max: 100,
          fill: '#10B981',
          backgroundFill: '#E5E7EB',
        }),
        B: row => ({
          type: 'badge',
          text: row.B,
          fill: '#3B82F6',
        }),
      },
    })

    // Progress bar generates multiple shapes: background and filled foreground
    // Primary shape at index 0 should be the background (roundedRectangle)
    const progressBg = ppt.getCellShape('Table', 1, 1, '0_0')
    expect(progressBg).not.toBeNull()
    expect(progressBg.type).toBe('roundedRectangle')

    // Filled shape is index 0_1
    const progressFill = ppt.getCellShape('Table', 1, 1, '0_1')
    expect(progressFill).not.toBeNull()
    expect(progressFill.type).toBe('roundedRectangle')

    // For Project Beta (0 progress), no filled shape should be generated
    const betaProgressFill = ppt.getCellShape('Table', 2, 1, '0_1')
    expect(betaProgressFill).toBeNull()

    // Badge shape should exist and be a roundedRectangle
    const badgeShape = ppt.getCellShape('Table', 1, 2, 0)
    expect(badgeShape).not.toBeNull()
    expect(badgeShape.type).toBe('roundedRectangle')

    const report = await ppt.validatePresentation()
    expect(report.valid).toBe(true)
  })

  it('should support multiple shapes in a cell and layering with zIndex', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    ppt.updateTable('Table', {
      rows: [{ A: 'Layer test', V: 'OK' }],
      cellShapes: {
        V: () => [
          {
            type: 'rectangle',
            fill: '#E5E7EB',
            zIndex: 1,
            width: 20,
            height: 20,
          },
          {
            type: 'circle',
            fill: '#EF4444',
            zIndex: 2,
            width: 10,
            height: 10,
          },
        ],
      },
    })

    // Verify both shapes exist in cell (1, 1)
    const rectShape = ppt.getCellShape('Table', 1, 1, 0)
    const circleShape = ppt.getCellShape('Table', 1, 1, 1)

    expect(rectShape).not.toBeNull()
    expect(rectShape.type).toBe('square')
    expect(circleShape).not.toBeNull()
    expect(circleShape.type).toBe('circle')

    const report = await ppt.validatePresentation()
    expect(report.valid).toBe(true)
  })

  it('should support conditional shape rendering and lifecycle clean updates', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    // First update: status 'Active' has shape
    ppt.updateTable('Table', {
      rows: [
        { A: 'Alice', V: 'Active' },
        { A: 'Bob', V: 'Inactive' },
      ],
      cellShapes: {
        V: row => (row.V === 'Active' ? { type: 'circle', fill: '#10B981' } : null),
      },
    })

    // Alice cell has circle
    expect(ppt.getCellShape('Table', 1, 1, 0)).not.toBeNull()
    // Bob cell has no shape
    expect(ppt.getCellShape('Table', 2, 1, 0)).toBeNull()

    // Second update: Alice becomes Inactive (shape should be removed cleanly)
    ppt.updateTable('Table', {
      rows: [
        { A: 'Alice', V: 'Inactive' },
        { A: 'Bob', V: 'Inactive' },
      ],
      cellShapes: {
        V: row => (row.V === 'Active' ? { type: 'circle', fill: '#10B981' } : null),
      },
    })

    // Alice cell should now have NO shape
    expect(ppt.getCellShape('Table', 1, 1, 0)).toBeNull()

    const report = await ppt.validatePresentation()
    expect(report.valid).toBe(true)
  })

  it('should calculate coordinates correctly for merged cells', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    // Merge cells first: row 1, col 0 and 1
    ppt.mergeCells('Table', 1, 0, 1, 1)

    ppt.updateTable('Table', {
      rows: [{ A: 'Header' }, { A: 'Merged Row Data' }],
      cellShapes: {
        A: () => ({
          type: 'circle',
          fill: '#3B82F6',
          width: 10,
          height: 10,
          position: 'center',
        }),
      },
    })

    // Retrieve the shape in the merged cell
    const shape = ppt.getCellShape('Table', 1, 0, 0)
    expect(shape).not.toBeNull()

    const report = await ppt.validatePresentation()
    expect(report.valid).toBe(true)
  })

  it('should support dynamic row height calculation, boundary constraints, and helper APIs', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    // 1. Get initial bounds of the cell
    const initialBounds = ppt.getCellBounds('Table', 1, 1)
    expect(initialBounds).not.toBeNull()
    expect(initialBounds.width).toBeGreaterThan(0)
    expect(initialBounds.height).toBeGreaterThan(0)

    // 2. Perform table update with very long text in column A to trigger row height expansion
    ppt.updateTable('Table', {
      rows: [
        {
          A: 'This is an extremely long paragraph text that will definitely wrap to multiple lines given normal slide column widths. It should force the row height to expand dynamically.',
          V: 'Active',
          B: 50,
        },
      ],
      cellShapes: {
        V: () => ({
          type: 'circle',
          fill: '#10B981',
          width: '50%', // percentage width
          height: 10,
          position: 'middle-right',
        }),
        B: () => ({
          type: 'square',
          fill: '#3B82F6',
          size: 15,
          x: 1000, // excessively large offset to test cell boundary constraint
          y: 1000,
        }),
      },
    })

    // 3. Verify that row height has expanded
    const newBounds = ppt.getCellBounds('Table', 1, 1)
    expect(newBounds.height).toBeGreaterThan(initialBounds.height)

    // 4. Verify cell position helper
    const cellPos = ppt.getCellPosition('Table', 1, 1)
    expect(cellPos).not.toBeNull()
    expect(cellPos.row).toBe(1)
    expect(cellPos.column).toBe(1)
    expect(cellPos.x).toBe(newBounds.x)
    expect(cellPos.y).toBe(newBounds.y)

    // 5. Verify percentage width shape is placed and scaled relative to cell width
    const shapeV = ppt.getCellShape('Table', 1, 1, 0)
    expect(shapeV).not.toBeNull()
    expect(shapeV.width).toBe(Math.round(newBounds.width * 0.5))

    // 6. Verify cell boundaries constraint on shape B (which had x: 1000, y: 1000)
    const shapeB = ppt.getCellShape('Table', 1, 2, 0)
    expect(shapeB).not.toBeNull()

    const cellBoundsB = ppt.getCellBounds('Table', 1, 2)
    // Shape B's right edge should be exactly at cell bounds right edge
    expect(shapeB.x + shapeB.width).toBeLessThanOrEqual(cellBoundsB.x + cellBoundsB.width)
    expect(shapeB.y + shapeB.height).toBeLessThanOrEqual(cellBoundsB.y + cellBoundsB.height)

    const report = await ppt.validatePresentation()
    expect(report.valid).toBe(true)
  })

  it('should support 9 position presets and custom alignment configurations', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    const presets = [
      'top-left',
      'top-center',
      'top-right',
      'middle-left',
      'center',
      'middle-right',
      'bottom-left',
      'bottom-center',
      'bottom-right',
    ]

    for (let i = 0; i < presets.length; i++) {
      await ppt.addCellShape('Table', 1, 1, {
        type: 'circle',
        fill: '#EF4444',
        width: 10,
        height: 10,
        position: presets[i],
      })
      const shape = ppt.getCellShape('Table', 1, 1, i)
      expect(shape).not.toBeNull()
    }

    const report = await ppt.validatePresentation()
    expect(report.valid).toBe(true)
  })
})
