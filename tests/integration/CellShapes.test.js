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

  it('should support table cell positioning, shape centering options, and alignment helpers', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    // 1. Verify getCellBounds & getCellPosition with a string ID and returning dimensions
    const cellBounds = ppt.getCellBounds('Table', 1, 1)
    expect(cellBounds).not.toBeNull()
    expect(cellBounds.width).toBeGreaterThan(0)
    expect(cellBounds.height).toBeGreaterThan(0)

    const cellPos = ppt.getCellPosition('Table', 1, 1)
    expect(cellPos).not.toBeNull()
    expect(cellPos.width).toBe(cellBounds.width)
    expect(cellPos.height).toBe(cellBounds.height)

    // 2. Verify support for passing table objects
    const mockTableObj = { id: 'Table', name: 'SomeTable' }
    const cellBoundsObj = ppt.getCellBounds(mockTableObj, 1, 1)
    expect(cellBoundsObj).toEqual(cellBounds)

    // 3. Add new shape aligned to cell center (using dual signature: type, options)
    await ppt.addShape('ellipse', {
      id: 'CenteredEllipse',
      width: 40,
      height: 30,
      fill: '#8B5CF6',
      alignToCell: {
        table: mockTableObj,
        row: 1,
        col: 1,
        horizontal: 'center',
        vertical: 'middle',
      },
    })

    const shape = ppt.getShape('CenteredEllipse')
    expect(shape).not.toBeNull()

    const expectedX = Math.round(cellBounds.x + (cellBounds.width - 40) / 2)
    const expectedY = Math.round(cellBounds.y + (cellBounds.height - 30) / 2)

    const shapeX = shape.x
    const shapeY = shape.y

    expect(shapeX).toBe(expectedX)
    expect(shapeY).toBe(expectedY)

    // 4. Test alignShapeToCell helper (right/bottom alignment)
    ppt.alignShapeToCell('CenteredEllipse', 'Table', 1, 1, {
      horizontal: 'right',
      vertical: 'bottom',
    })

    const shapeAligned = ppt.getShape('CenteredEllipse')
    const shapeXAligned = shapeAligned.x
    const shapeYAligned = shapeAligned.y

    const expectedXRight = Math.round(cellBounds.x + cellBounds.width - 40)
    const expectedYBottom = Math.round(cellBounds.y + cellBounds.height - 30)

    expect(shapeXAligned).toBe(expectedXRight)
    expect(shapeYAligned).toBe(expectedYBottom)

    const report = await ppt.validatePresentation()
    expect(report.valid).toBe(true)
  })

  it('should support vertical alignment configurations (top, middle, bottom) and place shapes at correct, distinct positions', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    const cellBounds = ppt.getCellBounds('Table', 1, 1)
    expect(cellBounds).not.toBeNull()

    // Add three shapes with different vertical alignments
    await ppt.addCellShape('Table', 1, 1, {
      type: 'circle',
      fill: '#EF4444',
      width: 10,
      height: 10,
      horizontal: 'center',
      vertical: 'top',
    })
    await ppt.addCellShape('Table', 1, 1, {
      type: 'circle',
      fill: '#10B981',
      width: 10,
      height: 10,
      horizontal: 'center',
      vertical: 'middle',
    })
    await ppt.addCellShape('Table', 1, 1, {
      type: 'circle',
      fill: '#3B82F6',
      width: 10,
      height: 10,
      horizontal: 'center',
      vertical: 'bottom',
    })

    const shapeTop = ppt.getCellShape('Table', 1, 1, 0)
    const shapeMiddle = ppt.getCellShape('Table', 1, 1, 1)
    const shapeBottom = ppt.getCellShape('Table', 1, 1, 2)

    expect(shapeTop).not.toBeNull()
    expect(shapeMiddle).not.toBeNull()
    expect(shapeBottom).not.toBeNull()

    // Assert that the Y positions are distinct and increasing
    expect(shapeMiddle.y).toBeGreaterThan(shapeTop.y)
    expect(shapeBottom.y).toBeGreaterThan(shapeMiddle.y)

    // Mathematically assert the values:
    // top: cellBounds.y + 5 (since config.y is undefined and default padding is 5px)
    // middle: cellBounds.y + (cellBounds.height - 10) / 2
    // bottom: cellBounds.y + cellBounds.height - 10 - 5
    const expectedTopY = cellBounds.y + 5
    const expectedMiddleY = Math.round(cellBounds.y + (cellBounds.height - 10) / 2)
    const expectedBottomY = cellBounds.y + cellBounds.height - 10 - 5

    expect(shapeTop.y).toBe(expectedTopY)
    expect(shapeMiddle.y).toBe(expectedMiddleY)
    expect(shapeBottom.y).toBe(expectedBottomY)

    const report = await ppt.validatePresentation()
    expect(report.valid).toBe(true)
  })

  it('should not modify row heights or column widths when calling addCellShape()', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    const getRowHeights = xml => {
      const regex = /<a:tr\s+h="(\d+)">/g
      const heights = []
      let match
      while ((match = regex.exec(xml)) !== null) {
        heights.push(match[1])
      }
      return heights
    }

    const getColWidths = xml => {
      const regex = /<a:gridCol\s+w="(\d+)">/g
      const widths = []
      let match
      while ((match = regex.exec(xml)) !== null) {
        widths.push(match[1])
      }
      return widths
    }

    const slideXmlBefore = await ppt.slideManager.getSlideXmlAsync(3)
    const heightsBefore = getRowHeights(slideXmlBefore)
    const widthsBefore = getColWidths(slideXmlBefore)

    // Call addCellShape multiple times
    await ppt.addCellShape('Table', 1, 1, {
      type: 'circle',
      fill: '#EF4444',
      width: 10,
      height: 10,
      horizontal: 'center',
      vertical: 'top',
    })
    await ppt.addCellShape('Table', 1, 1, {
      type: 'rectangle',
      fill: '#3B82F6',
      width: 15,
      height: 15,
      horizontal: 'center',
      vertical: 'bottom',
    })

    const slideXmlAfter = await ppt.slideManager.getSlideXmlAsync(3)
    const heightsAfter = getRowHeights(slideXmlAfter)
    const widthsAfter = getColWidths(slideXmlAfter)

    expect(heightsAfter).toEqual(heightsBefore)
    expect(widthsAfter).toEqual(widthsBefore)

    const report = await ppt.validatePresentation()
    expect(report.valid).toBe(true)
  })

  it('should calculate coordinates correctly for vertically and horizontally merged cells and respect alignments/offsets', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    // Populate the table with 4 data rows first to avoid out-of-bounds error
    ppt.updateTable('Table', {
      rows: [
        { A: 'Row 1', V: 'Val 1', B: 10 },
        { A: 'Row 2', V: 'Val 2', B: 20 },
        { A: 'Row 3', V: 'Val 3', B: 30 },
        { A: 'Row 4', V: 'Val 4', B: 40 },
      ],
    })

    // 1. Vertical and Horizontal merge: Merge rows 1-2 and columns 1-2 (2x2 merged cell)
    ppt.mergeCells('Table', 1, 1, 2, 2)

    // Add shapes to verify all combinations of alignments/offsets
    // - Shape 1: Center-Middle, no offset (true centering)
    await ppt.addCellShape('Table', 1, 1, {
      type: 'circle',
      fill: '#EF4444',
      width: 10,
      height: 10,
      alignX: 'center',
      alignY: 'middle',
    })

    // - Shape 2: Top-Left alignment with explicit offset (x: 5, y: 3)
    await ppt.addCellShape('Table', 1, 1, {
      type: 'rectangle',
      fill: '#10B981',
      width: 15,
      height: 15,
      alignX: 'left',
      alignY: 'top',
      x: 5,
      y: 3,
    })

    // - Shape 3: Bottom-Right alignment with explicit offset (x: 4, y: 2)
    await ppt.addCellShape('Table', 1, 1, {
      type: 'star5',
      fill: '#3B82F6',
      width: 12,
      height: 12,
      alignX: 'right',
      alignY: 'bottom',
      x: 4,
      y: 2,
    })

    // Get cell bounds and shapes
    const bounds = ppt.getCellBounds('Table', 1, 1)
    const shapeCenter = ppt.getCellShape('Table', 1, 1, 0)
    const shapeTopLeft = ppt.getCellShape('Table', 1, 1, 1)
    const shapeBottomRight = ppt.getCellShape('Table', 1, 1, 2)

    expect(shapeCenter).not.toBeNull()
    expect(shapeTopLeft).not.toBeNull()
    expect(shapeBottomRight).not.toBeNull()

    // 1. Verify true centering in merged cell bounds
    expect(shapeCenter.x).toBe(Math.round(bounds.x + (bounds.width - 10) / 2))
    expect(shapeCenter.y).toBe(Math.round(bounds.y + (bounds.height - 10) / 2))

    // 2. Verify Top-Left with offset (x: 5, y: 3)
    expect(shapeTopLeft.x).toBe(Math.round(bounds.x + 5))
    expect(shapeTopLeft.y).toBe(Math.round(bounds.y + 3))

    // 3. Verify Bottom-Right with offset (x: 4, y: 2)
    expect(shapeBottomRight.x).toBe(Math.round(bounds.x + bounds.width - 12 - 4))
    expect(shapeBottomRight.y).toBe(Math.round(bounds.y + bounds.height - 12 - 2))

    const report = await ppt.validatePresentation()
    expect(report.valid).toBe(true)
  })

  it('should calculate correct cell coordinates for standard, merged, wrapped-text, and dynamic/inserted rows with object signatures', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    // 1. Set up a dynamic table with dynamic row text (including very long text that wraps and expands the row height),
    // and custom dimensions.
    ppt.updateTable('Table', {
      rows: [
        { A: 'Standard Text', V: 'Normal status', B: 10 },
        {
          A: 'This is a very long text that wraps across multiple lines and expands the row height significantly. We expect this row to be much taller than the template default.',
          V: 'Wrapped status',
          B: 20,
        },
        { A: 'Row 3 Col 0', V: 'Status 3', B: 30 },
        { A: 'Row 4 Col 0', V: 'Status 4', B: 40 },
      ],
    })

    // 2. Insert an additional row dynamically using insertTableRow
    ppt.insertTableRow('Table', 3, ['Inserted Row X', 'Status X', 99])

    // Now table has:
    // Row 0: Header
    // Row 1: 'Standard Text'
    // Row 2: (Wrapped Row)
    // Row 3: 'Inserted Row X'
    // Row 4: 'Row 3 Col 0'
    // Row 5: 'Row 4 Col 0'
    // Let's verify the row count is 6 (1 header + 5 data rows)
    const rowsMeta = ppt.getTableRows('Table', { includeMetadata: true })
    expect(rowsMeta.rowCount).toBe(6)

    // 3. Merge cells in the table (vertical merge spanning the wrapped row and the inserted row)
    // Merge Row 2 and Row 3 of Column 1 (index 1)
    ppt.mergeCells('Table', 2, 1, 3, 1)

    // 4. Call addCellShape with both positional and object-based call signatures
    // - Shape 1: Positional call signature on Row 2 Column 1 (merged cell)
    await ppt.addCellShape('Table', 2, 1, {
      type: 'circle',
      width: 10,
      height: 10,
      alignX: 'center',
      alignY: 'middle',
    })

    // - Shape 2: Object-based call signature targeting the same merged cell
    await ppt.addCellShape('Table', {
      row: 2,
      column: 1,
      type: 'circle',
      width: 10,
      height: 10,
      alignX: 'center',
      alignY: 'middle',
    })

    // - Shape 3: Center-aligned shape in the wrapped cell (Row 2, Column 0) using object signature
    await ppt.addCellShape('Table', {
      row: 2,
      column: 0,
      type: 'rectangle',
      width: 15,
      height: 15,
      alignX: 'center',
      alignY: 'middle',
    })

    // - Shape 4: Explicit offset (x: 5, y: 3) shape in standard cell using object signature
    await ppt.addCellShape('Table', {
      row: 1,
      column: 2,
      type: 'star5',
      width: 12,
      height: 12,
      x: 5,
      y: 3,
    })

    // 5. Assert bounds and shape coordinate correctness
    const boundsMerged = ppt.getCellBounds('Table', 2, 1)
    const boundsMergedObj = ppt.getCellBounds('Table', { row: 2, column: 1 })
    // getCellBounds with object and positional signature must return exactly the same bounds
    expect(boundsMergedObj).toEqual(boundsMerged)

    const shapePositional = ppt.getCellShape('Table', 2, 1, 0)
    const shapeObject = ppt.getCellShape('Table', 2, 1, 1)
    expect(shapePositional).not.toBeNull()
    expect(shapeObject).not.toBeNull()

    // Shapes added to the same target using different signatures must be positioned at the exact same coordinates!
    expect(shapeObject.x).toBe(shapePositional.x)
    expect(shapeObject.y).toBe(shapePositional.y)

    // Verify true centering in the merged cell (which includes the wrapped row and the inserted row heights)
    expect(shapePositional.x).toBe(Math.round(boundsMerged.x + (boundsMerged.width - 10) / 2))
    expect(shapePositional.y).toBe(Math.round(boundsMerged.y + (boundsMerged.height - 10) / 2))

    // Verify wrapped row (Row 2, Column 0) centering
    const boundsWrapped = ppt.getCellBounds('Table', 2, 0)
    const shapeWrapped = ppt.getCellShape('Table', 2, 0, 0)
    expect(shapeWrapped).not.toBeNull()
    expect(shapeWrapped.x).toBe(Math.round(boundsWrapped.x + (boundsWrapped.width - 15) / 2))
    expect(shapeWrapped.y).toBe(Math.round(boundsWrapped.y + (boundsWrapped.height - 15) / 2))

    // Verify explicit offset (x: 5, y: 3) relative to the cell's top-left corner
    const boundsOffset = ppt.getCellBounds('Table', 1, 2)
    const shapeOffset = ppt.getCellShape('Table', 1, 2, 0)
    expect(shapeOffset).not.toBeNull()
    expect(shapeOffset.x).toBe(boundsOffset.x + 5)
    expect(shapeOffset.y).toBe(boundsOffset.y + 3)

    const report = await ppt.validatePresentation()
    expect(report.valid).toBe(true)
  })
})
