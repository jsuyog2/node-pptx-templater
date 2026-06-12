import { describe, it, expect } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')

describe('Cell Shapes Layout Regression Tests', () => {
  const runTests = existsSync(FIXTURE_FILE)

  if (!runTests) {
    it.skip('Skipping: sample.pptx fixture not found', () => {})
    return
  }

  it('should never change row height or column width when adding cell shapes', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3) // Slide 3 contains 'Table'

    // Get initial cell bounds
    const initialBounds = ppt.getCellBounds('Table', 1, 1)
    expect(initialBounds).not.toBeNull()

    // Add a cell shape
    await ppt.addCellShape('Table', 1, 1, {
      type: 'circle',
      width: 12,
      height: 12,
      fill: '#10B981',
    })

    // Get bounds after adding cell shape
    const afterBounds = ppt.getCellBounds('Table', 1, 1)
    expect(afterBounds).not.toBeNull()

    // Row heights and column widths must remain exactly the same
    expect(afterBounds.height).toBe(initialBounds.height)
    expect(afterBounds.width).toBe(initialBounds.width)

    // Structural validation
    const validationResult = await ppt.validatePresentation()
    expect(validationResult.valid).toBe(true)
  })

  it('should support small shapes in narrow cells without changing dimensions', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    // Col 2 is typically narrower
    const initialBounds = ppt.getCellBounds('Table', 1, 2)

    await ppt.addCellShape('Table', 1, 2, {
      type: 'circle',
      width: 4,
      height: 4,
      fill: '#EF4444',
      position: 'center',
    })

    const afterBounds = ppt.getCellBounds('Table', 1, 2)
    expect(afterBounds.height).toBe(initialBounds.height)
    expect(afterBounds.width).toBe(initialBounds.width)

    const shape = ppt.getCellShape('Table', 1, 2, 0)
    expect(shape).not.toBeNull()
    expect(shape.width).toBe(4)
    expect(shape.height).toBe(4)

    const validationResult = await ppt.validatePresentation()
    expect(validationResult.valid).toBe(true)
  })

  it('should scale down large shapes proportionally to fit the cell and log a warning', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    const initialBounds = ppt.getCellBounds('Table', 1, 1)

    // Add a shape that is significantly larger than the cell (e.g. 500x500)
    await ppt.addCellShape('Table', 1, 1, {
      type: 'circle',
      width: 500,
      height: 500,
      fill: '#3B82F6',
    })

    const afterBounds = ppt.getCellBounds('Table', 1, 1)

    // Row height and col width must remain unchanged
    expect(afterBounds.height).toBe(initialBounds.height)
    expect(afterBounds.width).toBe(initialBounds.width)

    // The shape should be scaled down to fit inside the cell boundaries
    const shape = ppt.getCellShape('Table', 1, 1, 0)
    expect(shape).not.toBeNull()
    expect(shape.width).toBeLessThanOrEqual(initialBounds.width)
    expect(shape.height).toBeLessThanOrEqual(initialBounds.height)

    const validationResult = await ppt.validatePresentation()
    expect(validationResult.valid).toBe(true)
  })

  it('should respect cell boundaries in cells with wrapped text', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    // First update table with wrapped text to change initial layout
    await ppt.updateTable('Table', {
      rows: [
        {
          A: 'This is a long wrapped text paragraph that forces the row height to grow due to text wrapping.',
          B: 'Other cell',
        },
      ],
    })

    const boundsAfterTextUpdate = ppt.getCellBounds('Table', 1, 1)

    // Add a cell shape to the wrapped text row
    await ppt.addCellShape('Table', 1, 1, {
      type: 'circle',
      width: 10,
      height: 10,
      fill: '#FBBF24',
      position: 'middle-right',
    })

    const boundsAfterShape = ppt.getCellBounds('Table', 1, 1)

    // Row heights and column widths must remain exactly as they were after text update
    expect(boundsAfterShape.height).toBe(boundsAfterTextUpdate.height)
    expect(boundsAfterShape.width).toBe(boundsAfterTextUpdate.width)

    const validationResult = await ppt.validatePresentation()
    expect(validationResult.valid).toBe(true)
  })

  it('should position shapes correctly within merged cells without altering layout', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    // Merge row 1 column 0 and column 1
    ppt.mergeCells('Table', 1, 0, 1, 1)

    const initialBounds = ppt.getCellBounds('Table', 1, 0)

    // Add cell shape in merged cell
    await ppt.addCellShape('Table', 1, 0, {
      type: 'circle',
      width: 15,
      height: 15,
      fill: '#10B981',
      position: 'center',
    })

    const afterBounds = ppt.getCellBounds('Table', 1, 0)
    expect(afterBounds.height).toBe(initialBounds.height)
    expect(afterBounds.width).toBe(initialBounds.width)

    const shape = ppt.getCellShape('Table', 1, 0, 0)
    expect(shape).not.toBeNull()

    const validationResult = await ppt.validatePresentation()
    expect(validationResult.valid).toBe(true)
  })

  it('should support multiple shapes in the same cell and support relative positioning', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    const initialBounds = ppt.getCellBounds('Table', 1, 1)

    // Add first shape with relative offset x: 4, y: 2
    await ppt.addCellShape('Table', 1, 1, {
      type: 'circle',
      x: 4,
      y: 2,
      width: 10,
      height: 10,
      fill: '#3B82F6',
    })

    // Add second shape in the same cell
    await ppt.addCellShape('Table', 1, 1, {
      type: 'circle',
      x: 15,
      y: 10,
      width: 8,
      height: 8,
      fill: '#10B981',
    })

    const afterBounds = ppt.getCellBounds('Table', 1, 1)
    expect(afterBounds.height).toBe(initialBounds.height)
    expect(afterBounds.width).toBe(initialBounds.width)

    const shape1 = ppt.getCellShape('Table', 1, 1, 0)
    const shape2 = ppt.getCellShape('Table', 1, 1, 1)

    expect(shape1).not.toBeNull()
    expect(shape2).not.toBeNull()

    const validationResult = await ppt.validatePresentation()
    expect(validationResult.valid).toBe(true)
  })

  it('should support cell shapes after dynamic row insertion', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3)

    // Dynamically insert a row
    await ppt.insertTableRow('Table', 2, { A: 'Inserted Row', B: 'Value' })

    const postInsertBounds = ppt.getCellBounds('Table', 2, 1)

    // Add shape to the newly inserted row
    await ppt.addCellShape('Table', 2, 1, {
      type: 'circle',
      width: 12,
      height: 12,
      fill: '#8B5CF6',
    })

    const postShapeBounds = ppt.getCellBounds('Table', 2, 1)
    expect(postShapeBounds.height).toBe(postInsertBounds.height)
    expect(postShapeBounds.width).toBe(postInsertBounds.width)

    const validationResult = await ppt.validatePresentation()
    expect(validationResult.valid).toBe(true)
  })
})
