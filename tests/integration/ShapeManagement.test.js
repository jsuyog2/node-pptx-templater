import { describe, test, expect } from 'vitest'
const { PPTXTemplater } = require('../../src/index.js')

async function createTestEngine() {
  const ppt = await PPTXTemplater.create()
  ppt.addSlide({
    title: 'Shape Test Slide',
    elements: [],
  })
  const slideIndex = ppt.getInfo().slideCount
  return { ppt, slideIndex }
}

describe('Shape Management Integration Tests', () => {
  test('should validate shapes correctly', async () => {
    const { ppt } = await createTestEngine()

    // Valid options
    const errors1 = ppt.validateShape({
      id: 'valid-rect',
      type: 'rectangle',
      x: 10,
      y: 10,
      width: 100,
      height: 50,
    })
    expect(errors1).toHaveLength(0)

    // Missing ID
    const errors2 = ppt.validateShape({
      type: 'rectangle',
      x: 10,
      y: 10,
      width: 100,
      height: 50,
    })
    expect(errors2).toContain('Shape ID is missing or empty.')

    // Unsupported type
    const errors3 = ppt.validateShape({
      id: 'test-1',
      type: 'octagon',
      x: 10,
      y: 10,
      width: 100,
      height: 50,
    })
    expect(errors3).toContain(
      'Unsupported shape type: "octagon". Supported types are: rectangle, square, circle, ellipse, roundedRectangle, triangle, star5, upArrow, downArrow, leftArrow, rightArrow, diamond, hexagon, line.'
    )

    // Invalid square dimensions
    const errors4 = ppt.validateShape({
      id: 'test-2',
      type: 'square',
      x: 10,
      y: 10,
    })
    expect(errors4).toContain('Square "size" must be a positive number.')

    // Invalid colors
    const errors5 = ppt.validateShape({
      id: 'test-3',
      type: 'rectangle',
      x: 10,
      y: 10,
      width: 100,
      height: 50,
      fill: 'invalid-color',
    })
    expect(errors5).toContain('Invalid fill color: "invalid-color". Must be a valid hex color.')

    // Invalid border radius for wrong shape
    const errors6 = ppt.validateShape({
      id: 'test-4',
      type: 'rectangle',
      x: 10,
      y: 10,
      width: 100,
      height: 50,
      borderRadius: 10,
    })
    expect(errors6).toContain('Shape type "rectangle" does not support borderRadius.')
  })

  test('should add shapes of various types with styling and text', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    // 1. Add Rectangle
    await ppt.useSlide(slideIndex).addShape({
      type: 'rectangle',
      id: 'rect-1',
      x: 10,
      y: 10,
      width: 200,
      height: 100,
      fill: '#2563EB',
      border: {
        color: '#000000',
        width: 2,
      },
      transparency: 30,
      shadow: true,
      rotation: 45,
      text: 'Hello Rectangle',
      textStyle: {
        fontSize: 18,
        bold: true,
        color: '#FFFFFF',
        align: 'center',
      },
    })

    // 2. Add Square
    await ppt.addShape({
      type: 'square',
      id: 'square-1',
      x: 220,
      y: 10,
      size: 100,
    })

    // 3. Add Circle
    await ppt.addShape({
      type: 'circle',
      id: 'circle-1',
      x: 10,
      y: 120,
      radius: 50,
    })

    // 4. Add Ellipse
    await ppt.addShape({
      type: 'ellipse',
      id: 'ellipse-1',
      x: 120,
      y: 120,
      width: 150,
      height: 80,
    })

    // 5. Add Rounded Rectangle with border radius
    await ppt.addShape({
      type: 'roundedRectangle',
      id: 'card',
      x: 10,
      y: 230,
      width: 300,
      height: 120,
      borderRadius: 20,
      fill: {
        type: 'gradient',
        colors: ['#2563EB', '#7C3AED'],
      },
    })

    // Verify they are added
    const shapes = ppt.getShapes()
    expect(shapes.length).toBe(6) // 5 shapes + 1 title shape

    // Discover individual shape details
    const rect = ppt.getShape('rect-1')
    expect(rect).not.toBeNull()
    expect(rect.type).toBe('rectangle')
    expect(rect.x).toBe(10)
    expect(rect.y).toBe(10)
    expect(rect.width).toBe(200)
    expect(rect.height).toBe(100)

    const square = ppt.getShape('square-1')
    expect(square.type).toBe('square')
    expect(square.width).toBe(100)
    expect(square.height).toBe(100)

    const circle = ppt.getShape('circle-1')
    expect(circle.type).toBe('circle')
    expect(circle.width).toBe(100)
    expect(circle.height).toBe(100)

    const roundedRect = ppt.getShape('card')
    expect(roundedRect.type).toBe('roundedRectangle')
  })

  test('should update existing shapes correctly', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    await ppt.useSlide(slideIndex).addShape({
      type: 'rectangle',
      id: 'target-shape',
      x: 50,
      y: 50,
      width: 100,
      height: 100,
      text: 'Original Text',
    })

    // Verify initial state
    let shape = ppt.getShape('target-shape')
    expect(shape.x).toBe(50)
    expect(shape.width).toBe(100)

    // Update position and size
    await ppt.updateShape('target-shape', {
      x: 100,
      y: 150,
      width: 250,
      height: 120,
    })

    shape = ppt.getShape('target-shape')
    expect(shape.x).toBe(100)
    expect(shape.y).toBe(150)
    expect(shape.width).toBe(250)
    expect(shape.height).toBe(120)

    // Update fill, border, rotation, and text
    await ppt.updateShape('target-shape', {
      fill: '#10B981',
      border: {
        color: '#EF4444',
        width: 3,
      },
      rotation: 90,
      text: 'Updated Text',
    })

    // The text and styles are written inside XML object.
    // Let's do a discovery to ensure it still resolves.
    shape = ppt.getShape('target-shape')
    expect(shape.id).toBe('target-shape')
  })

  test('should remove shape and sync z-order', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    await ppt.useSlide(slideIndex).addShape({
      type: 'rectangle',
      id: 'shape-to-remove',
      x: 10,
      y: 10,
      width: 100,
      height: 100,
    })

    expect(ppt.getShape('shape-to-remove')).not.toBeNull()

    // Remove the shape
    await ppt.removeShape('shape-to-remove')

    expect(ppt.getShape('shape-to-remove')).toBeNull()
  })

  test('should control shape z-index layering', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    // Clear slide title or elements to keep z-order simple
    await ppt.useSlide(slideIndex).addShape({
      type: 'rectangle',
      id: 'layer-1',
      x: 10,
      y: 10,
      width: 50,
      height: 50,
    })

    await ppt.addShape({
      type: 'rectangle',
      id: 'layer-2',
      x: 20,
      y: 20,
      width: 50,
      height: 50,
    })

    // Order initially: [Title 1, layer-1, layer-2]
    let order = ppt.getObjectOrder(slideIndex)
    expect(order[1].id).toBe('layer-1')
    expect(order[2].id).toBe('layer-2')

    // Bring layer-1 to front
    await ppt.bringToFront('layer-1')
    order = ppt.getObjectOrder(slideIndex)
    expect(order[2].id).toBe('layer-1')
    expect(order[1].id).toBe('layer-2')

    // Send layer-1 backward
    await ppt.sendBackward('layer-1')
    order = ppt.getObjectOrder(slideIndex)
    expect(order[1].id).toBe('layer-1')
    expect(order[2].id).toBe('layer-2')
  })
})
