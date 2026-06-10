import { describe, test, expect } from 'vitest'
const { PPTXTemplater } = require('../../src/index.js')

/**
 * Helper: creates a new PPTXTemplater, adds a test slide with known elements,
 * and returns both the engine and the index of the newly added slide.
 */
async function createTestEngine() {
  const ppt = await PPTXTemplater.create()

  ppt.addSlide({
    title: 'Test Slide',
    elements: [
      { type: 'text', name: 'Logo', value: 'Company Logo' },
      { type: 'text', name: 'Background', value: 'Background Block' },
      { type: 'text', name: 'Chart', value: 'Performance Chart' },
    ],
  })

  // The newly added slide is always at the new slideCount
  const slideIndex = ppt.getInfo().slideCount
  return { ppt, slideIndex }
}

describe('Z-Order (Layer Management) Integration Tests', () => {
  test('should initialize and report correct Z-order for slide elements', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    const order = ppt.getObjectOrder(slideIndex)
    // Expect: Title 1, Logo, Background, Chart (4 elements)
    expect(order).toHaveLength(4)
    expect(order[0].id).toBe('Title 1')
    expect(order[1].id).toBe('Logo')
    expect(order[2].id).toBe('Background')
    expect(order[3].id).toBe('Chart')

    expect(order[0].zIndex).toBe(1)
    expect(order[1].zIndex).toBe(2)
    expect(order[2].zIndex).toBe(3)
    expect(order[3].zIndex).toBe(4)
  })

  test('should bring forward and send backward successfully', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    // Initial: [Title 1, Logo, Background, Chart]
    // Move 'Logo' forward (swaps with Background)
    ppt.bringForward({ slide: slideIndex, objectId: 'Logo' })
    let order = ppt.getObjectOrder(slideIndex)
    expect(order[1].id).toBe('Background')
    expect(order[2].id).toBe('Logo')

    // Move 'Logo' backward (swaps back with Background)
    ppt.sendBackward({ slide: slideIndex, objectId: 'Logo' })
    order = ppt.getObjectOrder(slideIndex)
    expect(order[1].id).toBe('Logo')
    expect(order[2].id).toBe('Background')
  })

  test('should bring to front and send to back successfully', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    // Initial: [Title 1, Logo, Background, Chart]
    // Send 'Chart' to back (index 0)
    ppt.sendToBack({ slide: slideIndex, objectId: 'Chart' })
    let order = ppt.getObjectOrder(slideIndex)
    expect(order[0].id).toBe('Chart')
    expect(order[1].id).toBe('Title 1')
    expect(order[2].id).toBe('Logo')
    expect(order[3].id).toBe('Background')

    // Bring 'Logo' to front (index 3)
    ppt.bringToFront({ slide: slideIndex, objectId: 'Logo' })
    order = ppt.getObjectOrder(slideIndex)
    expect(order[3].id).toBe('Logo')
    expect(order[0].id).toBe('Chart')
    expect(order[1].id).toBe('Title 1')
    expect(order[2].id).toBe('Background')
  })

  test('should set absolute zIndex correctly', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    // Initial: [Title 1, Logo, Background, Chart]
    // Set 'Logo' to zIndex 4 (top-most position)
    ppt.setZIndex({ slide: slideIndex, objectId: 'Logo', zIndex: 4 })
    let order = ppt.getObjectOrder(slideIndex)
    expect(order[3].id).toBe('Logo')

    // Set 'Chart' to zIndex 1 (bottom-most)
    ppt.setZIndex({ slide: slideIndex, objectId: 'Chart', zIndex: 1 })
    order = ppt.getObjectOrder(slideIndex)
    expect(order[0].id).toBe('Chart')
  })

  test('should move elements relative to other targets (moveObjectBefore / moveObjectAfter)', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    // Initial: [Title 1, Logo, Background, Chart]
    // Move 'Chart' before 'Logo' → [Title 1, Chart, Logo, Background]
    ppt.moveObjectBefore({ slide: slideIndex, objectId: 'Chart', targetId: 'Logo' })
    let order = ppt.getObjectOrder(slideIndex)
    expect(order[1].id).toBe('Chart')
    expect(order[2].id).toBe('Logo')

    // Move 'Title 1' after 'Background' → [Chart, Logo, Background, Title 1]
    ppt.moveObjectAfter({ slide: slideIndex, objectId: 'Title 1', targetId: 'Background' })
    order = ppt.getObjectOrder(slideIndex)
    expect(order[3].id).toBe('Title 1')
  })

  test('should reorder objects in bulk', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    // Reorder bulk: specify all 4 objects in desired order
    ppt.reorderObjects({
      slide: slideIndex,
      order: ['Background', 'Chart', 'Logo', 'Title 1'],
    })

    const order = ppt.getObjectOrder(slideIndex)
    expect(order[0].id).toBe('Background')
    expect(order[1].id).toBe('Chart')
    expect(order[2].id).toBe('Logo')
    expect(order[3].id).toBe('Title 1')
  })

  test('should support template configs in applyZOrder', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    // Initial: [Title 1, Logo, Background, Chart]
    // Apply multiple layer instructions — operations are sequential:
    // 1. sendToBack Background   → [Background, Title 1, Logo, Chart]
    // 2. setZIndex Logo to 4     → [Background, Title 1, Chart, Logo]
    // 3. bringForward Chart      → [Background, Title 1, Logo, Chart]
    ppt.applyZOrder(slideIndex, [
      { id: 'Background', sendToBack: true },
      { id: 'Logo', zIndex: 4 },
      { id: 'Chart', bringForward: true },
    ])

    const order = ppt.getObjectOrder(slideIndex)
    // Background is bottom-most (index 0)
    expect(order[0].id).toBe('Background')
    // Chart ends up at top (index 3) after bringForward
    expect(order[3].id).toBe('Chart')
  })

  test('should support layer utility helper functions', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    // Initial: [Title 1, Logo, Background, Chart]
    expect(ppt.getTopMostObject(slideIndex).id).toBe('Chart')
    expect(ppt.getBottomMostObject(slideIndex).id).toBe('Title 1')

    // Swap Logo and Chart: [Title 1, Chart, Background, Logo]
    ppt.swapObjects(slideIndex, 'Logo', 'Chart')
    const afterSwap = ppt.getObjectOrder(slideIndex)
    expect(afterSwap[1].id).toBe('Chart')
    expect(afterSwap[3].id).toBe('Logo')

    // Sort by descending alphabetical (b.id.localeCompare(a.id))
    // T > L > C > B, so sorted descending → [Title 1, Logo, Chart, Background]
    ppt.sortObjects(slideIndex, (a, b) => b.id.localeCompare(a.id))
    const sorted = ppt.getObjectOrder(slideIndex)
    // Descending sort: Title 1 is first (index 0), Background is last (index 3)
    expect(sorted[0].id).toBe('Title 1')
    expect(sorted[3].id).toBe('Background')
  })

  test('should validate slide structure successfully without errors', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    ppt.bringToFront({ slide: slideIndex, objectId: 'Logo' })

    const validation = await ppt.validate()
    expect(validation.valid).toBe(true)
    expect(validation.errors).toHaveLength(0)
  })

  test('should normalize z-order without corrupting structure', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    // Scramble order then normalize
    ppt.bringToFront({ slide: slideIndex, objectId: 'Title 1' })
    ppt.normalizeZOrder(slideIndex)

    // After normalize, we should still be able to get order
    const order = ppt.getObjectOrder(slideIndex)
    expect(order).toHaveLength(4)
    expect(order.every(o => o.zIndex > 0)).toBe(true)
  })

  test('should correctly identify element types', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    const order = ppt.getObjectOrder(slideIndex)
    // All added as 'text' type elements
    const types = order.map(o => o.type)
    expect(types).toContain('text')
  })

  test('should handle getObjectOrder via fluent useSlide API', async () => {
    const { ppt, slideIndex } = await createTestEngine()

    ppt.useSlide(slideIndex)
    const order = ppt.getObjectOrder()
    expect(order).toHaveLength(4)
  })
})
