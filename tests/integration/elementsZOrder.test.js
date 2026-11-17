import { describe, it, expect } from 'vitest'
const { PPTXTemplater } = require('../../src/index.js')
const JSZip = require('jszip')

/**
 * Creates a blank PPTX with 3 shapes on Slide 1 in order:
 * [BoxA, BoxB, BoxC]
 */
async function createTestDeck() {
  const ppt = await PPTXTemplater.createBlank()
  ppt.useSlide(1)

  await ppt.addShape('rectangle', { name: 'BoxA', x: 20, y: 20, width: 80, height: 80 })
  await ppt.addShape('rectangle', { name: 'BoxB', x: 40, y: 40, width: 80, height: 80 })
  await ppt.addShape('rectangle', { name: 'BoxC', x: 60, y: 60, width: 80, height: 80 })

  return ppt
}

describe('Elements Z-Order & Layering Subsystem (TASK-ELEM-004)', () => {
  it('should initialize and report correct initial stacking order', async () => {
    const ppt = await createTestDeck()
    const elements = ppt.getElements(1)

    // Expected order: BoxA, BoxB, BoxC
    expect(elements).toHaveLength(3)
    expect(elements[0].name).toBe('BoxA')
    expect(elements[1].name).toBe('BoxB')
    expect(elements[2].name).toBe('BoxC')
  })

  it('should bring forward and send backward successfully', async () => {
    const ppt = await createTestDeck()

    // Move BoxA forward (swaps with BoxB) -> [BoxB, BoxA, BoxC]
    ppt.bringForward('BoxA')
    let elements = ppt.getElements(1)
    expect(elements[0].name).toBe('BoxB')
    expect(elements[1].name).toBe('BoxA')
    expect(elements[2].name).toBe('BoxC')

    // Move BoxA backward (swaps back with BoxB) -> [BoxA, BoxB, BoxC]
    ppt.sendBackward('BoxA')
    elements = ppt.getElements(1)
    expect(elements[0].name).toBe('BoxA')
    expect(elements[1].name).toBe('BoxB')
    expect(elements[2].name).toBe('BoxC')
  })

  it('should bring to front and send to back successfully', async () => {
    const ppt = await createTestDeck()

    // Send BoxC to back -> [BoxC, BoxA, BoxB]
    ppt.sendToBack('BoxC')
    let elements = ppt.getElements(1)
    expect(elements[0].name).toBe('BoxC')
    expect(elements[1].name).toBe('BoxA')
    expect(elements[2].name).toBe('BoxB')

    // Bring BoxC to front -> [BoxA, BoxB, BoxC]
    ppt.bringToFront('BoxC')
    elements = ppt.getElements(1)
    expect(elements[2].name).toBe('BoxC')
    expect(elements[0].name).toBe('BoxA')
    expect(elements[1].name).toBe('BoxB')
  })

  it('should set explicit z-order via setZOrder', async () => {
    const ppt = await createTestDeck()

    // Initial: [BoxA, BoxB, BoxC]
    // Move BoxA to position 3 (front) -> [BoxB, BoxC, BoxA]
    ppt.setZOrder('BoxA', 3)
    let elements = ppt.getElements(1)
    expect(elements[2].name).toBe('BoxA')

    // Move BoxA to position 1 (back) -> [BoxA, BoxB, BoxC]
    ppt.setZOrder('BoxA', 1)
    elements = ppt.getElements(1)
    expect(elements[0].name).toBe('BoxA')
  })

  it('should safely handle boundary shifts without error (already at front / back)', async () => {
    const ppt = await createTestDeck()

    // BoxC is already at the top (front)
    ppt.bringToFront('BoxC')
    let elements = ppt.getElements(1)
    expect(elements[2].name).toBe('BoxC')

    // Calling bringForward on top-most element is a safe no-op
    ppt.bringForward('BoxC')
    elements = ppt.getElements(1)
    expect(elements[2].name).toBe('BoxC')

    // BoxA is at bottom (back)
    ppt.sendToBack('BoxA')
    elements = ppt.getElements(1)
    expect(elements[0].name).toBe('BoxA')

    // Calling sendBackward on bottom-most element is a safe no-op
    ppt.sendBackward('BoxA')
    elements = ppt.getElements(1)
    expect(elements[0].name).toBe('BoxA')
  })

  it('should support fluent chaining on all 5 layering methods', async () => {
    const ppt = await createTestDeck()

    const instance = ppt
      .useSlide(1)
      .bringToFront('BoxA')
      .sendBackward('BoxA')
      .sendToBack('BoxC')
      .bringForward('BoxC')
      .setZOrder('BoxB', 2)

    expect(instance).toBe(ppt)
  })

  it('should preserve OpenXML container headers <p:nvGrpSpPr> and <p:grpSpPr> after z-ordering', async () => {
    const ppt = await createTestDeck()

    ppt.bringToFront('BoxA')
    ppt.sendToBack('BoxB')
    ppt.setZOrder('BoxC', 2)

    const buf = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buf)
    const slideXml = await zip.files['ppt/slides/slide1.xml'].async('text')

    // Headers must precede drawing children
    const nvIdx = slideXml.indexOf('<p:nvGrpSpPr>')
    const grpIdx = slideXml.indexOf('<p:grpSpPr>')
    const boxAIdx = slideXml.indexOf('BoxA')

    expect(nvIdx).toBeGreaterThan(-1)
    expect(grpIdx).toBeGreaterThan(-1)
    expect(nvIdx).toBeLessThan(grpIdx)
    expect(grpIdx).toBeLessThan(boxAIdx)
  })
})
