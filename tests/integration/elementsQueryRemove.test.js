import { describe, it, expect } from 'vitest'
const { PPTXTemplater } = require('../../src/index.js')
const JSZip = require('jszip')

describe('Elements Query & Removal Subsystem (TASK-ELEM-003)', () => {
  it('should query all visual element types with normalized metadata using getElements', async () => {
    const ppt = await PPTXTemplater.createBlank()

    // Add shape and table to blank slide 1
    ppt.useSlide(1)
    await ppt.addShape('rectangle', {
      name: 'CardShape',
      x: 100,
      y: 100,
      width: 300,
      height: 200,
    })
    ppt.addTable(1, 2, 2, [
      [{ text: 'A1' }, { text: 'B1' }],
      [{ text: 'A2' }, { text: 'B2' }],
    ])

    const elements = ppt.getElements(1)
    expect(elements).toHaveLength(2)

    // Check each descriptor has normalized metadata
    for (const el of elements) {
      expect(el).toHaveProperty('id')
      expect(el).toHaveProperty('name')
      expect(el).toHaveProperty('type')
      expect(el).toHaveProperty('bounds')
      expect(typeof el.bounds.x).toBe('number')
      expect(typeof el.bounds.y).toBe('number')
      expect(typeof el.bounds.width).toBe('number')
      expect(typeof el.bounds.height).toBe('number')
      expect(el).toHaveProperty('slideIndex', 1)
      expect(el).toHaveProperty('element')
    }

    // Verify type normalization
    const types = elements.map(e => e.type)
    expect(types).toContain('shape')
    expect(types).toContain('table')
  })

  it('should inspect single element by ID or name using getElement', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('roundRect', {
      id: 2001,
      name: 'TargetPill',
      x: 50,
      y: 50,
      width: 150,
      height: 40,
    })

    // Retrieve by Name
    const byName = ppt.getElement('TargetPill')
    expect(byName).not.toBeNull()
    expect(byName.name).toBe('TargetPill')
    expect(byName.id).toBe('2001')
    expect(byName.type).toBe('shape')

    // Retrieve by numerical ID
    const byId = ppt.getElement('2001')
    expect(byId).not.toBeNull()
    expect(byId.name).toBe('TargetPill')

    // Non-existent element returns null
    const notFound = ppt.getElement('NonExistent')
    expect(notFound).toBeNull()
  })

  it('should filter elements by type and bounds', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rectangle', { name: 'Box1', x: 10, y: 10, width: 50, height: 50 })
    await ppt.addShape('ellipse', { name: 'Circle1', x: 500, y: 500, width: 100, height: 100 })

    const shapesOnly = ppt.getElements(1, { type: 'shape' })
    expect(shapesOnly).toHaveLength(2)

    const withinFilter = ppt.getElements(1, {
      within: { x: 0, y: 0, width: 100, height: 100 },
    })
    const foundBox = withinFilter.find(e => e.name === 'Box1')
    expect(foundBox).toBeDefined()
    const foundCircle = withinFilter.find(e => e.name === 'Circle1')
    expect(foundCircle).toBeUndefined()
  })

  it('should safely remove an element and unlink orphaned relationship references', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    // 1x1 transparent PNG buffer
    const pngBuffer = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
      'base64'
    )

    await ppt.addImage(1, pngBuffer, {
      name: 'DisposableImage',
      x: 50,
      y: 50,
      width: 100,
      height: 100,
    })

    // Confirm element exists and has an image type
    const beforeRemove = ppt.getElement('DisposableImage')
    expect(beforeRemove).not.toBeNull()
    expect(beforeRemove.type).toBe('image')

    // Inspect slide rels before removal
    const bufBefore = await ppt.toBuffer()
    const zipBefore = await JSZip.loadAsync(bufBefore)
    const relsXmlBefore = await zipBefore.files['ppt/slides/_rels/slide1.xml.rels'].async('text')
    expect(relsXmlBefore).toContain('image')

    // Remove the element
    ppt.useSlide(1).removeElement('DisposableImage')

    // Confirm element is gone
    const afterRemove = ppt.getElement('DisposableImage')
    expect(afterRemove).toBeNull()

    // Inspect slide rels after removal: orphaned image relationship must be purged
    const bufAfter = await ppt.toBuffer()
    const zipAfter = await JSZip.loadAsync(bufAfter)
    const relsXmlAfter = await zipAfter.files['ppt/slides/_rels/slide1.xml.rels'].async('text')
    expect(relsXmlAfter).not.toContain('image')
    // Slide layout relationship must still be intact!
    expect(relsXmlAfter).toContain('slideLayout')
  })

  it('should remove all elements on a slide while strictly preserving group headers and layout bindings', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rectangle', { name: 'ElemA', x: 20, y: 20, width: 100, height: 100 })
    await ppt.addShape('ellipse', { name: 'ElemB', x: 150, y: 20, width: 100, height: 100 })

    expect(ppt.getElements(1)).toHaveLength(2)

    // Clear all elements
    ppt.removeAllElements(1)

    const remaining = ppt.getElements(1)
    expect(remaining).toHaveLength(0)

    // Verify OpenXML package integrity
    const buf = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buf)

    // Slide XML must preserve <p:nvGrpSpPr> and <p:grpSpPr>
    const slideXml = await zip.files['ppt/slides/slide1.xml'].async('text')
    expect(slideXml).toContain('<p:nvGrpSpPr>')
    expect(slideXml).toContain('</p:nvGrpSpPr>')
    expect(slideXml).toContain('<p:grpSpPr>')
    expect(slideXml).toContain('</p:grpSpPr>')

    // Must NOT contain removed shapes
    expect(slideXml).not.toContain('ElemA')
    expect(slideXml).not.toContain('ElemB')

    // Slide rels must preserve slideLayout
    const relsXml = await zip.files['ppt/slides/_rels/slide1.xml.rels'].async('text')
    expect(relsXml).toContain('slideLayout')
  })

  it('should support removeAllElements with preservePlaceholders option', async () => {
    const ppt = await PPTXTemplater.createBlank()

    // addSlide creates a slide with a Title placeholder
    ppt.addSlide({ title: 'Section Header Slide' })
    ppt.useSlide(2)

    await ppt.addShape('rectangle', { name: 'UserShape1', x: 50, y: 50, width: 200, height: 100 })
    expect(ppt.getElements(2).length).toBe(2)

    ppt.removeAllElements(2, { preservePlaceholders: true })

    const remaining = ppt.getElements(2)
    // Only placeholder element(s) should remain
    expect(remaining).toHaveLength(1)
    expect(remaining[0].name).toContain('Title')
    expect(remaining.some(e => e.name === 'UserShape1')).toBe(false)
  })
})
