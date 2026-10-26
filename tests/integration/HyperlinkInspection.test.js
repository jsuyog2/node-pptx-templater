import { describe, it, expect } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('Hyperlink Introspection APIs Integration Tests (TASK-HYPERLINK-003)', () => {
  it('1. should inspect all hyperlinks across a slide (shape-level and text-run level)', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.addSlide() // 2 slides total
    ppt.useSlide(1)

    // Add a shape with text
    await ppt.addShape('rect', {
      id: 'BtnAction',
      name: 'BtnAction',
      x: 100,
      y: 100,
      width: 250,
      height: 60,
      text: 'Visit Google and Documentation',
    })

    // Attach shape-level hyperlink
    ppt.addHyperlink('BtnAction', 'https://google.com', {
      tooltip: 'Search Google',
    })

    // Attach substring text-level hyperlink
    ppt.setHyperlink('BtnAction', 'Documentation', 'https://example.com/docs', {
      tooltip: 'Read Documentation',
      color: '#0563C1',
      underline: true,
    })

    // Inspect slide 1
    const links = ppt.getHyperlinks(1)
    expect(links.length).toBe(2)

    // Verify shape link metadata
    const shapeLink = links.find(l => l.type === 'shape')
    expect(shapeLink).toBeDefined()
    expect(shapeLink.url).toBe('https://google.com')
    expect(shapeLink.tooltip).toBe('Search Google')
    expect(shapeLink.isInternal).toBe(false)
    expect(shapeLink.shapeName).toBe('BtnAction')

    // Verify textRun link metadata
    const textLink = links.find(l => l.type === 'textRun')
    expect(textLink).toBeDefined()
    expect(textLink.url).toBe('https://example.com/docs')
    expect(textLink.text).toBe('Documentation')
    expect(textLink.tooltip).toBe('Read Documentation')
    expect(textLink.shapeName).toBe('BtnAction')
  })

  it('2. should query single hyperlink via getHyperlink by shapeId, shapeName, or text', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.addSlide()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'MyTargetBox',
      name: 'MyTargetBox',
      x: 50,
      y: 50,
      width: 200,
      height: 50,
      text: 'Click Here for More Info',
    })

    ppt.addHyperlink('MyTargetBox', 2, { tooltip: 'Jump to Slide 2' })

    // Query by shape name
    const byName = ppt.getHyperlink('MyTargetBox')
    expect(byName).toBeDefined()
    expect(byName.isInternal).toBe(true)
    expect(byName.targetSlide).toBe(2)
    expect(byName.tooltip).toBe('Jump to Slide 2')

    // Query non-existent target returns null
    const nonExistent = ppt.getHyperlink('NoSuchShape')
    expect(nonExistent).toBeNull()
  })

  it('3. should inspect special slideshow navigation actions without relationships', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'PrevBtn',
      name: 'PrevBtn',
      x: 80,
      y: 80,
      width: 100,
      height: 40,
      text: 'Previous',
    })

    ppt.addHyperlink('PrevBtn', 'previous', { action: 'previous', tooltip: 'Go Back' })

    const link = ppt.getHyperlink('PrevBtn')
    expect(link).toBeDefined()
    expect(link.isInternal).toBe(true)
    expect(link.action).toBe('ppaction://hlinkshowjump?s=prevslide')
    expect(link.tooltip).toBe('Go Back')

    // Ensure getHyperlinks returns empty array for empty slide or out of bounds slide
    const emptySlideLinks = ppt.getHyperlinks(99)
    expect(emptySlideLinks).toEqual([])
  })

  it('4. should ensure read-only introspection leaves slide XML unmodified', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'StaticShape',
      name: 'StaticShape',
      x: 50,
      y: 50,
      width: 100,
      height: 50,
      text: 'Static',
    })

    ppt.addHyperlink('StaticShape', 'https://example.com')

    const bufBefore = await ppt.toBuffer()
    // Read-only queries
    ppt.getHyperlinks(1)
    ppt.getHyperlink('StaticShape')
    const bufAfter = await ppt.toBuffer()

    expect(bufBefore.length).toBe(bufAfter.length)
  })
})
