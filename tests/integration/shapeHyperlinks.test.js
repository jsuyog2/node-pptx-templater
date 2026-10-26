import { describe, it, expect } from 'vitest'
import { PPTXTemplater, PPTXError } from '../../src/index.js'
import JSZip from 'jszip'

describe('Shape and Target-Level Hyperlinks Integration Tests (TASK-HYPERLINK-002)', () => {
  it('1. should add an external hyperlink to a shape with a tooltip', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'ActionBtn',
      name: 'ActionBtn',
      x: 100,
      y: 100,
      width: 200,
      height: 50,
      text: 'Click to Learn More',
    })

    ppt.addHyperlink('ActionBtn', 'https://github.com/action', {
      tooltip: 'Go to Action Repository',
    })

    const links = ppt.getHyperlinks(1)
    const btnLink = links.find(l => l.shapeName === 'ActionBtn' || l.targetId === 'ActionBtn')
    expect(btnLink).toBeDefined()
    expect(btnLink.url).toBe('https://github.com/action')
    expect(btnLink.tooltip).toBe('Go to Action Repository')
    expect(btnLink.isInternal).toBe(false)
    expect(btnLink.type).toBe('shape')

    // Verify raw OpenXML structures inside zip
    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')
    const relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')

    expect(slideXml).toContain('<a:hlinkClick')
    expect(slideXml).toContain('tooltip="Go to Action Repository"')
    expect(relsXml).toContain('Target="https://github.com/action"')
    expect(relsXml).toContain('TargetMode="External"')
  })

  it('2. should add internal slide jump hyperlinks with ppaction://hlinksldjump', async () => {
    const ppt = await PPTXTemplater.createBlank()
    // Add slide 2 and slide 3
    ppt.addSlide()
    ppt.addSlide()
    expect(ppt.slideCount).toBe(3)

    ppt.useSlide(1)
    await ppt.addShape('rect', {
      id: 'JumpToSlide3Btn',
      name: 'JumpToSlide3Btn',
      x: 50,
      y: 50,
      width: 150,
      height: 40,
      text: 'Jump to Slide 3',
    })

    ppt.addHyperlink('JumpToSlide3Btn', 3, { tooltip: 'Jump to Slide 3' })

    const links = ppt.getHyperlinks(1)
    const jumpLink = links.find(
      l => l.shapeName === 'JumpToSlide3Btn' || l.targetId === 'JumpToSlide3Btn'
    )
    expect(jumpLink).toBeDefined()
    expect(jumpLink.isInternal).toBe(true)
    expect(jumpLink.targetSlide).toBe(3)
    expect(jumpLink.tooltip).toBe('Jump to Slide 3')

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')
    const relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')

    expect(slideXml).toContain('action="ppaction://hlinksldjump"')
    expect(relsXml).toContain('Target="slide3.xml"')
    expect(relsXml).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"'
    )
  })

  it('3. should add slideshow navigation actions (next, prev, first, last)', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'NextArrow',
      name: 'NextArrow',
      x: 300,
      y: 300,
      width: 80,
      height: 40,
    })

    ppt.addHyperlink('NextArrow', 'next', { action: 'next', tooltip: 'Next Slide' })

    const links = ppt.getHyperlinks(1)
    const navLink = links.find(l => l.shapeName === 'NextArrow' || l.targetId === 'NextArrow')
    expect(navLink).toBeDefined()
    expect(navLink.isInternal).toBe(true)
    expect(navLink.action).toBe('ppaction://hlinkshowjump?s=nextslide')
    expect(navLink.tooltip).toBe('Next Slide')

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')

    expect(slideXml).toContain('action="ppaction://hlinkshowjump?s=nextslide"')
  })

  it('4. should update an existing hyperlink URL and tooltip via updateHyperlink', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'PortalBtn',
      name: 'PortalBtn',
      x: 100,
      y: 200,
      width: 200,
      height: 50,
      text: 'Visit Portal',
    })

    ppt.addHyperlink('PortalBtn', 'https://old-portal.com', { tooltip: 'Old Portal' })
    let links = ppt.getHyperlinks(1)
    expect(links.find(l => l.shapeName === 'PortalBtn').url).toBe('https://old-portal.com')

    // Update URL and tooltip
    ppt.updateHyperlink('PortalBtn', 'https://new-portal.com/dashboard', {
      tooltip: 'New Dashboard',
    })

    links = ppt.getHyperlinks(1)
    const updated = links.find(l => l.shapeName === 'PortalBtn')
    expect(updated).toBeDefined()
    expect(updated.url).toBe('https://new-portal.com/dashboard')
    expect(updated.tooltip).toBe('New Dashboard')

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')
    expect(relsXml).toContain('Target="https://new-portal.com/dashboard"')
    expect(relsXml).not.toContain('https://old-portal.com')
  })

  it('5. should remove a hyperlink and purge orphaned relationship from .rels', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'DiscardBtn',
      name: 'DiscardBtn',
      x: 150,
      y: 150,
      width: 180,
      height: 45,
      text: 'Temporary Link',
    })

    ppt.addHyperlink('DiscardBtn', 'https://temporary-link.org', { tooltip: 'Temp' })
    let links = ppt.getHyperlinks(1)
    expect(links.some(l => l.shapeName === 'DiscardBtn')).toBe(true)

    // Remove the hyperlink
    ppt.removeHyperlink('DiscardBtn')

    links = ppt.getHyperlinks(1)
    expect(links.some(l => l.shapeName === 'DiscardBtn')).toBe(false)

    // Verify .rels no longer has the orphaned URL
    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')
    expect(relsXml).not.toContain('temporary-link.org')
  })

  it('6. should reject unsafe javascript: URLs with PPTXError', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'MaliciousBtn',
      name: 'MaliciousBtn',
      x: 100,
      y: 100,
      width: 100,
      height: 50,
    })

    expect(() => {
      ppt.addHyperlink('MaliciousBtn', 'javascript:alert(1)')
    }).toThrow(PPTXError)
  })
})
