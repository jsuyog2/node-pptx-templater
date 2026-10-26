import { describe, it, expect } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'
import JSZip from 'jszip'

describe('Hyperlink Action Decoupling Integration Tests (TASK-ACTION-001 & TASK-ACTION-002)', () => {
  it('1. should allow passing preset actions directly without nesting inside options bag', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'BtnNext',
      name: 'BtnNext',
      x: 100,
      y: 100,
      width: 150,
      height: 50,
      text: 'Next',
    })

    await ppt.addShape('rect', {
      id: 'BtnPrev',
      name: 'BtnPrev',
      x: 300,
      y: 100,
      width: 150,
      height: 50,
      text: 'Previous',
    })

    // Decoupled action parameters: standalone strings
    ppt.addHyperlink('BtnNext', 'nextSlide')
    ppt.addHyperlink('BtnPrev', 'previousSlide')

    const links = ppt.getHyperlinks(1)
    const nextLink = links.find(l => l.targetId === 'BtnNext')
    const prevLink = links.find(l => l.targetId === 'BtnPrev')

    expect(nextLink).toBeDefined()
    expect(nextLink.isInternal).toBe(true)
    expect(nextLink.action).toBe('ppaction://hlinkshowjump?jump=nextslide')

    expect(prevLink).toBeDefined()
    expect(prevLink.isInternal).toBe(true)
    expect(prevLink.action).toBe('ppaction://hlinkshowjump?jump=previousslide')

    // Verify raw OpenXML
    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')
    const relsXml = (await zip.file('ppt/slides/_rels/slide1.xml.rels')?.async('text')) || ''

    expect(slideXml).toContain('action="ppaction://hlinkshowjump?jump=nextslide"')
    expect(slideXml).toContain('action="ppaction://hlinkshowjump?jump=previousslide"')
    // Showjump actions should have r:id="" and NOT create relationships in .rels
    expect(slideXml).toContain('r:id=""')
    expect(relsXml).not.toContain('ppaction')
  })

  it('2. should support all standard action preset strings (firstSlide, lastSlide, endShow)', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'FirstBtn',
      name: 'FirstBtn',
      x: 50,
      y: 50,
      width: 100,
      height: 40,
    })
    await ppt.addShape('rect', {
      id: 'LastBtn',
      name: 'LastBtn',
      x: 200,
      y: 50,
      width: 100,
      height: 40,
    })
    await ppt.addShape('rect', {
      id: 'EndBtn',
      name: 'EndBtn',
      x: 350,
      y: 50,
      width: 100,
      height: 40,
    })

    ppt.addHyperlink('FirstBtn', 'firstSlide')
    ppt.addHyperlink('LastBtn', 'lastSlide')
    ppt.addHyperlink('EndBtn', 'endShow')

    const links = ppt.getHyperlinks(1)
    expect(links.find(l => l.targetId === 'FirstBtn').action).toBe(
      'ppaction://hlinkshowjump?jump=firstslide'
    )
    expect(links.find(l => l.targetId === 'LastBtn').action).toBe(
      'ppaction://hlinkshowjump?jump=lastslide'
    )
    expect(links.find(l => l.targetId === 'EndBtn').action).toBe(
      'ppaction://hlinkshowjump?jump=endshow'
    )
  })

  it('3. should support direct slide number target without options bag', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.addSlide() // Slide 2
    ppt.addSlide() // Slide 3
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'JumpSlide3',
      name: 'JumpSlide3',
      x: 100,
      y: 100,
      width: 200,
      height: 50,
      text: 'Go to 3',
    })

    // Decoupled slide number passed directly as 2nd argument
    ppt.addHyperlink('JumpSlide3', 3)

    const links = ppt.getHyperlinks(1)
    const link = links.find(l => l.targetId === 'JumpSlide3')
    expect(link).toBeDefined()
    expect(link.isInternal).toBe(true)
    expect(link.targetSlide).toBe(3)

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')
    const relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')

    expect(slideXml).toContain('action="ppaction://hlinksldjump"')
    expect(relsXml).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"'
    )
    expect(relsXml).toContain('Target="slide3.xml"')
  })

  it('4. should maintain backward compatibility when action is nested inside options', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'LegacyBtn',
      name: 'LegacyBtn',
      x: 50,
      y: 50,
      width: 100,
      height: 40,
    })

    // Old API bag style: passed inside options or as 2nd arg options bag
    ppt.addHyperlink('LegacyBtn', { action: 'nextSlide', tooltip: 'Next Page' })

    const link = ppt.getHyperlink('LegacyBtn')
    expect(link).toBeDefined()
    expect(link.action).toBe('ppaction://hlinkshowjump?jump=nextslide')
    expect(link.tooltip).toBe('Next Page')
  })

  it('5. should allow decoupled action in setHyperlink and updateHyperlink', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'TextContainer',
      name: 'TextContainer',
      x: 50,
      y: 50,
      width: 400,
      height: 80,
      text: 'Click here to proceed to next slide now.',
    })

    // Substring link with direct action preset
    ppt.setHyperlink('TextContainer', 'proceed to next slide', 'nextSlide', {
      color: '#2563EB',
      underline: true,
      bold: true,
    })

    const links = ppt.getHyperlinks(1)
    const textLink = links.find(l => l.type === 'textRun')
    expect(textLink).toBeDefined()
    expect(textLink.text).toBe('proceed to next slide')
    expect(textLink.action).toBe('ppaction://hlinkshowjump?jump=nextslide')

    // Update hyperlink directly to previousSlide
    ppt.updateHyperlink('TextContainer', 'previousSlide')
    const updatedShapeLink = ppt.getHyperlink('TextContainer')
    expect(updatedShapeLink).toBeDefined()
    expect(updatedShapeLink.action).toBe('ppaction://hlinkshowjump?jump=previousslide')
  })
})
