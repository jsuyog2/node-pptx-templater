import { describe, it, expect } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'
import JSZip from 'jszip'

describe('Hyperlink OpenXML Schema Compliance & Repair Prevention Tests (TASK-ACTION-003 & TASK-ACTION-004)', () => {
  it('1. should enforce strict ECMA-376 child element sequence within <a:rPr>', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'StyledBox',
      name: 'StyledBox',
      x: 100,
      y: 100,
      width: 400,
      height: 60,
      text: 'Explore our Architecture and Documentation right here.',
    })

    // Apply hyperlink with custom typography (color, fontFace, underline, bold)
    ppt.setHyperlink('StyledBox', 'Architecture', 'https://example.com/arch', {
      color: '#10B981',
      fontFace: 'Arial',
      bold: true,
      underline: true,
      tooltip: 'Architecture Docs',
    })

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')

    // Find the <a:rPr> containing the hyperlink
    const rPrMatch = slideXml.match(/<a:rPr[^>]*>[\s\S]*?<\/a:rPr>/g)
    const targetRPr = rPrMatch.find(r => r.includes('Architecture Docs') || r.includes('10B981'))
    expect(targetRPr).toBeDefined()

    // ECMA-376 sequence: solidFill -> latin -> hlinkClick
    const solidFillIdx = targetRPr.indexOf('<a:solidFill')
    const latinIdx = targetRPr.indexOf('<a:latin')
    const hlinkClickIdx = targetRPr.indexOf('<a:hlinkClick')

    expect(solidFillIdx).toBeGreaterThan(-1)
    expect(latinIdx).toBeGreaterThan(-1)
    expect(hlinkClickIdx).toBeGreaterThan(-1)

    // Verify ordering: solidFill MUST appear before latin, and latin MUST appear before hlinkClick
    expect(solidFillIdx).toBeLessThan(latinIdx)
    expect(latinIdx).toBeLessThan(hlinkClickIdx)

    // Ensure NO redundant xmlns:r is generated on a:hlinkClick
    expect(targetRPr).not.toContain('xmlns:r=')
  })

  it('2. should enforce correct child ordering inside <p:cNvPr> with hlinkClick before extLst', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'ActionContainer',
      name: 'ActionContainer',
      x: 200,
      y: 200,
      width: 150,
      height: 50,
      text: 'Click Target',
    })

    ppt.addHyperlink('ActionContainer', 'nextSlide')

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')

    const cNvPrMatches = slideXml.match(/<p:cNvPr[^>]*>[\s\S]*?<\/p:cNvPr>/g) || []
    const targetCnvPr = cNvPrMatches.find(c => c.includes('ActionContainer'))
    expect(targetCnvPr).toBeDefined()

    const hlinkClickIdx = targetCnvPr.indexOf('<a:hlinkClick')
    const extLstIdx = targetCnvPr.indexOf('<a:extLst')

    expect(hlinkClickIdx).toBeGreaterThan(-1)
    if (extLstIdx !== -1) {
      expect(hlinkClickIdx).toBeLessThan(extLstIdx)
    }

    // Verify NO redundant xmlns:r in shape-level link
    expect(targetCnvPr).not.toContain('xmlns:r=')
  })

  it('3. should deduplicate slide and hyperlink relationships in slideN.xml.rels', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.addSlide() // Slide 2
    ppt.addSlide() // Slide 3
    ppt.useSlide(1)

    await ppt.addShape('rect', { id: 'Btn1', name: 'Btn1', x: 50, y: 50, width: 100, height: 40 })
    await ppt.addShape('rect', { id: 'Btn2', name: 'Btn2', x: 200, y: 50, width: 100, height: 40 })

    // Both buttons point to Slide 3
    ppt.addHyperlink('Btn1', 3)
    ppt.addHyperlink('Btn2', 3)

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')

    // Slide 3 relationship should only be listed once
    const slide3Matches = relsXml.match(/Target="slide3\.xml"/g) || []
    expect(slide3Matches.length).toBe(1)
  })

  it('4. should apply text styling to shape text when options include styling properties', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'StyledShapeBtn',
      name: 'StyledShapeBtn',
      x: 100,
      y: 100,
      width: 180,
      height: 50,
      text: 'Submit Order',
    })

    ppt.addHyperlink('StyledShapeBtn', 'nextSlide', {
      color: '#FFFFFF',
      bold: true,
      underline: true,
      tooltip: 'Go to payment',
    })

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')

    // Check that text inside the shape received color and underline
    expect(slideXml).toContain('val="FFFFFF"')
    expect(slideXml).toContain('u="sng"')
    expect(slideXml).toContain('b="1"')
  })

  it('5. should safely purge relationships only when no other element references them', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.addSlide() // Slide 2
    ppt.addSlide() // Slide 3
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'Share1',
      name: 'Share1',
      x: 50,
      y: 50,
      width: 100,
      height: 40,
    })
    await ppt.addShape('rect', {
      id: 'Share2',
      name: 'Share2',
      x: 200,
      y: 50,
      width: 100,
      height: 40,
    })

    ppt.addHyperlink('Share1', 3)
    ppt.addHyperlink('Share2', 3)

    // Remove link from Share1; Share2 still uses Slide 3
    ppt.removeHyperlink('Share1')

    let buffer = await ppt.toBuffer()
    let zip = await JSZip.loadAsync(buffer)
    let relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')
    expect(relsXml).toContain('Target="slide3.xml"')

    // Now remove link from Share2; now orphaned relationship should be purged
    ppt.removeHyperlink('Share2')

    buffer = await ppt.toBuffer()
    zip = await JSZip.loadAsync(buffer)
    relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')
    expect(relsXml).not.toContain('Target="slide3.xml"')
  })
})
