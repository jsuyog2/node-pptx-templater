import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import { PPTXTemplater, ValidationEngine } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('TASK-REPLACE-TAG-002: replaceTextByTag Full Text Box Replacement Integration Tests', () => {
  const outputPath = resolve(__dirname, '../fixtures/output_replace_text_by_tag.pptx')

  beforeEach(() => {
    if (existsSync(outputPath)) {
      try {
        unlinkSync(outputPath)
      } catch (_) {}
    }
  })

  afterEach(() => {
    if (existsSync(outputPath)) {
      try {
        unlinkSync(outputPath)
      } catch (_) {}
    }
  })

  it('1. should replace full text in a text box targeted by shape name', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('Original draft text that will be replaced', {
      id: 'HeroTitleBox',
      x: 100,
      y: 100,
      width: 500,
      height: 80,
    })

    const initialXml = ppt.slideManager.getSlideXml(1)
    expect(initialXml).toContain('Original draft text that will be replaced')

    // Call replaceTextByTag targeting by shape name
    const result = ppt.replaceTextByTag('HeroTitleBox', 'Final Executive Presentation Title')
    expect(result).toBe(ppt) // Fluent chaining

    const updatedXml = ppt.slideManager.getSlideXml(1)
    expect(updatedXml).toContain('Final Executive Presentation Title')
    expect(updatedXml).not.toContain('Original draft text that will be replaced')

    const elements = ppt.getTextElements()
    const targetEl = elements.find(el => el.shapeName === 'HeroTitleBox')
    expect(targetEl).toBeDefined()
    expect(targetEl.text).toBe('Final Executive Presentation Title')
  })

  it('2. should replace full text in a text box targeted by numeric shape ID string', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('First shape content', {
      id: 'FirstBox',
      x: 50,
      y: 50,
      width: 300,
      height: 50,
    })

    const slideObj = ppt.slideManager.getSlideObj(1)
    const shapes = slideObj['p:sld']['p:cSld']['p:spTree']['p:sp']
    const shapeList = Array.isArray(shapes) ? shapes : [shapes]
    const createdShape = shapeList.find(s => s['p:nvSpPr']?.['p:cNvPr']?.['@_name'] === 'FirstBox')
    expect(createdShape).toBeDefined()
    const shapeId = String(createdShape['p:nvSpPr']['p:cNvPr']['@_id'])

    // Target via numeric ID string
    ppt.replaceTextByTag(shapeId, 'Completely Replaced via Numeric ID')

    const updatedXml = ppt.slideManager.getSlideXml(1)
    expect(updatedXml).toContain('Completely Replaced via Numeric ID')
    expect(updatedXml).not.toContain('First shape content')
  })

  it('3. should handle multi-line string replacement by splitting into separate paragraphs', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('Single line text', {
      id: 'MultiLineBox',
      x: 80,
      y: 80,
      width: 400,
      height: 120,
    })

    const multiLineContent =
      'Line 1: Strategic Planning\nLine 2: Operational Execution\nLine 3: Financial Settlement'
    ppt.replaceTextByTag('MultiLineBox', multiLineContent)

    const slideObj = ppt.slideManager.getSlideObj(1)
    const shapes = slideObj['p:sld']['p:cSld']['p:spTree']['p:sp']
    const shapeList = Array.isArray(shapes) ? shapes : [shapes]
    const targetShape = shapeList.find(
      s => s['p:nvSpPr']?.['p:cNvPr']?.['@_name'] === 'MultiLineBox'
    )

    expect(targetShape).toBeDefined()
    const paras = Array.isArray(targetShape['p:txBody']['a:p'])
      ? targetShape['p:txBody']['a:p']
      : [targetShape['p:txBody']['a:p']]

    expect(paras.length).toBe(3)
    expect(paras[0]['a:r'][0]['a:t']).toBe('Line 1: Strategic Planning')
    expect(paras[1]['a:r'][0]['a:t']).toBe('Line 2: Operational Execution')
    expect(paras[2]['a:r'][0]['a:t']).toBe('Line 3: Financial Settlement')
  })

  it('4. should preserve original typography and paragraph formatting', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    // Create a styled text box
    ppt.addText('Initial Styled Content', {
      id: 'StyledBox',
      x: 100,
      y: 100,
      width: 450,
      height: 70,
      fontSize: 24,
      bold: true,
      italic: true,
      color: '#1E40AF',
      alignment: 'center',
    })

    ppt.replaceTextByTag('StyledBox', 'Preserved Style Title')

    const xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('Preserved Style Title')
    expect(xml).not.toContain('Initial Styled Content')

    // Verify font size in hundredths of a point (24pt = 2400)
    expect(xml).toContain('sz="2400"')
    // Verify bold and italic attributes
    expect(xml).toContain('b="1"')
    expect(xml).toContain('i="1"')
    // Verify color hex
    expect(xml).toContain('1E40AF')
    // Verify paragraph alignment
    expect(xml).toContain('algn="ctr"')
  })

  it('5. should support optional style overrides during replacement', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('Initial Text', {
      id: 'OverrideBox',
      x: 100,
      y: 100,
      width: 400,
      height: 60,
      fontSize: 14,
      color: '#64748B',
    })

    // Override styling to red 28pt bold
    ppt.replaceTextByTag('OverrideBox', 'Important Overridden Header', {
      fontSize: 28,
      bold: true,
      color: '#DC2626',
    })

    const xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('Important Overridden Header')
    expect(xml).toContain('sz="2800"')
    expect(xml).toContain('b="1"')
    expect(xml).toContain('DC2626')
  })

  it('6. should gracefully skip and log a warning if target ID is not found', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('Existing text', {
      id: 'ExistingBox',
      x: 50,
      y: 50,
      width: 200,
      height: 50,
    })

    // Should not throw
    expect(() => {
      const res = ppt.replaceTextByTag('NonExistent_Shape_9999', 'Ghost Text')
      expect(res).toBe(ppt)
    }).not.toThrow()

    const xml = ppt.slideManager.getSlideXml(1)
    expect(xml).not.toContain('Ghost Text')
    expect(xml).toContain('Existing text')
  })

  it('7. should respect options.slide to target a specific slide', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)
    ppt.addText('Slide 1 Box', { id: 'TargetBox', x: 50, y: 50, width: 200, height: 50 })

    await ppt.addSlide({ title: 'Slide 2' })
    ppt.useSlide(2)
    ppt.addText('Slide 2 Box', { id: 'TargetBox', x: 50, y: 50, width: 200, height: 50 })

    // Replace on slide 2 explicitly via options.slide
    ppt.replaceTextByTag('TargetBox', 'Updated Slide 2 Only', { slide: 2 })

    const slide1Xml = ppt.slideManager.getSlideXml(1)
    const slide2Xml = ppt.slideManager.getSlideXml(2)

    expect(slide1Xml).toContain('Slide 1 Box')
    expect(slide1Xml).not.toContain('Updated Slide 2 Only')
    expect(slide2Xml).toContain('Updated Slide 2 Only')
    expect(slide2Xml).not.toContain('Slide 2 Box')
  })

  it('8. should generate valid PPTX package passing presentation validation', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('Original Box', { id: 'Box1', x: 50, y: 50, width: 300, height: 50 })
    ppt.replaceTextByTag('Box1', 'Validated Content Line 1\nValidated Content Line 2')

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const report = await ValidationEngine.validatePresentation(ppt)
    expect(report.valid).toBe(true)
    expect(report.errors).toHaveLength(0)
  })
})
