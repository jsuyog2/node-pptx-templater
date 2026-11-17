import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('TASK-TEXT-003: Core Text & Placeholder Methods Integration Tests', () => {
  const outputPath = resolve(__dirname, '../fixtures/output_text_operations.pptx')

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

  it('1. addText: supports pixel, EMU, and percentage string coordinates', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    // Percentage coordinates (e.g. 10% x, 20% y, 50% width)
    ppt.addText('Percentage Positioned Text', {
      id: 'PctShape',
      x: '10%',
      y: '20%',
      width: '50%',
      height: 80,
      fontSize: 18,
      color: '#059669',
    })

    // Pixel coordinates
    ppt.addText('Pixel Positioned Text', {
      id: 'PxShape',
      x: 100,
      y: 200,
      width: 300,
      height: 50,
      fontSize: 14,
    })

    const xml = ppt.slideManager.getSlideXml(1)
    // 10% of 12192000 = 1219200
    expect(xml).toContain('x="1219200"')
    // 20% of 6858000 = 1371600
    expect(xml).toContain('y="1371600"')
    // 50% of 12192000 = 6096000
    expect(xml).toContain('cx="6096000"')
    expect(xml).toContain('Percentage Positioned Text')
    expect(xml).toContain('Pixel Positioned Text')
  })

  it('2. replaceText & replaceTextByTag: replaces placeholders and full text box content', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('Initial Header Content', {
      id: 'ClientBox',
      x: 50,
      y: 50,
      width: 400,
      height: 60,
    })

    ppt.addText('Report for {{client}} on {{date}}', {
      id: 'DateBox',
      x: 50,
      y: 120,
      width: 400,
      height: 60,
    })

    // replaceTextByTag targets text box by ID/name and replaces full content
    ppt.replaceTextByTag('ClientBox', 'Acme Global Enterprises')
    let xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('Acme Global Enterprises')
    expect(xml).not.toContain('Initial Header Content')

    // replaceText replaces template placeholder keys across text runs
    ppt.replaceText({ '{{client}}': 'Wayne Enterprises', '{{date}}': 'October 2026' })
    xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('Report for Wayne Enterprises on October 2026')
    expect(xml).not.toContain('{{date}}')
    expect(xml).not.toContain('{{client}}')
  })

  it('3. replaceText & findText: performs AST run healing across fragmented runs', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    // Manually inject a shape with fragmented text runs split in XML
    const fragmentedSlideXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="200" name="FragTestBox"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="3000000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:lstStyle/>
          <a:p>
            <a:r><a:rPr b="1"/><a:t>{{Frag</a:t></a:r>
            <a:r><a:rPr b="1"/><a:t>mented</a:t></a:r>
            <a:r><a:rPr b="1"/><a:t>Key}}</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`

    ppt.slideManager.setSlideXml(1, fragmentedSlideXml)

    // findText should heal runs and find 'FragmentedKey'
    const matches = ppt.findText('FragmentedKey')
    expect(matches.length).toBe(1)
    expect(matches[0].shapeName).toBe('FragTestBox')
    expect(matches[0].text).toContain('{{FragmentedKey}}')

    // replaceText should heal fragmented runs and replace seamlessly
    ppt.replaceText({ '{{FragmentedKey}}': 'Healed AST Run Success' })
    const updatedXml = ppt.slideManager.getSlideXml(1)
    expect(updatedXml).toContain('Healed AST Run Success')
    expect(updatedXml).not.toContain('{{Frag')
    expect(updatedXml).not.toContain('mented')
    expect(updatedXml).not.toContain('Key}}')
  })

  it('4. findText: respects caseSensitive option', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('Case Sensitivity Test TARGET and target', {
      id: 'CaseBox',
      x: 100,
      y: 100,
      width: 400,
      height: 80,
    })

    // Case-insensitive by default
    const insensitiveMatches = ppt.findText('target')
    expect(insensitiveMatches.length).toBeGreaterThan(0)

    // Case-sensitive search
    const sensitiveUpper = ppt.findText('TARGET', { caseSensitive: true })
    expect(sensitiveUpper.length).toBe(1)

    const sensitiveLower = ppt.findText('target', { caseSensitive: true })
    expect(sensitiveLower.length).toBe(1)

    const sensitiveNotFound = ppt.findText('TaRgEt', { caseSensitive: true })
    expect(sensitiveNotFound.length).toBe(0)
  })

  it('5. getText: extracts plain text content by shape name or tag', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('Hello World Plain Text Content', {
      id: 'TargetBox',
      x: 50,
      y: 50,
      width: 300,
      height: 60,
    })

    const text = ppt.getText('TargetBox')
    expect(text).toBe('Hello World Plain Text Content')
  })

  it('6. clearText: clears text runs while retaining shape container and geometry coordinates', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('Text To Be Cleared', {
      id: 'ClearBox',
      x: 500000,
      y: 600000,
      width: 2500000,
      height: 400000,
    })

    let xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('Text To Be Cleared')
    expect(xml).toContain('x="500000"')
    expect(xml).toContain('y="600000"')

    // Clear text
    ppt.clearText('ClearBox')
    xml = ppt.slideManager.getSlideXml(1)

    // Text must be purged
    expect(xml).not.toContain('Text To Be Cleared')
    // Shape container and geometry coordinates must remain 100% intact
    expect(xml).toContain('ClearBox')
    expect(xml).toContain('x="500000"')
    expect(xml).toContain('y="600000"')
    expect(xml).toContain('cx="2500000"')
    expect(xml).toContain('cy="400000"')
    expect(xml).toContain('prst="rect"')
  })

  it('7. removeText: removes shape matching tag from slide completely', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('Removable Shape Text', {
      id: 'DeleteMe',
      x: 100,
      y: 100,
      width: 200,
      height: 50,
    })

    let xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('DeleteMe')

    ppt.removeText('DeleteMe')
    xml = ppt.slideManager.getSlideXml(1)
    expect(xml).not.toContain('DeleteMe')
    expect(xml).not.toContain('Removable Shape Text')
  })

  it('8. getTextElements: discovers all text elements across shapes on slide', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('First Box Text', { id: 'Box1', x: 50, y: 50, width: 200, height: 40 })
    ppt.addText('Second Box Text', { id: 'Box2', x: 50, y: 150, width: 200, height: 40 })

    const elements = ppt.getTextElements()
    expect(elements.length).toBeGreaterThanOrEqual(2)

    const box1 = elements.find(el => el.shapeName === 'Box1')
    expect(box1).toBeDefined()
    expect(box1.text).toBe('First Box Text')

    const box2 = elements.find(el => el.shapeName === 'Box2')
    expect(box2).toBeDefined()
    expect(box2.text).toBe('Second Box Text')

    // Persist and verify round-trip integrity
    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    const reloadedElements = reloaded.getTextElements()
    expect(reloadedElements.length).toBeGreaterThanOrEqual(2)
  })
})
