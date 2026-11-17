import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'
import { normalizeHexColor, applyTextCasing } from '../../src/parser/StyleBuilder.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('TASK-TEXT-002: Comprehensive 31-Property OpenXML Style Engine Integration Tests', () => {
  const outputPath = resolve(__dirname, '../fixtures/output_styling_31.pptx')

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

  it('1. StyleBuilder unit checks for helpers and casing transformations', () => {
    expect(normalizeHexColor('#1a73e8')).toBe('1A73E8')
    expect(normalizeHexColor('abc')).toBe('AABBCC')
    expect(normalizeHexColor('')).toBe('000000')

    expect(applyTextCasing('hello world', { uppercase: true })).toBe('HELLO WORLD')
    expect(applyTextCasing('HELLO WORLD', { lowercase: true })).toBe('hello world')
    expect(applyTextCasing('hello world title', { capitalize: true })).toBe('Hello World Title')
  })

  it('2. Maps all Typography properties to <a:rPr> (fontSize, color, highlight, underline, bold, italic, superscript, subscript, strikethrough)', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    // Add text shape with typography properties
    ppt.addText('Typography Benchmark Text', {
      x: 50,
      y: 50,
      width: 400,
      height: 80,
      id: 'TypeShape',
      fontSize: 24,
      color: '#1A73E8',
      highlight: '#FFFF00',
      underline: true,
      bold: true,
      italic: true,
      superscript: true,
      strikethrough: true,
    })

    const xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('sz="2400"') // fontSize (24 * 100)
    expect(xml).toContain('val="1A73E8"') // color
    expect(xml).toContain('<a:highlight><a:srgbClr val="FFFF00"') // highlight
    expect(xml).toContain('u="sng"') // underline
    expect(xml).toContain('b="1"') // bold
    expect(xml).toContain('i="1"') // italic
    expect(xml).toContain('baseline="30000"') // superscript
    expect(xml).toContain('strike="sngStrike"') // strikethrough

    // Test subscript
    ppt.addText('Chemical Formula H2O', {
      x: 50,
      y: 150,
      width: 300,
      height: 60,
      id: 'SubShape',
      subscript: true,
    })
    const subXml = ppt.slideManager.getSlideXml(1)
    expect(subXml).toContain('baseline="-25000"')
  })

  it('3. Maps Casing & Spacing properties (uppercase, lowercase, capitalize, smallCaps, charSpacing, lineSpacing)', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('casing and spacing demo', {
      x: 50,
      y: 50,
      width: 500,
      height: 100,
      id: 'CasingShape',
      smallCaps: true,
      charSpacing: 2.5,
      lineSpacing: 1.5,
      capitalize: true,
    })

    const xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('cap="small"')
    expect(xml).toContain('spc="250"') // 2.5 * 100
    expect(xml).toContain('spcPct val="150000"') // 1.5 lines = 150000
    expect(xml).toContain('Casing And Spacing Demo')
  })

  it('4. Maps Alignment & Direction properties (textAlignment, alignment, indentLevel, verticalAlignment, textDirection)', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('Alignment & Direction Demo', {
      x: 100,
      y: 100,
      width: 300,
      height: 120,
      id: 'AlignShape',
      alignment: 'center',
      indentLevel: 2,
      verticalAlignment: 'middle',
      textDirection: 'vertical',
    })

    const xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('algn="ctr"')
    expect(xml).toContain('lvl="2"')
    expect(xml).toContain('anchor="ctr"')
    expect(xml).toContain('vert="vert"')
  })

  it('5. Maps Bullets properties (bulletChar, bulletColor, bulletFontSize, bulletStyle, bulletLevel)', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addLists(['Custom Bullet Alpha', 'Custom Bullet Beta'], {
      x: 80,
      y: 80,
      width: 400,
      height: 150,
      id: 'BulletShape',
      bulletChar: '★',
      bulletColor: '#00AAFF',
      bulletFontSize: 120, // percentage
      bulletLevel: 1,
    })

    const xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('char="★"')
    expect(xml).toContain('val="00AAFF"')
    expect(xml).toContain('buSzPct')
    expect(xml).toContain('val="120000"')
    expect(xml).toContain('lvl="1"')

    // Ordered / numbered list
    ppt.addLists(['Step 1', 'Step 2'], {
      x: 80,
      y: 250,
      width: 300,
      height: 100,
      id: 'NumShape',
      bulletStyle: 'ordered',
    })

    const numXml = ppt.slideManager.getSlideXml(1)
    expect(numXml).toContain('buAutoNum')
    expect(numXml).toContain('type="arabicPeriod"')
  })

  it('6. Maps Shadows & Effects properties (shadowColor, textShadow, blur, shadowRadius, shadowOffsetX, shadowOffsetY)', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('Shadow Effect Text', {
      x: 60,
      y: 60,
      width: 400,
      height: 80,
      id: 'ShadowShape',
      textShadow: true,
      shadowColor: '#333333',
      blur: 4,
      shadowOffsetX: 3,
      shadowOffsetY: 3,
    })

    const xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('<a:effectLst>')
    expect(xml).toContain('<a:outerShdw')
    expect(xml).toContain('val="333333"')
    expect(xml).toContain('blurRad=')
    expect(xml).toContain('dist=')
  })

  it('7. Modifies styles non-destructively using addStyle, updateStyle, and removeStyle', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('Base Styled Headline', {
      id: 'Headline',
      x: 50,
      y: 50,
      width: 400,
      height: 80,
      fontSize: 20,
      bold: true,
      color: '#112233',
    })

    let xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('sz="2000"')
    expect(xml).toContain('b="1"')

    // Add / Update styles
    ppt.addStyle('Headline', {
      italic: true,
      underline: true,
      color: '#FF0055',
    })

    xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('i="1"')
    expect(xml).toContain('u="sng"')
    expect(xml).toContain('val="FF0055"')
    // original text content must remain intact
    expect(xml).toContain('Base Styled Headline')

    // Remove styles
    ppt.removeStyle('Headline', ['bold', 'italic', 'underline'])
    xml = ppt.slideManager.getSlideXml(1)
    expect(xml).not.toContain('b="1"')
    expect(xml).not.toContain('i="1"')
    expect(xml).not.toContain('u="sng"')
    expect(xml).toContain('Base Styled Headline')
  })

  it('8. Successfully writes modified presentations with all 31 properties to disk and reloads cleanly', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addText('Full 31 Property Test Suite', {
      id: 'MasterShape',
      x: 100,
      y: 100,
      width: 500,
      height: 120,
      fontSize: 28,
      color: '#2563EB',
      highlight: '#FEF08A',
      underline: true,
      bold: true,
      italic: true,
      strikethrough: false,
      textDirection: 'horizontal',
      textAlignment: 'center',
      indentLevel: 0,
      lineSpacing: 1.25,
      uppercase: true,
      charSpacing: 1.5,
      textShadow: true,
      shadowColor: '#1E293B',
      blur: 2,
      shadowOffsetX: 2,
      shadowOffsetY: 2,
      verticalAlignment: 'middle',
    })

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    expect(reloaded.slideCount).toBe(1)
    const reloadedXml = reloaded.slideManager.getSlideXml(1)
    expect(reloadedXml).toContain('FULL 31 PROPERTY TEST SUITE')
    expect(reloadedXml).toContain('sz="2800"')
    expect(reloadedXml).toContain('val="2563EB"')
    expect(reloadedXml).toContain('val="FEF08A"')
  })
})
