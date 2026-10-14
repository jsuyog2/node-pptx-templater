import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Typography, Text Formatting & Text Runs', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_typography.pptx')

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

  it('should create shapes with advanced typography and text styles', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape('rectangle', {
      id: 'advancedTextShape',
      x: 50,
      y: 50,
      width: 300,
      height: 150,
      text: 'Formatted Headline\nSubheading Text',
      textStyle: {
        fontSize: 18,
        bold: true,
        italic: true,
        underline: 'sng',
        strike: 'sngStrike',
        fontFamily: 'Helvetica',
        color: '#1A73E8',
        characterSpacing: 2,
        align: 'center',
        lineSpacing: 1.25,
        spaceBefore: 6,
        spaceAfter: 12,
        verticalAlign: 'middle',
        padding: { left: 10, top: 10, right: 10, bottom: 10 },
      },
    })

    const slideXml = ppt.slideManager.getSlideXml(1)
    expect(slideXml).toContain('u="sng"')
    expect(slideXml).toContain('strike="sngStrike"')
    expect(slideXml).toContain('typeface="Helvetica"')
    expect(slideXml).toContain('val="1A73E8"')
    expect(slideXml).toContain('spc="200"')
    expect(slideXml).toContain('algn="ctr"')
    expect(slideXml).toContain('anchor="ctr"')
    expect(slideXml).toContain('lIns="95250"')

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    const reloadedXml = reloaded.slideManager.getSlideXml(1)
    expect(reloadedXml).toContain('u="sng"')
    expect(reloadedXml).toContain('strike="sngStrike"')
  })

  it('should support subscript and superscript styling in shapes', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape('rectangle', {
      id: 'superSubShape',
      x: 100,
      y: 100,
      width: 200,
      height: 80,
      text: 'E = mc2',
      textStyle: {
        superscript: true,
      },
    })

    const slideXml = ppt.slideManager.getSlideXml(1)
    expect(slideXml).toContain('baseline="30000"')
  })

  it('should append and prepend formatted text runs preserving structure', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape('rectangle', {
      id: 'runTestShape',
      x: 50,
      y: 50,
      width: 250,
      height: 80,
      text: 'Original Text',
      textStyle: {
        fontSize: 14,
        color: '#333333',
      },
    })

    // Append a run with different styling
    ppt.appendTextRun('runTestShape', ' [Appended]', {
      style: {
        bold: true,
        color: '#00AA00',
      },
    })

    // Prepend a run
    ppt.prependTextRun('runTestShape', '[Prepended] ', {
      style: {
        italic: true,
        underline: true,
        color: '#AA0000',
      },
    })

    const slideXml = ppt.slideManager.getSlideXml(1)
    expect(slideXml).toContain('[Prepended] ')
    expect(slideXml).toContain('Original Text')
    expect(slideXml).toContain(' [Appended]')
    expect(slideXml).toContain('val="00AA00"')
    expect(slideXml).toContain('val="AA0000"')
    expect(slideXml).toContain('u="sng"')

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    const reloadedXml = reloaded.slideManager.getSlideXml(1)
    expect(reloadedXml).toContain('[Prepended] ')
    expect(reloadedXml).toContain(' [Appended]')
  })
})
