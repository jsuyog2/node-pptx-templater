import { describe, it, expect, afterEach } from 'vitest'
import fs from 'fs'
import path from 'path'
import {
  PPTXTemplater,
  isFlatOpc,
  isPkZip,
} from '../../src/index.js'

describe('Flat OPC XML Ingestion & Manipulation Integration Tests', () => {
  const sampleXmlPath = path.resolve(__dirname, '../fixtures/sample.xml')
  const tempOutputDir = path.resolve(__dirname, '../output')
  const tempPptxPath = path.join(tempOutputDir, 'test-flat-opc-export.pptx')

  afterEach(() => {
    if (fs.existsSync(tempPptxPath)) {
      try {
        fs.unlinkSync(tempPptxPath)
      } catch {
        // ignore cleanup error
      }
    }
  })

  it('1. should detect Flat OPC XML format correctly with isFlatOpc()', () => {
    const rawXml = fs.readFileSync(sampleXmlPath, 'utf8')
    const rawBuffer = fs.readFileSync(sampleXmlPath)

    expect(isFlatOpc(rawXml)).toBe(true)
    expect(isFlatOpc(rawBuffer)).toBe(true)
    expect(isFlatOpc(new Uint8Array(rawBuffer))).toBe(true)
    expect(isFlatOpc('<?xml version="1.0"?><note><body>Not flat opc</body></note>')).toBe(false)
    expect(isFlatOpc(Buffer.from('PK\x03\x041234'))).toBe(false)
    expect(isFlatOpc(null)).toBe(false)
  })

  it('2. should load sample.xml directly from file path via PPTXTemplater.load()', async () => {
    const ppt = await PPTXTemplater.load(sampleXmlPath)
    expect(ppt).toBeInstanceOf(PPTXTemplater)
    expect(ppt.slideCount).toBe(6)

    const dims = ppt.getSlideDimensions()
    expect(dims.width).toBe(12192000)
    expect(dims.height).toBe(6858000)
  })

  it('3. should load sample.xml from in-memory Buffer via PPTXTemplater.load()', async () => {
    const buffer = fs.readFileSync(sampleXmlPath)
    const ppt = await PPTXTemplater.load(buffer)
    expect(ppt).toBeInstanceOf(PPTXTemplater)
    expect(ppt.slideCount).toBe(6)
  })

  it('4. should load sample.xml from raw XML string via PPTXTemplater.load()', async () => {
    const xmlContent = fs.readFileSync(sampleXmlPath, 'utf8')
    const ppt = await PPTXTemplater.load(xmlContent)
    expect(ppt).toBeInstanceOf(PPTXTemplater)
    expect(ppt.slideCount).toBe(6)
  })

  it('5. should load via explicit static PPTXTemplater.fromFlatOpc()', async () => {
    const ppt = await PPTXTemplater.fromFlatOpc(sampleXmlPath)
    expect(ppt).toBeInstanceOf(PPTXTemplater)
    expect(ppt.slideCount).toBe(6)
  })

  it('6. should execute text and table replacements and export to a valid .pptx file', async () => {
    const ppt = await PPTXTemplater.load(sampleXmlPath)

    // Slide 1: replace {{title}}
    ppt.useSlide(1)
    ppt.replaceText({ '{{title}}': 'World Class Presentation' })

    // Slide 2: replace {{title2}}
    ppt.useSlide(2)
    ppt.replaceText({ '{{title2}}': 'Quarterly Insights' })

    // Slide 3: update table
    ppt.useSlide(3)
    ppt.replaceText({ '{{title}}': 'Performance Metrics' })
    ppt.updateTable('Table', [
      ['Metric', 'Target', 'Actual', 'Variance', 'Growth', 'Status'],
      ['Revenue', '$10M', '$12.4M', '+$2.4M', '+24%', 'Exceeded'],
      ['Gross Margin', '75%', '78.2%', '+3.2%', '+4%', 'On Track'],
    ])

    // Export to Buffer
    const pptxBuffer = await ppt.toBuffer()
    expect(Buffer.isBuffer(pptxBuffer)).toBe(true)
    expect(pptxBuffer.length).toBeGreaterThan(10000)
    expect(isPkZip(pptxBuffer)).toBe(true)

    // Save to disk
    if (!fs.existsSync(tempOutputDir)) {
      fs.mkdirSync(tempOutputDir, { recursive: true })
    }
    await ppt.saveToFile(tempPptxPath)
    expect(fs.existsSync(tempPptxPath)).toBe(true)

    // Reload exported file with standard load to verify OpenXML validity
    const reloaded = await PPTXTemplater.load(tempPptxPath)
    expect(reloaded.slideCount).toBe(6)

    // Search for replaced text on slide 1
    reloaded.useSlide(1)
    const matches = reloaded.findText('World Class Presentation')
    expect(matches.length).toBeGreaterThanOrEqual(1)

    // Search for replaced text on slide 2
    reloaded.useSlide(2)
    const matches2 = reloaded.findText('Quarterly Insights')
    expect(matches2.length).toBeGreaterThanOrEqual(1)
  })

  it('7. should support round-trip serialization back to Flat OPC XML via toFlatXml()', async () => {
    const ppt = await PPTXTemplater.load(sampleXmlPath)
    ppt.useSlide(1)
    ppt.replaceText({ '{{title}}': 'Round Trip Test' })

    const flatXml = await ppt.toFlatXml()
    expect(typeof flatXml).toBe('string')
    expect(flatXml).toContain('<pkg:package')
    expect(flatXml).toContain('<pkg:part')
    expect(flatXml).toContain('Round Trip Test')

    // Re-load the generated flat XML string to verify validity
    const reloadedFromFlatXml = await PPTXTemplater.load(flatXml)
    expect(reloadedFromFlatXml.slideCount).toBe(6)
    reloadedFromFlatXml.useSlide(1)
    const matches = reloadedFromFlatXml.findText('Round Trip Test')
    expect(matches.length).toBeGreaterThanOrEqual(1)
  })
})
