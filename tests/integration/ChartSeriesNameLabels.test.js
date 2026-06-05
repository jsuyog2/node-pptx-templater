import { describe, it, expect, beforeAll } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')
const TEMPLATE_FILE = resolve(__dirname, '../../templates/sample.pptx')

let PPTXTemplater

beforeAll(async () => {
  const module = await import('../../src/index.js')
  PPTXTemplater = module.PPTXTemplater
})

describe('Chart Series Name Labels (External Text Boxes)', () => {
  const testFile = existsSync(TEMPLATE_FILE)
    ? TEMPLATE_FILE
    : existsSync(FIXTURE_FILE)
      ? FIXTURE_FILE
      : null

  if (!testFile) {
    it.skip('Skipping: template or fixture sample.pptx not found', () => {})
    return
  }

  it('should validate seriesNameLabels options and reject invalid positions', async () => {
    const ppt = await PPTXTemplater.load(testFile)

    // Valid left position
    const resLeft = await ppt.useSlide(1).validateSeriesNameLabels('Chart', {
      enabled: true,
      position: 'left',
    })
    expect(resLeft.valid).toBe(true)

    // Valid right position
    const resRight = await ppt.useSlide(1).validateSeriesNameLabels('Chart', {
      enabled: true,
      position: 'right',
    })
    expect(resRight.valid).toBe(true)

    // Invalid positions should be rejected
    for (const pos of ['top', 'bottom', 'center', 'invalid']) {
      const res = await ppt.useSlide(1).validateSeriesNameLabels('Chart', {
        enabled: true,
        position: pos,
      })
      expect(res.valid).toBe(false)
      expect(res.errors[0]).toContain('Only "left" and "right" are supported')
    }
  })

  it('should generate text boxes positioned on the left and inherit styling', async () => {
    const ppt = await PPTXTemplater.load(testFile)

    ppt.useSlide(1).updateChart('Chart', {
      categories: ['Sales'],
      seriesNameLabels: {
        enabled: true,
        position: 'left',
        autoFit: true,
      },
      series: [
        { name: 'Product A', values: [100] },
        { name: 'Product B', values: [150] },
      ],
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)
    const slideXml = await ppt2.zipManager.readFile('ppt/slides/slide1.xml')

    // Expect generated text boxes with names SeriesNameLabel-Chart-0 and SeriesNameLabel-Chart-1
    expect(slideXml).toContain('name="SeriesNameLabel-Chart-0"')
    expect(slideXml).toContain('name="SeriesNameLabel-Chart-1"')

    // Expect alignment to be right-aligned (algn="r") for left position
    expect(slideXml).toContain('algn="r"')

    // Check that we generated shape components with correct content
    expect(slideXml).toContain('<a:t>Product A</a:t>')
    expect(slideXml).toContain('<a:t>Product B</a:t>')
  })

  it('should generate text boxes positioned on the right', async () => {
    const ppt = await PPTXTemplater.load(testFile)

    ppt.useSlide(1).updateChart('Chart', {
      categories: ['Sales'],
      seriesNameLabels: {
        enabled: true,
        position: 'right',
        autoFit: true,
      },
      series: [
        { name: 'Product A', values: [100] },
        { name: 'Product B', values: [150] },
      ],
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)
    const slideXml = await ppt2.zipManager.readFile('ppt/slides/slide1.xml')

    // Expect generated text boxes
    expect(slideXml).toContain('name="SeriesNameLabel-Chart-0"')
    // Expect alignment to be left-aligned (algn="l") for right position
    expect(slideXml).toContain('algn="l"')
  })

  it('should cleanly replace previous labels on multiple chart updates', async () => {
    const ppt = await PPTXTemplater.load(testFile)

    // First update
    ppt.useSlide(1).updateChart('Chart', {
      categories: ['Sales'],
      seriesNameLabels: {
        enabled: true,
        position: 'left',
      },
      series: [
        { name: 'Old A', values: [100] },
        { name: 'Old B', values: [150] },
      ],
    })

    // Second update
    ppt.useSlide(1).updateChart('Chart', {
      categories: ['Sales'],
      seriesNameLabels: {
        enabled: true,
        position: 'right',
      },
      series: [
        { name: 'New A', values: [200] },
        { name: 'New B', values: [250] },
      ],
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)
    const slideXml = await ppt2.zipManager.readFile('ppt/slides/slide1.xml')

    // The old names must not exist, only new ones
    expect(slideXml).not.toContain('Old A')
    expect(slideXml).not.toContain('Old B')
    expect(slideXml).toContain('New A')
    expect(slideXml).toContain('New B')
  })

  it('should support long series names with auto-wrap height estimation', async () => {
    const ppt = await PPTXTemplater.load(testFile)

    ppt.useSlide(1).updateChart('Chart', {
      categories: ['Sales'],
      seriesNameLabels: {
        enabled: true,
        position: 'left',
        autoFit: true,
      },
      series: [
        {
          name: 'Super Long Series Name That Should Trigger Wrap Height Adjustment',
          values: [100],
        },
      ],
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)
    const slideXml = await ppt2.zipManager.readFile('ppt/slides/slide1.xml')

    // The shape should be generated successfully
    expect(slideXml).toContain('name="SeriesNameLabel-Chart-0"')
    expect(slideXml).toContain('Super Long Series Name')
  })

  it('should validate and apply custom style options to seriesNameLabels', async () => {
    const ppt = await PPTXTemplater.load(testFile)

    // Validation checks for styles
    const resInvalidColor = await ppt.useSlide(1).validateSeriesNameLabels('Chart', {
      enabled: true,
      position: 'left',
      style: {
        color: '#ZZZZZZ', // Invalid hex
      },
    })
    expect(resInvalidColor.valid).toBe(false)
    expect(resInvalidColor.errors[0]).toContain('is not a valid hex color')

    const resInvalidAlign = await ppt.useSlide(1).validateSeriesNameLabels('Chart', {
      enabled: true,
      position: 'left',
      style: {
        align: 'justified', // Invalid align
      },
    })
    expect(resInvalidAlign.valid).toBe(false)
    expect(resInvalidAlign.errors[0]).toContain('Supported alignments are')

    const resValidStyle = await ppt.useSlide(1).validateSeriesNameLabels('Chart', {
      enabled: true,
      position: 'left',
      style: {
        fontSize: 14,
        fontFamily: 'Calibri',
        bold: true,
        italic: true,
        color: '#FF5500',
        align: 'center',
      },
    })
    expect(resValidStyle.valid).toBe(true)

    // Run chart update with custom styling options
    ppt.useSlide(1).updateChart('Chart', {
      categories: ['Sales'],
      seriesNameLabels: {
        enabled: true,
        position: 'left',
        style: {
          fontSize: 14,
          fontFamily: 'Calibri',
          bold: true,
          italic: true,
          color: '#FF5500',
          align: 'center',
        },
      },
      series: [{ name: 'Product A', values: [100] }],
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)
    const slideXml = await ppt2.zipManager.readFile('ppt/slides/slide1.xml')

    // Verify properties injected into <a:rPr> of the textbox
    // 14pt in PowerPoint/DrawingML is 1400 sz
    expect(slideXml).toContain('sz="1400"')
    expect(slideXml).toContain('b="1"')
    expect(slideXml).toContain('i="1"')
    expect(slideXml).toContain('val="FF5500"')
    expect(slideXml).toContain('typeface="Calibri"')
    // Alignment "ctr" in <a:pPr>
    expect(slideXml).toContain('algn="ctr"')
  })

  it('should support getChartLabelPositions and getChartBarPositions APIs', async () => {
    const ppt = await PPTXTemplater.load(testFile)

    ppt.useSlide(1).updateChart('Chart', {
      categories: ['Sales'],
      series: [
        { name: 'Product A', values: [100] },
        { name: 'Product B', values: [150] },
      ],
    })

    const labels = await ppt.getChartLabelPositions('Chart')
    const bars = await ppt.getChartBarPositions('Chart')

    expect(labels).toBeInstanceOf(Array)
    expect(labels.length).toBe(2)
    expect(labels[0].series).toBe('Product A')
    expect(labels[0].category).toBe('Sales')
    expect(labels[0].x).toBeGreaterThan(0)
    expect(labels[0].y).toBeGreaterThan(0)
    expect(labels[0].width).toBeGreaterThan(0)
    expect(labels[0].height).toBeGreaterThan(0)

    expect(bars).toBeInstanceOf(Array)
    expect(bars.length).toBe(2)
    expect(bars[0].series).toBe('Product A')
    expect(bars[0].category).toBe('Sales')
    expect(bars[0].x).toBeGreaterThan(0)
    expect(bars[0].y).toBeGreaterThan(0)
    expect(bars[0].width).toBeGreaterThan(0)
    expect(bars[0].height).toBeGreaterThan(0)
  })

  it('should support addTextAtPosition and addTextNearChartLabel APIs', async () => {
    const ppt = await PPTXTemplater.load(testFile)

    ppt.useSlide(1).updateChart('Chart', {
      categories: ['Sales'],
      series: [
        { name: 'Product A', values: [100] },
        { name: 'Product B', values: [150] },
      ],
    })

    ppt.useSlide(1).addTextAtPosition({
      text: 'Custom Text At Position',
      x: 1000000,
      y: 2000000,
      width: 1500000,
      height: 400000,
      style: {
        fontSize: 12,
        fontFamily: 'Arial',
        bold: true,
      },
    })

    ppt.addTextNearChartLabel({
      chart: 'Chart',
      text: ({ series }) => `Near ${series}`,
      position: 'right',
      style: {
        fontSize: 11,
        fontFamily: 'Calibri',
        italic: true,
      },
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)
    const slideXml = await ppt2.zipManager.readFile('ppt/slides/slide1.xml')

    expect(slideXml).toContain('TextBoxAtPosition')
    expect(slideXml).toContain('Custom Text At Position')
    expect(slideXml).toContain('sz="1200"')
    expect(slideXml).toContain('b="1"')

    expect(slideXml).toContain('Near Product A')
    expect(slideXml).toContain('Near Product B')
    expect(slideXml).toContain('sz="1100"')
    expect(slideXml).toContain('i="1"')

    // Ensure b="0" or i="0" are never generated in slide XML (PPT Repair violation)
    expect(slideXml).not.toContain('b="0"')
    expect(slideXml).not.toContain('i="0"')
  })
})
