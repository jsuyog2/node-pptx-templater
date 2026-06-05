import { describe, it, expect, beforeAll } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')
const TEMPLATE_FILE = resolve(__dirname, '../../templates/sample.pptx')

let PPTXTemplater
let ValidationEngine

beforeAll(async () => {
  const module = await import('../../src/index.js')
  PPTXTemplater = module.PPTXTemplater
  ValidationEngine = module.ValidationEngine
})

describe('Chart Label Inheritance and Series Name Tests', () => {
  const testFile = existsSync(TEMPLATE_FILE)
    ? TEMPLATE_FILE
    : existsSync(FIXTURE_FILE)
      ? FIXTURE_FILE
      : null

  if (!testFile) {
    it.skip('Skipping: template or fixture sample.pptx not found', () => {})
    return
  }

  it('should inherit styling from existing txPr for custom rich text labels', async () => {
    const ppt = await PPTXTemplater.load(testFile)

    // Update chart to have custom data labels
    ppt.useSlide(1).updateChart('Chart', {
      categories: ['Sales', 'Revenue'],
      series: [
        {
          name: 'Series 1',
          values: [
            { data: 120, label: 'Custom Sales Label' },
            { data: 240, label: 'Custom Revenue Label' },
          ],
        },
      ],
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)

    // Retrieve chart XML
    const xml = await ppt2.zipManager.readFile('ppt/charts/chart1.xml')

    // We expect the custom text labels inside <c:rich> to have inherited the <a:rPr> elements
    // Let's verify that a:r elements in the rich text contain an a:rPr element
    expect(xml).toContain('<c:rich>')
    expect(xml).toContain('<a:rPr')
  })

  it('should support showSeriesNameInBar option globally', async () => {
    const ppt = await PPTXTemplater.load(testFile)

    // Update chart with showSeriesNameInBar: true at the top level
    ppt.useSlide(1).updateChart('Chart', {
      showSeriesNameInBar: true,
      categories: ['Sales', 'Revenue'],
      series: [
        { name: 'Series A', values: [100, 200] },
        { name: 'Series B', values: [150, 250] },
      ],
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)

    const xml = await ppt2.zipManager.readFile('ppt/charts/chart1.xml')

    // Native showSerName should be val="1" and showVal val="0"
    expect(xml).toContain('<c:showSerName val="1"/>')
    expect(xml).toContain('<c:showVal val="0"/>')
    // Default position for showSeriesNameInBar should be ctr
    expect(xml).toContain('<c:dLblPos val="ctr"/>')
  })

  it('should support showSeriesNameInBar option per-series', async () => {
    const ppt = await PPTXTemplater.load(testFile)

    // Update chart with showSeriesNameInBar: true only on the first series
    ppt.useSlide(1).updateChart('Chart', {
      categories: ['Sales', 'Revenue'],
      series: [
        { name: 'Series A', values: [100, 200], showSeriesNameInBar: true },
        { name: 'Series B', values: [150, 250], showSeriesNameInBar: false },
      ],
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)

    const xml = await ppt2.zipManager.readFile('ppt/charts/chart1.xml')

    // Series A (index 0) should have showSerName val="1"
    // Series B (index 1) should have showSerName val="0"
    // Let's verify that both showSerName val="1" and val="0" exist in the XML
    expect(xml).toContain('<c:showSerName val="1"/>')
    expect(xml).toContain('<c:showSerName val="0"/>')
  })

  it('should validate chart label settings successfully', async () => {
    const ppt = await PPTXTemplater.load(testFile)

    // Validate valid config
    const result = await ppt.useSlide(1).validateChartLabels('Chart', {
      labels: ['A', 'B', 'C', 'D'],
      showSeriesNameInBar: true,
    })
    expect(result.valid).toBe(true)
    expect(result.errors).toEqual([])

    // Validate invalid config (e.g. mismatched label count)
    const resultInvalid = await ppt.useSlide(1).validateChartLabels('Chart', {
      labels: ['A', 'B'], // Mismatched count
    })
    expect(resultInvalid.valid).toBe(false)
    expect(resultInvalid.errors.length).toBeGreaterThan(0)
    expect(resultInvalid.errors[0]).toContain('does not match chart data points')
  })
})
