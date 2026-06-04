import { describe, it, expect, beforeAll } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')

let PPTXTemplater
let ValidationEngine

beforeAll(async () => {
  const module = await import('../../src/index.js')
  PPTXTemplater = module.PPTXTemplater
  ValidationEngine = module.ValidationEngine
})

describe('Chart Data Labels Integration Tests', () => {
  const runTests = existsSync(FIXTURE_FILE)

  if (!runTests) {
    it.skip('Skipping: sample.pptx fixture not found', () => {})
    return
  }

  it('should apply literal label arrays and write to spreadsheet', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1).updateDataLabels('Chart', {
      series: 0,
      labels: ['Excellent', 'Good', 'Average', 'Poor']
    })

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeDefined()
    expect(buffer[0]).toBe(0x50) // PK zip header

    // Read back and assert values match
    const ppt2 = await PPTXTemplater.load(buffer)
    const labels = await ppt2.useSlide(1).getDataLabels('Chart', { series: 0 })
    expect(labels.length).toBeGreaterThan(0)
    expect(labels[0].value).toBe('Excellent')
    expect(labels[1].value).toBe('Good')
  })

  it('should support dynamic label mapping based on categories', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1).updateDataLabels('Chart', {
      series: 0,
      labelMap: {
        'Category 1': 'Tier 1',
        'Category 2': 'Tier 2'
      }
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)
    const labels = await ppt2.useSlide(1).getDataLabels('Chart', { series: 0 })
    expect(labels.length).toBeGreaterThan(0)
    expect(labels[0].value).toBe('Tier 1')
  })

  it('should evaluate label templates correctly', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    // First, update chart with known series values
    ppt.useSlide(1).updateChart('Chart', {
      categories: ['Sales', 'Revenue'],
      series: [{ name: '2026', values: [100, 200] }]
    })

    // Now set data label template
    ppt.updateDataLabels('Chart', {
      series: 0,
      template: '{category}: {value}'
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)
    const labels = await ppt2.useSlide(1).getDataLabels('Chart', { series: 0 })
    expect(labels).toContainEqual({ point: 0, value: 'Sales: 100' })
    expect(labels).toContainEqual({ point: 1, value: 'Revenue: 200' })
  })

  it('should support positioning and advanced text styling', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1).updateDataLabels('Chart', {
      series: 0,
      position: 'insideEnd',
      labelStyle: {
        fontFamily: 'Arial',
        fontSize: 14,
        bold: true,
        italic: true,
        underline: true,
        color: '#FF0000'
      }
    })

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeDefined()
  })

  it('should validate chart data label configurations', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    
    // Valid labels configuration
    const result1 = await ppt.useSlide(1).validateDataLabels('Chart', {
      labels: ['A', 'B', 'C', 'D']
    })
    expect(result1.valid).toBe(true)

    // Invalid labels configuration (length mismatch)
    const result2 = await ppt.useSlide(1).validateDataLabels('Chart', {
      labels: ['A', 'B'] // matches 4 points typically in Chart, so should be invalid
    })
    expect(result2.valid).toBe(false)
    expect(result2.errors.length).toBeGreaterThan(0)
    expect(result2.errors[0]).toContain('does not match chart data points')
  })

  it('should support inline object values with custom labels under updateChart', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1).updateChart('Chart', {
      categories: ['Q1', 'Q2', 'Q3', 'Q4'],
      series: [
        {
          name: 'Product A',
          values: [
            { data: 145, label: '145 (30%)' },
            { data: 210, label: '210 (40%)' },
            { data: 190, label: '190 (20%)' },
            { data: 250, label: '250 (10%)' }
          ]
        }
      ]
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)
    const labels = await ppt2.useSlide(1).getDataLabels('Chart', { series: 0 })
    expect(labels.length).toBe(4)
    expect(labels[0].value).toBe('145 (30%)')
    expect(labels[3].value).toBe('250 (10%)')
  })

  it('should support multi-series custom labels and mixed series configurations', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1).updateChart('Chart', {
      categories: ['Q1', 'Q2', 'Q3', 'Q4'],
      series: [
        {
          name: 'Product A',
          values: [
            { data: 10, label: 'Low A' },
            { data: 20, label: 'Mid A' },
            { data: 30, label: 'High A' },
            { data: 40, label: 'Peak A' }
          ]
        },
        {
          name: 'Product B',
          values: [100, 200, 300, 400] // primitive values (no custom labels for Product B)
        }
      ]
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)
    const labelsA = await ppt2.useSlide(1).getDataLabels('Chart', { series: 0 })
    const labelsB = await ppt2.useSlide(1).getDataLabels('Chart', { series: 1 })

    expect(labelsA[0].value).toBe('Low A')
    expect(labelsB.length).toBe(0) // No custom labels generated for series 1
  })

  it('should throw validation error on label count mismatch or invalid data types', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    
    // Label count mismatch
    expect(() => {
      ppt.useSlide(1).updateChart('Chart', {
        categories: ['Q1', 'Q2'],
        series: [
          {
            name: 'Product A',
            values: [
              { data: 100, label: 'L1' },
              { data: 200 } // Missing label
            ]
          }
        ]
      })
    }).toThrow('Label count mismatch for series Product A')

    // Invalid data value (not a number)
    expect(() => {
      ppt.useSlide(1).updateChart('Chart', {
        categories: ['Q1', 'Q2'],
        series: [
          {
            name: 'Product A',
            values: [
              { data: 'not-a-number', label: 'L1' },
              { data: 200, label: 'L2' }
            ]
          }
        ]
      })
    }).toThrow('Data value must be numeric in series Product A')

    // Series length mismatch
    expect(() => {
      ppt.useSlide(1).updateChart('Chart', {
        categories: ['Q1', 'Q2'],
        series: [
          {
            name: 'Product A',
            values: [100, 200, 300] // 3 values instead of 2
          }
        ]
      })
    }).toThrow('Series lengths mismatch')
  })

  it('should preserve existing styles (txPr, dLblPos, show tags) when updating labels', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    
    // Check original XML for labels properties if possible, or perform update
    // with a styling option, then update again without styling, and verify style is preserved!
    ppt.useSlide(1).updateDataLabels('Chart', {
      series: 0,
      labels: ['Val A', 'Val B', 'Val C', 'Val D'],
      position: 'insideEnd',
      labelStyle: {
        fontFamily: 'Courier New',
        fontSize: 15,
        bold: true,
        color: '#00FF00'
      }
    })

    // Save and load
    const buffer1 = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer1)

    // Now update chart data using inline custom labels but WITHOUT overriding the styling options
    ppt2.useSlide(1).updateChart('Chart', {
      categories: ['Q1', 'Q2', 'Q3', 'Q4'],
      series: [
        {
          name: 'Product A',
          values: [
            { data: 12, label: 'Label A' },
            { data: 22, label: 'Label B' },
            { data: 32, label: 'Label C' },
            { data: 42, label: 'Label D' }
          ]
        }
      ]
    })

    const buffer2 = await ppt2.toBuffer()
    const ppt3 = await PPTXTemplater.load(buffer2)

    // Verify custom labels have been updated
    const labels = await ppt3.useSlide(1).getDataLabels('Chart', { series: 0 })
    expect(labels[0].value).toBe('Label A')

    // Inspect the chart XML of ppt3 to verify `<c:txPr>` and `<c:dLblPos>` style features are still there
    const zipManager = ppt3.zipManager
    const chartPath = 'ppt/charts/chart1.xml'
    const chartXml = await zipManager.readFile(chartPath)
    expect(chartXml).toContain('typeface="Courier New"')
    expect(chartXml).toContain('sz="1500"')
    expect(chartXml).toContain('b="1"')
    expect(chartXml).toContain('val="00FF00"')
    expect(chartXml).toContain('val="inEnd"') // insideEnd maps to inEnd
  })
})
