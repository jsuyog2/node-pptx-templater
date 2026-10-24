import { describe, it, expect, beforeAll } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'
import JSZip from 'jszip'
import { ChartWorkbookUpdater } from '../../src/elements/charts/ChartWorkbookUpdater.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')

let PPTXTemplater

beforeAll(async () => {
  const module = await import('../../src/index.js')
  PPTXTemplater = module.PPTXTemplater
})

describe('TASK-CHART-004: Dual-Sync Engine for Standard & Complex Chart Types', () => {
  const runTests = existsSync(FIXTURE_FILE)
  if (!runTests) {
    it.skip('Skipping: sample.pptx fixture not found', () => {})
    return
  }

  it('should maintain Dual-Sync invariant across chart XML cache and embedded Excel sheet for standard Column chart', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)

    // Add a new column chart
    ppt.addChart(
      1,
      'column',
      {
        categories: ['Alpha', 'Beta', 'Gamma'],
        series: [{ name: 'Velocity', values: [12.5, 24.8, 38.2] }],
      },
      { name: 'DualSyncColumn' }
    )

    // Update with new data using updateChartData
    ppt.useSlide(1).updateChartData('DualSyncColumn', {
      categories: ['Alpha-New', 'Beta-New', 'Gamma-New', 'Delta-New'],
      series: [{ name: 'Velocity-Updated', values: [55.1, 66.2, 77.3, 88.4] }],
    })

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)

    // 1. Verify DrawingML Chart XML cache
    const chartXmlFiles = Object.keys(zip.files).filter(
      f => f.startsWith('ppt/charts/chart') && f.endsWith('.xml') && !f.includes('_rels')
    )
    expect(chartXmlFiles.length).toBeGreaterThan(0)

    let foundChartXml = null
    for (const f of chartXmlFiles) {
      const xml = await zip.file(f).async('text')
      if (xml.includes('Alpha-New')) {
        foundChartXml = xml
        break
      }
    }
    expect(foundChartXml).toBeDefined()
    expect(foundChartXml).toContain('Alpha-New')
    expect(foundChartXml).toContain('Velocity-Updated')
    expect(foundChartXml).toContain('88.4')

    // 2. Verify embedded Excel sheet
    const embeddedXlsxFiles = Object.keys(zip.files).filter(
      f => f.startsWith('ppt/embeddings/') && f.endsWith('.xlsx')
    )
    expect(embeddedXlsxFiles.length).toBeGreaterThan(0)

    let foundWorkbookMatch = false
    for (const wbPath of embeddedXlsxFiles) {
      const wbBuffer = await zip.file(wbPath).async('nodebuffer')
      const wbZip = await JSZip.loadAsync(wbBuffer)
      const sheetFile = wbZip.file('xl/worksheets/sheet1.xml')
      if (sheetFile) {
        const sheetXml = await sheetFile.async('text')
        const sharedStrings = await ChartWorkbookUpdater.getSharedStrings(wbZip)
        const cells = ChartWorkbookUpdater.parseWorksheetCells(sheetXml, sharedStrings)

        if (cells['A2'] === 'Alpha-New' || cells['B1'] === 'Velocity-Updated') {
          foundWorkbookMatch = true
          expect(cells['A2']).toBe('Alpha-New')
          expect(cells['A3']).toBe('Beta-New')
          expect(cells['A4']).toBe('Gamma-New')
          expect(cells['A5']).toBe('Delta-New')
          expect(cells['B1']).toBe('Velocity-Updated')
          expect(cells['B2']).toBe(55.1)
          expect(cells['B5']).toBe(88.4)
          break
        }
      }
    }
    expect(foundWorkbookMatch).toBe(true)
  })

  it('should support Dual-Sync on Office 2016 extended Waterfall chart', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)

    ppt.addChart(
      1,
      'waterfall',
      {
        categories: ['Base', 'Revenue', 'Costs', 'Taxes', 'Net'],
        series: [{ name: 'Cash Flow', values: [1000, 400, -300, -100, 1000] }],
      },
      { name: 'WaterfallSync' }
    )

    // Mutate data
    ppt.useSlide(1).updateChart('WaterfallSync', {
      categories: ['Base-2026', 'Revenue-2026', 'Costs-2026', 'Taxes-2026', 'Net-2026'],
      series: [{ name: 'Cash Flow 2026', values: [2000, 800, -500, -200, 2100] }],
      title: 'FY26 Cash Waterfall',
    })

    const buffer = await ppt.toBuffer()
    const reloaded = await PPTXTemplater.load(buffer)
    const chart = reloaded.getChart('WaterfallSync')

    expect(chart).toBeDefined()
    expect(chart.type).toBe('waterfall')
    expect(chart.title).toBe('FY26 Cash Waterfall')
    expect(chart.categories).toContain('Base-2026')
    expect(chart.series[0].values).toContain(2000)
    expect(chart.series[0].values).toContain(2100)
  })

  it('should support Dual-Sync on Combo chart combining Column and Line series', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)

    ppt.addChart(
      1,
      'combo',
      {
        categories: ['2023', '2024', '2025', '2026'],
        series: [
          { name: 'Volume', values: [1000, 1500, 2200, 3100], type: 'column' },
          { name: 'Margin', values: [22.5, 25.1, 28.0, 31.4], type: 'line' },
        ],
      },
      { name: 'ComboSync' }
    )

    ppt.updateChartData('ComboSync', {
      categories: ['2024', '2025', '2026', '2027'],
      series: [
        { name: 'Volume', values: [1600, 2300, 3200, 4500] },
        { name: 'Margin', values: [26.0, 29.5, 33.0, 37.5] },
      ],
    })

    const buffer = await ppt.toBuffer()
    const reloaded = await PPTXTemplater.load(buffer)
    const chart = reloaded.getChart('ComboSync')

    expect(chart.type).toBe('combo')
    expect(chart.categories).toEqual(['2024', '2025', '2026', '2027'])
    expect(chart.series[0].values).toEqual([1600, 2300, 3200, 4500])
    expect(chart.series[1].values).toEqual([26.0, 29.5, 33.0, 37.5])
  })

  it('should cleanse Excel tables to prevent PowerPoint repair prompts', async () => {
    // Generate new workbook and update it
    const initialWb = await ChartWorkbookUpdater.createNewWorkbook({
      categories: ['A', 'B'],
      series: [{ name: 'S1', values: [1, 2] }],
    })

    const zip = await JSZip.loadAsync(initialWb)
    // Verify no xl/tables/ remain
    const tableFiles = Object.keys(zip.files).filter(f => f.startsWith('xl/tables/'))
    expect(tableFiles.length).toBe(0)

    const updatedWb = await ChartWorkbookUpdater.updateWorkbook(initialWb, {
      categories: ['A', 'B', 'C'],
      series: [{ name: 'S1', values: [10, 20, 30] }],
    })

    const updatedZip = await JSZip.loadAsync(updatedWb)
    const updatedTableFiles = Object.keys(updatedZip.files).filter(f => f.startsWith('xl/tables/'))
    expect(updatedTableFiles.length).toBe(0)
  })
})
