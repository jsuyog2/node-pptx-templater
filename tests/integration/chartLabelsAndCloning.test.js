import { describe, it, expect, beforeAll } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'
import JSZip from 'jszip'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')

describe('TASK-CHART-006: Chart Labels & Cloning Subsystems', () => {
  beforeAll(async () => {
    if (!existsSync(FIXTURE_FILE)) {
      const ppt = PPTXTemplater.create()
      ppt.addSlide(1)
      await ppt.saveAs(FIXTURE_FILE)
    }
  })

  it('should deeply clone a chart with fresh XML, relationship, and workbook', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)

    // Add source chart
    ppt.addChart(
      1,
      'column',
      {
        categories: ['2023', '2024', '2025'],
        series: [{ name: 'ARR', values: [100, 250, 500] }],
      },
      { name: 'OriginalARR', x: '1in', y: '1in', width: '4in', height: '3in' }
    )

    // Clone the chart to a new position
    ppt.useSlide(1).cloneChart('OriginalARR', 'ClonedARR', {
      x: '5.5in',
      y: '1in',
      width: '4in',
      height: '3in',
    })

    // Update the cloned chart with different data
    ppt.useSlide(1).updateChartData('ClonedARR', {
      categories: ['2023', '2024', '2025'],
      series: [{ name: 'ARR-Projected', values: [120, 300, 650] }],
    })

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)

    // Verify both charts exist in the presentation
    const reloaded = await PPTXTemplater.load(buffer)
    const charts = reloaded.getCharts(1)
    const original = charts.find(c => c.name === 'OriginalARR')
    const cloned = charts.find(c => c.name === 'ClonedARR')

    expect(original).toBeDefined()
    expect(cloned).toBeDefined()
    expect(original.zipPath).not.toBe(cloned.zipPath)
    expect(original.shapeId).not.toBe(cloned.shapeId)
    expect(original.rId).not.toBe(cloned.rId)

    // Verify independent series values
    expect(original.series[0].values).toEqual([100, 250, 500])
    expect(cloned.series[0].values).toEqual([120, 300, 650])
    expect(cloned.series[0].name).toBe('ARR-Projected')

    // Verify both chart XML files and workbooks exist in zip
    expect(zip.file(original.zipPath)).not.toBeNull()
    expect(zip.file(cloned.zipPath)).not.toBeNull()
  })

  it('should manipulate data labels via getChartLabel, setChartLabel, and updateChartLabel', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)

    ppt.addChart(
      1,
      'bar',
      {
        categories: ['Product A', 'Product B', 'Product C'],
        series: [{ name: 'Units Sold', values: [450, 720, 310] }],
      },
      { name: 'LabelTestBar' }
    )

    // 1. setChartLabel with boolean & position
    ppt.useSlide(1).setChartLabel('LabelTestBar', true, { position: 'outEnd' })

    let labelConfig = await ppt.getChartLabel('LabelTestBar')
    expect(labelConfig.showVal).toBe(true)
    expect(labelConfig.position).toBe('outEnd')

    // 2. updateChartLabel with custom label array and category display
    ppt.useSlide(1).updateChartLabel('LabelTestBar', {
      showCatName: true,
      labels: ['Low (450)', 'High (720)', 'Medium (310)'],
    })

    labelConfig = await ppt.getChartLabel('LabelTestBar')
    expect(labelConfig.showCatName).toBe(true)
    expect(labelConfig.labels.map(l => l.value || l)).toEqual([
      'Low (450)',
      'High (720)',
      'Medium (310)',
    ])

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)

    // Verify the chart XML has custom label definitions
    const chartInfo = ppt.getChart('LabelTestBar')
    const chartXml = await zip.file(chartInfo.zipPath).async('text')
    expect(chartXml).toContain('High (720)')
    expect(chartXml).toContain('showCatName val="1"')
    expect(chartXml).toContain('outEnd')
  })
})
