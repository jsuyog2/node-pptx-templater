import { describe, it, expect, beforeAll } from 'vitest'
import path from 'path'
import fs from 'fs-extra'
import JSZip from 'jszip'
import { PPTXTemplater } from '../../src/index.js'

describe('TASK-CHART-CLEARCHART-FIX-002: clearCharts Parameter Handling & Label Consolidation', () => {
  const templatePath = path.join(__dirname, '../../showcase/templates/ChartPowerPointSample.pptx')
  const testOutputDir = path.join(__dirname, '../../showcase/output/vitest_clearcharts')

  beforeAll(async () => {
    await fs.ensureDir(testOutputDir)
  })

  it('should remove all charts on a slide when clearCharts() is called with no arguments', async () => {
    const ppt = await PPTXTemplater.load(templatePath)

    // Add charts to Slide 1
    await ppt.addChart(
      1,
      'bar',
      {
        categories: ['A', 'B'],
        series: [{ name: 'S1', values: [10, 20] }],
      },
      { name: 'Slide1Chart1' }
    )
    await ppt.addChart(
      1,
      'line',
      {
        categories: ['A', 'B'],
        series: [{ name: 'S2', values: [30, 40] }],
      },
      { name: 'Slide1Chart2' }
    )

    const initialCharts = ppt.useSlide(1).getCharts()
    expect(initialCharts.length).toBeGreaterThanOrEqual(2)

    // Clear all charts on Slide 1
    ppt.useSlide(1).clearCharts()

    const remainingCharts = ppt.useSlide(1).getCharts()
    expect(remainingCharts.length).toBe(0)

    const outPath = path.join(testOutputDir, 'clear_all_slide1.pptx')
    await ppt.saveToFile(outPath)

    const buffer = await fs.readFile(outPath)
    const zip = await JSZip.loadAsync(buffer)
    const slide1Xml = await zip.file('ppt/slides/slide1.xml').async('text')

    // DrawingML chart references should be gone from slide1.xml
    expect(slide1Xml).not.toContain('name="Slide1Chart1"')
    expect(slide1Xml).not.toContain('name="Slide1Chart2"')
  })

  it('should remove all charts on a designated slide when clearCharts(slideIndex) is called', async () => {
    const ppt = await PPTXTemplater.load(templatePath)

    await ppt.addChart(
      1,
      'bar',
      {
        categories: ['A', 'B'],
        series: [{ name: 'S1', values: [10, 20] }],
      },
      { name: 'Slide1Bar' }
    )

    const initialCharts = ppt.getCharts(1)
    expect(initialCharts.length).toBeGreaterThan(0)

    // Clear charts on Slide 1 by passing slide index 1
    ppt.clearCharts(1)

    const remainingCharts = ppt.getCharts(1)
    expect(remainingCharts.length).toBe(0)
  })

  it('should remove a specific chart without throwing when clearCharts(chartId) is called', async () => {
    const ppt = await PPTXTemplater.load(templatePath)

    // Add two distinct charts on Slide 1
    await ppt.addChart(
      1,
      'bar',
      {
        categories: ['A', 'B', 'C'],
        series: [{ name: 'S1', values: [10, 20, 30] }],
      },
      { name: 'TargetChartToClear' }
    )

    await ppt.addChart(
      1,
      'line',
      {
        categories: ['Jan', 'Feb', 'Mar'],
        series: [{ name: 'Trend', values: [100, 200, 300] }],
      },
      { name: 'ChartToKeep' }
    )

    expect(ppt.getChart('TargetChartToClear')).not.toBeNull()
    expect(ppt.getChart('ChartToKeep')).not.toBeNull()

    // Invoke clearCharts with specific chartId (polymorphic handling)
    expect(() => {
      ppt.clearCharts('TargetChartToClear')
    }).not.toThrow()

    // TargetChartToClear is evacuated, ChartToKeep remains
    expect(() => ppt.getChart('TargetChartToClear')).toThrow(/not found/)
    expect(ppt.getChart('ChartToKeep')).not.toBeNull()

    const outPath = path.join(testOutputDir, 'clear_by_chart_id.pptx')
    await ppt.saveToFile(outPath)

    const buffer = await fs.readFile(outPath)
    const zip = await JSZip.loadAsync(buffer)
    const slide1Xml = await zip.file('ppt/slides/slide1.xml').async('text')

    expect(slide1Xml).not.toContain('name="TargetChartToClear"')
    expect(slide1Xml).toContain('name="ChartToKeep"')
  })

  it('should gracefully handle non-existent chartId without throwing in clearCharts(chartId)', async () => {
    const ppt = await PPTXTemplater.load(templatePath)

    expect(() => {
      ppt.clearCharts('NonExistentChartId12345')
    }).not.toThrow()
  })

  it('should update chart data labels via updateChart(chartId, { labels: { showVal: true } })', async () => {
    const ppt = await PPTXTemplater.load(templatePath)

    await ppt.addChart(
      1,
      'column',
      {
        categories: ['Q1', 'Q2', 'Q3', 'Q4'],
        series: [{ name: 'Revenue', values: [50, 75, 100, 125] }],
      },
      { name: 'LabelUpdateChart' }
    )

    // Update labels using updateChart
    ppt.updateChart('LabelUpdateChart', {
      labels: {
        showVal: true,
        showCatName: true,
        position: 'outEnd',
      },
    })

    const chart = ppt.getChart('LabelUpdateChart')
    expect(chart).not.toBeNull()
    expect(chart.dataLabels).toBeDefined()
    expect(chart.dataLabels.showVal).toBe(true)
    expect(chart.dataLabels.showCatName).toBe(true)
    expect(chart.dataLabels.position).toBe('outEnd')

    const outPath = path.join(testOutputDir, 'labels_via_update_chart.pptx')
    await ppt.saveToFile(outPath)

    const buffer = await fs.readFile(outPath)
    const zip = await JSZip.loadAsync(buffer)
    const chartXml = await zip.file(chart.zipPath).async('text')

    expect(chartXml).toContain('showVal val="1"')
    expect(chartXml).toContain('showCatName val="1"')
    expect(chartXml).toContain('outEnd')
  })
})
