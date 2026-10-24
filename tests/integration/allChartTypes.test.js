import { describe, it, expect, beforeAll } from 'vitest'
import path from 'path'
import fs from 'fs-extra'
import JSZip from 'jszip'
import { PPTXTemplater } from '../../src/index.js'
import {
  testBarChart,
  testColumnChart,
  testLineChart,
  testPieChart,
  testAreaChart,
  testScatterChart,
  testMapChart,
  testStockChart,
  testSurfaceChart,
  testRadarChart,
  testTreemapChart,
  testSunburstChart,
  testHistogramChart,
  testBoxWhiskerChart,
  testWaterfallChart,
  testFunnelChart,
  testComboChart,
} from '../../showcase/charts_operations.js'

describe('TASK-CHART-VERIFY-005: Comprehensive All 17 Chart Types Verification Suite', () => {
  const templatePath = path.join(__dirname, '../../showcase/templates/ChartPowerPointSample.pptx')
  const testOutputDir = path.join(__dirname, '../../showcase/output/vitest_all_charts')

  beforeAll(async () => {
    await fs.ensureDir(testOutputDir)
  })

  const standardTypes = [
    { name: 'bar', fn: testBarChart, file: 'chart_bar.pptx', expectedTag: '<c:barChart>' },
    { name: 'column', fn: testColumnChart, file: 'chart_column.pptx', expectedTag: '<c:barChart>' },
    { name: 'line', fn: testLineChart, file: 'chart_line.pptx', expectedTag: '<c:lineChart>' },
    { name: 'pie', fn: testPieChart, file: 'chart_pie.pptx', expectedTag: '<c:pieChart>' },
    { name: 'area', fn: testAreaChart, file: 'chart_area.pptx', expectedTag: '<c:areaChart>' },
    {
      name: 'scatter',
      fn: testScatterChart,
      file: 'chart_scatter.pptx',
      expectedTag: '<c:scatterChart>',
    },
    { name: 'stock', fn: testStockChart, file: 'chart_stock.pptx', expectedTag: '<c:stockChart>' },
    {
      name: 'surface',
      fn: testSurfaceChart,
      file: 'chart_surface.pptx',
      expectedTag: '<c:surfaceChart>',
    },
    { name: 'radar', fn: testRadarChart, file: 'chart_radar.pptx', expectedTag: '<c:radarChart>' },
    { name: 'combo', fn: testComboChart, file: 'chart_combo.pptx', expectedTag: '<c:barChart>' },
  ]

  const extendedTypes = [
    { name: 'treemap', fn: testTreemapChart, file: 'chart_treemap.pptx', layoutId: 'treemap' },
    { name: 'sunburst', fn: testSunburstChart, file: 'chart_sunburst.pptx', layoutId: 'sunburst' },
    {
      name: 'histogram',
      fn: testHistogramChart,
      file: 'chart_histogram.pptx',
      layoutId: 'clusteredColumn',
    },
    {
      name: 'boxWhisker',
      fn: testBoxWhiskerChart,
      file: 'chart_boxWhisker.pptx',
      layoutId: 'boxWhisker',
    },
    {
      name: 'waterfall',
      fn: testWaterfallChart,
      file: 'chart_waterfall.pptx',
      layoutId: 'waterfall',
    },
    { name: 'funnel', fn: testFunnelChart, file: 'chart_funnel.pptx', layoutId: 'funnel' },
    { name: 'map', fn: testMapChart, file: 'chart_map.pptx', layoutId: 'regionMap' },
  ]

  for (const { name, fn, expectedTag } of standardTypes) {
    it(`should generate a repair-free presentation package for Standard Chart: "${name}"`, async () => {
      const outPath = await fn(templatePath, testOutputDir)
      expect(fs.existsSync(outPath)).toBe(true)

      const buffer = await fs.readFile(outPath)
      const zip = await JSZip.loadAsync(buffer)

      // 1. Slide XML contains graphicFrame
      const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')
      expect(slideXml).toContain('<p:graphicFrame>')
      expect(slideXml).toContain('http://schemas.openxmlformats.org/drawingml/2006/chart')

      // 2. Chart XML exists and contains required tags
      const slideRelsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')
      const targetMatch = slideRelsXml.match(/Target="\.\.\/charts\/(chart\d+\.xml)"/)
      const chartFiles = Object.keys(zip.files).filter(f => /^ppt\/charts\/chart\d+\.xml$/.test(f))
      expect(chartFiles.length).toBeGreaterThan(0)
      const chartFileName = targetMatch ? targetMatch[1] : chartFiles[chartFiles.length - 1]
      const chartXml = await zip.file(`ppt/charts/${chartFileName}`).async('text')
      expect(chartXml).toContain(expectedTag)
      expect(chartXml).toContain('xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"')
      expect(chartXml).toContain('<c:externalData')

      // 3. Embedded Excel workbook exists
      const wbFiles = Object.keys(zip.files).filter(f =>
        /^ppt\/embeddings\/Microsoft_Excel_Worksheet\d*\.xlsx$/.test(f)
      )
      expect(wbFiles.length).toBeGreaterThan(0)

      // 4. Content Types overrides
      const contentTypes = await zip.file('[Content_Types].xml').async('text')
      expect(contentTypes).toContain(
        'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
      )
      expect(contentTypes).toContain(
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      )
    })
  }

  for (const { name, fn, layoutId } of extendedTypes) {
    it(`should generate a repair-free presentation package for Office 2016 Extended Chart: "${name}"`, async () => {
      const outPath = await fn(templatePath, testOutputDir)
      expect(fs.existsSync(outPath)).toBe(true)

      const buffer = await fs.readFile(outPath)
      const zip = await JSZip.loadAsync(buffer)

      // 1. Slide XML contains graphicFrame with chartex URI and namespace
      const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')
      expect(slideXml).toContain('<p:graphicFrame>')
      expect(slideXml).toContain('http://schemas.microsoft.com/office/drawing/2014/chartex')
      expect(slideXml).toContain(
        'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"'
      )

      // 2. ChartEx XML exists and contains layoutId
      const slideRelsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')
      const targetMatch = slideRelsXml.match(/Target="\.\.\/charts\/(chartEx\d+\.xml)"/)
      const chartExFiles = Object.keys(zip.files).filter(f =>
        /^ppt\/charts\/chartEx\d+\.xml$/.test(f)
      )
      expect(chartExFiles.length).toBeGreaterThan(0)
      const chartFileName = targetMatch ? targetMatch[1] : chartExFiles[chartExFiles.length - 1]
      const chartExXml = await zip.file(`ppt/charts/${chartFileName}`).async('text')
      expect(chartExXml).toContain(`layoutId="${layoutId}"`)
      expect(chartExXml).toContain(
        'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"'
      )
      expect(chartExXml).toContain('<cx:externalData')

      // 3. Embedded Excel workbook exists
      const wbFiles = Object.keys(zip.files).filter(f =>
        /^ppt\/embeddings\/Microsoft_Excel_Worksheet\d*\.xlsx$/.test(f)
      )
      expect(wbFiles.length).toBeGreaterThan(0)

      // 4. Content Types overrides
      const contentTypes = await zip.file('[Content_Types].xml').async('text')
      expect(contentTypes).toContain('application/vnd.ms-office.chartex+xml')
      expect(contentTypes).toContain(
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      )
    })
  }

  it('should support slide-level context and optional slideId parameter in addChart', async () => {
    const ppt = await PPTXTemplater.load(templatePath)
    const initialS1 = ppt.getCharts(1).length
    const initialS2 = ppt.getCharts(2).length

    // Pattern 1: Scoped slide context
    ppt
      .useSlide(1)
      .addChart('bar', { categories: ['A', 'B'], series: [{ name: 'S1', values: [1, 2] }] })

    // Pattern 2: Presentation-level without slideId
    ppt.addChart('line', { categories: ['X', 'Y'], series: [{ name: 'S2', values: [3, 4] }] })

    // Pattern 3: Presentation-level with explicit slideId
    ppt.addChart(2, 'pie', { categories: ['P1', 'P2'], series: [{ name: 'S3', values: [50, 50] }] })

    const s1Charts = ppt.getCharts(1)
    const s2Charts = ppt.getCharts(2)

    expect(s1Charts.length).toBe(initialS1 + 2)
    expect(s2Charts.length).toBe(initialS2 + 1)
  })
})
