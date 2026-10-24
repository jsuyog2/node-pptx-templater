import { describe, it, expect, beforeAll } from 'vitest'
import path from 'path'
import fs from 'fs-extra'
import JSZip from 'jszip'
import { PPTXTemplater } from '../../src/index.js'
import {
  testBoxWhiskerChart,
  testFunnelChart,
  testHistogramChart,
  testMapChart,
  testSunburstChart,
  testTreemapChart,
  testWaterfallChart,
} from '../../showcase/charts_operations.js'

describe('TASK-CX-REPAIR-004: Office 2016+ Extended Charts Reverse-Engineered Schema Verification', () => {
  const templatePath = path.join(__dirname, '../../showcase/templates/ChartPowerPointSample.pptx')
  const testOutputDir = path.join(__dirname, '../../showcase/output/vitest_extended_charts')

  beforeAll(async () => {
    await fs.ensureDir(testOutputDir)
  })

  const extendedCharts = [
    {
      type: 'box_and_whisker',
      fn: testBoxWhiskerChart,
      layoutId: 'boxWhisker',
      axesExpected: ['0', '1'],
      expectedLayoutPr: 'quartileMethod="exclusive"',
    },
    {
      type: 'funnel',
      fn: testFunnelChart,
      layoutId: 'funnel',
      axesExpected: ['1'],
    },
    {
      type: 'histogram',
      fn: testHistogramChart,
      layoutId: 'clusteredColumn',
      axesExpected: ['0', '1'],
      expectedLayoutPr: '<cx:binning',
    },
    {
      type: 'sunburst',
      fn: testSunburstChart,
      layoutId: 'sunburst',
      axesExpected: [],
    },
    {
      type: 'treemap',
      fn: testTreemapChart,
      layoutId: 'treemap',
      axesExpected: [],
      expectedLayoutPr: 'parentLabelLayout val="overlapping"',
    },
    {
      type: 'waterfall',
      fn: testWaterfallChart,
      layoutId: 'waterfall',
      axesExpected: ['0', '1'],
    },
    {
      type: 'map',
      fn: testMapChart,
      layoutId: 'regionMap',
      axesExpected: [],
      expectedLayoutPr: 'attribution="Powered by Bing"',
    },
  ]

  for (const { type, fn, layoutId, axesExpected, expectedLayoutPr } of extendedCharts) {
    it(`should build valid, repair-free OpenXML package for extended chart: "${type}"`, async () => {
      const outPath = await fn(templatePath, testOutputDir)
      expect(fs.existsSync(outPath)).toBe(true)

      const buffer = await fs.readFile(outPath)
      const zip = await JSZip.loadAsync(buffer)

      // 1. Slide Relationships: locate the chart added to slide 1
      const slideRelsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')
      expect(slideRelsXml).toContain(
        'http://schemas.microsoft.com/office/2014/relationships/chartEx'
      )
      const targetMatch = slideRelsXml.match(/Target="\.\.\/charts\/(chartEx\d+\.xml)"/)
      expect(targetMatch).not.toBeNull()
      const chartFileName = targetMatch[1]

      // 2. Content Types Registration (chartEx, style, colors, workbook)
      const contentTypesXml = await zip.file('[Content_Types].xml').async('text')
      expect(contentTypesXml).toContain('application/vnd.ms-office.chartex+xml')
      expect(contentTypesXml).toContain('application/vnd.ms-office.chartstyle+xml')
      expect(contentTypesXml).toContain('application/vnd.ms-office.chartcolorstyle+xml')
      expect(contentTypesXml).toContain(`/ppt/charts/${chartFileName}`)
      expect(contentTypesXml).toContain(
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      )

      // 3. Slide GraphicFrame enclosed in mc:AlternateContent
      const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')
      expect(slideXml).toContain('<mc:AlternateContent')
      expect(slideXml).toContain('Requires="cx1"')
      expect(slideXml).toContain('<p:graphicFrame>')
      expect(slideXml).toContain('http://schemas.microsoft.com/office/drawing/2014/chartex')
      expect(slideXml).toContain('<cx:chart')
      expect(slideXml).toContain(
        'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"'
      )
      expect(slideXml).toContain('<mc:Fallback>')

      // 4. ChartEx Part Relationships pointing to Excel workbook, style, and colors
      const chartRelsXml = await zip.file(`ppt/charts/_rels/${chartFileName}.rels`).async('text')
      expect(chartRelsXml).toContain(
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package'
      )
      expect(chartRelsXml).toContain(
        'http://schemas.microsoft.com/office/2011/relationships/chartStyle'
      )
      expect(chartRelsXml).toContain(
        'http://schemas.microsoft.com/office/2011/relationships/chartColorStyle'
      )

      // 5. ChartEx XML Schema Integrity
      const chartExXml = await zip.file(`ppt/charts/${chartFileName}`).async('text')
      expect(chartExXml).toContain(
        'xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"'
      )
      expect(chartExXml).toContain(`layoutId="${layoutId}"`)
      expect(chartExXml).toContain('cx:autoUpdate="0"')

      // externalData MUST be inside chartData before data id="0"
      expect(chartExXml).toContain('<cx:chartData>')
      const extDataIdx = chartExXml.indexOf('<cx:externalData')
      const dataIdIdx = chartExXml.indexOf('<cx:data id="0">')
      const chartEndIdx = chartExXml.indexOf('</cx:chart>')
      expect(extDataIdx).toBeGreaterThan(-1)
      expect(dataIdIdx).toBeGreaterThan(extDataIdx)
      expect(chartEndIdx).toBeGreaterThan(dataIdIdx)

      // Ensure NO trailing externalData after </cx:chart>
      const trailingExtDataIdx = chartExXml.indexOf('<cx:externalData', chartEndIdx)
      expect(trailingExtDataIdx).toBe(-1)

      // Verify layout properties
      if (expectedLayoutPr) {
        expect(chartExXml).toContain(expectedLayoutPr)
      }

      // Verify axes presence or omission
      if (axesExpected.length === 0) {
        expect(chartExXml).not.toContain('<cx:axis id=')
      } else {
        for (const axId of axesExpected) {
          expect(chartExXml).toContain(`<cx:axis id="${axId}">`)
        }
      }

      // 6. Backing Excel Workbook Integrity
      const wbRelMatch = chartRelsXml.match(
        /Target="\.\.\/embeddings\/(Microsoft_Excel_Worksheet\d*\.xlsx)"/
      )
      expect(wbRelMatch).not.toBeNull()
      const wbFileName = wbRelMatch[1]
      const wbFile = zip.file(`ppt/embeddings/${wbFileName}`)
      expect(wbFile).not.toBeNull()
      const wbBuffer = await wbFile.async('nodebuffer')
      const wbZip = await JSZip.loadAsync(wbBuffer)
      expect(wbZip.file('xl/workbook.xml')).toBeDefined()
      expect(wbZip.file('xl/worksheets/sheet1.xml')).toBeDefined()
    })
  }

  it('should support multi-level category hierarchies for Treemap and Sunburst', async () => {
    const ppt = await PPTXTemplater.load(templatePath)
    ppt.useSlide(1).addChart(
      'treemap',
      {
        hierarchyLevels: [
          ['North', 'North', 'South', 'South'],
          ['Laptops', 'Phones', 'Laptops', 'Phones'],
        ],
        series: [{ name: 'Revenue', values: [1200, 800, 950, 700] }],
      },
      { name: 'HierarchicalTreemap' }
    )

    const outPath = path.join(testOutputDir, 'chart_hierarchical_treemap.pptx')
    await ppt.saveToFile(outPath)

    const buffer = await fs.readFile(outPath)
    const zip = await JSZip.loadAsync(buffer)
    const slideRelsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')
    const targetMatch = slideRelsXml.match(/Target="\.\.\/charts\/(chartEx\d+\.xml)"/)
    expect(targetMatch).not.toBeNull()
    const chartFileName = targetMatch[1]
    const chartXml = await zip.file(`ppt/charts/${chartFileName}`).async('text')

    expect(chartXml).toContain('ptCount="4"')
    expect(chartXml).toContain('Laptops')
    expect(chartXml).toContain('Phones')
  })
})
