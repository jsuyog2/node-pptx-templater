import { describe, it, expect, beforeAll } from 'vitest'
import path from 'path'
import fs from 'fs-extra'
import JSZip from 'jszip'
import { PPTXTemplater } from '../../src/index.js'

describe('TASK-MULTI-CHART: Multi-Chart Batch Addition & Non-Colliding Package Integrity', () => {
  const templatePath = path.join(__dirname, '../../showcase/templates/ChartPowerPointSample.pptx')
  const testOutputDir = path.join(__dirname, '../../showcase/output/vitest_multi_charts')

  beforeAll(async () => {
    await fs.ensureDir(testOutputDir)
  })

  it('should add multiple charts on the same slide without rId, shape ID, workbook, or content type collisions', async () => {
    const ppt = await PPTXTemplater.load(templatePath)

    // Add 3 different charts to Slide 1
    ppt
      .useSlide(1)
      .addChart(
        'bar',
        { categories: ['Q1', 'Q2'], series: [{ name: 'Revenue', values: [100, 150] }] },
        { name: 'MultiBarChart', x: 50, y: 50, width: 300, height: 200 }
      )
    ppt
      .useSlide(1)
      .addChart(
        'line',
        { categories: ['Q1', 'Q2'], series: [{ name: 'Growth', values: [10, 25] }] },
        { name: 'MultiLineChart', x: 400, y: 50, width: 300, height: 200 }
      )
    ppt
      .useSlide(1)
      .addChart(
        'pie',
        { categories: ['Alpha', 'Beta'], series: [{ name: 'Share', values: [60, 40] }] },
        { name: 'MultiPieChart', x: 50, y: 300, width: 300, height: 200 }
      )

    const outPath = path.join(testOutputDir, 'multi_charts_same_slide.pptx')
    await ppt.saveToFile(outPath)
    expect(fs.existsSync(outPath)).toBe(true)

    const buffer = await fs.readFile(outPath)
    const zip = await JSZip.loadAsync(buffer)

    // 1. Check slide1 relationships
    const slide1Rels = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')
    const chartRelMatches = [
      ...slide1Rels.matchAll(/<Relationship[^>]*Id="([^"]+)"[^>]*Target="([^"]+)"/g),
    ]
    const rIds = chartRelMatches.map(m => m[1])
    const uniqueRIds = new Set(rIds)
    expect(uniqueRIds.size).toBe(rIds.length) // Zero duplicate rIds

    // 2. Check slide1 shape IDs
    const slide1Xml = await zip.file('ppt/slides/slide1.xml').async('text')
    const shapeIdMatches = [...slide1Xml.matchAll(/<p:cNvPr[^>]*id="(\d+)"/g)].map(m => m[1])
    const uniqueShapeIds = new Set(shapeIdMatches)
    expect(uniqueShapeIds.size).toBe(shapeIdMatches.length) // Zero duplicate shape IDs

    // 3. Check each chart has its own dedicated workbook
    const chartTargets = chartRelMatches
      .filter(m => m[2].includes('charts/'))
      .map(m => m[2].replace('../', 'ppt/'))

    expect(chartTargets.length).toBe(3)

    const workbookTargets = []
    for (const chartPath of chartTargets) {
      const relsPath = chartPath.replace('ppt/charts/', 'ppt/charts/_rels/') + '.rels'
      const chartRelContent = await zip.file(relsPath).async('text')
      const wbMatch = /Target="([^"]*embeddings\/[^"]+)"/.exec(chartRelContent)
      expect(wbMatch).not.toBeNull()
      workbookTargets.push(wbMatch[1])
    }

    const uniqueWorkbooks = new Set(workbookTargets)
    expect(uniqueWorkbooks.size).toBe(3) // All 3 workbooks must be distinct

    // 4. Verify Content_Types has an Override for each distinct chart and workbook
    const ct = await zip.file('[Content_Types].xml').async('text')
    for (const chartPath of chartTargets) {
      expect(ct).toContain(`PartName="/${chartPath}"`)
    }
    for (const wbRelTarget of workbookTargets) {
      const wbZipPath = wbRelTarget.replace('../', 'ppt/')
      expect(ct).toContain(`PartName="/${wbZipPath}"`)
      expect(zip.file(wbZipPath)).not.toBeNull()
    }
  })

  it('should support mixing standard and Office 2016+ extended charts across sequential slides', async () => {
    const ppt = await PPTXTemplater.load(templatePath)

    // Slide 1: Column + Waterfall
    ppt
      .useSlide(1)
      .addChart(
        'column',
        { categories: ['Mon', 'Tue'], series: [{ name: 'Clicks', values: [50, 80] }] },
        { name: 'Slide1Column' }
      )
    ppt.useSlide(1).addChart(
      'waterfall',
      {
        categories: ['Start', 'Gain', 'End'],
        series: [{ name: 'Flow', values: [100, 20, 120] }],
      },
      { name: 'Slide1Waterfall' }
    )

    // Slide 2: Treemap + Sunburst + Area
    ppt
      .useSlide(2)
      .addChart(
        'treemap',
        { categories: ['A', 'B'], series: [{ name: 'Size', values: [30, 70] }] },
        { name: 'Slide2Treemap' }
      )
    ppt
      .useSlide(2)
      .addChart(
        'sunburst',
        { categories: ['X', 'Y'], series: [{ name: 'Share', values: [40, 60] }] },
        { name: 'Slide2Sunburst' }
      )
    ppt
      .useSlide(2)
      .addChart(
        'area',
        { categories: ['Jan', 'Feb'], series: [{ name: 'Trend', values: [5, 15] }] },
        { name: 'Slide2Area' }
      )

    const outPath = path.join(testOutputDir, 'multi_charts_mixed_slides.pptx')
    await ppt.saveToFile(outPath)
    expect(fs.existsSync(outPath)).toBe(true)

    const buffer = await fs.readFile(outPath)
    const zip = await JSZip.loadAsync(buffer)

    // Check Content_Types integrity for extended and standard types
    const ct = await zip.file('[Content_Types].xml').async('text')
    expect(ct).toContain('application/vnd.ms-office.chartex+xml')
    expect(ct).toContain('application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
    expect(ct).toContain('application/vnd.ms-office.chartstyle+xml')
    expect(ct).toContain('application/vnd.ms-office.chartcolorstyle+xml')

    // Check slide 2 has non-colliding graphic frames
    const slide2Xml = await zip.file('ppt/slides/slide2.xml').async('text')
    // Extract shape IDs specifically from the graphic frames / shapes
    const gfIds = [...slide2Xml.matchAll(/<p:graphicFrame>[\s\S]*?<p:cNvPr[^>]*id="(\d+)"/g)].map(
      m => m[1]
    )
    expect(gfIds.length).toBe(3)
    expect(new Set(gfIds).size).toBe(3) // 3 distinct shape IDs for the 3 added charts
    expect(gfIds).toEqual(['4', '5', '6'])
  })
})
