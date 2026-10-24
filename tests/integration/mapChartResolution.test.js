import { describe, it, expect, beforeAll } from 'vitest'
import path from 'path'
import fs from 'fs-extra'
import JSZip from 'jszip'
import { testMapChart } from '../../showcase/charts_operations.js'

describe('TASK-MAP-DATA-RESOLVE: Map Chart Bing Maps Offline/Online Resolution Verification', () => {
  const templatePath = path.join(__dirname, '../../showcase/templates/ChartPowerPointSample.pptx')
  const testOutputDir = path.join(__dirname, '../../showcase/output/vitest_map_resolution')

  beforeAll(async () => {
    await fs.ensureDir(testOutputDir)
  })

  it('should generate a map chart with pre-cached Bing geometry and valid chartex namespace to prevent online retrieval failure', async () => {
    const outPath = await testMapChart(templatePath, testOutputDir)
    expect(fs.existsSync(outPath)).toBe(true)

    const buffer = await fs.readFile(outPath)
    const zip = await JSZip.loadAsync(buffer)

    // 1. Check Slide 1 contains cx4 choice for map chart
    const slide1Xml = await zip.file('ppt/slides/slide1.xml').async('text')
    expect(slide1Xml).toContain('Requires="cx4"')
    expect(slide1Xml).toContain('http://schemas.microsoft.com/office/drawing/2016/5/10/chartex')

    // 2. Identify the chartEx part
    const slide1Rels = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')
    const chartMatch = slide1Rels.match(/Target="\.\.\/charts\/(chartEx\d+\.xml)"/)
    expect(chartMatch).not.toBeNull()
    const chartExPath = `ppt/charts/${chartMatch[1]}`

    // 3. Check chartEx XML has layoutId="regionMap", numDim type="colorVal", and geoCache
    const chartXml = await zip.file(chartExPath).async('text')
    expect(chartXml).toContain('layoutId="regionMap"')
    expect(chartXml).toContain('type="colorVal"')
    expect(chartXml).toContain('<cx:geography')
    expect(chartXml).toContain('attribution="Powered by Bing"')
    expect(chartXml).toContain('<cx:geoCache')
    expect(chartXml).toContain('provider="{E9337A44-BEBE-4D9F-B70C-5C5E7DAFC167}"')
    expect(chartXml).toContain('<cx:binary>')

    // 4. Verify embedded Excel sheet has matching dimension and entity rows
    const chartRels = await zip.file(`ppt/charts/_rels/${chartMatch[1]}.rels`).async('text')
    const wbMatch = chartRels.match(
      /Target="\.\.\/embeddings\/(Microsoft_Excel_Worksheet\d*\.xlsx)"/
    )
    expect(wbMatch).not.toBeNull()
    const wbBuffer = await zip.file(`ppt/embeddings/${wbMatch[1]}`).async('nodebuffer')
    const wbZip = await JSZip.loadAsync(wbBuffer)
    const sheet1Xml = await wbZip.file('xl/worksheets/sheet1.xml').async('text')
    expect(sheet1Xml).toContain('United States')
    expect(sheet1Xml).toContain('Germany')
    expect(sheet1Xml).toContain('85')
  })
})
