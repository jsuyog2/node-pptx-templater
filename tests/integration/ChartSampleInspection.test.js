import { describe, it, expect } from 'vitest'
import path from 'path'
import { fileURLToPath } from 'url'
import { PPTXTemplater } from '../../src/index.js'
import { ZipManager } from '../../src/managers/ZipManager.js'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

describe('ChartPowerPointSample Inspection and Optional slideId Verification', () => {
  const samplePath = path.resolve(__dirname, '../../showcase/templates/ChartPowerPointSample.pptx')

  it('TASK-CHART-DIAG-001: should inspect all slides and chart structures in ChartPowerPointSample.pptx', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const slides = ppt.getSlides()

    expect(slides.length).toBeGreaterThan(0)
    console.log(`Inspected ChartPowerPointSample.pptx: ${slides.length} slide(s) found.`)

    for (let i = 0; i < slides.length; i++) {
      const slideNum = i + 1
      const charts = ppt.getCharts(slideNum)
      console.log(
        `Slide ${slideNum} (id: ${slides[i].slideId}, layout: ${slides[i].layout}): ${charts.length} chart(s)`
      )
      for (const c of charts) {
        console.log(`  - Chart "${c.name}" [${c.type}], shapeId: ${c.shapeId}, rId: ${c.rId}`)
      }
    }
  })

  it('TASK-CHART-OPT-002: should support optional slideId via useSlide(n).addChart() and ppt.addChart()', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Pattern A: Scoped slide context
    ppt.useSlide(1).addChart(
      'bar',
      {
        categories: ['Item A', 'Item B', 'Item C'],
        series: [{ name: 'Revenue', values: [120, 200, 150] }],
      },
      {
        name: 'ScopedBarChart',
        title: 'Scoped Context Revenue',
        x: 100,
        y: 100,
        width: 500,
        height: 350,
      }
    )

    // Pattern B: Direct presentation-level call without slideId (defaults to active slide / slide 1)
    ppt.addChart(
      'column',
      {
        categories: ['2023', '2024', '2025'],
        series: [{ name: 'Growth', values: [10, 25, 40] }],
      },
      {
        name: 'DirectColumnChart',
        title: 'Direct Call Growth',
        x: 650,
        y: 100,
        width: 500,
        height: 350,
      }
    )

    const slide1Charts = ppt.getCharts(1)
    const names = slide1Charts.map(c => c.name)
    expect(names).toContain('ScopedBarChart')
    expect(names).toContain('DirectColumnChart')

    const barChart = ppt.getChart('ScopedBarChart')
    expect(barChart.type).toBe('bar')
    expect(barChart.categories).toEqual(['Item A', 'Item B', 'Item C'])
  })

  it('TASK-CHART-FIX-003 & TASK-CHART-FIX-004: should properly insert graphicFrame with valid namespace, ordering, and ContentTypes', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    ppt.useSlide(1).addChart(
      'bar',
      {
        categories: ['Alpha', 'Beta'],
        series: [{ name: 'Values', values: [10, 20] }],
      },
      { name: 'RepairCheckBar', title: 'Integrity Check' }
    )

    const buffer = await ppt.toBuffer()
    const zip = new ZipManager()
    await zip.load(buffer)
    const slide1Xml = await zip.readFile('ppt/slides/slide1.xml')

    // 1. Check graphicFrame is present in slide1.xml
    expect(slide1Xml).toContain('RepairCheckBar')
    expect(slide1Xml).toContain('<p:graphicFrame>')

    // 2. Check namespace binding on <c:chart>
    expect(slide1Xml).toMatch(
      /<c:chart\b[^>]*xmlns:c="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/chart"/
    )

    // 3. Check schema order: if <p:extLst> exists in <p:spTree>, graphicFrame must be before <p:extLst>
    const spTreeEndIdx = slide1Xml.indexOf('</p:spTree>')
    const gfIdx = slide1Xml.indexOf('RepairCheckBar')
    expect(gfIdx).toBeLessThan(spTreeEndIdx)

    const spTreeSub = slide1Xml.slice(0, spTreeEndIdx)
    const extLstIdx = spTreeSub.lastIndexOf('<p:extLst')
    if (extLstIdx !== -1) {
      expect(gfIdx).toBeLessThan(extLstIdx)
    }

    // 4. Check ContentTypes contains chart and workbook
    const ctXml = await zip.readFile('[Content_Types].xml')
    expect(ctXml).toContain('application/vnd.openxmlformats-officedocument.drawingml.chart+xml')
    expect(ctXml).toContain('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    // 5. Check chart relationships point to embedded workbook
    const chartFiles = zip.listFiles('ppt/charts/').filter(f => !f.includes('_rels'))
    expect(chartFiles.length).toBeGreaterThan(0)
    const latestChart = chartFiles[chartFiles.length - 1]
    const relsXml = await zip.readFile(`ppt/charts/_rels/${latestChart.split('/').pop()}.rels`)
    expect(relsXml).toContain(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package'
    )
    expect(relsXml).toContain('Microsoft_Excel_Worksheet')
  })
})
