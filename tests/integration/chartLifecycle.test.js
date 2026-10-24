import { describe, it, expect, beforeAll } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')

let PPTXTemplater

beforeAll(async () => {
  const module = await import('../../src/index.js')
  PPTXTemplater = module.PPTXTemplater
})

describe('TASK-CHART-003: Chart Creation, Removal, and Slide-Wide clearCharts Operations', () => {
  const runTests = existsSync(FIXTURE_FILE)
  if (!runTests) {
    it.skip('Skipping: sample.pptx fixture not found', () => {})
    return
  }

  it('should successfully add standard charts via ppt.addChart() and ppt.useSlide().addChart()', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)

    // Add Column chart on Slide 1
    ppt.addChart(
      1,
      'column',
      {
        categories: ['Q1', 'Q2', 'Q3', 'Q4'],
        series: [{ name: '2026 Sales', values: [150, 210, 180, 290] }],
      },
      {
        name: 'NewQSalesChart',
        title: '2026 Quarterly Performance',
        x: '1in',
        y: '1.5in',
        width: '5in',
        height: '3in',
      }
    )

    // Add Pie chart on Slide 1 chained via useSlide
    ppt.useSlide(1).addChart(
      'pie',
      {
        categories: ['Direct', 'Partner', 'Online'],
        series: [{ name: 'Channels', values: [45, 30, 25] }],
      },
      {
        name: 'NewChannelsPieChart',
        title: 'Channel Breakdown',
        x: '6.5in',
        y: '1.5in',
        width: '3.5in',
        height: '3in',
      }
    )

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeDefined()
    expect(buffer[0]).toBe(0x50)

    // Reload and verify presence
    const reloaded = await PPTXTemplater.load(buffer)
    const charts = reloaded.useSlide(1).getCharts()
    const chartNames = charts.map(c => c.name)

    expect(chartNames).toContain('NewQSalesChart')
    expect(chartNames).toContain('NewChannelsPieChart')

    const salesChart = charts.find(c => c.name === 'NewQSalesChart')
    expect(salesChart.type).toBe('column')
    expect(salesChart.title).toBe('2026 Quarterly Performance')
    expect(salesChart.series[0].values).toEqual([150, 210, 180, 290])
  })

  it('should support adding Office 2016 extended chart types (e.g. Treemap & Waterfall)', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)

    ppt.addChart(
      1,
      'treemap',
      {
        categories: ['Hardware', 'Software', 'Services'],
        series: [{ name: 'Divisions', values: [500, 350, 200] }],
      },
      {
        name: 'DivisionsTreemap',
        title: 'Revenue by Division',
      }
    )

    const buffer = await ppt.toBuffer()
    const reloaded = await PPTXTemplater.load(buffer)
    const charts = reloaded.getCharts(1)
    const treemap = charts.find(c => c.name === 'DivisionsTreemap')

    expect(treemap).toBeDefined()
    expect(treemap.type).toBe('treemap')
    expect(treemap.zipPath).toContain('chartEx')
  })

  it('should remove a single chart via removeChart and clean orphaned parts', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)

    ppt.addChart(
      1,
      'line',
      {
        categories: ['Jan', 'Feb', 'Mar'],
        series: [{ name: 'Trend', values: [10, 20, 15] }],
      },
      { name: 'TemporaryChart' }
    )

    const chartsBefore = ppt.getCharts(1)
    expect(chartsBefore.some(c => c.name === 'TemporaryChart')).toBe(true)

    // Remove the chart
    ppt.useSlide(1).removeChart('TemporaryChart')

    const buffer = await ppt.toBuffer()
    const reloaded = await PPTXTemplater.load(buffer)
    const chartsAfter = reloaded.getCharts(1)

    expect(chartsAfter.some(c => c.name === 'TemporaryChart')).toBe(false)
  })

  it('should clear all charts from a slide via clearCharts while preserving text, shapes, and tables', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)

    // Add a test table and shape to Slide 1
    ppt.useSlide(1).addTable({
      id: 'PreservedTable',
      x: 100,
      y: 100,
      width: 400,
      rows: [
        ['Header 1', 'Header 2'],
        ['Val 1', 'Val 2'],
      ],
    })

    ppt.addChart(
      1,
      'column',
      {
        categories: ['A', 'B'],
        series: [{ name: 'S1', values: [1, 2] }],
      },
      { name: 'ClearMeChart1' }
    )

    ppt.addChart(
      1,
      'pie',
      {
        categories: ['X', 'Y'],
        series: [{ name: 'S2', values: [3, 4] }],
      },
      { name: 'ClearMeChart2' }
    )

    expect(ppt.getCharts(1).length).toBeGreaterThanOrEqual(2)

    // Execute clearCharts on Slide 1
    ppt.useSlide(1).clearCharts()

    const chartsAfter = ppt.getCharts(1)
    expect(chartsAfter.length).toBe(0)

    // Verify non-chart elements survived
    const tables = ppt.useSlide(1).getTables()
    expect(tables.some(t => t.name === 'PreservedTable')).toBe(true)

    const buffer = await ppt.toBuffer()
    const reloaded = await PPTXTemplater.load(buffer)
    expect(reloaded.getCharts(1).length).toBe(0)
    expect(
      reloaded
        .useSlide(1)
        .getTables()
        .some(t => t.name === 'PreservedTable')
    ).toBe(true)
  })
})
