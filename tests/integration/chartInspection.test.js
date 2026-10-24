import { describe, it, expect, beforeAll } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')

describe('TASK-CHART-005: Chart Metadata Querying & Inspection Engine', () => {
  beforeAll(async () => {
    if (!existsSync(FIXTURE_FILE)) {
      const ppt = PPTXTemplater.create()
      ppt.addSlide(1)
      await ppt.saveAs(FIXTURE_FILE)
    }
  })

  it('should discover all charts across slides or for a targeted slide', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)

    // All slides charts
    const allCharts = ppt.getCharts()
    expect(Array.isArray(allCharts)).toBe(true)

    // Add a chart on slide 1
    ppt.addChart(
      1,
      'column',
      {
        categories: ['A', 'B'],
        series: [{ name: 'S1', values: [1, 2] }],
      },
      { name: 'DiscoveredCol' }
    )

    // Slide 1 charts
    const slide1Charts = ppt.getCharts(1)
    expect(Array.isArray(slide1Charts)).toBe(true)
    expect(slide1Charts.some(c => c.name === 'DiscoveredCol')).toBe(true)
    for (const c of slide1Charts) {
      expect(c.slideIndex).toBe(1)
    }

    // useSlide scoped
    const chainedCharts = ppt.useSlide(1).getCharts()
    expect(chainedCharts.length).toBe(slide1Charts.length)
  })

  it('should inspect normalized chart descriptors with all metadata properties', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)

    ppt.addChart(
      1,
      'bar',
      {
        categories: ['North', 'South', 'East', 'West'],
        series: [{ name: 'Regional Target', values: [120, 150, 90, 210] }],
      },
      {
        name: 'RegionalBarChart',
        title: 'Q1 Regional Target',
        x: '2in',
        y: '1.5in',
        width: '6in',
        height: '4in',
      }
    )

    const chart = ppt.getChart('RegionalBarChart')

    expect(chart).toBeDefined()
    expect(chart.chartId).toBe('RegionalBarChart')
    expect(chart.name).toBe('RegionalBarChart')
    expect(chart.type).toBe('bar')
    expect(chart.slideIndex).toBe(1)
    expect(chart.title).toBe('Q1 Regional Target')
    expect(chart.categories).toEqual(['North', 'South', 'East', 'West'])
    expect(chart.series.length).toBe(1)
    expect(chart.series[0].name).toBe('Regional Target')
    expect(chart.series[0].values).toEqual([120, 150, 90, 210])
    expect(chart.coordinates).toBeDefined()
    expect(typeof chart.coordinates.x).toBe('number')
    expect(typeof chart.coordinates.y).toBe('number')
    expect(typeof chart.coordinates.cx).toBe('number')
    expect(typeof chart.coordinates.cy).toBe('number')
    expect(chart.zipPath).toMatch(/^ppt\/charts\/chart\d+\.xml$/)
    expect(chart.workbookPath).toMatch(/^ppt\/embeddings\/Microsoft_Excel_Worksheet\d*\.xlsx$/)
  })

  it('should inspect Office 2016 extended chart types (e.g. Waterfall)', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)

    ppt.addChart(
      1,
      'waterfall',
      {
        categories: ['Gross', 'COGS', 'OpEx', 'Net'],
        series: [{ name: 'P&L', values: [5000, -2000, -1000, 2000] }],
      },
      {
        name: 'WaterfallInspection',
        title: 'FY26 Waterfall Breakdown',
      }
    )

    const chart = ppt.useSlide(1).getChart('WaterfallInspection')
    expect(chart).toBeDefined()
    expect(chart.name).toBe('WaterfallInspection')
    expect(chart.type).toBe('waterfall')
    expect(chart.title).toBe('FY26 Waterfall Breakdown')
    expect(chart.categories).toEqual(['Gross', 'COGS', 'OpEx', 'Net'])
    expect(chart.series[0].name).toBe('P&L')
    expect(chart.series[0].values).toEqual([5000, -2000, -1000, 2000])
  })

  it('should throw ChartNotFoundError when chart does not exist', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)

    expect(() => {
      ppt.getChart('NonExistentChartId')
    }).toThrow(/Chart "NonExistentChartId" not found/)

    expect(() => {
      ppt.useSlide(1).getChart('GhostChart')
    }).toThrow(/Chart "GhostChart" not found/)
  })
})
