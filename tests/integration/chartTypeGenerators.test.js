import { describe, it, expect } from 'vitest'
import { ChartCacheGenerator } from '../../src/elements/charts/ChartCacheGenerator.js'
import { ExtendedChartGenerator } from '../../src/elements/charts/ExtendedChartGenerator.js'

describe('TASK-CHART-002: Extended Chart Types Architecture & XML Generator Support', () => {
  const sampleData = {
    categories: ['North', 'South', 'East', 'West'],
    series: [
      { name: 'Revenue', values: [120, 180, 240, 300] },
      { name: 'Profit', values: [40, 60, 90, 110] },
    ],
  }

  const all17Types = [
    { type: 'column', expectedTag: '<c:barChart>', expectedDir: '<c:barDir val="col"/>' },
    { type: 'col', expectedTag: '<c:barChart>', expectedDir: '<c:barDir val="col"/>' }, // alias
    { type: 'bar', expectedTag: '<c:barChart>', expectedDir: '<c:barDir val="bar"/>' },
    { type: 'line', expectedTag: '<c:lineChart>' },
    { type: 'pie', expectedTag: '<c:pieChart>' },
    { type: 'area', expectedTag: '<c:areaChart>' },
    { type: 'scatter', expectedTag: '<c:scatterChart>' },
    { type: 'xy', expectedTag: '<c:scatterChart>' }, // alias
    { type: 'xyScatter', expectedTag: '<c:scatterChart>' }, // alias
    { type: 'map', expectedTag: '<cx:chartSpace', layoutId: 'regionMap' },
    { type: 'stock', expectedTag: '<c:stockChart>' },
    { type: 'surface', expectedTag: '<c:surfaceChart>' },
    { type: 'radar', expectedTag: '<c:radarChart>' },
    { type: 'treemap', expectedTag: '<cx:chartSpace', layoutId: 'treemap' },
    { type: 'sunburst', expectedTag: '<cx:chartSpace', layoutId: 'sunburst' },
    { type: 'histogram', expectedTag: '<cx:chartSpace', layoutId: 'clusteredColumn' },
    { type: 'boxWhisker', expectedTag: '<cx:chartSpace', layoutId: 'boxWhisker' },
    { type: 'boxAndWhisker', expectedTag: '<cx:chartSpace', layoutId: 'boxWhisker' }, // alias
    { type: 'box_and_whisker', expectedTag: '<cx:chartSpace', layoutId: 'boxWhisker' }, // alias
    { type: 'waterfall', expectedTag: '<cx:chartSpace', layoutId: 'waterfall' },
    { type: 'funnel', expectedTag: '<cx:chartSpace', layoutId: 'funnel' },
    { type: 'combo', expectedTag: '<c:barChart>' },
  ]

  it('should normalize all chart types and aliases to canonical names', () => {
    expect(ChartCacheGenerator.normalizeChartType('col')).toBe('column')
    expect(ChartCacheGenerator.normalizeChartType('COLUMN')).toBe('column')
    expect(ChartCacheGenerator.normalizeChartType('xy')).toBe('scatter')
    expect(ChartCacheGenerator.normalizeChartType('xyScatter')).toBe('scatter')
    expect(ChartCacheGenerator.normalizeChartType('boxAndWhisker')).toBe('boxWhisker')
    expect(ChartCacheGenerator.normalizeChartType('box_and_whisker')).toBe('boxWhisker')
    expect(ChartCacheGenerator.normalizeChartType('line')).toBe('line')
    expect(ChartCacheGenerator.normalizeChartType('treemap')).toBe('treemap')
  })

  it('should correctly identify Office 2016 extended chart types', () => {
    expect(ChartCacheGenerator.isExtendedChart('treemap')).toBe(true)
    expect(ChartCacheGenerator.isExtendedChart('sunburst')).toBe(true)
    expect(ChartCacheGenerator.isExtendedChart('histogram')).toBe(true)
    expect(ChartCacheGenerator.isExtendedChart('boxWhisker')).toBe(true)
    expect(ChartCacheGenerator.isExtendedChart('boxAndWhisker')).toBe(true)
    expect(ChartCacheGenerator.isExtendedChart('waterfall')).toBe(true)
    expect(ChartCacheGenerator.isExtendedChart('funnel')).toBe(true)
    expect(ChartCacheGenerator.isExtendedChart('map')).toBe(true)
    expect(ChartCacheGenerator.isExtendedChart('regionMap')).toBe(true)
    expect(ChartCacheGenerator.isExtendedChart('geochart')).toBe(true)

    expect(ChartCacheGenerator.isExtendedChart('column')).toBe(false)
    expect(ChartCacheGenerator.isExtendedChart('bar')).toBe(false)
    expect(ChartCacheGenerator.isExtendedChart('line')).toBe(false)
    expect(ChartCacheGenerator.isExtendedChart('pie')).toBe(false)
    expect(ChartCacheGenerator.isExtendedChart('scatter')).toBe(false)
    expect(ChartCacheGenerator.isExtendedChart('combo')).toBe(false)
  })

  for (const { type, expectedTag, expectedDir, layoutId } of all17Types) {
    it(`should generate valid schema-compliant XML for chart type: "${type}"`, () => {
      const xml = ChartCacheGenerator.generateChartSpaceXml({
        type,
        data: sampleData,
        title: `Test ${type} Chart`,
      })

      expect(xml).toBeDefined()
      expect(xml).toContain(expectedTag)
      expect(xml).toContain(`Test ${type} Chart`)

      if (expectedDir) {
        expect(xml).toContain(expectedDir)
      }
      if (layoutId) {
        expect(xml).toContain(`layoutId="${layoutId}"`)
        expect(xml).toContain('xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"')
      } else {
        expect(xml).toContain('xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"')
      }
    })
  }

  it('should generate multi-plot layout for Combo charts', () => {
    const comboXml = ChartCacheGenerator.generateChartSpaceXml({
      type: 'combo',
      data: {
        categories: ['2023', '2024', '2025', '2026'],
        series: [
          { name: 'Revenue', values: [100, 200, 300, 400], type: 'column' },
          { name: 'Growth', values: [10, 15, 20, 25], type: 'line' },
        ],
      },
      title: 'Revenue & Growth Combo',
    })

    expect(comboXml).toContain('<c:barChart>')
    expect(comboXml).toContain('<c:lineChart>')
    expect(comboXml).toContain('Revenue')
    expect(comboXml).toContain('Growth')
  })

  it('should generate valid Scatter charts with xVal and yVal numeric references', () => {
    const scatterXml = ChartCacheGenerator.generateChartSpaceXml({
      type: 'scatter',
      data: {
        categories: ['1', '2', '3', '4'],
        series: [{ name: 'Points', values: [10, 25, 40, 80] }],
      },
    })

    expect(scatterXml).toContain('<c:scatterChart>')
    expect(scatterXml).toContain('<c:xVal>')
    expect(scatterXml).toContain('<c:yVal>')
  })

  it('should support updating extended chart spaces through ExtendedChartGenerator', () => {
    const initialXml = ExtendedChartGenerator.generateChartSpaceXml({
      type: 'treemap',
      data: sampleData,
      title: 'Original Title',
    })

    const updatedXml = ExtendedChartGenerator.updateChartSpaceXml(
      initialXml,
      {
        categories: ['North-Updated', 'South-Updated', 'East-Updated', 'West-Updated'],
        series: [{ name: 'Revenue-2026', values: [500, 600, 700, 800] }],
      },
      { title: 'New Title' }
    )

    expect(updatedXml).toContain('North-Updated')
    expect(updatedXml).toContain('Revenue-2026')
    expect(updatedXml).toContain('800')
    expect(updatedXml).toContain('New Title')
  })
})
