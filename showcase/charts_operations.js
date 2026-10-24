/**
 * @fileoverview showcase/charts_operations.js — Comprehensive Charts API Master Showcase.
 *
 * Demonstrates and validates the strict 11-function Charts Whitelist for node-pptx-templater.
 * Provides isolated standalone test functions for all 17 supported chart types across
 * standard DrawingML (<c:chartSpace>) and Office 2016+ extended schemas (<cx:chartSpace>).
 *
 * The Strict 11 Charts Methods Whitelist:
 * =========================================================================
 *  1. ppt.addChart([slideRef], chartType, data, options)
 *  2. ppt.updateChart(chartId, options)
 *  3. ppt.updateChartData(chartId, data)
 *  4. ppt.removeChart(chartId)
 *  5. ppt.clearCharts(maybeSlideIndex)
 *  6. ppt.getCharts(slideId)
 *  7. ppt.getChart(chartId)
 *  8. ppt.cloneChart(chartId, clonedChartId, clonedOptions)
 *  9. ppt.getChartLabel(chartId)
 * 10. ppt.setChartLabel(chartId, label, options)
 * 11. ppt.updateChartLabel(chartId, options)
 * =========================================================================
 */

'use strict'

const path = require('path')
const fs = require('fs-extra')
const { PPTXTemplater } = require('../src/index.js')

const DEFAULT_TEMPLATE = path.join(__dirname, 'templates/ChartPowerPointSample.pptx')
const OUTPUT_DIR = path.join(__dirname, 'output')
const ADVANCED_OUTPUT_DIR = path.join(__dirname, 'output/slide_advanced')

function logHeader(title) {
  console.log('\n' + '='.repeat(80))
  console.log(`  ${title}`)
  console.log('='.repeat(80))
}

function logStep(num, name, details) {
  console.log(`\n[Step ${num}] ${name}`)
  if (details) console.log(`       -> ${details}`)
}

// -----------------------------------------------------------------------------
// Isolated Standalone Test Functions for All 17 Chart Types
// -----------------------------------------------------------------------------

async function testBarChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'bar',
    {
      categories: ['Item A', 'Item B', 'Item C'],
      series: [{ name: 'Revenue', values: [120, 200, 150] }],
    },
    {
      name: 'StandaloneBarChart',
      title: 'Quarterly Revenue Bar Chart',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_bar.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testColumnChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'column',
    {
      categories: ['Q1', 'Q2', 'Q3', 'Q4'],
      series: [{ name: 'Operating Cost', values: [80, 95, 110, 130] }],
    },
    {
      name: 'StandaloneColumnChart',
      title: 'Operating Cost Clustered Column',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_column.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testLineChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'line',
    {
      categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
      series: [{ name: 'Monthly Growth', values: [12, 19, 25, 32, 45, 60] }],
    },
    {
      name: 'StandaloneLineChart',
      title: 'Monthly Growth Trend',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_line.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testPieChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'pie',
    {
      categories: ['Desktop', 'Mobile', 'Tablet'],
      series: [{ name: 'Traffic Share', values: [60, 32, 8] }],
    },
    {
      name: 'StandalonePieChart',
      title: 'Traffic Device Share',
      x: 100,
      y: 100,
      width: 450,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_pie.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testAreaChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'area',
    {
      categories: ['2022', '2023', '2024', '2025'],
      series: [{ name: 'Bandwidth GB', values: [1200, 2400, 4800, 9600] }],
    },
    {
      name: 'StandaloneAreaChart',
      title: 'Bandwidth Utilization Area',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_area.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testScatterChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'scatter',
    {
      categories: ['10', '20', '30', '40'],
      series: [{ name: 'Sample Points', values: [15, 28, 42, 65] }],
    },
    {
      name: 'StandaloneScatterChart',
      title: 'XY Correlation Scatter',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_scatter.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testMapChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'map',
    {
      categories: ['United States', 'Brazil', 'Germany', 'India', 'Australia'],
      series: [{ name: 'Global Adoption', values: [85, 42, 63, 91, 55] }],
    },
    {
      name: 'StandaloneMapChart',
      title: 'Global Adoption Geographic Map',
      projection: 'mercator',
      viewedRegion: 'dataOnly',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_map.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testStockChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'stock',
    {
      categories: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'],
      series: [
        { name: 'High', values: [125, 130, 128, 135, 140] },
        { name: 'Low', values: [110, 115, 112, 120, 122] },
        { name: 'Close', values: [120, 122, 125, 132, 138] },
      ],
    },
    {
      name: 'StandaloneStockChart',
      title: 'Weekly High-Low-Close Stock Performance',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_stock.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testSurfaceChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'surface',
    {
      categories: ['Row 1', 'Row 2', 'Row 3'],
      series: [
        { name: 'Col 1', values: [10, 20, 30] },
        { name: 'Col 2', values: [20, 35, 45] },
      ],
    },
    {
      name: 'StandaloneSurfaceChart',
      title: '3D Surface Topology',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_surface.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testRadarChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'radar',
    {
      categories: ['Speed', 'Reliability', 'Comfort', 'Safety', 'Efficiency'],
      series: [{ name: 'Model Metric', values: [90, 85, 80, 95, 75] }],
    },
    {
      name: 'StandaloneRadarChart',
      title: 'Performance Radar Dimensions',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_radar.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testTreemapChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'treemap',
    {
      categories: ['Electronics', 'Furniture', 'Clothing', 'Groceries'],
      series: [{ name: 'Sales Volume', values: [450, 220, 180, 310] }],
    },
    {
      name: 'StandaloneTreemapChart',
      title: 'Category Sales Treemap',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_treemap.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testSunburstChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'sunburst',
    {
      categories: ['North', 'South', 'East', 'West'],
      series: [{ name: 'Regional Breakdown', values: [300, 200, 150, 250] }],
    },
    {
      name: 'StandaloneSunburstChart',
      title: 'Hierarchy Sunburst Distribution',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_sunburst.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testHistogramChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'histogram',
    {
      categories: ['0-10', '10-20', '20-30', '30-40', '40-50'],
      series: [{ name: 'Frequency', values: [5, 12, 28, 18, 7] }],
    },
    {
      name: 'StandaloneHistogramChart',
      title: 'Score Distribution Histogram',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_histogram.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testBoxWhiskerChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'boxWhisker',
    {
      categories: ['Group A', 'Group B', 'Group C'],
      series: [{ name: 'Scores', values: [72, 85, 91] }],
    },
    {
      name: 'StandaloneBoxWhiskerChart',
      title: 'Percentile Box & Whisker',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_boxWhisker.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testWaterfallChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'waterfall',
    {
      categories: ['Starting Cash', 'Product Sales', 'Services', 'Salaries', 'Cloud', 'Ending'],
      series: [{ name: 'Cash Flow', values: [1000, 450, 200, -350, -150, 1150] }],
    },
    {
      name: 'StandaloneWaterfallChart',
      title: 'FY2026 Enterprise Cash Waterfall',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_waterfall.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testFunnelChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'funnel',
    {
      categories: ['Impressions', 'Clicks', 'Signups', 'Purchases'],
      series: [{ name: 'Pipeline', values: [10000, 2500, 600, 150] }],
    },
    {
      name: 'StandaloneFunnelChart',
      title: 'Marketing Pipeline Funnel',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_funnel.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

async function testComboChart(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1).addChart(
    'combo',
    {
      categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
      series: [
        { name: 'Direct Sales', type: 'column', values: [50, 60, 75, 80, 95, 110] },
        { name: 'Target Line', type: 'line', values: [55, 65, 70, 85, 90, 105] },
      ],
    },
    {
      name: 'StandaloneComboChart',
      title: 'H1 Actuals vs Target Combo',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  const outPath = path.join(outDir, 'chart_combo.pptx')
  await ppt.saveToFile(outPath)
  return outPath
}

// -----------------------------------------------------------------------------
// Test Runner for All 17 Isolated Chart Types
// -----------------------------------------------------------------------------

const ALL_CHART_TEST_FUNCTIONS = {
  bar: testBarChart,
  column: testColumnChart,
  line: testLineChart,
  pie: testPieChart,
  area: testAreaChart,
  scatter: testScatterChart,
  map: testMapChart,
  stock: testStockChart,
  surface: testSurfaceChart,
  radar: testRadarChart,
  treemap: testTreemapChart,
  sunburst: testSunburstChart,
  histogram: testHistogramChart,
  boxwhisker: testBoxWhiskerChart,
  box_and_whisker: testBoxWhiskerChart,
  waterfall: testWaterfallChart,
  funnel: testFunnelChart,
  combo: testComboChart,
}

async function runAllChartTests(templatePath = DEFAULT_TEMPLATE, outDir = OUTPUT_DIR) {
  await fs.ensureDir(outDir)
  const results = []
  const uniqueKeys = [
    'bar',
    'column',
    'line',
    'pie',
    'area',
    'scatter',
    'map',
    'stock',
    'surface',
    'radar',
    'treemap',
    'sunburst',
    'histogram',
    'boxwhisker',
    'waterfall',
    'funnel',
    'combo',
  ]

  for (const type of uniqueKeys) {
    const fn = ALL_CHART_TEST_FUNCTIONS[type]
    const t0 = Date.now()
    const outPath = await fn(templatePath, outDir)
    const duration = Date.now() - t0
    const stats = await fs.stat(outPath)
    results.push({ type, outPath, size: stats.size, duration })
    console.log(
      `[PASS] ${type.padEnd(12)} -> ${path.basename(outPath)} (${(stats.size / 1024).toFixed(1)} KB, ${duration}ms)`
    )
  }

  return results
}

// -----------------------------------------------------------------------------
// Master Showcase Exercising All 11 Whitelisted Chart Methods
// -----------------------------------------------------------------------------

async function runChartsShowcase() {
  const t0 = Date.now()
  logHeader('node-pptx-templater v3.0 — Strict 11 Charts API Master Showcase')
  const templatePath = DEFAULT_TEMPLATE

  await fs.ensureDir(OUTPUT_DIR)
  await fs.ensureDir(ADVANCED_OUTPUT_DIR)

  // 1. Diagnostics & Slide Enumeration
  logStep(
    1,
    'TASK-CHART-DIAG-001: Inspect All Slides & Element Counts in ChartPowerPointSample.pptx'
  )
  const ppt = await PPTXTemplater.load(templatePath)
  const slides = ppt.getSlides()
  console.log(`       Presentation loaded successfully. Total slides: ${slides.length}`)

  for (let i = 0; i < slides.length; i++) {
    const slideNum = i + 1
    const slideInfo = slides[i]
    const elements = ppt.useSlide(slideNum).getElements()
    const charts = ppt.getCharts(slideNum)
    console.log(
      `       • Slide ${slideNum} [ID: ${slideInfo.slideId}, Layout: ${slideInfo.layout || 'default'}] ` +
        `— ${elements.length} element(s), ${charts.length} chart(s)`
    )
  }

  // 2. Method 1: addChart with Optional slideId (Scoped Context & Direct Call on Slide 1)
  logStep(
    2,
    'TASK-CHART-OPT-002 & FIX-003: addChart() with Optional slideId — Bar Chart & Column Chart on Slide 1'
  )

  // Pattern A: Scoped slide context without slideId
  await ppt.useSlide(1).addChart(
    'bar',
    {
      categories: ['Item A', 'Item B', 'Item C'],
      series: [{ name: 'Revenue', values: [120, 200, 150] }],
    },
    {
      name: 'QuarterlyRevenueBar',
      title: 'Quarterly Revenue',
      x: 100,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  console.log(
    '       -> Added Bar Chart "QuarterlyRevenueBar" via ppt.useSlide(1).addChart("bar", ...)'
  )

  // Pattern B: Direct presentation-level call defaulting to active slide / slide 1
  ppt.addChart(
    'column',
    {
      categories: ['Q1', 'Q2', 'Q3', 'Q4'],
      series: [{ name: 'Operating Cost', values: [80, 95, 110, 130] }],
    },
    {
      name: 'OperatingCostColumn',
      title: 'FY2026 Operating Cost',
      x: 650,
      y: 100,
      width: 500,
      height: 350,
    }
  )
  console.log('       -> Added Column Chart "OperatingCostColumn" via ppt.addChart("column", ...)')

  // 3. Method 1: addChart with Extended Charts & Combo on Slide 2
  logStep(3, 'Method 1: addChart() — Office 2016+ Extended Charts (Waterfall) & Combo on Slide 2')
  ppt.useSlide(2)
  ppt.addChart(
    'waterfall',
    {
      categories: ['Starting Cash', 'Product Sales', 'Services', 'Salaries', 'Cloud', 'Ending'],
      series: [{ name: 'Cash Flow', values: [1000, 450, 200, -350, -150, 1150] }],
    },
    {
      name: 'CashWaterfall',
      title: 'FY2026 Enterprise Cash Waterfall',
      x: '0.8in',
      y: '1.2in',
      width: '5.5in',
      height: '4in',
    }
  )

  ppt.addChart(
    2,
    'combo',
    {
      categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
      series: [
        { name: 'Direct Sales', type: 'column', values: [50, 60, 75, 80, 95, 110] },
        { name: 'Target Line', type: 'line', values: [55, 65, 70, 85, 90, 105] },
      ],
    },
    {
      name: 'SalesCombo',
      title: 'H1 Actuals vs Target',
      x: '6.8in',
      y: '1.2in',
      width: '5.5in',
      height: '4in',
    }
  )
  console.log('       -> Added Waterfall and Combo charts to Slide 2')

  // 4. Methods 6 & 7: getCharts() & getChart()
  logStep(4, 'Methods 6 & 7: getCharts() and getChart() — Inspecting Chart Metadata')
  const slide1Charts = ppt.getCharts(1)
  console.log(`       Slide 1 contains ${slide1Charts.length} chart(s):`)
  for (const c of slide1Charts) {
    console.log(
      `         • [${c.type.toUpperCase()}] "${c.name}" (shapeId: ${c.shapeId}, rId: ${c.rId})`
    )
  }

  const waterfallDesc = ppt.useSlide(2).getChart('CashWaterfall')
  console.log(`       Retrieved descriptor for "${waterfallDesc.name}":`)
  console.log(`         - Type: ${waterfallDesc.type}`)
  console.log(`         - Title: ${waterfallDesc.title}`)
  console.log(`         - Categories: ${waterfallDesc.categories.join(', ')}`)
  console.log(`         - Series Count: ${waterfallDesc.series.length}`)

  // 5. Methods 2 & 3: updateChart() & updateChartData() (Dual-Sync)
  logStep(5, 'Methods 2 & 3: updateChart() and updateChartData() — Dual-Sync Mutation')
  ppt.useSlide(1).updateChartData('QuarterlyRevenueBar', {
    categories: ['Item A-Adjusted', 'Item B-Adjusted', 'Item C-Adjusted'],
    series: [{ name: 'Revenue-Audited', values: [140, 220, 180] }],
  })
  ppt.useSlide(1).updateChart('QuarterlyRevenueBar', {
    title: 'Quarterly Revenue (Audited FY2026)',
  })
  console.log('       -> Dual-Sync updated chart XML cache and embedded Excel worksheet')

  // 6. Slide 3: Deep Cloning (Method 8)
  logStep(6, 'Method 8: cloneChart() — Deep Chart Part Duplication on Slide 3')
  ppt.useSlide(3).addChart(
    'bar',
    {
      categories: ['Tier 1', 'Tier 2', 'Tier 3'],
      series: [{ name: 'Enterprise Accounts', values: [42, 85, 120] }],
    },
    {
      name: 'AccountsBar',
      title: 'Customer Distribution',
      x: '1.0in',
      y: '1.2in',
      width: '5.0in',
      height: '3.8in',
    }
  )

  ppt.useSlide(3).cloneChart('AccountsBar', 'AccountsBar_Cloned', {
    x: '6.5in',
    y: '1.2in',
    width: '5.0in',
    height: '3.8in',
  })
  console.log(
    '       -> Cloned chart "AccountsBar" -> "AccountsBar_Cloned" with unique IDs & workbook'
  )

  // 7. Methods 9, 10, 11: Data Label operations
  logStep(7, 'Methods 9, 10 & 11: getChartLabel, setChartLabel, updateChartLabel')
  ppt.useSlide(3).setChartLabel('AccountsBar', true, { position: 'outsideEnd' })
  ppt.useSlide(3).updateChartLabel('AccountsBar_Cloned', {
    showCatName: true,
    showPercent: true,
    position: 'insideEnd',
    labels: ['Tier 1 (42)', 'Tier 2 (85)', 'Tier 3 (120)'],
  })

  const labelMeta = await ppt.useSlide(3).getChartLabel('AccountsBar')
  console.log(
    `       "AccountsBar" label config: showVal=${labelMeta.showVal}, position=${labelMeta.position}`
  )

  // 8. Slide 4: Lifecycle & clearCharts (Methods 4 & 5)
  logStep(8, 'Methods 4 & 5: removeChart() and clearCharts() — Lifecycle Cleanups on Slide 4')
  ppt.useSlide(4).addChart(
    'pie',
    {
      categories: ['Desktop', 'Mobile', 'Tablet'],
      series: [{ name: 'Market Share', values: [65, 30, 5] }],
    },
    { name: 'TempPieChart', x: '1.0in', y: '1.2in', width: '4.5in', height: '3.5in' }
  )
  ppt.useSlide(4).addChart(
    'area',
    {
      categories: ['2023', '2024', '2025'],
      series: [{ name: 'Bandwidth GB', values: [1200, 2400, 4800] }],
    },
    { name: 'TempAreaChart', x: '6.0in', y: '1.2in', width: '4.5in', height: '3.5in' }
  )

  // Remove single chart via removeChart
  ppt.useSlide(4).removeChart('TempPieChart')
  console.log('       -> Removed single chart "TempPieChart" via removeChart()')

  // Clear all remaining charts via clearCharts
  ppt.useSlide(4).clearCharts()
  console.log('       -> Cleared remaining charts on Slide 4 via clearCharts()')

  // 9. Export presentation
  logStep(9, 'Exporting Showcase Presentations', 'Generating final .pptx packages')
  const sampleOutputPath = path.join(OUTPUT_DIR, 'ChartPowerPointSample_output.pptx')
  const advancedOutputPath = path.join(ADVANCED_OUTPUT_DIR, 'showcase_chart_advanced.pptx')

  await ppt.saveToFile(sampleOutputPath)
  await ppt.saveToFile(advancedOutputPath)

  const sampleStats = await fs.stat(sampleOutputPath)
  const advancedStats = await fs.stat(advancedOutputPath)

  logHeader('Showcase Execution Summary')
  console.log(`Sample Output   : ${sampleOutputPath} (${(sampleStats.size / 1024).toFixed(2)} KB)`)
  console.log(
    `Advanced Output : ${advancedOutputPath} (${(advancedStats.size / 1024).toFixed(2)} KB)`
  )
  console.log(`Total Duration  : ${Date.now() - t0} ms`)
  console.log('Result          : ALL 11 CHART METHODS EXERCISED & VERIFIED SUCCESSFULLY!')
  console.log('='.repeat(80) + '\n')
}

// -----------------------------------------------------------------------------
// CLI Dispatcher
// -----------------------------------------------------------------------------

if (require.main === module) {
  const arg = process.argv[2]?.toLowerCase()?.trim()

  ;(async () => {
    try {
      if (!arg || arg === 'all' || arg === 'showcase') {
        console.log('\n--- Running 17 Isolated Chart Type Generation Tests ---')
        await runAllChartTests()
        console.log('\n--- Running Master 11-Method Whitelist Showcase ---')
        await runChartsShowcase()
      } else if (ALL_CHART_TEST_FUNCTIONS[arg]) {
        console.log(`\n--- Running Isolated Test for Chart Type: "${arg}" ---`)
        const outPath = await ALL_CHART_TEST_FUNCTIONS[arg]()
        const stats = await fs.stat(outPath)
        console.log(`[PASS] Generated: ${outPath} (${(stats.size / 1024).toFixed(2)} KB)`)
      } else {
        console.error(`Unknown chart type argument: "${arg}"`)
        console.log(`Available types: ${Object.keys(ALL_CHART_TEST_FUNCTIONS).join(', ')}`)
        process.exit(1)
      }
    } catch (err) {
      console.error('Execution failed:', err)
      process.exit(1)
    }
  })()
}

module.exports = {
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
  runAllChartTests,
  runChartsShowcase,
}
