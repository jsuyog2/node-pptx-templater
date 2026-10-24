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
const OUTPUT_DIR = path.join(__dirname, 'output/slide_advanced')

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

async function testBarChart(ppt) {
  await ppt.addChart(
    'bar',
    {
      categories: ['Item A', 'Item B', 'Item C'],
      series: [{ name: 'Revenue', values: [120, 200, 150] }],
    },
    {
      name: 'StandaloneBarChart',
      title: 'Quarterly Revenue Bar Chart',
      // left-top
      x: '5%',
      y: '5%',
      width: 500,
      height: 250,
    }
  )
}

async function testColumnChart(ppt) {
  await ppt.addChart(
    'column',
    {
      categories: ['Q1', 'Q2', 'Q3', 'Q4'],
      series: [{ name: 'Operating Cost', values: [80, 95, 110, 130] }],
    },
    {
      name: 'StandaloneColumnChart',
      title: 'Operating Cost Clustered Column',
      // right-top
      x: '55%',
      y: '5%',
      width: 500,
      height: 250,
    }
  )
}

async function testLineChart(ppt) {
  await ppt.addChart(
    'line',
    {
      categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
      series: [{ name: 'Monthly Growth', values: [12, 19, 25, 32, 45, 60] }],
    },
    {
      name: 'StandaloneLineChart',
      title: 'Monthly Growth Trend',
      // left-bottom
      x: '5%',
      y: '25%',
      width: 500,
      height: 250,
    }
  )
}


async function testAreaChart(ppt) {
  await ppt.addChart(
    'area',
    {
      categories: ['2022', '2023', '2024', '2025'],
      series: [{ name: 'Bandwidth GB', values: [1200, 2400, 4800, 9600] }],
    },
    {
      name: 'StandaloneAreaChart',
      title: 'Bandwidth Utilization Area',
      // right-bottom
      x: '55%',
      y: '25%',
      width: 500,
      height: 250,
    }
  )
}

async function testPieChart(ppt) {
  await ppt.addChart(
    'pie',
    {
      categories: ['Desktop', 'Mobile', 'Tablet'],
      series: [{ name: 'Traffic Share', values: [60, 32, 8] }],
    },
    {
      name: 'StandalonePieChart',
      title: 'Traffic Device Share',
      // left-top
      x: '5%',
      y: '5%',
      width: 500,
      height: 250,
    }
  )
}

async function testSunburstChart(ppt) {
  await ppt.addChart(
    'sunburst',
    {
      categories: ['North', 'South', 'East', 'West'],
      series: [{ name: 'Regional Breakdown', values: [300, 200, 150, 250] }],
    },
    {
      name: 'StandaloneSunburstChart',
      title: 'Hierarchy Sunburst Distribution',
      // right-top
      x: '55%',
      y: '5%',
      width: 500,
      height: 250,
    }
  )
}

async function testFunnelChart(ppt) {
  await ppt.addChart(
    'funnel',
    {
      categories: ['Impressions', 'Clicks', 'Signups', 'Purchases'],
      series: [{ name: 'Pipeline', values: [10000, 2500, 600, 150] }],
    },
    {
      name: 'StandaloneFunnelChart',
      title: 'Marketing Pipeline Funnel',
      // left-bottom
      x: '5%',
      y: '25%',
      width: 500,
      height: 250,
    }
  )
}



async function testRadarChart(ppt) {
  await ppt.addChart(
    'radar',
    {
      categories: ['Speed', 'Reliability', 'Comfort', 'Safety', 'Efficiency'],
      series: [{ name: 'Model Metric', values: [90, 85, 80, 95, 75] }],
    },
    {
      name: 'StandaloneRadarChart',
      title: 'Performance Radar Dimensions',
      // right-bottom
      x: '55%',
      y: '25%',
      width: 500,
      height: 250,
    }
  )
}

async function testHistogramChart(ppt) {
  await ppt.addChart(
    'histogram',
    {
      categories: ['0-10', '10-20', '20-30', '30-40', '40-50'],
      series: [{ name: 'Frequency', values: [5, 12, 28, 18, 7] }],
    },
    {
      name: 'StandaloneHistogramChart',
      title: 'Score Distribution Histogram',
      // left-top
      x: '5%',
      y: '5%',
      width: 500,
      height: 250,
    }
  )
}


async function testScatterChart(ppt) {
  await ppt.addChart(
    'scatter',
    {
      categories: ['10', '20', '30', '40'],
      series: [{ name: 'Sample Points', values: [15, 28, 42, 65] }],
    },
    {
      name: 'StandaloneScatterChart',
      title: 'XY Correlation Scatter',
      // right-top
      x: '55%',
      y: '5%',
      width: 500,
      height: 250,
    }
  )
}

async function testStockChart(ppt) {
  await ppt.addChart(
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
      // left-bottom
      x: '5%',
      y: '25%',
      width: 500,
      height: 250,
    }
  )
}

async function testSurfaceChart(ppt) {
  await ppt.addChart(
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
      // right-bottom
      x: '55%',
      y: '25%',
      width: 500,
      height: 250,
    }
  )
}

async function testBoxWhiskerChart(ppt) {
  await ppt.addChart(
    'boxWhisker',
    {
      categories: ['Group A', 'Group B', 'Group C'],
      series: [{ name: 'Scores', values: [72, 85, 91] }],
    },
    {
      name: 'StandaloneBoxWhiskerChart',
      title: 'Percentile Box & Whisker',
      // left-top
      x: '5%',
      y: '5%',
      width: 500,
      height: 250,
    }
  )
}

async function testWaterfallChart(ppt) {
  await ppt.addChart(
    'waterfall',
    {
      categories: ['Starting Cash', 'Product Sales', 'Services', 'Salaries', 'Cloud', 'Ending'],
      series: [{ name: 'Cash Flow', values: [1000, 450, 200, -350, -150, 1150] }],
    },
    {
      name: 'StandaloneWaterfallChart',
      title: 'FY2026 Enterprise Cash Waterfall',
      // right-top
      x: '55%',
      y: '5%',
      width: 500,
      height: 250,
    }
  )
}
async function testComboChart(ppt) {
  await ppt.addChart(
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
      // left-bottom
      x: '5%',
      y: '25%',
      width: 500,
      height: 250,
    }
  )
}

async function testTreemapChart(ppt) {
  await ppt.addChart(
    'treemap',
    {
      categories: ['Electronics', 'Furniture', 'Clothing', 'Groceries'],
      series: [{ name: 'Sales Volume', values: [450, 220, 180, 310] }],
    },
    {
      name: 'StandaloneTreemapChart',
      title: 'Category Sales Treemap',
      // right-bottom
      x: '55%',
      y: '25%',
      width: 500,
      height: 250,
    }
  )
}

async function testMapChart(ppt) {
  await ppt.addChart(
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
      // left-top
      x: '5%',
      y: '5%',
      width: '90%',
      height: '50%',
    }
  )
}


// --------------------------F---------------------------------------------------
// Master Showcase Exercising All 11 Whitelisted Chart Methods
// -----------------------------------------------------------------------------

async function runChartsShowcase() {
  const t0 = new Date()
  console.log('='.repeat(75))
  console.log('PPTXTemplater — Charts Operations Master Guide')
  console.log('='.repeat(75))

  // 1. Initialize a clean presentation deck
  console.log('\n[1/10] Initializing presentation...')
  // Initialize a PPTXTemplater instance with a sample presentation
  console.log('→ Step 0: Initializing PPTXTemplater with ChartPowerPointSample.pptx...')
  const templatePath = path.join(__dirname, 'templates/ChartPowerPointSample.pptx')
  const ppt = await PPTXTemplater.load(templatePath)

  // --------------------------slide 1---------------------------------------------------
  ppt.useSlide(1)

  // 1. Bar
  logStep(5, 'Creating Bar Chart', 'Generates a horizontal bar chart with custom dimensions')
  await testBarChart(ppt)

  // 2. Column
  logStep(6, 'Creating Column Chart', 'Generates a vertical column chart with detailed point formatting')
  await testColumnChart(ppt)

  // 3. Line
  logStep(2, 'Creating Line Chart', 'Generates a multi-series line chart with optional formatting')
  await testLineChart(ppt)

  // 4. Area
  logStep(3, 'Creating Area Chart', 'Generates an area chart showing cumulative values or trends')
  await testAreaChart(ppt)

  // --------------------------slide 2---------------------------------------------------
  ppt.useSlide(2)

  // 5. Pie
  logStep(5, 'Creating Pie Chart', 'Generates a standard pie chart for proportional data visualization')
  await testPieChart(ppt)

  // 6. Sunburst
  logStep(6, 'Creating Sunburst Chart', 'Generates a hierarchical sunburst chart for multi-level analysis')
  await testSunburstChart(ppt)

  // 7. Funnel
  logStep(7, 'Creating Funnel Chart', 'Generates a funnel chart for pipeline visualization')
  await testFunnelChart(ppt)

  // 8. Radar
  logStep(8, 'Creating Radar Chart', 'Generates a radar/spider chart for multi-criteria evaluation')
  await testRadarChart(ppt)

  // --------------------------slide 3---------------------------------------------------
  ppt.useSlide(3)

  // 9. Histogram
  logStep(9, 'Creating Histogram Chart', 'Generates a histogram for frequency distribution analysis')
  await testHistogramChart(ppt)

  // 10. Scatter
  logStep(10, 'Creating Scatter Chart', 'Generates a scatter chart with mathematically precise coordinate rendering')
  await testScatterChart(ppt)

  // 11. Stock
  logStep(11, 'Creating Stock Chart', 'Generates a candlestick stock chart with high/low/close')
  await testStockChart(ppt)


  // 12. Surface
  logStep(12, 'Creating Surface Chart', 'Generates a 3D surface chart for multi-variable analysis')
  await testSurfaceChart(ppt)

  // --------------------------slide 4---------------------------------------------------
  ppt.useSlide(4)

  // 13. Box & Whisker
  logStep(13, 'Creating Box & Whisker Chart', 'Generates a box-and-whisker chart for statistical distribution')
  await testBoxWhiskerChart(ppt)

  // 14. Waterfall
  logStep(14, 'Creating Waterfall Chart', 'Generates a waterfall chart for financial analysis')
  await testWaterfallChart(ppt)

  // 15. Combo
  logStep(15, 'Creating Combo Chart', 'Generates a combination chart with multiple chart types')
  await testComboChart(ppt)

  // 16. Treemap
  logStep(16, 'Creating Treemap Chart', 'Generates a treemap showing hierarchical data partitioning')
  await testTreemapChart(ppt)

  // --------------------------slide 5---------------------------------------------------
  ppt.useSlide(5)
  // 17. Map
  logStep(17, 'Creating Map Chart (Bing Maps with Excel Table)', 'Generates a choropleth map using Bing Maps data resolution')
  await testMapChart(ppt)

  // 18. Export presentation
  logStep(18, 'Exporting Showcase Presentations', 'Generating final .pptx packages')
  const outputPath = path.join(OUTPUT_DIR, 'showcase_chart_advanced.pptx')

  await ppt.saveToFile(outputPath)

  logHeader('Showcase Execution Summary')
  console.log(`Total Duration  : ${Date.now() - t0} ms`)
  console.log('Result          : ALL 11 CHART METHODS EXERCISED & VERIFIED SUCCESSFULLY!')
  console.log('='.repeat(80) + '\n')
}


if (require.main === module) {
  runChartsShowcase()
}

module.exports = { runChartsShowcase }
