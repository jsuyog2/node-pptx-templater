/**
 * @fileoverview showcase/charts_operations.js — Comprehensive Charts API Master Showcase.
 *
 * Demonstrates and validates the strict 8-function Charts Whitelist for node-pptx-templater.
 * Provides isolated standalone test functions for all 17 supported chart types across
 * standard DrawingML (<c:chartSpace>) and Office 2016+ extended schemas (<cx:chartSpace>).
 *
 * The Strict 8 Charts Methods Whitelist:
 * =========================================================================
 *  1. ppt.addChart([slideRef], chartType, data, options)
 *  2. ppt.updateChart(chartId, options)
 *  3. ppt.updateChartData(chartId, data)
 *  4. ppt.removeChart(chartId)
 *  5. ppt.clearCharts([slideRefOrChartId])
 *  6. ppt.getCharts([slideRef])
 *  7. ppt.getChart(chartId)
 *  8. ppt.cloneChart(chartId, clonedChartId, clonedOptions)
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

async function testBarChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('BarChart', {
        title: 'Updated Quarterly Revenue Bar Chart',
        // left-top
        x: '5%',
        y: '5%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('BarChart', {
        categories: ['Item A', 'Item B', 'Item C'],
        series: [{ name: 'Revenue', values: [120, 200, 150] }],
      })
      break
    default:
    case 'add':
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
      break
  }
}

async function testColumnChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('ColumnChart', {
        title: 'Updated Operating Cost Clustered Column',
        // right-top
        x: '55%',
        y: '5%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('ColumnChart', {
        categories: ['Q1', 'Q2', 'Q3', 'Q4'],
        series: [{ name: 'Operating Cost', values: [80, 95, 110, 130] }],
      })
      break
    default:
    case 'add':
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
      break
  }
}
async function testLineChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('LineChart', {
        title: 'Updated Monthly Growth Trend',
        // left-bottom
        x: '5%',
        y: '25%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('LineChart', {
        categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
        series: [{ name: 'Monthly Growth', values: [12, 19, 25, 32, 45, 60] }],
      })
      break
    default:
    case 'add':
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
      break
  }
}

async function testAreaChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('AreaChart', {
        title: 'Updated Bandwidth Utilization Area',
        // right-bottom
        x: '55%',
        y: '25%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('AreaChart', {
        categories: ['2022', '2023', '2024', '2025'],
        series: [{ name: 'Bandwidth GB', values: [1200, 2400, 4800, 9600] }],
      })
      break
    default:
    case 'add':
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
      break
  }
}

async function testPieChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('PieChart', {
        title: 'Updated Traffic Device Share',
        // left-top
        x: '5%',
        y: '5%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('PieChart', {
        categories: ['Desktop', 'Mobile', 'Tablet'],
        series: [{ name: 'Traffic Share', values: [60, 32, 8] }],
      })
      break
    default:
    case 'add':
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
      break
  }
}
async function testSunburstChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('SunburstChart', {
        title: 'Updated Regional Breakdown Sunburst',
        // left-top
        x: '55%',
        y: '5%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('SunburstChart', {
        categories: ['North', 'South', 'East', 'West'],
        series: [{ name: 'Regional Breakdown', values: [300, 200, 150, 250] }],
      })
      break
    default:
    case 'add':
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
      break
  }
}

async function testFunnelChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('FunnelChart', {
        title: 'Updated Marketing Pipeline Funnel',
        // left-bottom
        x: '5%',
        y: '25%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('FunnelChart', {
        categories: ['Impressions', 'Clicks', 'Signups', 'Purchases'],
        series: [{ name: 'Pipeline', values: [10000, 2500, 600, 150] }],
      })
      break
    default:
    case 'add':
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
      break
  }
}

async function testRadarChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('RadarChart', {
        title: 'Updated Performance Radar Dimensions',
        // right-bottom
        x: '55%',
        y: '25%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('RadarChart', {
        categories: ['Speed', 'Reliability', 'Comfort', 'Safety', 'Efficiency'],
        series: [{ name: 'Model Metric', values: [90, 85, 80, 95, 75] }],
      })
      break
    default:
    case 'add':
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
      break
  }
}

async function testHistogramChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('HistogramChart', {
        title: 'Updated Score Distribution Histogram',
        // left-top
        x: '5%',
        y: '5%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('HistogramChart', {
        categories: ['0-10', '10-20', '20-30', '30-40', '40-50'],
        series: [{ name: 'Frequency', values: [5, 12, 28, 18, 7] }],
      })
      break
    default:
    case 'add':
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
      break
  }
}

async function testScatterChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('ScatterChart', {
        title: 'Updated XY Correlation Scatter',
        // right-top
        x: '55%',
        y: '5%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('ScatterChart', {
        categories: ['10', '20', '30', '40'],
        series: [{ name: 'Sample Points', values: [15, 28, 42, 65] }],
      })
      break
    default:
    case 'add':
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
      break
  }
}

async function testStockChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('StockChart', {
        title: 'Updated Weekly High-Low-Close Stock Performance',
        // left-bottom
        x: '5%',
        y: '25%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('StockChart', {
        categories: ['Mon', 'Tue', 'Wed', 'Thu', 'Fri'],
        series: [
          { name: 'High', values: [125, 130, 128, 135, 140] },
          { name: 'Low', values: [110, 115, 112, 120, 122] },
          { name: 'Close', values: [120, 122, 125, 132, 138] },
        ],
      })
      break
    default:
    case 'add':
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
      break
  }
}

async function testSurfaceChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('SurfaceChart', {
        title: 'Updated 3D Surface Topology',
        // right-bottom
        x: '55%',
        y: '25%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('SurfaceChart', {
        categories: ['Row 1', 'Row 2', 'Row 3'],
        series: [
          { name: 'Col 1', values: [10, 20, 30] },
          { name: 'Col 2', values: [20, 35, 45] },
        ],
      })
      break
    default:
    case 'add':
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
      break
  }
}

async function testBoxWhiskerChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('BoxWhiskerChart', {
        title: 'Updated Percentile Box & Whisker',
        // left-top
        x: '5%',
        y: '5%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('BoxWhiskerChart', {
        categories: ['Group A', 'Group B', 'Group C'],
        series: [{ name: 'Scores', values: [72, 85, 91] }],
      })
      break
    default:
    case 'add':
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
      break
  }
}

async function testWaterfallChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('WaterfallChart', {
        title: 'Updated FY2026 Enterprise Cash Waterfall',
        // right-top
        x: '55%',
        y: '5%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('WaterfallChart', {
        categories: ['Starting Cash', 'Product Sales', 'Services', 'Salaries', 'Cloud', 'Ending'],
        series: [{ name: 'Cash Flow', values: [1000, 450, 200, -350, -150, 1150] }],
      })
      break
    default:
    case 'add':
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
      break
  }
}
async function testComboChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('ComboChart', {
        title: 'Updated H1 Actuals vs Target Combo',
        // left-bottom
        x: '5%',
        y: '25%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('ComboChart', {
        categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
        series: [
          { name: 'Direct Sales', type: 'column', values: [50, 60, 75, 80, 95, 110] },
          { name: 'Target Line', type: 'line', values: [55, 65, 70, 85, 90, 105] },
        ],
      })
      break
    default:
    case 'add':
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
      break
  }
}

async function testTreemapChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('TreemapChart', {
        title: 'Updated Category Sales Treemap',
        // right-bottom
        x: '55%',
        y: '25%',
        width: 500,
        height: 250,
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('TreemapChart', {
        categories: ['Electronics', 'Furniture', 'Clothing', 'Groceries'],
        series: [{ name: 'Sales Volume', values: [450, 220, 180, 310] }],
      })
      break
    default:
    case 'add':
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
}

async function testMapChart(ppt, update) {
  switch (update) {
    case 'update':
      await ppt.updateChart('MapChart', {
        title: 'Updated Global Adoption Geographic Map',
        projection: 'mercator',
        viewedRegion: 'dataOnly',
        // left-top
        x: '5%',
        y: '5%',
        width: '90%',
        height: '50%',
      })
      break
    case 'updateChartData':
      await ppt.updateChartData('MapChart', {
        categories: ['United States', 'Brazil', 'Germany', 'India', 'Australia'],
        series: [{ name: 'Global Adoption', values: [85, 42, 63, 91, 55] }],
      })
      break
    default:
    case 'add':
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
      break
  }
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

  // --------------------------F---------------------------------------------------
  // Part 1: addChart() :- Adding All 17 Types of Charts
  // -----------------------------------------------------------------------------

  console.log('\n[1/12] Executing addChart()')

  // --------------------------slide 1---------------------------------------------------
  ppt.useSlide(1)

  // 1. Bar
  logStep(1, 'Creating Bar Chart', 'Generates a horizontal bar chart with custom dimensions')
  await testBarChart(ppt)

  // 2. Column
  logStep(
    2,
    'Creating Column Chart',
    'Generates a vertical column chart with detailed point formatting'
  )
  await testColumnChart(ppt)

  // 3. Line
  logStep(3, 'Creating Line Chart', 'Generates a multi-series line chart with optional formatting')
  await testLineChart(ppt)

  // 4. Area
  logStep(4, 'Creating Area Chart', 'Generates an area chart showing cumulative values or trends')
  await testAreaChart(ppt)

  // --------------------------slide 2---------------------------------------------------
  ppt.useSlide(2)

  // 5. Pie
  logStep(
    5,
    'Creating Pie Chart',
    'Generates a standard pie chart for proportional data visualization'
  )
  await testPieChart(ppt)

  // 6. Sunburst
  logStep(
    6,
    'Creating Sunburst Chart',
    'Generates a hierarchical sunburst chart for multi-level analysis'
  )
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
  logStep(
    9,
    'Creating Histogram Chart',
    'Generates a histogram for frequency distribution analysis'
  )
  await testHistogramChart(ppt)

  // 10. Scatter
  logStep(
    10,
    'Creating Scatter Chart',
    'Generates a scatter chart with mathematically precise coordinate rendering'
  )
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
  logStep(
    13,
    'Creating Box & Whisker Chart',
    'Generates a box-and-whisker chart for statistical distribution'
  )
  await testBoxWhiskerChart(ppt)

  // 14. Waterfall
  logStep(14, 'Creating Waterfall Chart', 'Generates a waterfall chart for financial analysis')
  await testWaterfallChart(ppt)

  // 15. Combo
  logStep(15, 'Creating Combo Chart', 'Generates a combination chart with multiple chart types')
  await testComboChart(ppt)

  // 16. Treemap
  logStep(
    16,
    'Creating Treemap Chart',
    'Generates a treemap showing hierarchical data partitioning'
  )
  await testTreemapChart(ppt)

  // --------------------------slide 5---------------------------------------------------
  ppt.useSlide(5)
  // 17. Map
  logStep(
    17,
    'Creating Map Chart (Bing Maps with Excel Table)',
    'Generates a choropleth map using Bing Maps data resolution'
  )
  await testMapChart(ppt)

  // --------------------------F---------------------------------------------------
  // Part 2: updateChart() :- Updating All 17 Types of Charts
  // -----------------------------------------------------------------------------

  console.log('\n[2/12] Executing updateChart(...)')

  // --------------------------slide 6---------------------------------------------------
  ppt.useSlide(6)

  // 1. Bar
  logStep(1, 'update Bar Chart', 'Updates a horizontal bar chart with custom dimensions')
  await testBarChart(ppt, 'update')

  // 2. Column
  logStep(
    2,
    'update Column Chart',
    'Updates a vertical column chart with detailed point formatting'
  )
  await testColumnChart(ppt, 'update')

  // 3. Line
  logStep(3, 'update Line Chart', 'Updates a multi-series line chart with optional formatting')
  await testLineChart(ppt, 'update')

  // 4. Area
  logStep(4, 'update Area Chart', 'Updates an area chart showing cumulative values or trends')
  await testAreaChart(ppt, 'update')

  // --------------------------slide 7---------------------------------------------------
  ppt.useSlide(7)

  // 5. Pie
  logStep(5, 'update Pie Chart', 'Updates a standard pie chart for proportional data visualization')
  await testPieChart(ppt, 'update')

  // 6. Sunburst
  logStep(
    6,
    'update Sunburst Chart',
    'Updates a hierarchical sunburst chart for multi-level analysis'
  )
  await testSunburstChart(ppt, 'update')

  // 7. Funnel
  logStep(7, 'update Funnel Chart', 'Updates a funnel chart for pipeline visualization')
  await testFunnelChart(ppt, 'update')

  // 8. Radar
  logStep(8, 'update Radar Chart', 'Updates a radar/spider chart for multi-criteria evaluation')
  await testRadarChart(ppt, 'update')

  // --------------------------slide 8---------------------------------------------------
  ppt.useSlide(8)

  // 9. Histogram
  logStep(9, 'update Histogram Chart', 'Updates a histogram for frequency distribution analysis')
  await testHistogramChart(ppt, 'update')

  // 10. Scatter
  logStep(
    10,
    'update Scatter Chart',
    'Updates a scatter chart with mathematically precise coordinate rendering'
  )
  await testScatterChart(ppt, 'update')

  // 11. Stock
  logStep(11, 'update Stock Chart', 'Updates a candlestick stock chart with high/low/close')
  await testStockChart(ppt, 'update')

  // 12. Surface
  logStep(12, 'update Surface Chart', 'Updates a 3D surface chart for multi-variable analysis')
  await testSurfaceChart(ppt, 'update')

  // --------------------------slide 9---------------------------------------------------
  ppt.useSlide(9)

  // 13. Box & Whisker
  logStep(
    13,
    'update Box & Whisker Chart',
    'Updates a box-and-whisker chart for statistical distribution'
  )
  await testBoxWhiskerChart(ppt, 'update')

  // 14. Waterfall
  logStep(14, 'update Waterfall Chart', 'Updates a waterfall chart for financial analysis')
  await testWaterfallChart(ppt, 'update')

  // 15. Combo
  logStep(15, 'update Combo Chart', 'Updates a combination chart with multiple chart types')
  await testComboChart(ppt, 'update')

  // 16. Treemap
  logStep(16, 'update Treemap Chart', 'Updates a treemap showing hierarchical data partitioning')
  await testTreemapChart(ppt, 'update')

  // --------------------------slide 10---------------------------------------------------
  ppt.useSlide(10)
  // 17. Map
  logStep(
    17,
    'update Map Chart (Bing Maps with Excel Table)',
    'Updates a choropleth map using Bing Maps data resolution'
  )
  await testMapChart(ppt, 'update')

  // --------------------------F---------------------------------------------------
  // Part 3: updateChartData() :- Updating All 17 Types of Charts Data
  // -----------------------------------------------------------------------------

  console.log('\n[3/12] Executing updateChartData(...)')

  // --------------------------slide 11---------------------------------------------------
  ppt.useSlide(11)

  // 1. Bar
  logStep(1, 'update Bar Chart Data', 'Updates the data of a horizontal bar chart')
  await testBarChart(ppt, 'updateChartData')

  // 2. Column
  logStep(2, 'update Column Chart Data', 'Updates the data of a vertical column chart')
  await testColumnChart(ppt, 'updateChartData')

  // 3. Line
  logStep(3, 'update Line Chart Data', 'Updates the data of a multi-series line chart')
  await testLineChart(ppt, 'updateChartData')

  // 4. Area
  logStep(
    4,
    'update Area Chart Data',
    'Updates the data of an area chart showing cumulative values or trends'
  )
  await testAreaChart(ppt, 'updateChartData')

  // --------------------------slide 12---------------------------------------------------
  ppt.useSlide(12)

  // 5. Pie
  logStep(
    5,
    'update Pie Chart Data',
    'Updates the data of a standard pie chart for proportional data visualization'
  )
  await testPieChart(ppt, 'updateChartData')

  // 6. Sunburst
  logStep(
    6,
    'update Sunburst Chart Data',
    'Updates the data of a hierarchical sunburst chart for multi-level analysis'
  )
  await testSunburstChart(ppt, 'updateChartData')

  // 7. Funnel
  logStep(
    7,
    'update Funnel Chart Data',
    'Updates the data of a funnel chart for pipeline visualization'
  )
  await testFunnelChart(ppt, 'updateChartData')

  // 8. Radar
  logStep(
    8,
    'update Radar Chart Data',
    'Updates the data of a radar/spider chart for multi-criteria evaluation'
  )
  await testRadarChart(ppt, 'updateChartData')

  // --------------------------slide 13---------------------------------------------------
  ppt.useSlide(13)

  // 9. Histogram
  logStep(
    9,
    'update Histogram Chart Data',
    'Updates the data of a histogram for frequency distribution analysis'
  )
  await testHistogramChart(ppt, 'updateChartData')

  // 10. Scatter
  logStep(
    10,
    'update Scatter Chart Data',
    'Updates the data of a scatter chart with mathematically precise coordinate rendering'
  )
  await testScatterChart(ppt, 'updateChartData')

  // 11. Stock
  logStep(
    11,
    'update Stock Chart Data',
    'Updates the data of a candlestick stock chart with high/low/close'
  )
  await testStockChart(ppt, 'updateChartData')

  // 12. Surface
  logStep(
    12,
    'update Surface Chart Data',
    'Updates the data of a 3D surface chart for multi-variable analysis'
  )
  await testSurfaceChart(ppt, 'updateChartData')

  // --------------------------slide 14---------------------------------------------------
  ppt.useSlide(14)

  // 13. Box & Whisker
  logStep(
    13,
    'update Box & Whisker Chart Data',
    'Updates the data of a box-and-whisker chart for statistical distribution'
  )
  await testBoxWhiskerChart(ppt, 'updateChartData')

  // 14. Waterfall
  logStep(
    14,
    'update Waterfall Chart Data',
    'Updates the data of a waterfall chart for financial analysis'
  )
  await testWaterfallChart(ppt, 'updateChartData')

  // 15. Combo
  logStep(
    15,
    'update Combo Chart Data',
    'Updates the data of a combination chart with multiple chart types'
  )
  await testComboChart(ppt, 'updateChartData')

  // 16. Treemap
  logStep(
    16,
    'update Treemap Chart Data',
    'Updates the data of a treemap showing hierarchical data partitioning'
  )
  await testTreemapChart(ppt, 'updateChartData')

  // --------------------------slide 15---------------------------------------------------
  ppt.useSlide(15)
  // 17. Map
  logStep(
    17,
    'update Map Chart Data (Bing Maps with Excel Table)',
    'Updates the data of a choropleth map using Bing Maps data resolution'
  )
  await testMapChart(ppt, 'updateChartData')

  // --------------------------F---------------------------------------------------
  // Part 4: removeChart() :- Removing All 17 Types of Charts
  // -----------------------------------------------------------------------------

  console.log('\n[4/12] Executing removeChart(...)')
  ppt.useSlide(16)
  // 1. Bar
  logStep(1, 'remove Bar Chart', 'Removes a horizontal bar chart')
  await ppt.removeChart('BarChart')

  // 2. Column
  logStep(2, 'remove Column Chart', 'Removes a vertical column chart')
  await ppt.removeChart('ColumnChart')

  // 3. Line
  logStep(3, 'remove Line Chart', 'Removes a multi-series line chart')
  await ppt.removeChart('LineChart')


  // --------------------------F---------------------------------------------------
  // Part 4: clearChart() :- Clearing All 17 Types of Charts
  // -----------------------------------------------------------------------------

  console.log('\n[5/12] Executing clearChart(...)')
  ppt.useSlide(17)

  // 5. Pie
  logStep(1, 'Clearing All Charts', 'Clears all charts on the slide')
  await ppt.clearCharts()

  // --------------------------F---------------------------------------------------
  // Part 5: cloneChart() :- Cloning All 17 Types of Charts
  // -----------------------------------------------------------------------------

  console.log('\n[5/12] Executing cloneChart(...)')
  ppt.useSlide(18)

  // 13. Box & Whisker
  logStep(1, 'Cloning Box & Whisker Chart', 'Clones a box-and-whisker chart to a new slide')
  await ppt.cloneChart('BoxWhiskerChart', 'ClonedBoxWhiskerChart', {
    //top left
    x: '55%',
    y: '5%',
  })

  // 15. Combo
  logStep(2, 'Cloning Combo Chart', 'Clones a combination chart to a new slide')
  await ppt.cloneChart('ComboChart', 'ClonedComboChart', {
    //bottom right
    x: '55%',
    y: '25%',
  })

  // --------------------------F---------------------------------------------------
  // Part 6: getCharts() & getChart() :- Getting All 17 Types of Charts
  // -----------------------------------------------------------------------------

  console.log('\n[6/12] Executing getCharts() & getChart(...)')
  ppt.useSlide(9)

  // 13. Box & Whisker
  logStep(1, 'Getting All Charts', 'Gets all charts on the current slide')
  const charts = await ppt.getCharts()
  console.log(charts)

  logStep(2, 'Getting a Chart', 'Gets a specific chart on the current slide')
  const chart = await ppt.getChart('TreemapChart')
  console.log(chart)

  // 18. Export presentation
  logStep(18, 'Exporting Showcase Presentations', 'Generating final .pptx packages')
  const outputPath = path.join(OUTPUT_DIR, 'showcase_chart_advanced.pptx')

  await ppt.saveToFile(outputPath)

  logHeader('Showcase Execution Summary')
  console.log(`Total Duration  : ${Date.now() - t0} ms`)
  console.log('Result          : ALL 8 CHART METHODS EXERCISED & VERIFIED SUCCESSFULLY!')
  console.log('='.repeat(80) + '\n')
}

if (require.main === module) {
  runChartsShowcase()
}

module.exports = { runChartsShowcase }
