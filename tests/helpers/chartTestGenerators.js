import path from 'path'
import { PPTXTemplater } from '../../src/index.js'

async function generateSingleChart(chartType, data, options, templatePath, outputDir, fileName) {
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1)
  await ppt.addChart(chartType, data, options)
  const outPath = path.join(outputDir, fileName)
  await ppt.saveToFile(outPath)
  return outPath
}

export async function testBarChart(templatePath, outputDir) {
  return generateSingleChart(
    'bar',
    {
      categories: ['Item A', 'Item B', 'Item C'],
      series: [{ name: 'Revenue', values: [120, 200, 150] }],
    },
    {
      name: 'StandaloneBarChart',
      title: 'Quarterly Revenue Bar Chart',
      x: '5%',
      y: '5%',
      width: 500,
      height: 350,
    },
    templatePath,
    outputDir,
    'chart_bar.pptx'
  )
}

export async function testColumnChart(templatePath, outputDir) {
  return generateSingleChart(
    'column',
    {
      categories: ['Q1', 'Q2', 'Q3', 'Q4'],
      series: [{ name: 'Operating Cost', values: [80, 95, 110, 130] }],
    },
    {
      name: 'StandaloneColumnChart',
      title: 'Operating Cost Clustered Column',
      x: '55%',
      y: '5%',
      width: 500,
      height: 350,
    },
    templatePath,
    outputDir,
    'chart_column.pptx'
  )
}

export async function testLineChart(templatePath, outputDir) {
  return generateSingleChart(
    'line',
    {
      categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
      series: [{ name: 'Monthly Growth', values: [12, 19, 25, 32, 45, 60] }],
    },
    {
      name: 'StandaloneLineChart',
      title: 'Monthly Growth Trend',
      x: '5%',
      y: '25%',
      width: 500,
      height: 350,
    },
    templatePath,
    outputDir,
    'chart_line.pptx'
  )
}

export async function testPieChart(templatePath, outputDir) {
  return generateSingleChart(
    'pie',
    {
      categories: ['Desktop', 'Mobile', 'Tablet'],
      series: [{ name: 'Traffic Share', values: [60, 32, 8] }],
    },
    {
      name: 'StandalonePieChart',
      title: 'Traffic Device Share',
      x: '5%',
      y: '45%',
      width: 500,
      height: 350,
    },
    templatePath,
    outputDir,
    'chart_pie.pptx'
  )
}

export async function testAreaChart(templatePath, outputDir) {
  return generateSingleChart(
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
    },
    templatePath,
    outputDir,
    'chart_area.pptx'
  )
}

export async function testScatterChart(templatePath, outputDir) {
  return generateSingleChart(
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
    },
    templatePath,
    outputDir,
    'chart_scatter.pptx'
  )
}

export async function testMapChart(templatePath, outputDir) {
  return generateSingleChart(
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
    },
    templatePath,
    outputDir,
    'chart_map.pptx'
  )
}

export async function testStockChart(templatePath, outputDir) {
  return generateSingleChart(
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
    },
    templatePath,
    outputDir,
    'chart_stock.pptx'
  )
}

export async function testSurfaceChart(templatePath, outputDir) {
  return generateSingleChart(
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
    },
    templatePath,
    outputDir,
    'chart_surface.pptx'
  )
}

export async function testRadarChart(templatePath, outputDir) {
  return generateSingleChart(
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
    },
    templatePath,
    outputDir,
    'chart_radar.pptx'
  )
}

export async function testTreemapChart(templatePath, outputDir) {
  return generateSingleChart(
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
    },
    templatePath,
    outputDir,
    'chart_treemap.pptx'
  )
}

export async function testSunburstChart(templatePath, outputDir) {
  return generateSingleChart(
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
    },
    templatePath,
    outputDir,
    'chart_sunburst.pptx'
  )
}

export async function testHistogramChart(templatePath, outputDir) {
  return generateSingleChart(
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
    },
    templatePath,
    outputDir,
    'chart_histogram.pptx'
  )
}

export async function testBoxWhiskerChart(templatePath, outputDir) {
  return generateSingleChart(
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
    },
    templatePath,
    outputDir,
    'chart_boxWhisker.pptx'
  )
}

export async function testWaterfallChart(templatePath, outputDir) {
  return generateSingleChart(
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
    },
    templatePath,
    outputDir,
    'chart_waterfall.pptx'
  )
}

export async function testFunnelChart(templatePath, outputDir) {
  return generateSingleChart(
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
    },
    templatePath,
    outputDir,
    'chart_funnel.pptx'
  )
}

export async function testComboChart(templatePath, outputDir) {
  return generateSingleChart(
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
    },
    templatePath,
    outputDir,
    'chart_combo.pptx'
  )
}
