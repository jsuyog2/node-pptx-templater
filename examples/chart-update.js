/**
 * @fileoverview Chart update example.
 *
 * Demonstrates updating bar, line, and pie chart data
 * while preserving all chart styling from the template.
 *
 * To run: node examples/chart-update.js
 */

import { PPTXTemplater } from '../src/index.js';
import { existsSync } from 'fs';
import { resolve, dirname } from 'path';
import { fileURLToPath } from 'url';

const __dirname = dirname(fileURLToPath(import.meta.url));
const TEMPLATE = resolve(__dirname, '../templates/charts-template.pptx');
const OUTPUT = resolve(__dirname, './output/chart-output.pptx');

async function main() {
  if (!existsSync(TEMPLATE)) {
    console.log('📊 Chart Update API Demo (no template file needed)\n');
    showChartApiExample();
    return;
  }

  const ppt = await PPTXTemplater.load(TEMPLATE);
  console.log(`Loaded template with ${ppt.slideCount} slides`);

  // ── Slide 1: Bar chart ─────────────────────────────────────────────
  ppt.useSlide(1).updateChart('revenue-bar', {
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [
      { name: 'Product A', values: [145, 210, 190, 250] },
      { name: 'Product B', values: [90, 130, 160, 200] },
      { name: 'Product C', values: [70, 85, 100, 120] },
    ],
  });
  console.log('✅ Updated bar chart');

  // ── Slide 2: Line chart ────────────────────────────────────────────
  ppt.useSlide(2).updateChart('trend-line', {
    categories: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'],
    series: [
      {
        name: 'Revenue ($K)',
        values: [120, 135, 128, 160, 185, 210],
      },
      {
        name: 'Target ($K)',
        values: [130, 140, 140, 155, 170, 200],
      },
    ],
  });
  console.log('✅ Updated line chart');

  // ── Slide 3: Pie chart ─────────────────────────────────────────────
  ppt.useSlide(3).updateChart('market-pie', {
    categories: ['North America', 'Europe', 'Asia Pacific', 'Other'],
    series: [
      { name: 'Market Share', values: [42, 28, 22, 8] },
    ],
  });
  console.log('✅ Updated pie chart');

  await ppt.saveToFile(OUTPUT);
  console.log('💾 Saved to:', OUTPUT);
}

function showChartApiExample() {
  console.log(`
// Update a bar chart
ppt.updateChart('my-chart', {
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'Revenue', values: [145, 210, 190, 250] },
    { name: 'Expenses', values: [90, 110, 130, 150] },
  ],
});

// Update a pie chart
ppt.updateChart('pie-chart', {
  categories: ['Segment A', 'Segment B', 'Segment C'],
  series: [{ name: 'Share', values: [40, 35, 25] }],
});

// Supported chart types:
// - bar (barChart)
// - line (lineChart)
// - pie (pieChart)
// - area (areaChart)
// - scatter (scatterChart)
// - doughnut (doughnutChart)
  `);
}

main().catch(console.error);
