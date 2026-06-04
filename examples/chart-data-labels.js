/**
 * @fileoverview Chart Data Labels example.
 *
 * Demonstrates all data label configurations:
 * - Basic Labels
 * - Labels From Cells
 * - Label Arrays
 * - Label Mapping
 * - Label Templates
 * - Pie Chart Labels
 * - Percentage Labels
 * - Advanced Formatting
 * - Large Dataset Labels
 *
 * To run: node examples/chart-data-labels.js
 */

const { PPTXTemplater } = require('../src/index.js')
const { existsSync } = require('fs')
const { resolve } = require('path')
const fsExtra = require('fs-extra')

const TEMPLATE = resolve(__dirname, '../tests/fixtures/sample.pptx')
const OUTPUT = resolve(__dirname, './output/chart-data-labels-output.pptx')

async function main() {
  console.log('📊 Chart Data Labels API Demo')

  const hasTemplate = existsSync(TEMPLATE)
  let ppt;

  if (hasTemplate) {
    ppt = await PPTXTemplater.load(TEMPLATE)
    console.log(`Loaded template with ${ppt.slideCount} slides`)
  } else {
    // If no template exists in the environment, show console code examples
    showCodeShowcase()
    return
  }

  // 1. Basic Labels & Label Arrays
  // Update Chart 1 with explicit string label array
  ppt.useSlide(1).updateChart('Chart', {
    categories: ['Excellent', 'Good', 'Average', 'Poor'],
    series: [
      { name: 'Feedback', values: [85, 95, 45, 12] }
    ]
  })
  ppt.useSlide(1).updateDataLabels('Chart', {
    series: 0,
    labels: ['A+ Grade', 'B Grade', 'C Grade', 'F Grade']
  })
  console.log('✅ Applied literal label arrays to Slide 1 chart')

  // 2. Label Mapping & Templates
  // Update series and categories on slide 1 and map label keys dynamically
  ppt.useSlide(1).updateDataLabels('Chart', {
    series: 0,
    labelMap: {
      'Excellent': 'Top performer',
      'Poor': 'Needs improvement'
    },
    template: '{category}: {customLabel} ({value})'
  })
  console.log('✅ Applied dynamic label templates and mapping to Slide 1 chart')

  // 3. Label Styling & Positioning
  // Style and position the data labels with custom DrawingML properties
  ppt.useSlide(1).updateDataLabels('Chart', {
    series: 0,
    position: 'insideEnd',
    labelStyle: {
      fontFamily: 'Century Gothic',
      fontSize: 12,
      color: '#0055A5',
      bold: true,
      italic: true,
      underline: true
    }
  })
  console.log('✅ Applied position insideEnd and custom styles')

  // 4. Value From Cells ("Labels From Cells")
  // Pull values from cell coordinates (Sheet1!D2:D5)
  ppt.useSlide(1).updateDataLabels('Chart', {
    series: 0,
    labelsFromCells: 'Sheet1!D2:D5'
  })
  console.log('✅ Configured Labels From Cells referencing Sheet1!D2:D5')

  // Ensure output directory exists
  await fsExtra.ensureDir(resolve(__dirname, './output'))
  await ppt.saveToFile(OUTPUT)
  console.log('💾 Saved output to:', OUTPUT)
}

function showCodeShowcase() {
  console.log(`
// 1. Basic Labels & Label Arrays
ppt.useSlide(1).updateDataLabels('MyChart', {
  series: 0,
  labels: ['Top Performance', 'Target Met', 'At Risk']
});

// 2. Labels From Cells
ppt.useSlide(1).updateDataLabels('MyChart', {
  series: 0,
  labelsFromCells: 'Sheet1!D2:D4'
});

// 3. Label Mapping
ppt.useSlide(1).updateDataLabels('MyChart', {
  series: 0,
  labelMap: {
    'Category 1': 'Peak',
    'Category 2': 'Normal'
  }
});

// 4. Label Templates
ppt.useSlide(1).updateDataLabels('MyChart', {
  series: 0,
  template: '{category}: {value} ({percentage}%)'
});

// 5. Pie Chart Labels
ppt.useSlide(1).updateDataLabels('PieChart', {
  series: 0,
  position: 'bestFit',
  showPercent: true
});

// 6. Advanced Formatting & Styling
ppt.useSlide(1).updateDataLabels('MyChart', {
  series: 0,
  position: 'insideEnd',
  labelStyle: {
    fontFamily: 'Century Gothic',
    fontSize: 14,
    color: '#FF5500',
    bold: true,
    italic: true,
    underline: true
  }
});

// 7. Large Dataset Labels
ppt.useSlide(1).updateDataLabels('ScatterChart', {
  series: 0,
  template: '{value}' // minimal clutter
});
  `)
}

main().catch(console.error)
