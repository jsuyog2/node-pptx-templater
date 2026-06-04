/**
 * @fileoverview Basic usage example for pptx-templater.
 *
 * This example demonstrates the core workflow:
 *  1. Load a PPTX template
 *  2. Select specific slides
 *  3. Replace text placeholders
 *  4. Add a new slide
 *  5. Save to file and buffer
 *
 * To run: node examples/basic-usage.js
 */

// import { PPTXTemplater } from '../src/index.js';
// import { existsSync } from 'fs';
// import { resolve, dirname } from 'path';
// import { fileURLToPath } from 'url';
const { PPTXTemplater } = require('../src/index.js')
const { existsSync } = require('fs')
const { resolve } = require('path')

const TEMPLATE_PATH = resolve(__dirname, '../templates/sample.pptx')
const OUTPUT_PATH = resolve(__dirname, '../examples/output/basic-output.pptx')

async function main() {
  // Check if template exists
  if (!existsSync(TEMPLATE_PATH)) {
    console.log('ℹ Template file not found at:', TEMPLATE_PATH)
    console.log('  Place a PPTX file at templates/sample.pptx to run this example.')
    console.log('\nRunning API demonstration without file I/O...\n')
    await demonstrateAPI()
    return
  }

  try {
    console.log('📂 Loading template:', TEMPLATE_PATH)
    const ppt = await PPTXTemplater.load(TEMPLATE_PATH)

    console.log(`📊 Loaded ${ppt.slideCount} slides`)

    // Step 1: Work on specific slide
    ppt.useSlide(1)

    // Step 2: Replace text placeholders
    ppt.replaceText({
      '{{title}}': 'Quarterly Business Review',
      '{{subtitle}}': 'Q1 2026 Results',
      '{{date}}': new Date().toLocaleDateString('en-US', {
        year: 'numeric',
        month: 'long',
        day: 'numeric',
      }),
      '{{company}}': 'Acme Corporation',
      '{{presenter}}': 'Jane Smith',
    })

    console.log('✅ Text replaced')

    ppt.updateChart('Chart', {
      categories: ['Q1', 'Q2', 'Q3', 'Q4'],
      series: [
        { name: 'Product A', values: [145, 210, 190, 250] },
        { name: 'Product B', values: [90, 130, 160, 200] },
        { name: 'Product C', values: [70, 85, 100, 120] },
      ],
    })

    ppt.applyZOrder(1, [{ id: 'Shape', zIndex: 2 }])

    ppt.removeSlide(2)

    ppt.useSlide(2)
    ppt.replaceText({
      '{{title}}': 'New Title',
    })
    ppt.addHyperlink({
      text: 'Google',
      url: 'https://www.google.com',
    })

    ppt.addSlideLink({
      element: 'Hello New Title',
      sourceSlide: 2,
      targetSlide: 1,
    })

    ppt.updateTable('Table', [
      ['Name', '', 'Role', 'Dept'],
      ['Alice', '', 'Engineer', 'Platform'],
      ['Bob', '', 'Designer', 'Product'],
    ])

    // merge all cells of first column with all cells of second col
    ppt.mergeCells('Table', 0, 0, 0, 1)
    ppt.mergeCells('Table', 1, 0, 1, 1)
    ppt.mergeCells('Table', 2, 0, 2, 1)

    ppt.updateTable('Table2', [
      ['Name', '', 'Role', 'Dept'],
      ['Alice', '', 'Engineer', 'Platform'],
      ['Bob', { value: 'Alice', align: 'ctr' }, 'Designer', 'Product'],
    ])

    ppt.mergeCells('Table2', 0, 0, 0, 1)
    ppt.mergeCells('Table2', 1, 0, 1, 1)

    ppt.useSlide(3)

    ppt.updateChart('Label Chart', {
      categories: [''],
      series: [
        { name: 'Product A', values: [{ data: 145, label: '145 (18.24%)' }] },
        { name: 'Product B', values: [{ data: 210, label: '210 (26.42%)' }] },
        { name: 'Product C', values: [{ data: 190, label: '190 (23.9%)' }] },
        { name: 'Product D', values: [{ data: 250, label: '250 (31.45%)' }] }
      ],
    })


    // Step 3: Save to file
    await ppt.saveToFile(OUTPUT_PATH)
    console.log('💾 Saved to:', OUTPUT_PATH)

    // Step 4: Also get as buffer (for API responses)
    const buffer = await ppt.toBuffer()
    console.log(`📦 Buffer size: ${(buffer.length / 1024).toFixed(1)} KB`)

    // Step 5: Validate
    const validation = ppt.validate()
    if (validation.valid) {
      console.log('✅ Validation passed')
    } else {
      console.warn('⚠ Validation warnings:', validation.errors)
    }

    console.log('\n🎉 Example completed successfully!')
  } catch (err) {
    console.error('❌ Error:', err.message)
    if (process.env.DEBUG) console.error(err.stack)
    process.exit(1)
  }
}

/**
 * Demonstrates the API structure without needing an actual PPTX file.
 */
async function demonstrateAPI() {
  console.log('=== API Usage Patterns ===\n')

  // Show the API shape
  console.log(`
// Load a template
const ppt = await PPTXTemplater.load('template.pptx');

// Select slides to operate on
ppt.useSlide(1);                     // Single slide
ppt.useSlide(1, 3, 5);              // Multiple slides
ppt.useAllSlides();                  // All slides

// Replace {{placeholder}} text
ppt.replaceText({
  '{{title}}': 'Quarterly Report',
  '{{year}}':  '2026',
});

// Update chart data
ppt.updateChart('sales-chart', {
  categories: ['Jan', 'Feb', 'Mar'],
  series: [{ name: 'Revenue', values: [120, 150, 180] }],
});

// Update table rows
ppt.updateTable('data-table', [
  ['Name',  'Role',       'Dept'],
  ['Alice', 'Engineer',   'Platform'],
  ['Bob',   'Designer',   'Product'],
]);

// Add hyperlinks
ppt.addHyperlink({ text: 'Visit Us', url: 'https://example.com' });

// Link slide numbers
ppt.linkSlideNumber({ slide: 1, targetSlide: 5 });

// Add new slides
ppt.addSlide({
  title: 'New Slide',
  elements: [{ type: 'text', value: 'Hello World' }],
});

// Clone and remove slides
ppt.cloneSlide(1);          // Duplicate slide 1
ppt.removeSlide(3);         // Remove slide 3
ppt.reorderSlides([3,1,2]); // Reorder

// Export
await ppt.saveToFile('./output/report.pptx');
const buffer = await ppt.toBuffer();
const stream = await ppt.toStream();
  `)

  console.log('=== End of API Demo ===')
}

main()
