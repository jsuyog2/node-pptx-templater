const { PPTXTemplater } = require('../src/index.js')
const { existsSync, mkdirSync } = require('fs')
const { resolve } = require('path')

const TEMPLATE_PATH = resolve(__dirname, '../templates/sample.pptx')
const OUTPUT_DIR = resolve(__dirname, '../examples/output')
const OUTPUT_PATH = resolve(__dirname, '../examples/output/shape-cells-output.pptx')

async function main() {
  if (!existsSync(TEMPLATE_PATH)) {
    console.error(
      `Template not found at ${TEMPLATE_PATH}. Please run within the repository context.`
    )
    process.exit(1)
  }

  if (!existsSync(OUTPUT_DIR)) {
    mkdirSync(OUTPUT_DIR, { recursive: true })
  }

  console.log('📂 Loading template...')
  const ppt = await PPTXTemplater.load(TEMPLATE_PATH)

  // Use slide 3 (contains 'Table')
  ppt.useSlide(3)

  console.log('➕ Adding rows with Shape Cells...')

  // Row 1: KPI Badge Only (circle)
  await ppt.addTableRow('Table', [
    {
      type: 'circle',
      fill: '#10B981', // green
      radius: 6,
      position: 'center',
    },
    'Task Complete',
    'High Priority',
  ])

  // Row 2: Status Indicator (circle + text)
  await ppt.addTableRow('Table', [
    {
      type: 'circle',
      fill: '#EF4444', // red
      text: 'Blocked',
      radius: 6,
      position: 'left',
    },
    'DB Server Down',
    'Critical',
  ])

  // Row 3: Progress Marker (rectangle + text)
  await ppt.addTableRow('Table', [
    {
      type: 'rectangle',
      fill: '#3B82F6', // blue
      text: 'In Progress',
      width: 14,
      height: 14,
      position: 'left',
    },
    'API Implementation',
    'Medium',
  ])

  // Row 4: Custom shapes (diamond / hexagon)
  await ppt.addTableRow('Table', [
    {
      type: 'diamond',
      fill: '#F59E0B', // amber
      width: 12,
      height: 12,
      position: 'center',
    },
    {
      type: 'hexagon',
      fill: '#8B5CF6', // purple
      text: 'Review Needed',
      width: 12,
      height: 12,
      position: 'left',
    },
    'Normal',
  ])

  console.log(`💾 Saving output to: ${OUTPUT_PATH}`)
  await ppt.saveToFile(OUTPUT_PATH)

  console.log('✅ Validation checks...')
  const report = await ppt.validatePresentation()
  if (report.valid) {
    console.log('🎉 Shape cells example completed successfully!')
  } else {
    console.error('❌ Validation failed:', report.errors)
    process.exit(1)
  }
}

main().catch(err => {
  console.error('Error running example:', err)
  process.exit(1)
})
