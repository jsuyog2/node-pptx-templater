/**
 * Shape Operations Example
 *
 * Demonstrates: getShapes, addShape, updateShape, updateShapeText,
 *               updateShapePosition, cloneShape, deleteShape, removeShape
 *
 * Run: node examples/shape-operations.js
 */

'use strict'

const path = require('path')
const { PPTXTemplater } = require('../src/index.js')

const TEMPLATE = path.join(__dirname, '../templates/sample.pptx')
const OUTPUT = path.join(__dirname, 'output/shape-operations.pptx')

async function main() {
  const ppt = await PPTXTemplater.load(TEMPLATE, { logLevel: 'warn' })
  ppt.useSlide(1)

  // 1. List existing shapes
  const shapes = ppt.getShapes()
  console.log(`Found ${shapes.length} shape(s) on slide 1:`)
  shapes.forEach(s => console.log(`  • ${s.name || s.id} (${s.type})`))

  // 2. Add a blue rectangle
  await ppt.addShape({
    type: 'rectangle',
    x: 914400, // 1 inch
    y: 914400, // 1 inch
    width: 2743200, // 3 inches
    height: 457200, // 0.5 inches
    fill: '#3B82F6',
    id: 'demo-shape',
  })
  console.log('✓ Added shape: demo-shape (blue rectangle)')

  // 3. Update its fill to green
  await ppt.updateShape('demo-shape', { fill: '#10B981' })
  console.log('✓ Updated demo-shape fill to green')

  // 4. Add text to the shape
  ppt.updateShapeText('demo-shape', 'Hello from node-pptx-templater!')
  console.log('✓ Set text on demo-shape')

  // 5. Clone it to a new position
  ppt.cloneShape('demo-shape', 'demo-shape-2', { x: 914400, y: 1600000 })
  console.log('✓ Cloned demo-shape → demo-shape-2')

  // 6. Reposition the clone
  ppt.updateShapePosition('demo-shape-2', { y: 1800000 })
  console.log('✓ Repositioned demo-shape-2')

  // 7. Delete the clone
  ppt.deleteShape('demo-shape-2')
  console.log('✓ Deleted demo-shape-2')

  await ppt.saveToFile(OUTPUT)
  console.log('✓ Done:', OUTPUT)
}

main().catch(err => {
  console.error('Error:', err.message)
  process.exit(1)
})
