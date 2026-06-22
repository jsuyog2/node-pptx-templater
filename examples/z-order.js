/**
 * Layer Ordering (Z-Order) Example
 *
 * Demonstrates: getObjectOrder, getTopMostObject, getBottomMostObject,
 *               bringToFront, sendToBack, bringForward, sendBackward,
 *               swapObjects, setZIndex, reorderObjects
 *
 * Run: node examples/z-order.js
 */

'use strict'

const path = require('path')
const { PPTXTemplater } = require('../src/index.js')

const TEMPLATE = path.join(__dirname, '../templates/sample.pptx')
const OUTPUT = path.join(__dirname, 'output/z-order.pptx')

async function main() {
  const ppt = await PPTXTemplater.load(TEMPLATE, { logLevel: 'warn' })
  ppt.useSlide(1)

  // 1. Check the current stacking order
  const order = ppt.getObjectOrder(1)
  console.log(`Current stacking order (${order.length} objects):`)
  order.forEach((o, i) => console.log(`  [${i + 1}] ${o.name || o.id} (${o.type})`))

  // 2. Identify top and bottom objects
  const topObj = ppt.getTopMostObject(1)
  const bottomObj = ppt.getBottomMostObject(1)
  if (topObj) console.log(`✓ Top-most: ${topObj.name || topObj.id}`)
  if (bottomObj) console.log(`✓ Bottom-most: ${bottomObj.name || bottomObj.id}`)

  // 3. Add two overlapping shapes to demonstrate layering
  await ppt.addShape({
    type: 'rectangle',
    x: 500000,
    y: 500000,
    width: 1828800,
    height: 457200,
    fill: '#3B82F6',
    id: 'layer-shape-A',
  })
  await ppt.addShape({
    type: 'rectangle',
    x: 700000,
    y: 600000,
    width: 1828800,
    height: 457200,
    fill: '#EF4444',
    id: 'layer-shape-B',
  })
  console.log('✓ Added two overlapping shapes')

  // 4. Z-order operations
  ppt.bringToFront('layer-shape-A')
  console.log('✓ Brought layer-shape-A to front')

  ppt.sendToBack('layer-shape-B')
  console.log('✓ Sent layer-shape-B to back')

  ppt.bringForward('layer-shape-B')
  console.log('✓ Moved layer-shape-B forward one step')

  ppt.sendBackward('layer-shape-A')
  console.log('✓ Moved layer-shape-A backward one step')

  // 5. Swap two objects
  ppt.swapObjects(1, 'layer-shape-A', 'layer-shape-B')
  console.log('✓ Swapped layer-shape-A and layer-shape-B positions')

  // 6. Final order
  const newOrder = ppt.getObjectOrder(1)
  console.log(`Final stacking order (${newOrder.length} objects):`)
  newOrder.forEach((o, i) => console.log(`  [${i + 1}] ${o.name || o.id}`))

  await ppt.saveToFile(OUTPUT)
  console.log('✓ Done:', OUTPUT)
}

main().catch(err => {
  console.error('Error:', err.message)
  process.exit(1)
})
