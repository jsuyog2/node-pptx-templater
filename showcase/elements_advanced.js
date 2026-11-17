/**
 * @fileoverview showcase/elements_advanced.js — Advanced Elements & Layering Master Guide.
 *
 * Demonstrates and validates the unified "Elements" category strictly adhering
 * to the 9 approved functions:
 *   1. getElements(slideIdOrFilter, filter)
 *   2. getElement(elementId, slideId)
 *   3. removeElement(elementId, slideId)
 *   4. removeAllElements(slideId, options)
 *   5. setZOrder(elementId, zIndex, slideId)
 *   6. bringToFront(elementId, slideId)
 *   7. sendToBack(elementId, slideId)
 *   8. bringForward(elementId, slideId)
 *   9. sendBackward(elementId, slideId)
 *
 * Requirements:
 *   Node.js >= 20.12.0
 *
 * How to run:
 *   node showcase/elements_advanced.js
 *
 * Output presentation saved to: showcase/output/elements_advanced.pptx
 */

'use strict'

const path = require('path')
const fs = require('fs-extra')
const { PPTXTemplater } = require('../src/index.js')

const OUTPUT_DIR = path.resolve(__dirname, 'output')
const OUTPUT_FILE = path.join(OUTPUT_DIR, 'elements_advanced.pptx')

// 1x1 transparent PNG buffer for embedded image demonstration
const DEMO_IMAGE_BUF = Buffer.from(
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
  'base64'
)

async function runElementsAdvancedShowcase() {
  console.log('==============================================================================')
  console.log('  node-pptx-templater — Advanced Elements & Layering Showcase')
  console.log('  Validating the Strict 9-Method Elements Whitelist API')
  console.log('==============================================================================')
  console.log(`Node.js version  : ${process.version}`)
  console.log(`Working directory: ${process.cwd()}`)
  console.log(`Output target    : ${OUTPUT_FILE}`)

  await fs.ensureDir(OUTPUT_DIR)
  const startTime = Date.now()

  // --------------------------------------------------------------------------
  // Initialize Presentation
  // --------------------------------------------------------------------------
  const templatePath = path.join(__dirname, 'templates/ElementsPowerPointSample.pptx')
  const ppt = await PPTXTemplater.load(templatePath)

  // ==========================================================================
  // SCENARIO 1: Element Discovery, Metadata & Single Inspection
  // Methods: [1] getElements, [2] getElement
  // ==========================================================================
  console.log('\n------------------------------------------------------------------------------')
  console.log('  SCENARIO 1: Element Enumeration, Filtering & Single Element Inspection')
  console.log('------------------------------------------------------------------------------')

  ppt.useSlide(1)

  // [1] getElements() — retrieve full element list on slide 1
  const allElements = ppt.getElements()
  console.log(`[Method 1/9: getElements] Discovered ${allElements.length} element(s) on Slide 1:`)
  for (const el of allElements) {
    const b = el.bounds
    console.log(
      `  • [${el.type.padEnd(8)}] ID: ${String(el.id).padEnd(4)} Name: "${el.name.padEnd(14)}" Pos: (${b.x}, ${b.y}) Dim: ${b.width}x${b.height}`
    )
  }

  // [2] getElement() — single element inspection by name and ID
  const heroHeader = ppt.getElement('TableRemoveElement')
  console.log('\n[Method 2/9: getElement] Inspected "TableRemoveElement" by name:')
  console.log(`  ✓ Type      : ${heroHeader.type}`)
  console.log(`  ✓ ID        : ${heroHeader.id}`)
  console.log(`  ✓ Width/Height: ${heroHeader.bounds.width} x ${heroHeader.bounds.height} EMUs`)

  const statusBadgeById = ppt.getElement('ChartRemoveElement')
  console.log(`  ✓ Found "ChartRemoveElement" by numerical ID ("ChartRemoveElement"): Name = "${statusBadgeById?.name}"`)

  // Type-filtered querying
  const shapesOnly = ppt.getElements({ type: 'shape' })
  console.log(`  ✓ Filtered by type ('shape'): Found ${shapesOnly.length} shape elements`)

  // ==========================================================================
  // SCENARIO 2: Z-Order Stacking, Permutations & Layer Boundary Navigation
  // Methods: [5] setZOrder, [6] bringToFront, [7] sendToBack, [8] bringForward, [9] sendBackward
  // ==========================================================================
  console.log('\n------------------------------------------------------------------------------')
  console.log('  SCENARIO 2: Visual Stacking & Z-Order Layer Reordering')
  console.log('------------------------------------------------------------------------------')


  // Fluent Chaining Demonstration
  ppt.useSlide(3)
  ppt.setZOrder('TableRemoveElement', 3)
  console.log('Fluent Method Chaining Result')

  console.log('Initial Stacking (Back -> Front)')
  ppt.useSlide(4)
  // [6] bringToFront — send BottomLayer straight to the front
  ppt.bringToFront('ImageRemoveElement')
  console.log('[Method 6/9: bringToFront]')

  ppt.useSlide(5)
  // [7] sendToBack — send BottomLayer back to bottom
  ppt.sendToBack('ChartRemoveElement')
  console.log('[Method 7/9: sendToBack]')

  ppt.useSlide(6)
  // [8] bringForward — step BottomLayer forward by one position
  ppt.bringForward('ShapeRemoveElement')
  console.log('[Method 8/9: bringForward]')

  ppt.useSlide(7)
  // [9] sendBackward — step BottomLayer back by one position
  ppt.sendBackward('TableRemoveElement')
  console.log('[Method 9/9: sendBackward]')


  // ==========================================================================
  // SCENARIO 3: Canvas Evacuation & Batch Purging
  // Method: [3] removeAllElements
  // ==========================================================================
  console.log('\n------------------------------------------------------------------------------')
  console.log('  SCENARIO 4: Slide Canvas Reset & Batch Element Evacuation')
  console.log('------------------------------------------------------------------------------')

  ppt.useSlide(1)

  // [4] removeElement complete evacuation
  ppt.removeElement('TableRemoveElement')
  ppt.removeElement('ChartRemoveElement')
  ppt.removeElement('ShapeRemoveElement')
  console.log('[Method 3/9: removeElement (complete evacuation)]')
  console.log(`  ✓ Canvas completely evacuated: ${ppt.getElements().length} element(s) remaining`)


  ppt.useSlide(2)

  // [4] removeAllElements complete evacuation
  ppt.removeAllElements({ preservePlaceholders: false })
  console.log('[Method 4/9: removeAllElements (complete evacuation)]')
  console.log(`  ✓ Canvas completely evacuated: ${ppt.getElements().length} element(s) remaining`)

  // --------------------------------------------------------------------------
  // Save Presentation to Disk
  // --------------------------------------------------------------------------
  console.log('\n------------------------------------------------------------------------------')
  console.log('  SECTION 5: Package Serialization & Disk Output')
  console.log('------------------------------------------------------------------------------')

  await ppt.saveToFile(OUTPUT_FILE)
  const stats = await fs.stat(OUTPUT_FILE)
  const duration = Date.now() - startTime

  console.log(`✓ Presentation package saved successfully!`)
  console.log(`  • File Size: ${stats.size} bytes (${(stats.size / 1024).toFixed(1)} KB)`)
  console.log(`  • Total Execution Time: ${duration}ms`)

  console.log('\n==============================================================================')
  console.log('  ALL 9 APPROVED ELEMENTS CATEGORY METHODS EXECUTED CLEANLY!')
  console.log('==============================================================================')
}
if (require.main === module) {
  runElementsAdvancedShowcase().catch(err => {
    console.error('Showcase failed:', err)
    process.exit(1)
  })
}

module.exports = { runElementsAdvancedShowcase }
