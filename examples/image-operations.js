/**
 * Image Operations Example
 *
 * Demonstrates: getImages, addImage, replaceImage, removeImage
 *
 * Run: node examples/image-operations.js
 */

'use strict'

const path = require('path')
const fs = require('fs')
const { PPTXTemplater } = require('../src/index.js')

const TEMPLATE = path.join(__dirname, '../templates/sample.pptx')
const OUTPUT = path.join(__dirname, 'output/image-operations.pptx')
// Use a test fixture image if available, otherwise generate a tiny placeholder
const TEST_IMG = path.join(__dirname, '../tests/fixtures/test-image.png')

async function main() {
  const ppt = await PPTXTemplater.load(TEMPLATE, { logLevel: 'warn' })
  ppt.useSlide(1)

  // 1. List existing images on slide 1
  const images = ppt.getImages()
  console.log(`Found ${images.length} image(s) on slide 1:`)
  images.forEach(img => console.log(`  • ${img.name || img.rId} @ (${img.x}, ${img.y})`))

  // 2. Add a new image (4 inches wide × 3 inches tall, positioned at 1" × 1")
  const imgPath = fs.existsSync(TEST_IMG) ? TEST_IMG : TEMPLATE // fallback: embed PPTX as binary data demo
  if (fs.existsSync(TEST_IMG)) {
    await ppt.addImage(TEST_IMG, {
      x: 914400, // 1 inch in EMUs (1 inch = 914400 EMUs)
      y: 914400, // 1 inch
      width: 3657600, // 4 inches
      height: 2743200, // 3 inches
      name: 'demo-logo',
    })
    console.log('✓ Added new image: demo-logo')

    // 3. Replace that image with another copy of itself
    await ppt.replaceImage('demo-logo', TEST_IMG)
    console.log('✓ Replaced image: demo-logo')

    // 4. Remove it
    ppt.removeImage('demo-logo')
    console.log('✓ Removed image: demo-logo')
  } else {
    console.log('(Skipping add/replace/remove — no test image fixture found)')
  }

  await ppt.saveToFile(OUTPUT)
  console.log('✓ Done:', OUTPUT)
}

main().catch(err => {
  console.error('Error:', err.message)
  process.exit(1)
})
