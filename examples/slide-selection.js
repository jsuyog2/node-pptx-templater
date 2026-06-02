/**
 * @fileoverview Slide management example.
 *
 * Demonstrates slide operations: duplicating, inserting, moving,
 * deleting slides and keeping internal structures synchronized.
 *
 * To run: node examples/slide-selection.js
 */

const { PPTXTemplater } = require('../src/index.js')
const { existsSync } = require('fs')
const { resolve } = require('path')
const fsExtra = require('fs-extra')

const TEMPLATE = resolve(__dirname, '../templates/sample.pptx')
const OUTPUT = resolve(__dirname, './output/slide-output.pptx')

async function main() {
  if (!existsSync(TEMPLATE)) {
    console.log('ℹ Template file not found at:', TEMPLATE)
    return
  }

  try {
    console.log('📂 Loading template:', TEMPLATE)
    const ppt = await PPTXTemplater.load(TEMPLATE)
    console.log(`📊 Loaded ${ppt.slideCount} slides`)

    console.log('Duplicating Slide 1 to position 2...')
    ppt.duplicateSlide(1, 2)
    console.log(`📊 Current slide count: ${ppt.slideCount}`)

    console.log('Moving Slide 4 to position 1...')
    ppt.moveSlide(4, 1)

    console.log('Inserting a new slide with a title at position 2...')
    ppt.insertSlide(2, { title: 'Newly Inserted Slide' })
    console.log(`📊 Current slide count: ${ppt.slideCount}`)

    console.log('Deleting the slide at position 3...')
    ppt.deleteSlide(3)
    console.log(`📊 Current slide count: ${ppt.slideCount}`)

    // Ensure output directory exists
    await fsExtra.ensureDir(resolve(__dirname, './output'))

    await ppt.saveToFile(OUTPUT)
    console.log('💾 Saved slide output to:', OUTPUT)

    const validation = ppt.validate()
    if (validation.valid) {
      console.log('✅ Validation passed')
    } else {
      console.warn('⚠ Validation warnings:', validation.errors)
    }

    console.log('\n🎉 Slide management example completed successfully!')
  } catch (err) {
    console.error('❌ Error:', err.message)
    if (process.env.DEBUG) console.error(err.stack)
    process.exit(1)
  }
}

main()
