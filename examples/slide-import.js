/**
 * Slide Import & Export Example
 *
 * Demonstrates: importSlideFrom, importSlides, cloneSlide,
 *               exportSlides, duplicateSlide, deleteSlide,
 *               moveSlide, getSlides
 *
 * Run: node examples/slide-import.js
 */

'use strict'

const path = require('path')
const { PPTXTemplater } = require('../src/index.js')

const TEMPLATE = path.join(__dirname, '../templates/sample.pptx')
const OUTPUT_IMPORT = path.join(__dirname, 'output/slide-import.pptx')
const OUTPUT_SUBSET = path.join(__dirname, 'output/slide-subset.pptx')

async function main() {
  // 1. Load two engine instances
  const ppt = await PPTXTemplater.load(TEMPLATE, { logLevel: 'warn' })
  const source = await PPTXTemplater.load(TEMPLATE, { logLevel: 'warn' })

  console.log(`Loaded template with ${ppt.slideCount} slide(s)`)

  // 2. List existing slides
  const slides = ppt.getSlides()
  console.log('Slides:')
  slides.forEach(s => console.log(`  • Slide ${s.index}: ${s.title || '(no title)'}`))

  // 3. Clone slide 1 to the end
  ppt.cloneSlide(1)
  console.log(`✓ Cloned slide 1 — now ${ppt.slideCount} slides`)

  // 4. Duplicate slide 2 and insert it after slide 1
  ppt.duplicateSlide(2, 2)
  console.log(`✓ Duplicated slide 2 — now ${ppt.slideCount} slides`)

  // 5. Move the last slide to position 2
  ppt.moveSlide(ppt.slideCount, 2)
  console.log(`✓ Moved last slide to position 2`)

  // 6. Import slide 1 from another engine instance
  await ppt.importSlideFrom(source, 1)
  console.log(`✓ Imported slide 1 from source — now ${ppt.slideCount} slides`)

  // 7. Delete the extra slide
  ppt.deleteSlide(ppt.slideCount)
  console.log(`✓ Deleted last slide — now ${ppt.slideCount} slides`)

  await ppt.saveToFile(OUTPUT_IMPORT)
  console.log('✓ Saved:', OUTPUT_IMPORT)

  // 8. Export a subset of slides (only slides 1 and 2) to a new engine
  const ppt2 = await PPTXTemplater.load(TEMPLATE, { logLevel: 'warn' })
  const subset = await ppt2.exportSlides(1, 2)
  console.log(`✓ Exported subset: ${subset.slideCount} slide(s)`)

  await subset.saveToFile(OUTPUT_SUBSET)
  console.log('✓ Saved subset:', OUTPUT_SUBSET)
}

main().catch(err => {
  console.error('Error:', err.message)
  process.exit(1)
})
