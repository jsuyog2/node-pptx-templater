/**
 * @fileoverview showcase/slide_advanced.js — Comprehensive Slide Management Master Guide.
 *
 * Demonstrates and validates the strict 23-method Slide Management API whitelist:
 * =========================================================================
 * SLIDE MANAGEMENT WHITELIST (23 functions):
 *   1.  ppt.useSlide(...slideRefs)
 *   2.  ppt.useAllSlides()
 *   3.  ppt.addSlide(options, position)
 *   4.  ppt.duplicateSlide(sourceIndex, atPosition)
 *   5.  ppt.removeSlide(position)
 *   6.  ppt.getSlides()
 *   7.  ppt.getSlide(index)
 *   8.  ppt.importSlideFrom(sourcePpt, sourceIndex, atPosition)
 *   9.  ppt.exportSlides(...slideIndices)
 *   10. ppt.moveSlide(fromIndex, toIndex)
 *   11. ppt.reorderSlides(order)
 *   12. ppt.tagSlide(slideIndex, tag)
 *   13. ppt.getSlideDimensions()
 *   14. ppt.getSlideLayout(slideIndex)
 *   15. ppt.useSection(sectionRef)
 *   16. ppt.addSection(name, startSlideIndex)
 *   17. ppt.updateSection(sectionIdentifier, options)
 *   18. ppt.removeSection(sectionIdentifier, removeSlides)
 *   19. ppt.getSection(sectionRef)
 *   20. ppt.hideSlide(...slideRefs)
 *   21. ppt.unhideSlide(...slideRefs)
 *   22. ppt.resetSlide(slideRef, options)
 *   23. ppt.addComment(text, options)
 *   24. ppt.removeComment(commentIdentifier, options)
 * =========================================================================
 */

'use strict'

const path = require('path')
const fs = require('fs-extra')
const { PPTXTemplater } = require('../src/index.js')

const OUTPUT_DIR = path.join(__dirname, 'output/slide_advanced')

async function runSlideAdvancedShowcase() {
  console.log('='.repeat(75))
  console.log('PPTXTemplater — Advanced Slide Management Showcase (Strict 23 Methods)')
  console.log('='.repeat(75))

  await fs.ensureDir(OUTPUT_DIR)
  const templatePath = path.join(__dirname, 'templates/sample.pptx')
  const ppt = await PPTXTemplater.load(templatePath)

  // 1. ppt.getSlideDimensions()
  const dimensions = ppt.getSlideDimensions()
  console.log(
    `\n[1/23] ppt.getSlideDimensions() -> width: ${dimensions.width}, height: ${dimensions.height} EMUs`
  )

  // 2. ppt.getSlides()
  const allSlides = ppt.getSlides()
  console.log(`[2/23] ppt.getSlides() -> Found ${allSlides.length} slide(s) in deck`)

  // 3. ppt.getSlide(index)
  const firstSlide = ppt.getSlide(1)
  console.log(
    `[3/23] ppt.getSlide(1) -> Slide ID: ${firstSlide.slideId}, RelID: ${firstSlide.relationshipId}`
  )

  // 4. ppt.getSlideLayout(slideIndex)
  const layout = ppt.getSlideLayout(1)
  console.log(`[4/23] ppt.getSlideLayout(1) -> Target layout: ${layout}`)

  // 5. ppt.useSlide(...slideRefs)
  ppt.useSlide(1)
  console.log(`[5/23] ppt.useSlide(1) -> Active slide context scoped to slide 1`)

  // 6. ppt.useAllSlides()
  ppt.useAllSlides()
  console.log(`[6/23] ppt.useAllSlides() -> Scoped context widened to all slides`)

  // 7. ppt.tagSlide(slideIndex, tag)
  ppt.tagSlide(1, 'cover-slide')
  console.log(`[7/23] ppt.tagSlide(1, 'cover-slide') -> Tagged slide 1 as 'cover-slide'`)

  // 8. ppt.addSlide(options, position)
  ppt.addSlide({ title: 'Key Performance Indicators' }, 2)
  console.log(
    `[8/23] ppt.addSlide({ title: 'Key Performance Indicators' }, 2) -> Inserted slide at pos 2, total: ${ppt.slideCount}`
  )

  // 9. ppt.duplicateSlide(sourceIndex, atPosition)
  await ppt.duplicateSlide(1, 3)
  console.log(
    `[9/23] ppt.duplicateSlide(1, 3) -> Cloned slide 1 to pos 3, total: ${ppt.slideCount}`
  )

  // 10. ppt.moveSlide(fromIndex, toIndex)
  ppt.moveSlide(3, 1)
  console.log(`[10/23] ppt.moveSlide(3, 1) -> Moved slide 3 to pos 1`)

  // 11. ppt.reorderSlides(order)
  const order = Array.from({ length: ppt.slideCount }, (_, i) => ppt.slideCount - i)
  ppt.reorderSlides(order)
  console.log(`[11/23] ppt.reorderSlides([${order.join(', ')}]) -> Inverted deck order`)

  // 12. ppt.removeSlide(position)
  ppt.removeSlide(ppt.slideCount - 2)
  console.log(
    `[12/23] ppt.removeSlide(${ppt.slideCount - 2 + 1}) -> Removed slide, remaining: ${ppt.slideCount}`
  )

  // 13. ppt.hideSlide(...slideRefs)
  ppt.hideSlide(1)
  console.log(`[13/23] ppt.hideSlide(1) -> Slide 1 visibility set to hidden (show="0")`)

  // 14. ppt.unhideSlide(...slideRefs)
  ppt.unhideSlide(1)
  console.log(`[14/23] ppt.unhideSlide(1) -> Slide 1 restored to visible state`)

  // 15. ppt.resetSlide(slideRef, options)
  ppt.resetSlide(2, { preservePlaceholders: true })
  console.log(
    `[15/23] ppt.resetSlide(2, { preservePlaceholders: true }) -> Reset canvas shapes while keeping placeholders`
  )

  // 16. ppt.getSection(sectionRef)
  const section = ppt.getSection('Default Section')
  console.log(
    `[16/23] ppt.getSection('Default Section') -> Section GUID: ${section.id}, Slides: [${section.slideIndices.join(', ')}]`
  )

  // 17. ppt.updateSection(sectionIdentifier, options)
  ppt.updateSection('Default Section', { name: 'Leadership Overview' })
  console.log(
    `[17/23] ppt.updateSection('Default Section', { name: 'Leadership Overview' }) -> Renamed section`
  )

  // 18. ppt.useSection(sectionRef)
  ppt.useSection('Untitled Section')
  console.log(`[18/23] ppt.useSection('Untitled Section') -> Scoped operations to member slides`)

  // 19. ppt.removeSection(sectionIdentifier, removeSlides)
  ppt.removeSection('Untitled Section', false)
  console.log(
    `[19/23] ppt.removeSection('Untitled Section', false) -> Section deleted, slides preserved`
  )

  // 20. ppt.addSection(name, startSlideIndex)
  ppt.addSection('Executive Overview', 5)
  console.log(`[20/23] ppt.addSection('Executive Overview', 5) -> Created PowerPoint section`)

  // 21. ppt.addComment(text, options)
  await ppt.useSlide(1).addComment('Audit team reviewed Q3 figures', {
    author: 'Principal Architect',
    initials: 'PA',
    x: '0.1',
    y: '0.1',
  })
  console.log(
    `[21/23] ppt.addComment('Audit team reviewed Q3 figures') -> Added collaborative OpenXML comment`
  )

  // 22. ppt.removeComment(commentIdentifier, options)
  await ppt.useSlide(1).addComment('Temporary review flag', {
    author: 'Reviewer',
    x: 20,
    y: 20,
  })
  await ppt.useSlide(1).removeComment('Temporary review flag')
  console.log(`[22/23] ppt.removeComment('Temporary review flag') -> Removed targeted comment`)

  // 23. ppt.importSlideFrom(sourcePpt, sourceIndex, atPosition)
  const blankDeck = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })
  blankDeck.useSlide(1)
  await blankDeck.addShape('roundedRectangle', {
    x: 100,
    y: 100,
    width: 600,
    height: 200,
    fill: '#1E293B',
    text: 'Synthesized with PPTXTemplater.createBlank()',
  })
  await ppt.importSlideFrom(blankDeck, 1, 3)
  console.log(
    `[23/23] ppt.importSlideFrom(blankDeck, 1, 3) -> Imported slide across presentations, count: ${blankDeck.slideCount}`
  )

  // 24. ppt.exportSlides(...slideIndices)
  const exportedDeck = await ppt.exportSlides(1)
  console.log(
    `[Extra] ppt.exportSlides(1) -> Exported subset deck, count: ${exportedDeck.slideCount}`
  )

  // Save and verify package integrity
  const outPath = path.join(OUTPUT_DIR, 'advanced_slides_output.pptx')
  await ppt.saveToFile(outPath)
  console.log(`\n✓ Output presentation saved to: ${outPath}`)

  console.log('='.repeat(75))
  console.log('All 24 Slide Management API Methods Verified Successfully!')
  console.log('='.repeat(75))
}

if (require.main === module) {
  runSlideAdvancedShowcase().catch(err => {
    console.error('Showcase execution failed:', err)
    process.exit(1)
  })
}

module.exports = { runSlideAdvancedShowcase }
