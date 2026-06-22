/**
 * Text Search & Replace Example
 *
 * Demonstrates: findText, getTextElements, replaceTextByTag,
 *               replaceMultiple, updateText, getList
 *
 * Run: node examples/text-search.js
 */

'use strict'

const path = require('path')
const { PPTXTemplater } = require('../src/index.js')

const TEMPLATE = path.join(__dirname, '../templates/sample.pptx')
const OUTPUT = path.join(__dirname, 'output/text-search.pptx')

async function main() {
  const ppt = await PPTXTemplater.load(TEMPLATE, { logLevel: 'warn' })

  // 1. Search for text across all slides
  ppt.useAllSlides()
  const matches = ppt.findText('{{')
  console.log(`Found ${matches.length} placeholder(s) across all slides:`)
  matches
    .slice(0, 5)
    .forEach(m => console.log(`  • Slide ${m.slideIndex}: "${m.text}" in shape "${m.shapeName}"`))

  // 2. List all text elements (first 5)
  const elements = ppt.getTextElements()
  console.log(`\nTotal text elements: ${elements.length}`)
  elements
    .slice(0, 5)
    .forEach(el => console.log(`  • Slide ${el.slideIndex}: "${el.text.substring(0, 40)}"`))

  // 3. Replace a single tag across all slides
  ppt.replaceTextByTag('{{title}}', 'v1.1.0 Release')
  console.log('\n✓ Replaced {{title}} → "v1.1.0 Release"')

  // 4. Replace multiple tags at once
  ppt.replaceMultiple({
    '{{company}}': 'Acme Corp',
    '{{date}}': new Date().toLocaleDateString(),
    '{{version}}': '1.1.0',
  })
  console.log('✓ Replaced {{company}}, {{date}}, {{version}}')

  // 5. Update a specific named text shape
  ppt.useSlide(1)
  ppt.updateText('Title', 'node-pptx-templater v1.1.0')
  console.log('✓ Updated "Title" shape text')

  await ppt.saveToFile(OUTPUT)
  console.log('✓ Done:', OUTPUT)
}

main().catch(err => {
  console.error('Error:', err.message)
  process.exit(1)
})
