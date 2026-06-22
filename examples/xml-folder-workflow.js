/**
 * XML Folder Workflow Example
 *
 * Demonstrates the extract → edit → rebuild workflow using
 * extractPptx(), fromPresentationXml() / load(folder), saveToFolder(),
 * and buildPptx().
 *
 * This is useful for:
 *   - Inspecting PPTX internals as plain XML files
 *   - Version-controlling presentation XML
 *   - Batch editing across files using standard text tools
 *   - Building presentations from scratch via XML templates
 *
 * Run: node examples/xml-folder-workflow.js
 */

'use strict'

const path = require('path')
const fs = require('fs-extra')
const { PPTXTemplater } = require('../src/index.js')

const TEMPLATE = path.join(__dirname, '../templates/sample.pptx')
const WORK_FOLDER = path.join(__dirname, 'output/pptx-extracted')
const OUTPUT_REBUILT = path.join(__dirname, 'output/xml-folder-rebuilt.pptx')

async function main() {
  // Clean up any previous run's directory
  await fs.remove(WORK_FOLDER)

  // 1. Extract PPTX to a folder of XML files
  await PPTXTemplater.extractPptx(TEMPLATE, WORK_FOLDER)
  console.log(`✓ Extracted PPTX to: ${WORK_FOLDER}`)

  // 2. Verify the extracted structure
  const files = await fs.readdir(WORK_FOLDER)
  console.log(`   Contains: ${files.join(', ')}`)

  // 3. Load the engine from the extracted folder
  const ppt = await PPTXTemplater.load(WORK_FOLDER, { logLevel: 'warn' })
  console.log(`✓ Loaded from folder — ${ppt.slideCount} slide(s)`)

  // 4. Make modifications
  ppt.useAllSlides()
  ppt.replaceTextByTag('{{version}}', '1.1.0')
  ppt.replaceTextByTag('{{env}}', 'Production')
  console.log('✓ Replaced text placeholders')

  // 5. Save changes back to the folder (modifies XML in-place)
  await ppt.saveToFolder(WORK_FOLDER)
  console.log(`✓ Saved changes back to folder: ${WORK_FOLDER}`)

  // 6. Rebuild the folder into a PPTX file
  await PPTXTemplater.buildPptx(WORK_FOLDER, OUTPUT_REBUILT)
  console.log(`✓ Rebuilt PPTX: ${OUTPUT_REBUILT}`)

  // 7. Verify the rebuilt file works
  const verify = await PPTXTemplater.load(OUTPUT_REBUILT, { logLevel: 'warn' })
  console.log(`✓ Verified rebuilt PPTX — ${verify.slideCount} slide(s)`)

  // 8. Clean up extracted folder (optional)
  await fs.remove(WORK_FOLDER)
  console.log('✓ Cleaned up extracted folder')

  console.log('✓ Done:', OUTPUT_REBUILT)
}

main().catch(err => {
  console.error('Error:', err.message)
  process.exit(1)
})
