const { PPTXTemplater } = require('../src/index.js')
const path = require('path')

async function run() {
  const templatePath = path.resolve(__dirname, '../templates/officePPT.pptx')
  const outputPath = path.resolve(__dirname, '../examples/output/office-complex.pptx')

  console.log('📂 Loading template:', templatePath)
  const ppt = await PPTXTemplater.load(templatePath)
  console.log(`📊 Loaded ${ppt.slideCount} slides`)

  // 1. Duplicate slide 2 to position 4
  console.log('👥 Duplicating slide 2 to position 4...')
  await ppt.duplicateSlide(2, 4)
  console.log(`Slides: ${ppt.slideCount}`)

  // 2. Remove slide 1
  console.log('🗑️ Removing slide 1...')
  await ppt.removeSlide(1)
  console.log(`Slides: ${ppt.slideCount}`)

  // 3. Move slide 3 to 1
  console.log('🔄 Moving slide 3 to 1...')
  ppt.moveSlide(3, 1)
  console.log(`Slides: ${ppt.slideCount}`)

  console.log('💾 Saving to:', outputPath)
  await ppt.save(outputPath)

  // Run validation
  const { execSync } = require('child_process')
  try {
    const output = execSync(`node scratch/validate-pptx.js examples/output/office-complex.pptx`, { encoding: 'utf-8' })
    console.log(output)
  } catch (e) {
    console.error('Validation script execution failed:', e)
  }
}

run().catch(console.error)
