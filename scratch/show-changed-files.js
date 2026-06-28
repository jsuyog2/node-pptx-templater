const { PPTXTemplater } = require('../src/index.js')
const JSZip = require('jszip')
const fs = require('fs')
const path = require('path')

async function run() {
  const templatePath = path.resolve(__dirname, '../templates/officePPT.pptx')
  const savePath = path.resolve(__dirname, '../examples/output/compare-simple-save.pptx')

  const origZip = await JSZip.loadAsync(fs.readFileSync(templatePath))
  const modZip = await JSZip.loadAsync(fs.readFileSync(savePath))

  const files = [
    'docProps/app.xml',
    'ppt/slides/_rels/slide1.xml.rels',
    'ppt/notesSlides/notesSlide1.xml',
  ]

  for (const f of files) {
    console.log(`\n${'='.repeat(60)}`)
    console.log(`FILE: ${f}`)
    console.log('='.repeat(60))

    const origText = await origZip.file(f).async('text')
    const modText = await modZip.file(f).async('text')

    if (origText === modText) {
      console.log('  ✅ IDENTICAL')
      continue
    }

    console.log('\n--- ORIGINAL ---')
    console.log(origText)
    console.log('\n--- MODIFIED ---')
    console.log(modText)
  }
}

run().catch(console.error)
