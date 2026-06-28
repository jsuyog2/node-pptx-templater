const JSZip = require('jszip')
const fs = require('fs')
const path = require('path')

async function viewCustomProps() {
  const origPath = path.resolve(__dirname, '../templates/officePPT.pptx')
  const savedPath = path.resolve(__dirname, '../examples/output/office-saved.pptx')

  const origZip = await JSZip.loadAsync(fs.readFileSync(origPath))
  const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath))

  console.log('=== Original docProps/custom.xml ===')
  console.log(await origZip.file('docProps/custom.xml').async('text'))

  console.log('\n=== Saved docProps/custom.xml ===')
  console.log(await savedZip.file('docProps/custom.xml').async('text'))
}

viewCustomProps().catch(console.error)
