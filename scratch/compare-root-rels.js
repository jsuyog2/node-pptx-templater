const JSZip = require('jszip')
const fs = require('fs')
const path = require('path')

async function compare() {
  const origPath = path.resolve(__dirname, '../templates/officePPT.pptx')
  const savedPath = path.resolve(__dirname, '../examples/output/office-saved.pptx')

  const origZip = await JSZip.loadAsync(fs.readFileSync(origPath))
  const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath))

  console.log('=== Original _rels/.rels ===')
  console.log(await origZip.file('_rels/.rels').async('text'))

  console.log('\n=== Saved _rels/.rels ===')
  console.log(await savedZip.file('_rels/.rels').async('text'))
}

compare().catch(console.error)
