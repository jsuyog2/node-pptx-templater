const JSZip = require('jszip')
const fs = require('fs')
const path = require('path')

async function viewTags() {
  const filePath = path.resolve(__dirname, '../examples/output/office-modified.pptx')
  const zip = await JSZip.loadAsync(fs.readFileSync(filePath))

  const tag4 = await zip.file('ppt/tags/tag4.xml').async('text')
  console.log('=== ppt/tags/tag4.xml ===')
  console.log(tag4)

  const tag16 = await zip.file('ppt/tags/tag16.xml').async('text')
  console.log('\n=== ppt/tags/tag16.xml ===')
  console.log(tag16)
}

viewTags().catch(console.error)
