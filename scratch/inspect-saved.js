const JSZip = require('jszip')
const fs = require('fs')
const path = require('path')

async function inspect() {
  const filePath = path.resolve(__dirname, '../examples/output/office-modified.pptx')
  if (!fs.existsSync(filePath)) {
    console.error('File not found:', filePath)
    return
  }

  const data = fs.readFileSync(filePath)
  const zip = await JSZip.loadAsync(data)

  const presXml = await zip.file('ppt/presentation.xml').async('text')
  console.log('=== Saved ppt/presentation.xml ===')
  console.log(presXml)
}

inspect().catch(console.error)
