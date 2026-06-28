const JSZip = require('jszip')
const fs = require('fs')
const path = require('path')

async function viewCustomXml() {
  const origPath = path.resolve(__dirname, '../templates/officePPT.pptx')
  const savedPath = path.resolve(__dirname, '../examples/output/office-saved.pptx')

  const origZip = await JSZip.loadAsync(fs.readFileSync(origPath))
  const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath))

  const customXmlFiles = Object.keys(origZip.files).filter(f => f.startsWith('customXml/')).sort()

  for (const f of customXmlFiles) {
    if (origZip.files[f].dir) continue
    console.log(`\n=== File: ${f} ===`)
    const origText = await origZip.file(f).async('text')
    const savedText = await savedZip.file(f).async('text')
    
    if (origText === savedText) {
      console.log('  ✅ Identical!')
    } else {
      console.log('  ⚠️ Differs!')
      console.log('  Original:', origText)
      console.log('  Saved:   ', savedText)
    }
  }
}

viewCustomXml().catch(console.error)
