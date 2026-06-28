const JSZip = require('jszip')
const fs = require('fs')
const path = require('path')

async function diff() {
  const originalPath = path.resolve(__dirname, '../templates/officePPT.pptx')
  const modifiedPath = path.resolve(__dirname, '../examples/output/office-modified.pptx')

  const origZip = await JSZip.loadAsync(fs.readFileSync(originalPath))
  const modZip = await JSZip.loadAsync(fs.readFileSync(modifiedPath))

  const origFiles = new Set(Object.keys(origZip.files))
  const modFiles = new Set(Object.keys(modZip.files))

  console.log('=== Files in Original but NOT in Modified ===')
  for (const f of origFiles) {
    if (!modFiles.has(f)) {
      console.log(' -', f)
    }
  }

  console.log('\n=== Files in Modified but NOT in Original ===')
  for (const f of modFiles) {
    if (!origFiles.has(f)) {
      console.log(' +', f)
    }
  }
}

diff().catch(console.error)
