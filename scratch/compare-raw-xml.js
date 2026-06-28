const JSZip = require('jszip')
const fs = require('fs')
const path = require('path')

async function compare() {
  const origPath = path.resolve(__dirname, '../templates/officePPT.pptx')
  const savedPath = path.resolve(__dirname, '../examples/output/office-saved.pptx')

  const origZip = await JSZip.loadAsync(fs.readFileSync(origPath))
  const savedZip = await JSZip.loadAsync(fs.readFileSync(savedPath))

  async function compareFile(zipPath) {
    console.log(`\n=== Comparing ${zipPath} ===`)
    const origText = await origZip.file(zipPath).async('text')
    const savedText = await savedZip.file(zipPath).async('text')

    if (origText === savedText) {
      console.log('  ✅ Identical text!')
      return
    }

    console.log('  ⚠️ Differs!')
    console.log(`  Original length: ${origText.length}, Saved length: ${savedText.length}`)

    // Print first difference
    let diffIndex = -1
    for (let i = 0; i < Math.min(origText.length, savedText.length); i++) {
      if (origText[i] !== savedText[i]) {
        diffIndex = i
        break
      }
    }
    if (diffIndex !== -1) {
      console.log(`  First diff at character ${diffIndex}:`)
      console.log(`    Original around: "${origText.slice(Math.max(0, diffIndex - 20), diffIndex + 40).replace(/\n/g, '\\n')}"`)
      console.log(`    Saved around:    "${savedText.slice(Math.max(0, diffIndex - 20), diffIndex + 40).replace(/\n/g, '\\n')}"`)
    }
  }

  await compareFile('[Content_Types].xml')
  await compareFile('ppt/presentation.xml')
  await compareFile('ppt/_rels/presentation.xml.rels')
}

compare().catch(console.error)
