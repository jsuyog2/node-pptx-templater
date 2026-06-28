const { PPTXTemplater } = require('../src/index.js')
const JSZip = require('jszip')
const fs = require('fs')
const path = require('path')

// Comprehensive file-by-file comparison between original and any generated PPTX
async function deepCompare(origPath, modPath) {
  const origZip = await JSZip.loadAsync(fs.readFileSync(origPath))
  const modZip = await JSZip.loadAsync(fs.readFileSync(modPath))

  const origFiles = Object.keys(origZip.files).sort()
  const modFiles = Object.keys(modZip.files).sort()

  console.log('=== Files only in Original (MISSING from Modified) ===')
  origFiles.forEach(f => {
    if (!origZip.files[f].dir && !modZip.file(f)) {
      console.log(' ❌ MISSING:', f)
    }
  })

  console.log('\n=== Files only in Modified (EXTRA/NEW) ===')
  modFiles.forEach(f => {
    if (!modZip.files[f].dir && !origZip.file(f)) {
      console.log(' ➕ NEW:', f)
    }
  })

  console.log('\n=== Files in Both — Content Diff ===')
  const changed = []
  const unchanged = []

  for (const f of origFiles) {
    if (origZip.files[f].dir) continue
    if (!modZip.file(f)) continue

    const origContent = await origZip.file(f).async('nodebuffer')
    const modContent = await modZip.file(f).async('nodebuffer')

    // Compare as binary buffers
    if (origContent.equals(modContent)) {
      unchanged.push(f)
    } else {
      changed.push({ file: f, origSize: origContent.length, modSize: modContent.length })
    }
  }

  console.log(`\n${unchanged.length} files unchanged (binary identical)`)
  console.log(`\n${changed.length} files changed:`)
  changed.forEach(({ file, origSize, modSize }) => {
    const delta = modSize - origSize
    console.log(`  ⚠️  ${file} (${origSize} → ${modSize} bytes, ${delta >= 0 ? '+' : ''}${delta})`)
  })
}

async function run() {
  const templatePath = path.resolve(__dirname, '../templates/officePPT.pptx')

  // Case 1: Simple load + save (no slide operations)
  {
    const ppt = await PPTXTemplater.load(templatePath)
    const savePath = path.resolve(__dirname, '../examples/output/compare-simple-save.pptx')
    await ppt.save(savePath)
    console.log('\n========== CASE 1: Simple Load + Save ==========')
    await deepCompare(templatePath, savePath)
  }

  // Case 2: duplicateSlide(1, 2)
  {
    const ppt = await PPTXTemplater.load(templatePath)
    await ppt.duplicateSlide(1, 2)
    const savePath = path.resolve(__dirname, '../examples/output/compare-dup.pptx')
    await ppt.save(savePath)
    console.log('\n========== CASE 2: duplicateSlide(1, 2) ==========')
    await deepCompare(templatePath, savePath)
  }
}

run().catch(console.error)
