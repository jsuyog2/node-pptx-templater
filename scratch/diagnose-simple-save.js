const { PPTXTemplater } = require('../src/index.js')
const { XMLParser } = require('../src/parsers/XMLParser.js')
const JSZip = require('jszip')
const fs = require('fs')
const path = require('path')

async function diagnose() {
  const templatePath = path.resolve(__dirname, '../templates/officePPT.pptx')
  const outputPath = path.resolve(__dirname, '../examples/output/office-saved.pptx')
  
  if (!fs.existsSync(path.dirname(outputPath))) {
    fs.mkdirSync(path.dirname(outputPath), { recursive: true })
  }

  const ppt = await PPTXTemplater.load(templatePath)
  await ppt.save(outputPath)
  console.log('✅ Saved simple load-and-save to:', outputPath)

  const origZip = await JSZip.loadAsync(fs.readFileSync(templatePath))
  const modZip = await JSZip.loadAsync(fs.readFileSync(outputPath))

  const origFiles = Object.keys(origZip.files).sort()
  const modFiles = Object.keys(modZip.files).sort()

  console.log(`Original files count: ${origFiles.length}. Saved files count: ${modFiles.length}.`)

  const parser = new XMLParser()

  // Compare Content Types
  const origContentTypesXml = await origZip.file('[Content_Types].xml').async('text')
  const modContentTypesXml = await modZip.file('[Content_Types].xml').async('text')

  const origCT = parser.parse(origContentTypesXml, '[Content_Types].xml')
  const modCT = parser.parse(modContentTypesXml, '[Content_Types].xml')

  const getOverrides = (ctObj) => {
    let list = ctObj.Types.Override || []
    if (!Array.isArray(list)) list = [list]
    return list.map(o => ({ part: o['@_PartName'], type: o['@_ContentType'] }))
  }

  const origOvs = getOverrides(origCT)
  const modOvs = getOverrides(modCT)

  console.log(`Original has ${origOvs.length} overrides. Saved has ${modOvs.length} overrides.`)

  console.log('\n=== Overrides Diff ===')
  origOvs.forEach(o => {
    const found = modOvs.find(mo => mo.part === o.part)
    if (!found) {
      console.log(` ❌ Missing Override in Saved: ${o.part} (${o.type})`)
    } else if (found.type !== o.type) {
      console.log(` ⚠️ Mismatched Override Content Type for ${o.part}: Original=${o.type}, Saved=${found.type}`)
    }
  })

  modOvs.forEach(o => {
    const found = origOvs.find(oo => oo.part === o.part)
    if (!found) {
      console.log(` ➕ Extra Override in Saved: ${o.part} (${o.type})`)
    }
  })
}

diagnose().catch(console.error)
