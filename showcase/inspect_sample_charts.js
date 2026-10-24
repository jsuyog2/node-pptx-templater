const fs = require('fs')
const JSZip = require('jszip')

async function audit() {
  const data = fs.readFileSync('showcase/templates/ChartPowerPointSample.pptx')
  const zip = await JSZip.loadAsync(data)

  console.log('--- Content Types Overrides for Charts ---')
  const ct = await zip.file('[Content_Types].xml').async('text')
  const ctMatches = ct.match(/<Override[^>]+PartName="\/ppt\/charts[^"]*"[^>]*\/>/g) || []
  ctMatches.forEach(m => console.log(' ', m))

  console.log('\n--- Slide Relationships to Charts ---')
  for (let i = 1; i <= 12; i++) {
    const relFile = zip.file(`ppt/slides/_rels/slide${i}.xml.rels`)
    if (relFile) {
      const relXml = await relFile.async('text')
      const rels = relXml.match(/<Relationship[^>]+Target="[^"]*chart[^"]*"[^>]*\/>/g)
      if (rels) {
        console.log(`Slide ${i}:`)
        rels.forEach(r => console.log('  ', r))
      }
    }
  }

  console.log('\n--- ChartEx Files ---')
  for (let i = 1; i <= 7; i++) {
    const f = zip.file(`ppt/charts/chartEx${i}.xml`)
    if (f) {
      const xml = await f.async('text')
      const layoutIds = xml.match(/layoutId="[^"]+"/g)
      const relF = zip.file(`ppt/charts/_rels/chartEx${i}.xml.rels`)
      const rels = relF ? await relF.async('text') : 'none'
      console.log(`chartEx${i}.xml: layoutIds=${JSON.stringify(layoutIds)}`)
      console.log(`  rels: ${rels.replace(/\r?\n/g, ' ')}`)
    }
  }

  console.log('\n--- Slides 2, 3, 4, 5 GraphicFrames ---')
  for (let i = 2; i <= 5; i++) {
    const sFile = zip.file(`ppt/slides/slide${i}.xml`)
    if (sFile) {
      const sXml = await sFile.async('text')
      // Look for graphicFrame and graphicData
      const gfMatches = sXml.match(/<p:graphicFrame[\s\S]*?<\/p:graphicFrame>/g) || []
      console.log(`Slide ${i} has ${gfMatches.length} graphicFrame(s):`)
      gfMatches.forEach((gf, idx) => {
        const uri = gf.match(/uri="([^"]+)"/)?.[1]
        const chartTag = gf.match(/<(?:c|cx):chart[^>]+>/)?.[0]
        console.log(`  Frame ${idx + 1}: uri=${uri}, chartTag=${chartTag}`)
      })
    }
  }
}

audit().catch(console.error)
