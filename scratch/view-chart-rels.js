const JSZip = require('jszip')
const fs = require('fs')
const path = require('path')

async function viewChartRels() {
  const filePath = path.resolve(__dirname, '../examples/output/office-modified.pptx')
  const zip = await JSZip.loadAsync(fs.readFileSync(filePath))

  if (zip.file('ppt/charts/_rels/chart4.xml.rels')) {
    const text = await zip.file('ppt/charts/_rels/chart4.xml.rels').async('text')
    console.log('=== ppt/charts/_rels/chart4.xml.rels ===')
    console.log(text)
  } else {
    console.log('❌ ppt/charts/_rels/chart4.xml.rels DOES NOT EXIST')
  }

  if (zip.file('ppt/charts/chart4.xml')) {
    const text = await zip.file('ppt/charts/chart4.xml').async('text')
    console.log('=== ppt/charts/chart4.xml (first 500 chars) ===')
    console.log(text.slice(0, 500))
  }
}

viewChartRels().catch(console.error)
