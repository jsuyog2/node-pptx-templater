const JSZip = require('jszip')
const fs = require('fs')
const path = require('path')

async function checkChartXml() {
  const filePath = path.resolve(__dirname, '../examples/output/office-modified.pptx')
  const zip = await JSZip.loadAsync(fs.readFileSync(filePath))

  const chartXml = await zip.file('ppt/charts/chart4.xml').async('text')
  
  // Find all occurrences of r:id, r:embed, r:link, or similar r: attributes
  const matches = [...chartXml.matchAll(/r:(id|embed|link)="([^"]+)"/g)]
  console.log('=== Relationship IDs in chart4.xml ===')
  matches.forEach(m => {
    console.log(`  r:${m[1]}="${m[2]}"`)
  })
}

checkChartXml().catch(console.error)
