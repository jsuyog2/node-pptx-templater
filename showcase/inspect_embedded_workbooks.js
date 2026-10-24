const fs = require('fs')
const JSZip = require('jszip')

async function inspectWorkbooks() {
  const data = fs.readFileSync('showcase/templates/ChartPowerPointSample.pptx')
  const zip = await JSZip.loadAsync(data)

  const wbNames = [
    'Microsoft_Excel_Worksheet.xlsx',
    'Microsoft_Excel_Worksheet1.xlsx',
    'Microsoft_Excel_Worksheet2.xlsx',
    'Microsoft_Excel_Worksheet3.xlsx',
    'Microsoft_Excel_Worksheet4.xlsx',
    'Microsoft_Excel_Worksheet5.xlsx',
  ]

  for (const wbName of wbNames) {
    console.log(`\n================== ${wbName} ==================`)
    const wbFile = zip.file(`ppt/embeddings/${wbName}`)
    if (!wbFile) {
      console.log('File not found!')
      continue
    }
    const wbZip = await JSZip.loadAsync(await wbFile.async('nodebuffer'))
    const entries = Object.keys(wbZip.files)
    console.log('Entries:', entries)

    // Read sheet1.xml
    const sheet1 = wbZip.file('xl/worksheets/sheet1.xml')
    if (sheet1) {
      const sXml = await sheet1.async('text')
      // print row elements
      const rows = sXml.match(/<row[\s\S]*?<\/row>/g) || []
      console.log(`Row count in sheet1: ${rows.length}`)
      // Show first 5 rows
      rows.slice(0, 5).forEach((r, idx) => console.log(`  Row ${idx + 1}: ${r}`))
    }

    // Read sharedStrings if exists
    const ss = wbZip.file('xl/sharedStrings.xml')
    if (ss) {
      const ssXml = await ss.async('text')
      const count = ssXml.match(/<si>/g)?.length || 0
      console.log(`Shared strings count: ${count}`)
    } else {
      console.log('No sharedStrings.xml (inline strings or none)')
    }

    // Check table if exists
    const tables = entries.filter(e => e.includes('tables/table'))
    console.log('Tables:', tables)
    for (const t of tables) {
      console.log(`Content of ${t}:`)
      console.log(await wbZip.file(t).async('text'))
    }
  }
}

inspectWorkbooks().catch(console.error)
