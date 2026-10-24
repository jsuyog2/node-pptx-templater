const fs = require('fs')
const JSZip = require('jszip')

async function checkDetails() {
  const data = fs.readFileSync('showcase/templates/ChartPowerPointSample.pptx')
  const zip = await JSZip.loadAsync(data)

  for (let i = 1; i <= 7; i++) {
    const f = zip.file(`ppt/charts/chartEx${i}.xml`)
    if (f) {
      const xml = await f.async('text')
      const layoutIds = xml.match(/layoutId="[^"]+"/g)
      console.log(`\n=================== chartEx${i}.xml ===================`)
      console.log('Layout IDs:', layoutIds)
      // print first 1500 chars and structure
      console.log('--- Head of chartEx' + i + ' ---')
      console.log(xml.substring(0, 1000))
    }
  }
}

checkDetails().catch(console.error)
