const fs = require('fs')
const JSZip = require('jszip')
;(async () => {
  const zip = await JSZip.loadAsync(fs.readFileSync('templates/sample.pptx'))
  const ct = await zip.file('[Content_Types].xml').async('text')
  for (const m of ct.matchAll(/PartName="([^"]+)" ContentType="([^"]+)"/g)) {
    if (m[1].includes('chart') || m[1].includes('embed')) console.log(m[1], '=>', m[2])
  }
})()
