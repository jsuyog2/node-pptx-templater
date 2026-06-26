/**
 * Reproduce basic-usage slide ops and validate package structure.
 */
const fs = require('fs')
const { resolve } = require('path')
const JSZip = require('jszip')
const { PPTXTemplater } = require('../src/index.js')

async function validateBuffer(buffer, label) {
  const zip = await JSZip.loadAsync(buffer)
  const presXml = await zip.file('ppt/presentation.xml').async('text')
  const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('text')

  const slideRels = [...presRelsXml.matchAll(/Id="(rId\d+)"[^>]*Type="[^"]*\/slide"[^>]*Target="([^"]+)"/g)]
  const slideRelMap = new Map(slideRels.map(m => [m[1], m[2]]))
  const sldIdRefs = [...presXml.matchAll(/<p:sldId[^>]*id="(\d+)"[^>]*r:id="(rId\d+)"/g)]

  const errors = []
  for (const m of sldIdRefs) {
    if (!slideRelMap.has(m[2])) errors.push(`${label}: sldId ${m[1]} missing rId ${m[2]}`)
  }

  const slideFiles = Object.keys(zip.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f)).sort()
  const orphans = slideFiles.filter(f => ![...slideRelMap.values()].some(t => f.endsWith(t.split('/').pop())))

  if (orphans.length) errors.push(`${label}: orphan slides ${orphans.join(', ')}`)

  for (const f of slideFiles) {
    const xml = await zip.file(f).async('text')
    const relsPath = f.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
    if (!zip.file(relsPath)) continue
    const rels = await zip.file(relsPath).async('text')
    const valid = new Set([...rels.matchAll(/Id="(rId\d+)"/g)].map(m => m[1]))
    for (const ref of xml.matchAll(/\br:(?:embed|id|link)="(rId\d+)"/g)) {
      if (!valid.has(ref[1])) errors.push(`${label}: ${f} refs ${ref[1]} not in rels`)
    }
  }

  return errors
}

async function runBasicUsageFlow({ awaitDuplicate }) {
  const ppt = await PPTXTemplater.load(resolve(__dirname, '../templates/sample.pptx'))
  ppt.useSlide(1)
  ppt.replaceText({ '{{title}}': 'QBR', '{{subtitle}}': 'Q1' })
  ppt.updateChart('Chart', {
    categories: ['Q1', 'Q2'],
    series: [{ name: 'A', values: [1, 2] }],
  })
  ppt.removeSlide(2)

  const slideNumber = 5
  for (let index = 1; index <= 1; index++) {
    const p = ppt.duplicateSlide(slideNumber, slideNumber + index)
    if (awaitDuplicate) await p
  }

  return ppt.toBuffer()
}

async function main() {
  for (const awaitDuplicate of [true, false]) {
    const buf = await runBasicUsageFlow({ awaitDuplicate })
    const errors = await validateBuffer(buf, awaitDuplicate ? 'with-await' : 'NO-await')
    console.log(awaitDuplicate ? 'WITH await:' : 'WITHOUT await:', errors.length ? errors : 'OK')
  }
}

main().catch(e => {
  console.error(e)
  process.exit(1)
})
