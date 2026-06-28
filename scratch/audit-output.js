const fs = require('fs')
const { resolve } = require('path')
const JSZip = require('jszip')
const { PPTXTemplater } = require('../src/index.js')

async function audit(path) {
  const zip = await JSZip.loadAsync(fs.readFileSync(path))
  const pres = await zip.file('ppt/presentation.xml').async('text')
  const presRels = await zip.file('ppt/_rels/presentation.xml.rels').async('text')
  const ct = await zip.file('[Content_Types].xml').async('text')

  console.log('\n===', path, '===')
  const slideRels = [...presRels.matchAll(/Id="(rId\d+)"[^>]*Type="[^"]*\/slide"[^>]*Target="([^"]+)"/g)]
  console.log('pres slide rels:', slideRels.map(m => `${m[1]}->${m[2]}`).join(', '))
  const sldIds = [...pres.matchAll(/<p:sldId[^>]*id="(\d+)"[^>]*r:id="(rId\d+)"/g)]
  console.log('sldIdLst:', sldIds.map(m => `${m[1]}->${m[2]}`).join(', '))

  for (const m of sldIds) {
    if (!slideRels.find(r => r[1] === m[2])) console.log('BAD sldId ref', m[0])
  }

  const slides = Object.keys(zip.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f)).sort()
  console.log('slide files:', slides.join(', '))

  // chart target sharing across slides
  const chartTargets = new Map()
  for (const s of slides) {
    const rels = await zip.file(s.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels')?.async('text')
    if (!rels) continue
    for (const m of rels.matchAll(/Type="[^"]*\/chart"[^>]*Target="([^"]+)"/g)) {
      const t = m[1]
      if (!chartTargets.has(t)) chartTargets.set(t, [])
      chartTargets.get(t).push(s)
    }
  }
  for (const [t, users] of chartTargets) {
    if (users.length > 1) console.log('SHARED CHART', t, 'by', users.join(', '))
  }

  // content type overrides for missing parts
  for (const m of ct.matchAll(/PartName="(\/ppt\/slides\/slide\d+\.xml)"/g)) {
    const part = m[1].slice(1)
    if (!zip.file(part)) console.log('CT override missing part', part)
  }
}

async function main() {
  const ppt = await PPTXTemplater.load(resolve(__dirname, '../templates/sample.pptx'))
  ppt.removeSlide(2)
  await ppt.duplicateSlide(5, 6)
  const out = resolve(__dirname, 'audit-remove-dup.pptx')
  await ppt.saveToFile(out)
  await audit(out)

  // full basic usage output if exists
  const basic = resolve(__dirname, '../examples/output/basic-output.pptx')
  if (fs.existsSync(basic)) await audit(basic)
}

main().catch(console.error)
