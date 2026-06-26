/**
 * Diagnostic: verify duplicateSlide does not mutate the source slide.
 */
const fs = require('fs')
const { resolve } = require('path')
const JSZip = require('jszip')
const { PPTXTemplater } = require('../src/index.js')

async function readSlideParts(buffer, slidePath) {
  const zip = await JSZip.loadAsync(buffer)
  const xml = await zip.file(slidePath)?.async('text')
  const relsPath = slidePath.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
  const rels = await zip.file(relsPath)?.async('text')
  return { xml, rels }
}

async function main() {
  const template = resolve(__dirname, '../templates/sample.pptx')
  const ppt = await PPTXTemplater.load(template)

  const sourceIndex = 5
  const sourceInfo = ppt.slideManager.getSlideInfo(sourceIndex)
  const beforeXml = ppt.slideManager.getSlideXml(sourceIndex)
  const beforeRels = ppt.relationshipManager.getRelationships(sourceInfo.zipPath)

  const snapshotBefore = {
    xml: beforeXml,
    rels: JSON.stringify(beforeRels),
    zipPath: sourceInfo.zipPath,
  }

  ppt.duplicateSlide(sourceIndex, sourceIndex + 1)

  const afterXml = ppt.slideManager.getSlideXml(sourceIndex)
  const afterRels = ppt.relationshipManager.getRelationships(sourceInfo.zipPath)

  console.log('--- In-memory after duplicate (before save) ---')
  console.log('Source zipPath:', sourceInfo.zipPath)
  console.log('XML unchanged:', afterXml === snapshotBefore.xml)
  console.log('Rels unchanged:', JSON.stringify(afterRels) === snapshotBefore.rels)

  if (afterXml !== snapshotBefore.xml) {
    console.log('BEFORE len', snapshotBefore.xml.length, 'AFTER len', afterXml.length)
    for (let i = 0; i < Math.min(snapshotBefore.xml.length, afterXml.length); i++) {
      if (snapshotBefore.xml[i] !== afterXml[i]) {
        console.log('First diff at', i, snapshotBefore.xml.slice(i, i + 80))
        console.log('            vs', afterXml.slice(i, i + 80))
        break
      }
    }
  }

  const buf = await ppt.toBuffer()
  const parts = await readSlideParts(buf, 'ppt/slides/slide5.xml')
  console.log('\n--- After save: slide5.xml in output ---')
  console.log('Saved slide5 matches pre-duplicate snapshot:', parts.xml === snapshotBefore.xml)

  // Compare all slides in output
  const zip = await JSZip.loadAsync(buf)
  const slideFiles = Object.keys(zip.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
  console.log('\nSlide files:', slideFiles.sort().join(', '))

  // Check for duplicate rId references within each slide
  for (const f of slideFiles.sort()) {
    const xml = await zip.file(f).async('text')
    const relsPath = f.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
    const relsXml = await zip.file(relsPath)?.async('text')
    if (!relsXml) continue
    const validIds = [...relsXml.matchAll(/Id="(rId\d+)"/g)].map(m => m[1])
    const refs = [...xml.matchAll(/\br:(?:embed|id|link)="(rId\d+)"/g)].map(m => m[1])
    const missing = refs.filter(r => !validIds.includes(r))
    if (missing.length) {
      console.log(`${f}: INVALID refs`, missing)
    }
  }
}

main().catch(e => {
  console.error(e)
  process.exit(1)
})
