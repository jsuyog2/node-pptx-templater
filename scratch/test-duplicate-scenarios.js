const fs = require('fs')
const { resolve } = require('path')
const { PPTXTemplater } = require('../src/index.js')
const { validatePackageIntegrity } = require('../tests/helpers/packageIntegrity.js')

async function testDuplicate(label, setup, sourceIndex, atPosition) {
  const ppt = await PPTXTemplater.load(resolve(__dirname, '../templates/sample.pptx'))
  if (setup) await setup(ppt)

  const info = ppt.slideManager.getSlideInfo(sourceIndex)
  const beforeXml = ppt.slideManager.getSlideXml(sourceIndex)
  const beforeRels = JSON.stringify(ppt.relationshipManager.getRelationships(info.zipPath))

  ppt.duplicateSlide(sourceIndex, atPosition ?? sourceIndex + 1)

  const afterXml = ppt.slideManager.getSlideXml(sourceIndex)
  const afterRels = JSON.stringify(ppt.relationshipManager.getRelationships(info.zipPath))
  const memOk = afterXml === beforeXml && afterRels === beforeRels

  const buf = await ppt.toBuffer()
  const errors = await validatePackageIntegrity(buf)

  // find saved path for source after normalize (by slide index)
  const savedInfo = ppt.slideManager.getSlideInfo(sourceIndex)
  const savedXml = ppt.slideManager.getSlideXml(sourceIndex)
  const savedOk = savedXml === beforeXml

  console.log(`\n[${label}]`)
  console.log('  source mem unchanged:', memOk)
  console.log('  source saved unchanged:', savedOk)
  console.log('  package errors:', errors.length ? errors : 'none')

  // Check clone has unique rel ids vs source
  const cloneIndex = atPosition ?? ppt.slideCount
  const cloneInfo = ppt.slideManager.getSlideInfo(
    atPosition !== undefined ? atPosition : ppt.slideCount
  )
  const srcRels = ppt.relationshipManager.getRelationships(savedInfo.zipPath)
  const cloneRels = ppt.relationshipManager.getRelationships(cloneInfo.zipPath)
  const sharedTargets = cloneRels.filter(cr =>
    srcRels.some(sr => sr.target === cr.target && sr.type === cr.type)
  )
  console.log('  clone path:', cloneInfo.zipPath)
  console.log('  shared rel targets with source:', sharedTargets.map(r => `${r.type.split('/').pop()}->${r.target}`).join(', ') || 'none')

  return { memOk, savedOk, errors }
}

async function main() {
  let failed = 0

  const cases = [
    ['slide1 chart', null, 1, 2],
    ['slide6 images', null, 6, 7],
    ['slide1 after chart update', async ppt => {
      ppt.useSlide(1)
      ppt.updateChart('Chart', { categories: ['A'], series: [{ name: 'S', values: [1] }] })
    }, 1, 2],
    ['slide3 table after edit', async ppt => {
      ppt.useSlide(3)
      ppt.updateTable('Table', [['H1', 'H2', 'H3', 'H4', 'H5'], ['a', 'b', 'c', 'd', 'e']])
    }, 3, 4],
    ['slide5 basic-usage target', null, 5, 6],
    ['slide5 dup at end', null, 5, undefined],
  ]

  for (const [label, setup, src, pos] of cases) {
    const r = await testDuplicate(label, setup, src, pos)
    if (!r.memOk || !r.savedOk || r.errors.length) failed++
  }

  process.exit(failed > 0 ? 1 : 0)
}

main().catch(e => {
  console.error(e)
  process.exit(1)
})
