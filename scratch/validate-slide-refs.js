const { resolve } = require('path')
const fs = require('fs')
const JSZip = require('jszip')
const { PPTXTemplater } = require('../src/index.js')
const { XMLParser } = require('../src/parsers/XMLParser.js')

async function validatePackage(buffer, label) {
  const zip = await JSZip.loadAsync(buffer)
  const parser = new XMLParser()

  const presXml = await zip.file('ppt/presentation.xml').async('text')
  const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('text')
  const presObj = parser.parse(presXml)
  const presRelsObj = parser.parse(presRelsXml)

  const rels = Array.isArray(presRelsObj.Relationships.Relationship)
    ? presRelsObj.Relationships.Relationship
    : [presRelsObj.Relationships.Relationship]

  const slideRels = new Map(
    rels
      .filter(r => r['@_Type'].endsWith('/slide'))
      .map(r => [r['@_Id'], r['@_Target']])
  )

  const sldIds = presObj['p:presentation']['p:sldIdLst']['p:sldId']
  const sldIdArr = Array.isArray(sldIds) ? sldIds : [sldIds]

  const errors = []
  for (const sldId of sldIdArr) {
    const rId = sldId['@_r:id']
    const id = sldId['@_id']
    if (!slideRels.has(rId)) {
      errors.push(`sldId ${id} references missing rId ${rId}`)
    }
  }

  // Check slide XML r:embed references match slide .rels
  const slideFiles = Object.keys(zip.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
  for (const slidePath of slideFiles.sort()) {
    const slideXml = await zip.file(slidePath).async('text')
    const relsPath = slidePath.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
    if (!zip.file(relsPath)) continue
    const slideRelsXml = await zip.file(relsPath).async('text')
    const slideRelsObj = parser.parse(slideRelsXml)
    const sRels = Array.isArray(slideRelsObj.Relationships.Relationship)
      ? slideRelsObj.Relationships.Relationship
      : [slideRelsObj.Relationships.Relationship]
    const validIds = new Set(sRels.map(r => r['@_Id']))

    const embedMatches = slideXml.matchAll(/r:(?:embed|id|link)="(rId\d+)"/g)
    for (const m of embedMatches) {
      if (!validIds.has(m[1])) {
        errors.push(`${slidePath} references ${m[1]} not in .rels`)
      }
    }
  }

  // Stale slide files
  const relSlideTargets = new Set(
    [...slideRels.values()].map(t => `ppt/${t.replace(/^\.\.\//, '').replace(/^slides\//, 'slides/')}`)
  )
  for (const slidePath of slideFiles) {
    if (![...slideRels.values()].some(t => slidePath.endsWith(t.split('/').pop()))) {
      // check via resolve
    }
  }

  // Orphan slide parts
  const referencedSlides = new Set(
    [...slideRels.values()].map(t => {
      const name = t.split('/').pop()
      return `ppt/slides/${name}`
    })
  )
  for (const f of slideFiles) {
    if (!referencedSlides.has(f)) {
      errors.push(`Orphan slide part: ${f}`)
    }
  }

  // Content types overrides for slides
  const ctXml = await zip.file('[Content_Types].xml').async('text')
  for (const f of slideFiles) {
    const partName = `/${f}`
    if (!ctXml.includes(`PartName="${partName}"`)) {
      errors.push(`Missing content type override for ${f}`)
    }
  }

  console.log(`\n=== ${label} ===`)
  console.log(`Slides: ${sldIdArr.length}, pres slide rels: ${slideRels.size}`)
  console.log(
    'sldIdLst:',
    sldIdArr.map(s => `${s['@_id']}->${s['@_r:id']}`).join(', ')
  )
  console.log(
    'pres rels:',
    [...slideRels.entries()].map(([id, t]) => `${id}->${t}`).join(', ')
  )
  if (errors.length) {
    console.log('ERRORS:')
    errors.forEach(e => console.log('  -', e))
  } else {
    console.log('OK: no structural reference errors detected')
  }
  return errors
}

async function main() {
  const template = resolve(__dirname, '../templates/sample.pptx')
  if (!fs.existsSync(template)) {
    console.error('Template not found')
    process.exit(1)
  }

  const scenarios = []

  // Scenario 1: duplicate slide 5 at position 6 (like basic-usage)
  {
    const ppt = await PPTXTemplater.load(template)
    ppt.duplicateSlide(5, 6)
    const buf = await ppt.toBuffer()
    scenarios.push(['duplicate slide 5 at 6', buf])
  }

  // Scenario 2: remove + duplicate + reorder (SlideSectionsSync)
  {
    const ppt = await PPTXTemplater.load(template)
    ppt.removeSlide(2)
    ppt.duplicateSlide(1, 3)
    const buf = await ppt.toBuffer()
    scenarios.push(['remove 2, duplicate 1 at 3', buf])
  }

  // Scenario 3: multiple duplicates
  {
    const ppt = await PPTXTemplater.load(template)
    for (let i = 0; i < 3; i++) {
      ppt.duplicateSlide(1, ppt.slideCount + 1)
    }
    const buf = await ppt.toBuffer()
    scenarios.push(['triple duplicate slide 1', buf])
  }

  let totalErrors = 0
  for (const [label, buf] of scenarios) {
    const errors = await validatePackage(buf, label)
    totalErrors += errors.length
  }

  process.exit(totalErrors > 0 ? 1 : 0)
}

main().catch(err => {
  console.error(err)
  process.exit(1)
})
