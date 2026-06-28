const { PPTXTemplater } = require('../src/index.js')
const { XMLParser } = require('../src/parsers/XMLParser.js')
const JSZip = require('jszip')
const fs = require('fs')
const path = require('path')

async function diagnose() {
  const templatePath = path.resolve(__dirname, '../templates/officePPT.pptx')
  const outputPath = path.resolve(__dirname, '../examples/output/office-modified.pptx')
  
  if (!fs.existsSync(path.dirname(outputPath))) {
    fs.mkdirSync(path.dirname(outputPath), { recursive: true })
  }

  const ppt = await PPTXTemplater.load(templatePath)
  console.log(`📊 Loaded ${ppt.slideCount} slides`)

  // Step 1: Duplicate slide 1 to position 2
  console.log('\n👥 Duplicating slide 1 to position 2...')
  await ppt.duplicateSlide(1, 2)
  console.log(`Slide count after duplication: ${ppt.slideCount}`)

  // Step 2: Remove slide 3 (original slide 2)
  console.log('🗑️ Removing slide 3 (original slide 2)...')
  await ppt.removeSlide(3)
  console.log(`Slide count after removal: ${ppt.slideCount}`)

  console.log('💾 Saving to:', outputPath)
  await ppt.save(outputPath)

  const origZip = await JSZip.loadAsync(fs.readFileSync(templatePath))
  const modZip = await JSZip.loadAsync(fs.readFileSync(outputPath))
  const parser = new XMLParser()

  // 1. Check Content Types
  const origContentTypesXml = await origZip.file('[Content_Types].xml').async('text')
  const modContentTypesXml = await modZip.file('[Content_Types].xml').async('text')
  const origCT = parser.parse(origContentTypesXml, '[Content_Types].xml')
  const modCT = parser.parse(modContentTypesXml, '[Content_Types].xml')

  const getOverrides = (ctObj) => {
    let list = ctObj.Types.Override || []
    if (!Array.isArray(list)) list = [list]
    return list.map(o => ({ part: o['@_PartName'], type: o['@_ContentType'] }))
  }

  const origOvs = getOverrides(origCT)
  const modOvs = getOverrides(modCT)

  console.log('\n=== CONTENT TYPES OVERRIDES ===')
  console.log(`Original: ${origOvs.length} overrides. Modified: ${modOvs.length} overrides.`)
  
  console.log('\nOverrides only in Modified:')
  modOvs.forEach(o => {
    if (!origOvs.find(oo => oo.part === o.part)) {
      console.log(` + ${o.part} (${o.type})`)
    }
  })

  console.log('\nOverrides only in Original:')
  origOvs.forEach(o => {
    if (!modOvs.find(mo => mo.part === o.part)) {
      console.log(` - ${o.part} (${o.type})`)
    }
  })

  // 2. Check Slide IDs and Sections in ppt/presentation.xml
  console.log('\n=== PRESENTATION STRUCTURE ===')
  const origPresXml = await origZip.file('ppt/presentation.xml').async('text')
  const modPresXml = await modZip.file('ppt/presentation.xml').async('text')
  const origPres = parser.parse(origPresXml, 'ppt/presentation.xml')
  const modPres = parser.parse(modPresXml, 'ppt/presentation.xml')

  const getSldIds = (presObj) => {
    let list = presObj['p:presentation']?.['p:sldIdLst']?.['p:sldId'] || []
    if (!Array.isArray(list)) list = [list]
    return list.map(s => ({ id: s['@_id'], rId: s['@_r:id'] }))
  }

  const origSldIds = getSldIds(origPres)
  const modSldIds = getSldIds(modPres)

  console.log('Original Slide IDs:', origSldIds)
  console.log('Modified Slide IDs:', modSldIds)

  // Section check
  const getSections = (presObj) => {
    const extLst = presObj['p:presentation']?.['p:extLst']?.['p:ext'] || []
    const sectionExt = (Array.isArray(extLst) ? extLst : [extLst]).find(
      e => e['@_uri'] === '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}'
    )
    if (!sectionExt) return null
    let sectionLst = sectionExt['p14:sectionLst']?.['p14:section'] || []
    if (!Array.isArray(sectionLst)) sectionLst = [sectionLst]
    return sectionLst.map(sec => {
      let ids = sec['p14:sldIdLst']?.['p14:sldId'] || []
      if (!Array.isArray(ids)) ids = [ids]
      return {
        name: sec['@_name'],
        id: sec['@_id'],
        slideIds: ids.map(x => x['@_id'])
      }
    })
  }

  console.log('Original Sections:', getSections(origPres))
  console.log('Modified Sections:', getSections(modPres))

  // 3. Check presentation.xml.rels
  console.log('\n=== PRESENTATION RELATIONSHIPS ===')
  const getPresRels = async (zipObj) => {
    const text = await zipObj.file('ppt/_rels/presentation.xml.rels').async('text')
    const relsObj = parser.parse(text, 'ppt/_rels/presentation.xml.rels')
    let list = relsObj.Relationships?.Relationship || []
    if (!Array.isArray(list)) list = [list]
    return list.map(r => ({ id: r['@_Id'], type: r['@_Type'], target: r['@_Target'] }))
  }

  const origPresRels = await getPresRels(origZip)
  const modPresRels = await getPresRels(modZip)

  console.log('Original slide rels:')
  origPresRels.filter(r => r.type.endsWith('/slide')).forEach(r => console.log(`  id=${r.id} target=${r.target}`))
  console.log('Modified slide rels:')
  modPresRels.filter(r => r.type.endsWith('/slide')).forEach(r => console.log(`  id=${r.id} target=${r.target}`))

  // 4. Trace the Relationships of Each Slide in Modified
  console.log('\n=== ACTIVE SLIDES RELATIONSHIPS ===')
  for (let i = 1; i <= ppt.slideCount; i++) {
    const slideZipPath = `ppt/slides/slide${i}.xml`
    const relsPath = `ppt/slides/_rels/slide${i}.xml.rels`
    console.log(`\nSlide: ${slideZipPath}`)
    
    // Find slide override
    const ct = modOvs.find(o => o.part === '/' + slideZipPath)
    console.log(`  Content Type Override: ${ct ? ct.type : '❌ MISSING'}`)

    // Check relationship file
    if (modZip.file(relsPath)) {
      const text = await modZip.file(relsPath).async('text')
      const relsObj = parser.parse(text, relsPath)
      let list = relsObj.Relationships?.Relationship || []
      if (!Array.isArray(list)) list = [list]
      list.forEach(r => {
        const targetResolved = path.posix.join('ppt/slides', r['@_Target'])
        const exists = modZip.file(targetResolved) ? 'exists' : '❌ DOES NOT EXIST'
        console.log(`  Relation: id=${r['@_Id']} type=${r['@_Type'].split('/').pop()} target=${r['@_Target']} (${exists})`)
      })
    } else {
      console.log('  ❌ Relationship file missing')
    }
  }
}

diagnose().catch(console.error)
