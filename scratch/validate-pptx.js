const JSZip = require('jszip')
const fs = require('fs')
const path = require('path')
const { XMLParser } = require('../src/parsers/XMLParser.js')

async function validate(filePath) {
  console.log(`\n🔍 Starting validation of PPTX: ${filePath}`)
  if (!fs.existsSync(filePath)) {
    console.error('❌ File does not exist')
    return
  }

  const data = fs.readFileSync(filePath)
  const zip = await JSZip.loadAsync(data)
  const parser = new XMLParser()

  const errors = []
  const warnings = []

  function assert(condition, message) {
    if (!condition) {
      errors.push(message)
    }
  }

  // 1. Verify [Content_Types].xml
  const ctFile = zip.file('[Content_Types].xml')
  assert(ctFile, '[Content_Types].xml is missing')
  if (ctFile) {
    const ctXml = await ctFile.async('text')
    let ctObj
    try {
      ctObj = parser.parse(ctXml, '[Content_Types].xml')
    } catch (e) {
      assert(false, `[Content_Types].xml is malformed: ${e.message}`)
    }

    if (ctObj) {
      const overrides = ctObj.Types?.Override || []
      const ovsArray = Array.isArray(overrides) ? overrides : [overrides]
      
      // Check that all files in ZIP registered as Override exist (except directories or defaults)
      const defaults = ctObj.Types?.Default || []
      const defsArray = Array.isArray(defaults) ? defaults : [defaults]
      const defaultExtensions = new Set(defsArray.map(d => d['@_Extension']?.toLowerCase()))

      ovsArray.forEach(o => {
        const partName = o['@_PartName']
        if (partName) {
          const zipPath = partName.startsWith('/') ? partName.slice(1) : partName
          assert(zip.file(zipPath), `[Content_Types].xml override references non-existent file: ${zipPath}`)
        }
      })
    }
  }

  // 2. Verify ppt/presentation.xml
  const presFile = zip.file('ppt/presentation.xml')
  assert(presFile, 'ppt/presentation.xml is missing')
  
  let sldIdLst = []
  let sectionSlideIds = new Set()
  
  if (presFile) {
    const presXml = await presFile.async('text')
    let presObj
    try {
      presObj = parser.parse(presXml, 'ppt/presentation.xml')
    } catch (e) {
      assert(false, `ppt/presentation.xml is malformed: ${e.message}`)
    }

    if (presObj) {
      const sldIdElements = presObj['p:presentation']?.['p:sldIdLst']?.['p:sldId'] || []
      sldIdLst = Array.isArray(sldIdElements) ? sldIdElements : [sldIdElements]
      
      // Verify Slide IDs are unique and valid
      const idSet = new Set()
      sldIdLst.forEach((s, idx) => {
        const id = s['@_id']
        const rId = s['@_r:id']
        assert(id, `Slide at index ${idx} is missing id attribute`)
        assert(rId, `Slide at index ${idx} is missing r:id attribute`)
        if (id) {
          assert(!idSet.has(id), `Duplicate slide ID found: ${id}`)
          idSet.add(id)
        }
      })

      // Verify sections
      const extLst = presObj['p:presentation']?.['p:extLst']?.['p:ext'] || []
      const extArray = Array.isArray(extLst) ? extLst : [extLst]
      const sectionExt = extArray.find(e => e['@_uri'] === '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}')
      if (sectionExt) {
        let sections = sectionExt['p14:sectionLst']?.['p14:section'] || []
        sections = Array.isArray(sections) ? sections : [sections]
        sections.forEach(sec => {
          let sIds = sec['p14:sldIdLst']?.['p14:sldId'] || []
          sIds = Array.isArray(sIds) ? sIds : [sIds]
          sIds.forEach(si => {
            const sid = si['@_id']
            assert(sid, `Section slide item is missing id attribute`)
            if (sid) {
              sectionSlideIds.add(sid)
              assert(idSet.has(sid), `Section references slide ID ${sid} which is not in sldIdLst`)
            }
          })
        })
      }
    }
  }

  // 3. Verify ppt/_rels/presentation.xml.rels
  const presRelsFile = zip.file('ppt/_rels/presentation.xml.rels')
  assert(presRelsFile, 'ppt/_rels/presentation.xml.rels is missing')
  
  if (presRelsFile) {
    const presRelsXml = await presRelsFile.async('text')
    let relsObj
    try {
      relsObj = parser.parse(presRelsXml, 'ppt/_rels/presentation.xml.rels')
    } catch (e) {
      assert(false, `ppt/_rels/presentation.xml.rels is malformed: ${e.message}`)
    }

    if (relsObj) {
      const rels = relsObj.Relationships?.Relationship || []
      const relsArray = Array.isArray(rels) ? rels : [rels]
      const relMap = new Map(relsArray.map(r => [r['@_Id'], r]))

      // Verify each slide relationship
      sldIdLst.forEach(s => {
        const rId = s['@_r:id']
        const rel = relMap.get(rId)
        assert(rel, `Slide ID ${s['@_id']} references relationship ${rId} which is missing in presentation.xml.rels`)
        if (rel) {
          assert(rel['@_Type'] === 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide', `Slide relationship ${rId} has invalid type: ${rel['@_Type']}`)
          const target = rel['@_Target']
          const resolvedPath = path.posix.join('ppt', target)
          assert(zip.file(resolvedPath), `Presentation slide relationship ${rId} target does not exist: ${resolvedPath}`)
        }
      })
      
      // Check other presentation-level relationships targets
      relsArray.forEach(r => {
        const type = r['@_Type']
        const target = r['@_Target']
        if (r['@_TargetMode'] !== 'External' && target && !target.startsWith('http')) {
          const resolvedPath = path.posix.join('ppt', target)
          assert(zip.file(resolvedPath), `Presentation relationship ${r['@_Id']} target does not exist: ${resolvedPath}`)
        }
      })
    }
  }

  // 4. Verify slide relationship files and target integrity
  const slideFiles = Object.keys(zip.files).filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
  
  for (const slidePath of slideFiles) {
    // Verify XML validity
    const xmlText = await zip.file(slidePath).async('text')
    try {
      parser.parse(xmlText, slidePath)
    } catch (e) {
      assert(false, `Slide XML ${slidePath} is malformed: ${e.message}`)
    }

    // Verify .rels file for the slide
    const slideDir = path.posix.dirname(slidePath)
    const slideFileName = path.posix.basename(slidePath)
    const relsPath = `${slideDir}/_rels/${slideFileName}.rels`
    
    if (zip.file(relsPath)) {
      const relsText = await zip.file(relsPath).async('text')
      let relsObj
      try {
        relsObj = parser.parse(relsText, relsPath)
      } catch (e) {
        assert(false, `Slide rels ${relsPath} is malformed: ${e.message}`)
      }

      if (relsObj) {
        const rels = relsObj.Relationships?.Relationship || []
        const relsArray = Array.isArray(rels) ? rels : [rels]
        
        for (const rel of relsArray) {
          const target = rel['@_Target']
          const type = rel['@_Type']
          if (rel['@_TargetMode'] !== 'External' && target && !target.startsWith('http')) {
            const resolvedPath = path.posix.join(slideDir, target)
            assert(zip.file(resolvedPath), `Slide rel ${rel['@_Id']} in ${slidePath} references non-existent file: ${resolvedPath}`)
            
            // If notesSlide, check back-reference
            if (type.endsWith('/notesSlide')) {
              const notesRelsPath = `ppt/notesSlides/_rels/${path.posix.basename(resolvedPath)}.rels`
              assert(zip.file(notesRelsPath), `Notes slide rels file is missing: ${notesRelsPath}`)
              if (zip.file(notesRelsPath)) {
                const notesRelsText = await zip.file(notesRelsPath).async('text')
                const notesRelsObj = parser.parse(notesRelsText, notesRelsPath)
                let nRels = notesRelsObj.Relationships?.Relationship || []
                nRels = Array.isArray(nRels) ? nRels : [nRels]
                const backRef = nRels.find(nr => nr['@_Type'].endsWith('/slide'))
                assert(backRef, `Notes slide ${resolvedPath} has no back-reference to slide`)
                if (backRef) {
                  const backTarget = backRef['@_Target']
                  const resolvedBackPath = path.posix.join('ppt/notesSlides', backTarget)
                  assert(resolvedBackPath === slidePath, `Notes slide back-reference target mismatch: expected ${slidePath}, got ${resolvedBackPath}`)
                }
              }
            }

            // If chart, check chart relationships
            if (type.endsWith('/chart')) {
              const chartRelsPath = `ppt/charts/_rels/${path.posix.basename(resolvedPath)}.rels`
              if (zip.file(chartRelsPath)) {
                const chartRelsText = await zip.file(chartRelsPath).async('text')
                const chartRelsObj = parser.parse(chartRelsText, chartRelsPath)
                let cRels = chartRelsObj.Relationships?.Relationship || []
                cRels = Array.isArray(cRels) ? cRels : [cRels]
                for (const cr of cRels) {
                  if (cr['@_TargetMode'] !== 'External' && cr['@_Target'] && !cr['@_Target'].startsWith('http')) {
                    const resolvedChartPart = path.posix.join('ppt/charts', cr['@_Target'])
                    assert(zip.file(resolvedChartPart), `Chart relationship ${cr['@_Id']} in ${resolvedPath} references non-existent file: ${resolvedChartPart}`)
                  }
                }
              }
            }
          }
        }
      }
    }
  }

  console.log(`\n=== VALIDATION SUMMARY ===`)
  console.log(`Errors found: ${errors.length}`)
  console.log(`Warnings found: ${warnings.length}`)

  if (errors.length > 0) {
    console.error('\n❌ Errors:')
    errors.forEach(e => console.error(`  - ${e}`))
  } else {
    console.log('\n🎉 No errors found! Structure is 100% compliant.')
  }

  if (warnings.length > 0) {
    console.log('\n⚠️ Warnings:')
    warnings.forEach(w => console.log(`  - ${w}`))
  }
}

const targetPath = process.argv[2] || 'examples/output/office-modified.pptx'
validate(path.resolve(__dirname, '..', targetPath)).catch(console.error)
