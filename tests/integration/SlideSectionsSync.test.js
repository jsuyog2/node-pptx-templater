import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import { existsSync } from 'fs'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import fsExtra from 'fs-extra'
import JSZip from 'jszip'
import { XMLParser } from '../../src/parsers/XMLParser.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURES_DIR = resolve(__dirname, '../fixtures')
const OUTPUT_DIR = resolve(__dirname, '../output-sections')

let PPTXTemplater

beforeAll(async () => {
  const module = await import('../../src/index.js')
  PPTXTemplater = module.PPTXTemplater
  await fsExtra.ensureDir(OUTPUT_DIR)
})

afterAll(async () => {
  await fsExtra.remove(OUTPUT_DIR)
})

describe('PPTXTemplater - Slide Sections and Package Structural Normalization', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')
  const runFixtureTests = existsSync(fixtureFile)

  if (!runFixtureTests) {
    it.skip('Fixture file not found — skipping structural normalization tests.', () => {})
    return
  }

  it('should synchronize slide ordering in sections and sequentialize filenames & relationship IDs', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    expect(ppt.slideCount).toBe(3)

    // Remove slide 2 (leaves a gap in filenames and slide IDs in memory)
    ppt.removeSlide(2)
    expect(ppt.slideCount).toBe(2)

    // Add a slide link from slide 2 (originally slide 3) to slide 1
    ppt.addSlideLink({
      element: 'Hello {{title}}',
      sourceSlide: 2,
      targetSlide: 1,
    })

    // Duplicate slide 1 to position 3
    ppt.duplicateSlide(1, 3)
    expect(ppt.slideCount).toBe(3)

    const outPath = resolve(OUTPUT_DIR, 'structural-normalization-test.pptx')
    await ppt.saveToFile(outPath)

    // Load file to verify structure
    const zip = await JSZip.loadAsync(await fsExtra.readFile(outPath))
    const xmlParser = new XMLParser()

    const presentationXml = await zip.file('ppt/presentation.xml').async('text')
    const presObj = xmlParser.parse(presentationXml)

    // 1. Verify Slide IDs in presentation.xml are stable
    const sldIds = presObj['p:presentation']['p:sldIdLst']['p:sldId']
    const mainSldIds = (Array.isArray(sldIds) ? sldIds : [sldIds]).map(s => String(s['@_id']))
    expect(mainSldIds).toEqual(['256', '258', '259'])

    // 2. Verify Slide IDs in Sections match exactly and are ordered
    const sectionSldIds = []
    const extLst = presObj['p:presentation']['p:extLst']
    if (extLst?.['p:ext']) {
      const exts = Array.isArray(extLst['p:ext']) ? extLst['p:ext'] : [extLst['p:ext']]
      for (const ext of exts) {
        const sectionLst = ext['p14:sectionLst']
        if (sectionLst?.['p14:section']) {
          const sections = Array.isArray(sectionLst['p14:section'])
            ? sectionLst['p14:section']
            : [sectionLst['p14:section']]
          for (const section of sections) {
            const sldIds = section['p14:sldIdLst']?.['p14:sldId'] || []
            const ids = (Array.isArray(sldIds) ? sldIds : [sldIds]).map(s => String(s['@_id']))
            sectionSldIds.push(...ids)
          }
        }
      }
    }
    expect(sectionSldIds).toEqual(mainSldIds)

    // 3. Verify files in ZIP are named sequentially with NO gaps
    const files = Object.keys(zip.files)
    const slideFiles = files.filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f)).sort()
    expect(slideFiles).toEqual([
      'ppt/slides/slide1.xml',
      'ppt/slides/slide2.xml',
      'ppt/slides/slide3.xml'
    ])

    const relsFiles = files.filter(f => /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(f)).sort()
    expect(relsFiles).toEqual([
      'ppt/slides/_rels/slide1.xml.rels',
      'ppt/slides/_rels/slide2.xml.rels',
      'ppt/slides/_rels/slide3.xml.rels'
    ])

    // 4. Verify no gaps in relationship IDs in presentation.xml.rels
    const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('text')
    const presRelsObj = xmlParser.parse(presRelsXml)
    const presRels = Array.isArray(presRelsObj.Relationships.Relationship)
      ? presRelsObj.Relationships.Relationship
      : [presRelsObj.Relationships.Relationship]
    
    const rIds = presRels.map(r => r['@_Id'])
    const sortedNumIds = rIds.map(id => parseInt(id.replace('rId', ''), 10)).sort((a, b) => a - b)
    
    // Check that we have rId1, rId2, rId3, ..., rIdMax with no gaps
    const maxId = sortedNumIds[sortedNumIds.length - 1]
    const expectedNumIds = Array.from({ length: maxId }, (_, i) => i + 1)
    expect(sortedNumIds).toEqual(expectedNumIds)

    // 5. Verify slide-to-slide hyperlink target was updated to point to slide1.xml
    // Slide 2 is the source slide that contains the slide link to Slide 1
    const slide2RelsXml = await zip.file('ppt/slides/_rels/slide2.xml.rels').async('text')
    const slide2RelsObj = xmlParser.parse(slide2RelsXml)
    const slide2Rels = Array.isArray(slide2RelsObj.Relationships.Relationship)
      ? slide2RelsObj.Relationships.Relationship
      : [slide2RelsObj.Relationships.Relationship]
    
    const slideLinkRel = slide2Rels.find(r => r['@_Type'].endsWith('/slide'))
    expect(slideLinkRel).toBeDefined()
    expect(slideLinkRel['@_Target']).toBe('slide1.xml')
  })
})
