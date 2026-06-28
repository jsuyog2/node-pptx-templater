import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import { existsSync } from 'fs'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import fsExtra from 'fs-extra'
import JSZip from 'jszip'
import { XMLParser } from '../../src/parsers/XMLParser.js'
import { validatePackageIntegrity } from '../helpers/packageIntegrity.js'

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
    const initialCount = ppt.slideCount
    expect(initialCount).toBeGreaterThanOrEqual(3)

    const removedSlideId = ppt.slideManager.getSlideInfo(2).slideId

    // Remove slide 2 (leaves a gap in filenames and slide IDs in memory)
    ppt.removeSlide(2)
    expect(ppt.slideCount).toBe(initialCount - 1)

    // Add a slide link from slide 2 (formerly slide 3) to slide 1
    ppt.addSlideLink({
      element: 'Hello {{title}}',
      sourceSlide: 2,
      targetSlide: 1,
    })

    // Duplicate slide 1 to position 3
    await ppt.duplicateSlide(1, 3)
    expect(ppt.slideCount).toBe(initialCount)

    const outPath = resolve(OUTPUT_DIR, 'structural-normalization-test.pptx')
    await ppt.saveToFile(outPath)

    const buffer = await fsExtra.readFile(outPath)
    const integrityErrors = await validatePackageIntegrity(buffer)
    expect(integrityErrors).toEqual([])

    const zip = await JSZip.loadAsync(buffer)
    const xmlParser = new XMLParser()

    const presentationXml = await zip.file('ppt/presentation.xml').async('text')
    const presObj = xmlParser.parse(presentationXml)

    const sldIds = presObj['p:presentation']['p:sldIdLst']['p:sldId']
    const mainSldIds = (Array.isArray(sldIds) ? sldIds : [sldIds]).map(s => String(s['@_id']))
    expect(mainSldIds).not.toContain(String(removedSlideId))
    expect(mainSldIds).toHaveLength(initialCount)

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
            const sectionIds = section['p14:sldIdLst']?.['p14:sldId'] || []
            const ids = (Array.isArray(sectionIds) ? sectionIds : [sectionIds]).map(s =>
              String(s['@_id'])
            )
            sectionSldIds.push(...ids)
          }
        }
      }
    }
    expect(sectionSldIds).toEqual(mainSldIds)

    const files = Object.keys(zip.files)
    const slideFiles = files.filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f)).sort()
    const expectedSlideFiles = Array.from(
      { length: initialCount },
      (_, i) => `ppt/slides/slide${i + 1}.xml`
    )
    expect(slideFiles).toEqual(expectedSlideFiles)

    const relsFiles = files.filter(f => /^ppt\/slides\/_rels\/slide\d+\.xml\.rels$/.test(f)).sort()
    const expectedRelsFiles = expectedSlideFiles.map(
      f => f.replace('ppt/slides/', 'ppt/slides/_rels/') + '.rels'
    )
    expect(relsFiles).toEqual(expectedRelsFiles)

    const presRelsXml = await zip.file('ppt/_rels/presentation.xml.rels').async('text')
    const presRelsObj = xmlParser.parse(presRelsXml)
    const presRels = Array.isArray(presRelsObj.Relationships.Relationship)
      ? presRelsObj.Relationships.Relationship
      : [presRelsObj.Relationships.Relationship]

    const rIds = presRels.map(r => r['@_Id'])
    const uniqueRIds = new Set(rIds)
    expect(uniqueRIds.size).toBe(rIds.length)

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
