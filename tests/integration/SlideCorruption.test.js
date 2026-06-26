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
const OUTPUT_DIR = resolve(__dirname, '../output-corruption')

let PPTXTemplater

beforeAll(async () => {
  const module = await import('../../src/index.js')
  PPTXTemplater = module.PPTXTemplater
  await fsExtra.ensureDir(OUTPUT_DIR)
})

afterAll(async () => {
  await fsExtra.remove(OUTPUT_DIR)
})

describe('PPTXTemplater - Slide Management Corruption Prevention', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')
  const runFixtureTests = existsSync(fixtureFile)

  if (!runFixtureTests) {
    it.skip('Fixture file not found — skipping slide corruption tests.', () => {})
    return
  }

  it('should remove a slide, keep metadata synchronized, and clean sections without corruption', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    const initialCount = ppt.getInfo().slideCount
    expect(initialCount).toBeGreaterThanOrEqual(2)

    const removedSlideId = ppt.slideManager.getSlideInfo(2).slideId
    const remainingSlideIds = ppt.slideManager
      .getAllSlideInfo()
      .filter(s => s.index !== 2)
      .map(s => s.slideId)

    ppt.removeSlide(2)
    expect(ppt.getInfo().slideCount).toBe(initialCount - 1)

    const outPath = resolve(OUTPUT_DIR, 'slide-corruption-test.pptx')
    await ppt.saveToFile(outPath)

    const buffer = await fsExtra.readFile(outPath)
    const integrityErrors = await validatePackageIntegrity(buffer)
    expect(integrityErrors).toEqual([])

    const zip = await JSZip.loadAsync(buffer)
    const xmlParser = new XMLParser()

    const presentationXml = await zip.file('ppt/presentation.xml').async('text')
    const presObj = xmlParser.parse(presentationXml)

    const sldIds = presObj['p:presentation']['p:sldIdLst']['p:sldId']
    const sldIdArr = Array.isArray(sldIds) ? sldIds : [sldIds]
    const sldIdValues = sldIdArr.map(s => String(s['@_id']))
    expect(sldIdValues).not.toContain(String(removedSlideId))
    expect(sldIdValues).toHaveLength(initialCount - 1)
    for (const id of remainingSlideIds) {
      expect(sldIdValues).toContain(String(id))
    }

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
            const sectionSldIds = section['p14:sldIdLst']?.['p14:sldId'] || []
            const ids = (Array.isArray(sectionSldIds) ? sectionSldIds : [sectionSldIds]).map(s =>
              String(s['@_id'])
            )
            expect(ids).not.toContain(String(removedSlideId))
          }
        }
      }
    }

    const appXml = await zip.file('docProps/app.xml').async('text')
    const appObj = xmlParser.parse(appXml)
    const properties = appObj.Properties

    expect(Number(properties.Slides)).toBe(initialCount - 1)

    const variants = properties.HeadingPairs['vt:vector']['vt:variant']
    let slideTitlesCount = 0
    for (let i = 0; i < variants.length; i++) {
      if (variants[i]['vt:lpstr'] === 'Slide Titles') {
        slideTitlesCount = parseInt(variants[i + 1]['vt:i4'], 10)
        break
      }
    }
    expect(slideTitlesCount).toBe(initialCount - 1)

    const lpstrs = properties.TitlesOfParts['vt:vector']['vt:lpstr']
    const lpstrArray = Array.isArray(lpstrs) ? lpstrs : [lpstrs]
    expect(lpstrArray.length).toBeGreaterThanOrEqual(initialCount - 1)
    const slideTitles = lpstrArray.slice(-(initialCount - 1))
    expect(slideTitles.length).toBe(initialCount - 1)
  })
})
