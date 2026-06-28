import { describe, it, expect, beforeAll } from 'vitest'
import { existsSync } from 'fs'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import fsExtra from 'fs-extra'
import JSZip from 'jszip'
import { XMLParser } from '../../src/parsers/XMLParser.js'
import { validatePackageIntegrity } from '../helpers/packageIntegrity.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURES_DIR = resolve(__dirname, '../fixtures')
const TEMPLATE = resolve(__dirname, '../../templates/sample.pptx')

let PPTXTemplater

beforeAll(async () => {
  const module = await import('../../src/index.js')
  PPTXTemplater = module.PPTXTemplater
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')
  if (!existsSync(fixtureFile) && existsSync(TEMPLATE)) {
    await fsExtra.copy(TEMPLATE, fixtureFile)
  }
})

describe('PPTXTemplater - Slide Operations Compatibility', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')
  const runFixtureTests = existsSync(fixtureFile)

  if (!runFixtureTests) {
    it.skip('Fixture file not found — skipping slide operations compatibility tests.', () => {})
    return
  }

  // Helper to validate that slides and sections are in exact sync in the saved package
  async function validatePresentationStructure(buffer) {
    const zip = await JSZip.loadAsync(buffer)
    const xmlParser = new XMLParser()

    const presentationXml = await zip.file('ppt/presentation.xml').async('text')
    const presObj = xmlParser.parse(presentationXml)

    const sldIds = presObj['p:presentation']['p:sldIdLst']?.['p:sldId'] || []
    const sldIdArr = Array.isArray(sldIds) ? sldIds : [sldIds]
    const sldIdValues = sldIdArr.map(s => String(s['@_id']))

    const sectionSlideIds = []
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
            sectionSlideIds.push(...ids)
          }
        }
      }
    }

    // If there are sections defined, their slide ID sequence must match sldIdLst exactly
    if (sectionSlideIds.length > 0) {
      expect(sectionSlideIds).toEqual(sldIdValues)
    }

    return sldIdValues.length
  }

  it('should successfully remove the first slide', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    const originalCount = ppt.getInfo().slideCount

    ppt.removeSlide(1)
    expect(ppt.getInfo().slideCount).toBe(originalCount - 1)

    const buffer = await ppt.toBuffer()
    const integrityErrors = await validatePackageIntegrity(buffer)
    expect(integrityErrors).toEqual([])

    const validatedCount = await validatePresentationStructure(buffer)
    expect(validatedCount).toBe(originalCount - 1)
  })

  it('should successfully remove a middle slide', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    const originalCount = ppt.getInfo().slideCount
    const middleIndex = Math.floor(originalCount / 2)

    ppt.removeSlide(middleIndex)
    expect(ppt.getInfo().slideCount).toBe(originalCount - 1)

    const buffer = await ppt.toBuffer()
    const integrityErrors = await validatePackageIntegrity(buffer)
    expect(integrityErrors).toEqual([])

    const validatedCount = await validatePresentationStructure(buffer)
    expect(validatedCount).toBe(originalCount - 1)
  })

  it('should successfully remove the last slide', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    const originalCount = ppt.getInfo().slideCount

    ppt.removeSlide(originalCount)
    expect(ppt.getInfo().slideCount).toBe(originalCount - 1)

    const buffer = await ppt.toBuffer()
    const integrityErrors = await validatePackageIntegrity(buffer)
    expect(integrityErrors).toEqual([])

    const validatedCount = await validatePresentationStructure(buffer)
    expect(validatedCount).toBe(originalCount - 1)
  })

  it('should successfully duplicate the first slide and insert at various positions', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    const originalCount = ppt.getInfo().slideCount

    // Duplicate first slide to position 2 (after slide 1)
    await ppt.duplicateSlide(1, 2)
    // Duplicate first slide to the end of presentation
    await ppt.duplicateSlide(1)

    expect(ppt.getInfo().slideCount).toBe(originalCount + 2)

    const buffer = await ppt.toBuffer()
    const integrityErrors = await validatePackageIntegrity(buffer)
    expect(integrityErrors).toEqual([])

    const validatedCount = await validatePresentationStructure(buffer)
    expect(validatedCount).toBe(originalCount + 2)
  })

  it('should successfully duplicate a middle slide and insert at various positions', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    const originalCount = ppt.getInfo().slideCount
    const middleIndex = Math.floor(originalCount / 2)

    // Duplicate middle slide to position 1
    await ppt.duplicateSlide(middleIndex, 1)
    // Duplicate middle slide to the end of presentation
    await ppt.duplicateSlide(middleIndex + 1) // index shifted by 1 due to previous insert

    expect(ppt.getInfo().slideCount).toBe(originalCount + 2)

    const buffer = await ppt.toBuffer()
    const integrityErrors = await validatePackageIntegrity(buffer)
    expect(integrityErrors).toEqual([])

    const validatedCount = await validatePresentationStructure(buffer)
    expect(validatedCount).toBe(originalCount + 2)
  })

  it('should successfully duplicate the last slide and insert at various positions', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    const originalCount = ppt.getInfo().slideCount

    // Duplicate last slide to position 1
    await ppt.duplicateSlide(originalCount, 1)
    // Duplicate last slide to middle position
    await ppt.duplicateSlide(originalCount + 1, Math.floor(originalCount / 2))

    expect(ppt.getInfo().slideCount).toBe(originalCount + 2)

    const buffer = await ppt.toBuffer()
    const integrityErrors = await validatePackageIntegrity(buffer)
    expect(integrityErrors).toEqual([])

    const validatedCount = await validatePresentationStructure(buffer)
    expect(validatedCount).toBe(originalCount + 2)
  })
})
