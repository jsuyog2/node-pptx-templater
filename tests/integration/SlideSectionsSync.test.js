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

describe('PPTXTemplater - Slide Sections Sync', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')
  const runFixtureTests = existsSync(fixtureFile)

  if (!runFixtureTests) {
    it.skip('Fixture file not found — skipping slide sections test.', () => {})
    return
  }

  it('should synchronize slide ordering in sections after duplicating slide into different positions', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    expect(ppt.slideCount).toBe(3)

    // Duplicate slide 2 to position 1
    ppt.duplicateSlide(2, 1)
    expect(ppt.slideCount).toBe(4)

    const outPath = resolve(OUTPUT_DIR, 'slide-sections-sync-test.pptx')
    await ppt.saveToFile(outPath)

    // Load file to verify structure
    const zip = await JSZip.loadAsync(await fsExtra.readFile(outPath))
    const xmlParser = new XMLParser()

    const presentationXml = await zip.file('ppt/presentation.xml').async('text')
    const presObj = xmlParser.parse(presentationXml)

    // Get ordered slide IDs from main list
    const mainSldIds = presObj['p:presentation']['p:sldIdLst']['p:sldId'].map(s =>
      String(s['@_id'])
    )

    // Get ordered slide IDs from all sections combined
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

    expect(sectionSldIds.length).toBe(mainSldIds.length)
    expect(sectionSldIds).toEqual(mainSldIds)
  })
})
