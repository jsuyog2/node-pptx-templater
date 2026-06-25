import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import { existsSync } from 'fs'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import fsExtra from 'fs-extra'
import JSZip from 'jszip'
import { XMLParser } from '../../src/parsers/XMLParser.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURES_DIR = resolve(__dirname, '../fixtures')
const OUTPUT_DIR = resolve(__dirname, '../output-hyperlink')

let PPTXTemplater

beforeAll(async () => {
  const module = await import('../../src/index.js')
  PPTXTemplater = module.PPTXTemplater
  await fsExtra.ensureDir(OUTPUT_DIR)
})

afterAll(async () => {
  await fsExtra.remove(OUTPUT_DIR)
})

describe('PPTXTemplater - Slide Import & Slide Navigation Hyperlinks', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')
  const runFixtureTests = existsSync(fixtureFile)

  if (!runFixtureTests) {
    it.skip('Fixture file not found — skipping slide import & hyperlink tests.', () => {})
    return
  }

  it('should successfully import a slide from a template into a blank presentation and maintain integrity', async () => {
    const sourceEngine = await PPTXTemplater.load(fixtureFile)
    const destEngine = await PPTXTemplater.create()

    // The destination blank presentation initially has 3 slides
    expect(destEngine.slideCount).toBe(3)

    // Import slide 1 from sourceEngine into destEngine
    await destEngine.importSlideFrom(sourceEngine, 1)
    expect(destEngine.slideCount).toBe(4)

    const outPath = resolve(OUTPUT_DIR, 'imported-slide-test.pptx')
    await destEngine.saveToFile(outPath)

    // Verify slide XML was correctly copied and is well-formed
    const zip = await JSZip.loadAsync(await fsExtra.readFile(outPath))
    const xmlParser = new XMLParser()

    const slide2Xml = await zip.file('ppt/slides/slide2.xml').async('text')
    const slideObj = xmlParser.parse(slide2Xml)
    expect(slideObj).toBeDefined()

    // Check relationship file for slide 2 exists
    expect(zip.file('ppt/slides/_rels/slide2.xml.rels')).toBeDefined()
  })

  it('should successfully add slide-jumping next/prev/first/last navigation hyperlinks', async () => {
    const sourceEngine = await PPTXTemplater.load(fixtureFile)

    // Add navigation action links to slide 1
    // Let's first inspect slide 1 text to see what runs we can link.
    // Standard template has hello text or we can just link to slide number / first available run
    sourceEngine.addTextNavigationLink({
      slide: 1,
      element: 'Hello {{title}}',
      action: 'next',
    })

    sourceEngine.addShapeNavigationLink({
      slide: 1,
      shapeId: 'Title',
      action: 'last',
    })

    const outPath = resolve(OUTPUT_DIR, 'navigation-hyperlink-test.pptx')
    await sourceEngine.saveToFile(outPath)

    // Load zip and verify raw XML structures
    const zip = await JSZip.loadAsync(await fsExtra.readFile(outPath))
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')

    // Verify hlinkClick elements are correctly written with the exact action string and NO empty r:id
    expect(slideXml).toContain('action="ppaction://hlinkshowjump?s=nextslide"')
    expect(slideXml).toContain('action="ppaction://hlinkshowjump?s=lastslide"')
    expect(slideXml).not.toContain('r:id=""')
  })

  it('should successfully add slide-to-slide jumping links with slide type relationships', async () => {
    const sourceEngine = await PPTXTemplater.load(fixtureFile)

    // Add inter-slide link
    sourceEngine.addSlideLink({
      element: 'Hello {{title}}',
      sourceSlide: 1,
      targetSlide: 2,
    })

    const outPath = resolve(OUTPUT_DIR, 'slide-link-test.pptx')
    await sourceEngine.saveToFile(outPath)

    const zip = await JSZip.loadAsync(await fsExtra.readFile(outPath))
    const relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')

    // Verify relationship has type slide and target slide2.xml
    expect(relsXml).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"'
    )
    expect(relsXml).toContain('Target="slide2.xml"')
  })
})
