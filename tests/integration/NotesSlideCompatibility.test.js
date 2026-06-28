import { describe, it, expect, beforeAll } from 'vitest'
import { existsSync } from 'fs'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import fsExtra from 'fs-extra'
import JSZip from 'jszip'
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

describe('PPTXTemplater - Notes Slide Compatibility', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')
  const runFixtureTests = existsSync(fixtureFile)

  if (!runFixtureTests) {
    it.skip('Fixture file not found — skipping notes slide compatibility tests.', () => {})
    return
  }

  it('should duplicate slides with notes and correctly normalize and link them', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)

    // Programmatically add a notes slide for slide 1 to simulate a presentation with notes
    const sourceSlideZipPath = 'ppt/slides/slide1.xml'
    const notesSlideZipPath = 'ppt/notesSlides/notesSlide1.xml'
    const notesContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr>
        <p:cNvPr id="1" name=""/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr/>
    </p:spTree>
  </p:cSld>
</p:notes>`

    // 1. Add notes file to ZIP
    ppt.zipManager.writeFile(notesSlideZipPath, notesContent)

    // 2. Add override content type
    ppt.contentTypesManager.addOverride(notesSlideZipPath, 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml')

    // 3. Add relationships
    // Rel from slide1 to notesSlide1
    ppt.relationshipManager.addRelationship(
      sourceSlideZipPath,
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
      '../notesSlides/notesSlide1.xml'
    )
    // Rel from notesSlide1 back to slide1
    ppt.relationshipManager.addRelationship(
      notesSlideZipPath,
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
      '../slides/slide1.xml'
    )

    // Force preload notes slide cache for the newly added file
    await ppt.zipManager.readFile(notesSlideZipPath)
    await ppt.slideManager.preloadAll() // update SlideManager internal tracking

    // Verify that the slide has notes now
    const slide1Info = ppt.slideManager.getSlideInfo(1)
    const slide1Rels = ppt.relationshipManager.getRelationships(slide1Info.zipPath)
    expect(slide1Rels.find(r => r.type.endsWith('/notesSlide'))).toBeDefined()

    // Now duplicate slide 1 to slot 2
    await ppt.duplicateSlide(1, 2)

    // Save to buffer
    const buffer = await ppt.toBuffer()
    const integrityErrors = await validatePackageIntegrity(buffer)
    expect(integrityErrors).toEqual([])

    // Verify JSZip structure
    const zip = await JSZip.loadAsync(buffer)

    // Check that ppt/notesSlides/notesSlide2.xml was created
    expect(zip.file('ppt/notesSlides/notesSlide2.xml')).not.toBeNull()

    // Check relationships of slide 2
    const slide2RelsContent = await zip.file('ppt/slides/_rels/slide2.xml.rels').async('text')
    expect(slide2RelsContent).toContain('notesSlides/notesSlide2.xml')

    // Check relationships of notesSlide2
    const notes2RelsContent = await zip.file('ppt/notesSlides/_rels/notesSlide2.xml.rels').async('text')
    expect(notes2RelsContent).toContain('slides/slide2.xml')
  })

  it('should clean up notes slides and their relationships when a slide is removed', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)

    // Add notes slide for slide 1
    const sourceSlideZipPath = 'ppt/slides/slide1.xml'
    const notesSlideZipPath = 'ppt/notesSlides/notesSlide1.xml'
    const notesContent = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld/></p:notes>`

    ppt.zipManager.writeFile(notesSlideZipPath, notesContent)
    ppt.contentTypesManager.addOverride(notesSlideZipPath, 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml')
    ppt.relationshipManager.addRelationship(
      sourceSlideZipPath,
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
      '../notesSlides/notesSlide1.xml'
    )
    ppt.relationshipManager.addRelationship(
      notesSlideZipPath,
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
      '../slides/slide1.xml'
    )

    await ppt.zipManager.readFile(notesSlideZipPath)
    await ppt.slideManager.preloadAll()

    // Now remove slide 1
    ppt.removeSlide(1)

    const buffer = await ppt.toBuffer()
    const integrityErrors = await validatePackageIntegrity(buffer)
    expect(integrityErrors).toEqual([])

    const zip = await JSZip.loadAsync(buffer)
    // The notesSlide1.xml and its relationships should be gone
    expect(zip.file('ppt/notesSlides/notesSlide1.xml')).toBeNull()
    expect(zip.file('ppt/notesSlides/_rels/notesSlide1.xml.rels')).toBeNull()
  })
})
