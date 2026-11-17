import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'
import JSZip from 'jszip'
import fs from 'fs-extra'
import path from 'path'
import { XMLParser } from '../../src/parsers/XMLParser.js'

const TEMP_DIR = path.resolve(__dirname, '../temp-blank-export')
const parser = new XMLParser()

describe('PPTXTemplater.createBlank() OpenXML Export & Validation', () => {
  beforeEach(async () => {
    await fs.remove(TEMP_DIR)
    await fs.ensureDir(TEMP_DIR)
  })

  afterEach(async () => {
    await fs.remove(TEMP_DIR)
  })

  it('should create a 16:9 blank deck and validate cleanly without errors or warnings', async () => {
    const ppt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })
    expect(ppt.slideCount).toBe(1)

    const val = await ppt.validatePresentation()
    expect(val).toEqual({ valid: true, errors: [], warnings: [] })

    await expect(ppt.validateArchive()).resolves.toBe(ppt)
  })

  it('should add shapes and text, save to file, and verify valid XML in all ZIP parts', async () => {
    const ppt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })
    ppt.useSlide(1)

    await ppt.addShape('roundedRectangle', {
      id: 'headerBox',
      x: 100,
      y: 100,
      width: 400,
      height: 150,
      fill: '#1E293B',
      text: 'Verified OpenXML Clean Export',
    })

    const valBefore = await ppt.validatePresentation()
    expect(valBefore.valid).toBe(true)
    expect(valBefore.errors).toEqual([])

    const outPath = path.join(TEMP_DIR, 'blank-test.pptx')
    await ppt.saveToFile(outPath)
    expect(await fs.pathExists(outPath)).toBe(true)

    // Inspect the exported ZIP package
    const fileBytes = await fs.readFile(outPath)
    const zip = await JSZip.loadAsync(fileBytes)

    // 1. Content types inspection
    const contentTypesText = await zip.file('[Content_Types].xml').async('text')
    expect(contentTypesText).toContain('Extension="rels"')
    expect(contentTypesText).toContain('Extension="xml"')
    expect(contentTypesText).toContain('PartName="/ppt/presentation.xml"')
    expect(contentTypesText).toContain('PartName="/ppt/slideMasters/slideMaster1.xml"')
    expect(contentTypesText).toContain('PartName="/ppt/slideLayouts/slideLayout1.xml"')
    expect(contentTypesText).toContain('PartName="/ppt/theme/theme1.xml"')
    expect(contentTypesText).toContain('PartName="/ppt/slides/slide1.xml"')

    // 2. Root rels
    const rootRelsText = await zip.file('_rels/.rels').async('text')
    expect(rootRelsText).toContain('Target="ppt/presentation.xml"')
    expect(rootRelsText).toContain('relationships/officeDocument')

    // 3. Presentation rels
    const presRelsText = await zip.file('ppt/_rels/presentation.xml.rels').async('text')
    expect(presRelsText).toContain('Target="slideMasters/slideMaster1.xml"')
    expect(presRelsText).toContain('Target="slides/slide1.xml"')
    expect(presRelsText).toContain('Target="theme/theme1.xml"')

    // 4. Slide Master & Layout hierarchy
    const masterText = await zip.file('ppt/slideMasters/slideMaster1.xml').async('text')
    expect(masterText).toContain('<p:sldLayoutIdLst>')
    expect(masterText).toContain('r:id="rId1"')

    const masterRelsText = await zip
      .file('ppt/slideMasters/_rels/slideMaster1.xml.rels')
      .async('text')
    expect(masterRelsText).toContain('Target="../slideLayouts/slideLayout1.xml"')

    const layoutRelsText = await zip
      .file('ppt/slideLayouts/_rels/slideLayout1.xml.rels')
      .async('text')
    expect(layoutRelsText).toContain('Target="../slideMasters/slideMaster1.xml"')

    const slideRelsText = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')
    expect(slideRelsText).toContain('Target="../slideLayouts/slideLayout1.xml"')

    // 5. Theme validation: compound line must be sng, not s
    const themeText = await zip.file('ppt/theme/theme1.xml').async('text')
    expect(themeText).not.toContain('cmpd="s"')
    expect(themeText).toContain('cmpd="sng"')
    expect(themeText).toContain('<a:objectDefaults/>')
    expect(themeText).toContain('<a:extraClrSchemeLst/>')

    // 6. Ensure every XML file in the archive is syntactically valid XML
    for (const [name, entry] of Object.entries(zip.files)) {
      if (!entry.dir && (name.endsWith('.xml') || name.endsWith('.rels'))) {
        const text = await entry.async('text')
        expect(() => parser.parse(text)).not.toThrow()
      }
    }

    // 7. Re-load the saved presentation to ensure seamless round-trip
    const reloaded = await PPTXTemplater.load(outPath)
    expect(reloaded.slideCount).toBe(1)
    const valReloaded = await reloaded.validatePresentation()
    expect(valReloaded.valid).toBe(true)
  })

  it('should support adding multiple slides and maintaining unique collision-free relationship IDs', async () => {
    const ppt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })
    ppt.useSlide(1)
    await ppt.addShape('rectangle', { x: 50, y: 50, width: 200, height: 100, text: 'Slide 1' })

    ppt.addSlide({ title: 'Second Slide' })
    expect(ppt.slideCount).toBe(2)

    ppt.useSlide(2)
    await ppt.addShape('ellipse', { x: 100, y: 100, width: 200, height: 200, text: 'Slide 2' })

    const outPath = path.join(TEMP_DIR, 'two-slides-blank.pptx')
    await ppt.saveToFile(outPath)

    const reloaded = await PPTXTemplater.load(outPath)
    expect(reloaded.slideCount).toBe(2)

    const val = await reloaded.validatePresentation()
    expect(val.valid).toBe(true)
    expect(val.errors).toEqual([])
  })
})
