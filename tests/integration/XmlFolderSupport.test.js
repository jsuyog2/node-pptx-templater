import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import path from 'path'
import fs from 'fs-extra'
import JSZip from 'jszip'
import { fileURLToPath } from 'url'
import { PPTXTemplater, PPTXTemplate } from '../../src/index.js'

const __filename = fileURLToPath(import.meta.url)
const __dirname = path.dirname(__filename)

const FIXTURES_DIR = path.resolve(__dirname, '../fixtures')
const SAMPLE_PPTX = path.join(FIXTURES_DIR, 'sample.pptx')
const TEMP_DIR = path.resolve(__dirname, '../temp-xml-test')

async function extractPptxToFolder(pptxPath, destDir) {
  await fs.ensureDir(destDir)
  const data = await fs.readFile(pptxPath)
  const zip = await JSZip.loadAsync(data)
  for (const [filename, file] of Object.entries(zip.files)) {
    if (file.dir) continue
    const destPath = path.join(destDir, filename)
    await fs.ensureDir(path.dirname(destPath))
    const content = await file.async('nodebuffer')
    await fs.writeFile(destPath, content)
  }
}

describe('PowerPoint XML Folder Support Integration Tests', () => {
  const templateDir = path.join(TEMP_DIR, 'sample-template')
  const outputFolder = path.join(TEMP_DIR, 'output-folder')
  const outputPptx = path.join(TEMP_DIR, 'output.pptx')

  beforeAll(async () => {
    await fs.emptyDir(TEMP_DIR)
    // Extract sample.pptx to a folder structure for testing
    await extractPptxToFolder(SAMPLE_PPTX, templateDir)
  })

  afterAll(async () => {
    await fs.remove(TEMP_DIR)
  })

  it('should successfully load a presentation from a flat folder path', async () => {
    const ppt = await PPTXTemplater.load(templateDir)
    expect(ppt).toBeDefined()
    expect(ppt.slideManager.slideCount).toBeGreaterThan(0)
  })

  it('should successfully load a presentation using options config object', async () => {
    const ppt = await PPTXTemplater.load({
      presentation: path.join(templateDir, 'ppt/presentation.xml'),
      root: templateDir,
    })
    expect(ppt).toBeDefined()
    expect(ppt.slideManager.slideCount).toBeGreaterThan(0)
  })

  it('should successfully load a presentation using fromPresentationXml', async () => {
    const ppt = await PPTXTemplate.fromPresentationXml({
      presentation: path.join(templateDir, 'ppt/presentation.xml'),
      root: templateDir,
    })
    expect(ppt).toBeDefined()
    expect(ppt.slideManager.slideCount).toBeGreaterThan(0)
  })

  it('should support modifications and saving directly back to a folder', async () => {
    const ppt = await PPTXTemplate.fromPresentationXml(templateDir)

    // Select first slide and replace text
    ppt.useSlide(1)
    ppt.replaceText({ '{{title}}': 'Hello XML Folder' })

    // Save modifications to a new folder
    await ppt.saveToFolder(outputFolder)

    // Verify slide1.xml exists and has the replaced text
    const slide1XmlPath = path.join(outputFolder, 'ppt/slides/slide1.xml')
    expect(await fs.pathExists(slide1XmlPath)).toBe(true)

    const slideXml = await fs.readFile(slide1XmlPath, 'utf8')
    expect(slideXml).toContain('Hello XML Folder')
  })

  it('should support saving a folder template to a .pptx ZIP archive', async () => {
    const ppt = await PPTXTemplate.fromPresentationXml(outputFolder)

    // Save as .pptx file
    await ppt.save(outputPptx)
    expect(await fs.pathExists(outputPptx)).toBe(true)

    // Reload from file to verify
    const reloaded = await PPTXTemplater.load(outputPptx)
    expect(reloaded.slideManager.slideCount).toBe(ppt.slideManager.slideCount)

    // Run archive validation
    await reloaded.validateArchive()
  })

  it('should validate XML presentation files and relationships', async () => {
    const ppt = await PPTXTemplate.fromPresentationXml(outputFolder)
    const report = await ppt.validatePresentationXml()
    expect(report.valid).toBe(true)
    expect(report.errors).toHaveLength(0)
  })

  it('should detect missing relationship file in validatePresentationXml', async () => {
    const corruptFolder = path.join(TEMP_DIR, 'corrupt-folder')
    await fs.copy(outputFolder, corruptFolder)

    // Delete a critical slide layout relationship file
    const layoutRelsPath = path.join(corruptFolder, 'ppt/slides/_rels/slide1.xml.rels')
    await fs.remove(layoutRelsPath)

    const ppt = await PPTXTemplate.fromPresentationXml(corruptFolder)
    const report = await ppt.validatePresentationXml()
    expect(report.valid).toBe(false)
    expect(report.errors.some(e => e.includes('relationship'))).toBe(true)
  })
})
