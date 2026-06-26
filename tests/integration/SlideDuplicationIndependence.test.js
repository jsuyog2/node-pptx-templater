import { describe, it, expect, beforeAll, afterAll } from 'vitest'
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

describe('PPTXTemplater - Slide duplication independence', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')
  const runFixtureTests = existsSync(fixtureFile)

  if (!runFixtureTests) {
    it.skip('Fixture file not found — skipping duplication independence tests.', () => {})
    return
  }

  it('should not modify the source slide XML or relationships when duplicating', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    const sourceIndex = 1
    const sourceInfo = ppt.slideManager.getSlideInfo(sourceIndex)
    const beforeXml = ppt.slideManager.getSlideXml(sourceIndex)
    const beforeRels = JSON.stringify(
      ppt.relationshipManager.getRelationships(sourceInfo.zipPath)
    )

    await ppt.duplicateSlide(sourceIndex, 2)

    expect(ppt.slideManager.getSlideXml(sourceIndex)).toBe(beforeXml)
    expect(
      JSON.stringify(ppt.relationshipManager.getRelationships(sourceInfo.zipPath))
    ).toBe(beforeRels)

    const buffer = await ppt.toBuffer()
    const errors = await validatePackageIntegrity(buffer)
    expect(errors).toEqual([])
  })

  it('should deep-clone chart parts instead of sharing them with the source slide', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    const sourceIndex = 1
    const sourceInfo = ppt.slideManager.getSlideInfo(sourceIndex)
    const sourceChartRel = ppt.relationshipManager
      .getRelationships(sourceInfo.zipPath)
      .find(r => r.type.endsWith('/chart'))

    await ppt.duplicateSlide(sourceIndex)

    const cloneIndex = ppt.slideCount
    const cloneInfo = ppt.slideManager.getSlideInfo(cloneIndex)
    const cloneChartRel = ppt.relationshipManager
      .getRelationships(cloneInfo.zipPath)
      .find(r => r.type.endsWith('/chart'))

    expect(sourceChartRel).toBeDefined()
    expect(cloneChartRel).toBeDefined()
    expect(cloneChartRel.target).not.toBe(sourceChartRel.target)

    const sourceChartPath = ppt.relationshipManager.resolveTarget(
      sourceInfo.zipPath,
      sourceChartRel.target
    )
    const cloneChartPath = ppt.relationshipManager.resolveTarget(
      cloneInfo.zipPath,
      cloneChartRel.target
    )
    expect(cloneChartPath).not.toBe(sourceChartPath)
    expect(ppt.zipManager.hasFile(cloneChartPath)).toBe(true)
    expect(ppt.zipManager.hasFile(sourceChartPath)).toBe(true)
  })

  it('should deep-clone image media instead of sharing relationships with the source slide', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    const sourceIndex = 6
    const sourceInfo = ppt.slideManager.getSlideInfo(sourceIndex)
    const sourceImageRels = ppt.relationshipManager
      .getRelationships(sourceInfo.zipPath)
      .filter(r => r.type.endsWith('/image'))

    expect(sourceImageRels.length).toBeGreaterThan(0)

    await ppt.duplicateSlide(sourceIndex)

    const cloneInfo = ppt.slideManager.getSlideInfo(ppt.slideCount)
    const cloneImageRels = ppt.relationshipManager
      .getRelationships(cloneInfo.zipPath)
      .filter(r => r.type.endsWith('/image'))

    expect(cloneImageRels.length).toBe(sourceImageRels.length)
    for (const cloneRel of cloneImageRels) {
      expect(sourceImageRels.some(r => r.target === cloneRel.target)).toBe(false)
      const cloneAbs = ppt.relationshipManager.resolveTarget(cloneInfo.zipPath, cloneRel.target)
      for (const sourceRel of sourceImageRels) {
        const sourceAbs = ppt.relationshipManager.resolveTarget(sourceInfo.zipPath, sourceRel.target)
        expect(cloneAbs).not.toBe(sourceAbs)
      }
    }
  })

  it('should preserve the source slide byte-for-byte in the saved package after duplication', async () => {
    const baseline = await PPTXTemplater.load(fixtureFile)
    const sourceIndex = 4
    const sourceXmlBefore = baseline.slideManager.getSlideXml(sourceIndex)

    const ppt = await PPTXTemplater.load(fixtureFile)
    await ppt.duplicateSlide(sourceIndex, sourceIndex + 1)
    const buffer = await ppt.toBuffer()

    const zip = await JSZip.loadAsync(buffer)
    const normalizedSourcePath = `ppt/slides/slide${sourceIndex}.xml`
    const savedSourceXml = await zip.file(normalizedSourcePath).async('text')

    expect(savedSourceXml).toBe(sourceXmlBefore)

    const errors = await validatePackageIntegrity(buffer)
    expect(errors).toEqual([])
  })

  it('should keep package integrity when duplicating a slide after chart edits', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    ppt.useSlide(1)
    ppt.updateChart('Chart', {
      categories: ['A', 'B'],
      series: [{ name: 'Series 1', values: [10, 20] }],
    })

    const sourceXmlBefore = ppt.slideManager.getSlideXml(1)
    await ppt.duplicateSlide(1, 2)

    expect(ppt.slideManager.getSlideXml(1)).toBe(sourceXmlBefore)

    const errors = await validatePackageIntegrity(await ppt.toBuffer())
    expect(errors).toEqual([])
  })
})
