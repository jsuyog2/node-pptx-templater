import { describe, it, expect } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { PPTXTemplater } from '../../src/index.js'
import { validatePackageIntegrity } from '../helpers/packageIntegrity.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Section Management Methods (<p14:sectionLst>)', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')

  it('should add, get, update, and use sections', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Add first section
    ppt.addSection('Intro Section', 1)
    const sec1 = ppt.getSection('Intro Section')
    expect(sec1).not.toBeNull()
    expect(sec1.name).toBe('Intro Section')
    expect(sec1.id).toBeDefined()

    // Add second section
    if (ppt.slideCount >= 2) {
      ppt.addSection('Deep Dive Section', 2)
      const sec2 = ppt.getSection('Deep Dive Section')
      expect(sec2).not.toBeNull()
      expect(sec2.name).toBe('Deep Dive Section')
    }

    // Update section name
    ppt.updateSection('Intro Section', { name: 'Executive Intro' })
    const updated = ppt.getSection('Executive Intro')
    expect(updated).not.toBeNull()
    expect(updated.name).toBe('Executive Intro')

    // useSection should target member slides
    expect(ppt.useSection('Executive Intro')).toBe(ppt)

    // Verify package integrity
    const buf = await ppt.toBuffer()
    const errors = await validatePackageIntegrity(buf)
    expect(errors).toEqual([])
  })

  it('should remove sections without removing slides', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    ppt.addSection('Removable Section', 1)
    const countBefore = ppt.slideCount

    ppt.removeSection('Removable Section', false)
    expect(ppt.getSection('Removable Section')).toBeNull()
    expect(ppt.slideCount).toBe(countBefore)

    const buf = await ppt.toBuffer()
    const errors = await validatePackageIntegrity(buf)
    expect(errors).toEqual([])
  })

  it('should remove sections and their member slides when removeSlides is true', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    // Duplicate slide to guarantee enough slides
    await ppt.duplicateSlide(1)
    await ppt.duplicateSlide(1)

    ppt.addSection('Purge Section', ppt.slideCount)
    const countBefore = ppt.slideCount

    ppt.removeSection('Purge Section', true)
    expect(ppt.getSection('Purge Section')).toBeNull()
    expect(ppt.slideCount).toBeLessThan(countBefore)

    const buf = await ppt.toBuffer()
    const errors = await validatePackageIntegrity(buf)
    expect(errors).toEqual([])
  })
})
