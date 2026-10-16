import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'
import { validatePackageIntegrity } from '../helpers/packageIntegrity.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Slide Operations: Selection, Creation, Ordering & Transformation', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_slide_operations.pptx')

  beforeEach(() => {
    if (existsSync(outputPath)) {
      try {
        unlinkSync(outputPath)
      } catch (_) {}
    }
  })

  afterEach(() => {
    if (existsSync(outputPath)) {
      try {
        unlinkSync(outputPath)
      } catch (_) {}
    }
  })

  it('should support fluent slide selection via useSlide and useAllSlides', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Select specific slide
    expect(ppt.useSlide(1)).toBe(ppt)

    // Select multiple slides
    expect(ppt.useSlide(1, 2)).toBe(ppt)

    // Select all slides
    expect(ppt.useAllSlides()).toBe(ppt)
  })

  it('should add slides with options and position', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const initialCount = ppt.slideCount

    // Append to end
    ppt.addSlide({ title: 'Appended Slide' })
    expect(ppt.slideCount).toBe(initialCount + 1)
    expect(ppt.getSlide(ppt.slideCount).title).toBe('Appended Slide')

    // Insert at specific position (slide 2)
    ppt.addSlide({ title: 'Position 2 Slide' }, 2)
    expect(ppt.slideCount).toBe(initialCount + 2)
    expect(ppt.getSlide(2).title).toBe('Position 2 Slide')

    const buf = await ppt.toBuffer()
    const errors = await validatePackageIntegrity(buf)
    expect(errors).toEqual([])
  })

  it('should duplicate slides and insert at specified positions', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const initialCount = ppt.slideCount

    // Duplicate slide 1 to end
    await ppt.duplicateSlide(1)
    expect(ppt.slideCount).toBe(initialCount + 1)

    // Duplicate slide 1 to position 2
    await ppt.duplicateSlide(1, 2)
    expect(ppt.slideCount).toBe(initialCount + 2)

    const buf = await ppt.toBuffer()
    const errors = await validatePackageIntegrity(buf)
    expect(errors).toEqual([])
  })

  it('should remove slides and maintain valid numbering', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const initialCount = ppt.slideCount

    ppt.removeSlide(1)
    expect(ppt.slideCount).toBe(initialCount - 1)

    const buf = await ppt.toBuffer()
    const errors = await validatePackageIntegrity(buf)
    expect(errors).toEqual([])
  })

  it('should move slides within the presentation', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const initialFirst = ppt.getSlide(1).slideId

    ppt.moveSlide(1, 2)
    expect(ppt.getSlide(2).slideId).toBe(initialFirst)

    const buf = await ppt.toBuffer()
    const errors = await validatePackageIntegrity(buf)
    expect(errors).toEqual([])
  })

  it('should reorder slides according to array', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const count = ppt.slideCount
    if (count >= 2) {
      const order = []
      for (let i = count; i >= 1; i--) {
        order.push(i)
      }
      ppt.reorderSlides(order)

      const buf = await ppt.toBuffer()
      const errors = await validatePackageIntegrity(buf)
      expect(errors).toEqual([])
    }
  })

  it('should tag and resolve slides by tag', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    ppt.tagSlide(1, 'target-tag')

    const slide = ppt.getSlide(1)
    expect(slide.tags).toContain('target-tag')

    // Tag resolution in useSlide
    ppt.useSlide('target-tag')
  })

  it('should import slides from another PPTXTemplater instance', async () => {
    const sourcePpt = await PPTXTemplater.load(samplePath)
    const destPpt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })

    const destCount = destPpt.slideCount
    await destPpt.importSlideFrom(sourcePpt, 1)
    expect(destPpt.slideCount).toBe(destCount + 1)

    const buf = await destPpt.toBuffer()
    const errors = await validatePackageIntegrity(buf)
    expect(errors).toEqual([])
  })

  it('should export selected slides into a new deck', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const exportedDeck = await ppt.exportSlides(1)
    expect(exportedDeck).toBeInstanceOf(PPTXTemplater)
    expect(exportedDeck.slideCount).toBe(1)

    const buf = await exportedDeck.toBuffer()
    const errors = await validatePackageIntegrity(buf)
    expect(errors).toEqual([])
  })
})
