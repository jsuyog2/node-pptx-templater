import { describe, it, expect } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Slide Inspection, Layout & Geometry Methods', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')

  it('getSlides should return list of all slide metadata objects', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const slides = ppt.getSlides()

    expect(Array.isArray(slides)).toBe(true)
    expect(slides.length).toBe(ppt.slideCount)

    const first = slides[0]
    expect(first).toHaveProperty('index')
    expect(first).toHaveProperty('slideId')
    expect(first).toHaveProperty('relationshipId')
    expect(first).toHaveProperty('zipPath')
    expect(first.index).toBe(1)
  })

  it('getSlide should retrieve normalized metadata for a specific slide', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const slide = ppt.getSlide(1)

    expect(slide).toHaveProperty('index', 1)
    expect(slide).toHaveProperty('slideId')
    expect(slide).toHaveProperty('relationshipId')
    expect(slide).toHaveProperty('zipPath')
    expect(slide).toHaveProperty('layout')
    expect(slide).toHaveProperty('tags')
    expect(slide).toHaveProperty('hidden')
    expect(typeof slide.hidden).toBe('boolean')
  })

  it('getSlideDimensions should return width and height in EMUs', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const dims = ppt.getSlideDimensions()

    expect(dims).toHaveProperty('width')
    expect(dims).toHaveProperty('height')
    expect(dims.width).toBeGreaterThan(0)
    expect(dims.height).toBeGreaterThan(0)
  })

  it('getSlideLayout should return layout target path', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const layout = ppt.getSlideLayout(1)

    expect(typeof layout).toBe('string')
    expect(layout).toMatch(/slideLayout/i)
  })
})
