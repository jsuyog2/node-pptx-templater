import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Slide Dimensions and Layout Inspection', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_dimensions.pptx')

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

  it('should retrieve default slide dimensions and layout metadata', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const dims = ppt.getSlideDimensions()
    expect(dims).toHaveProperty('width')
    expect(dims).toHaveProperty('height')
    expect(typeof dims.width).toBe('number')
    expect(typeof dims.height).toBe('number')
    expect(dims.width).toBeGreaterThan(0)
    expect(dims.height).toBeGreaterThan(0)

    const sizeAlias = ppt.getSlideSize()
    expect(sizeAlias).toEqual(dims)

    const info = ppt.getInfo()
    expect(info.dimensions).toEqual(dims)
    expect(info.slides[0]).toHaveProperty('layout')

    const layout = ppt.getSlideLayout(1)
    expect(typeof layout).toBe('string')
    expect(layout).toMatch(/slideLayout/i)
  })

  it('should set standard aspect ratios (16:9, 16:10, 4:3)', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Set 4:3
    ppt.setAspectRatio('4:3')
    let dims = ppt.getSlideDimensions()
    expect(dims.width).toBe(9144000)
    expect(dims.height).toBe(6858000)

    // Set 16:10
    ppt.setSlideDimensions('16:10')
    dims = ppt.getSlideDimensions()
    expect(dims.width).toBe(9144000)
    expect(dims.height).toBe(5715000)

    // Set 16:9
    ppt.setSlideDimensions('16:9')
    dims = ppt.getSlideDimensions()
    expect(dims.width).toBe(9144000)
    expect(dims.height).toBe(5143500)
  })

  it('should set custom dimensions using string units and numbers', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Custom dimensions in inches
    ppt.setSlideDimensions({ width: '10in', height: '7.5in' })
    let dims = ppt.getSlideDimensions()
    expect(dims.width).toBe(9144000) // 10 * 914400
    expect(dims.height).toBe(6858000) // 7.5 * 914400

    // Custom dimensions with width and height arguments
    ppt.setSlideSize(12192000, 6858000)
    dims = ppt.getSlideDimensions()
    expect(dims.width).toBe(12192000)
    expect(dims.height).toBe(6858000)

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    // Reload saved file and verify persisted dimensions
    const reloaded = await PPTXTemplater.load(outputPath)
    expect(reloaded.getSlideDimensions()).toEqual({ width: 12192000, height: 6858000 })
  })
})
