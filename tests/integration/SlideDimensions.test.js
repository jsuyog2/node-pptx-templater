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

    const info = ppt.getInfo()
    expect(info.dimensions).toEqual(dims)
    expect(info.slides[0]).toHaveProperty('layout')

    const layout = ppt.getSlideLayout(1)
    expect(typeof layout).toBe('string')
    expect(layout).toMatch(/slideLayout/i)
  })

  it('should create blank presentations with custom aspect ratios', async () => {
    const ppt169 = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })
    expect(ppt169.getSlideDimensions()).toEqual({ width: 12192000, height: 6858000 })

    const ppt43 = await PPTXTemplater.createBlank({ aspectRatio: '4:3' })
    expect(ppt43.getSlideDimensions()).toEqual({ width: 9144000, height: 6858000 })
  })
})
