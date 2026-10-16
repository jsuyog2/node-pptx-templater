import { describe, it, expect } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'
import fs from 'fs-extra'
import path from 'path'

const TEMP_DIR = path.resolve(__dirname, '../temp-blank-presentation')

describe('Template-Free Blank Presentation Generation & Mutation', () => {
  it('should initialize 16:9 widescreen blank presentation with exact EMU dimensions', async () => {
    const ppt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })
    expect(ppt.slideCount).toBe(1)

    const dims = ppt.getSlideDimensions()
    expect(dims.width).toBe(12192000)
    expect(dims.height).toBe(6858000)
  })

  it('should initialize 4:3 standard blank presentation with exact EMU dimensions', async () => {
    const ppt = await PPTXTemplater.createBlank({ aspectRatio: '4:3' })
    expect(ppt.slideCount).toBe(1)

    const dims = ppt.getSlideDimensions()
    expect(dims.width).toBe(9144000)
    expect(dims.height).toBe(6858000)
  })

  it('should support custom dimensions and unit conversions', async () => {
    const ppt = await PPTXTemplater.createBlank({
      width: 1920,
      height: 1080,
      unit: 'px',
    })

    const dims = ppt.getSlideDimensions()
    // 1920 * 9525 = 18288000
    // 1080 * 9525 = 10287000
    expect(dims.width).toBe(18288000)
    expect(dims.height).toBe(10287000)
  })

  it('should allow adding slides, shapes, tables, and charts to blank presentation and produce valid output', async () => {
    await fs.remove(TEMP_DIR)
    await fs.ensureDir(TEMP_DIR)

    const ppt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })
    ppt.useSlide(1)

    // Add shape to blank slide 1
    await ppt.addShape('roundedRectangle', {
      id: 'titleBox',
      x: 100,
      y: 100,
      width: 300,
      height: 150,
      text: 'Title Box',
      fill: '#2563EB',
    })

    // Add table to slide 1
    ppt.addTable(
      [
        ['Header 1', 'Header 2'],
        ['Cell 1', 'Cell 2'],
      ],
      { x: 100, y: 300, width: 400, height: 100 }
    )

    // Add a second slide
    ppt.addSlide({ title: 'Analytics Slide' })
    expect(ppt.slideCount).toBe(2)

    // Save presentation
    const outPath = path.join(TEMP_DIR, 'blank-populated.pptx')
    await ppt.saveToFile(outPath)
    expect(await fs.pathExists(outPath)).toBe(true)

    // Verify it re-loads without error
    const reloaded = await PPTXTemplater.load(outPath)
    expect(reloaded.slideCount).toBe(2)

    await fs.remove(TEMP_DIR)
  })
})
