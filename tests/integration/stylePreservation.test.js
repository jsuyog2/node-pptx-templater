import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Non-Destructive Style Merging & Fill Preservation (TASK-STYLE-PRESERVE-002)', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_style_preservation.pptx')

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

  it('should preserve existing shape solid fill when updating text, fontSize, bold, and color', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // 1. Add a shape with explicit background fill #3B82F6
    ppt.addShape(1, 'rectangle', {
      id: 'CardShape',
      x: 50,
      y: 50,
      width: 200,
      height: 100,
      fill: '#3B82F6',
      text: 'Initial Text',
      textStyle: {
        fontSize: 14,
        color: '#FFFFFF',
      },
    })

    let shape = ppt.getShape('CardShape')
    expect(shape.fill).toBe('#3B82F6')
    expect(shape.text).toBe('Initial Text')

    // Verify raw XML has solidFill val="3B82F6"
    let slideXml = ppt.slideManager.getSlideXml(1)
    expect(slideXml).toContain('val="3B82F6"')

    // 2. Update ONLY text and fontSize without specifying fill
    ppt.updateShape('CardShape', {
      text: 'Updated Text Without Wiping Fill',
      fontSize: 24,
      bold: true,
      color: '#FBBF24', // Yellow text color, should NOT overwrite blue shape fill!
    })

    shape = ppt.getShape('CardShape')
    expect(shape.text).toBe('Updated Text Without Wiping Fill')
    // Background fill must remain preserved as #3B82F6!
    expect(shape.fill).toBe('#3B82F6')

    slideXml = ppt.slideManager.getSlideXml(1)
    // Shape spPr must still have solidFill 3B82F6
    expect(slideXml).toContain('val="3B82F6"')
    // Text run must have new color FBBF24 and 24pt (2400 sz)
    expect(slideXml).toContain('val="FBBF24"')
    expect(slideXml).toContain('sz="2400"')
    expect(slideXml).toContain('b="1"')
  })

  it('should preserve gradient fill when updating text or styling without fill', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape(1, 'roundedRectangle', {
      id: 'GradCard',
      x: 100,
      y: 100,
      width: 250,
      height: 120,
      fill: {
        type: 'gradient',
        colors: ['#4F46E5', '#06B6D4'],
      },
      text: 'Grad Content',
    })

    let slideXml = ppt.slideManager.getSlideXml(1)
    expect(slideXml).toContain('<a:gradFill')
    expect(slideXml).toContain('val="4F46E5"')
    expect(slideXml).toContain('val="06B6D4"')

    // Update text and font size only
    ppt.updateShape('GradCard', {
      text: 'New Header',
      fontSize: 20,
    })

    slideXml = ppt.slideManager.getSlideXml(1)
    // Gradient fill must still be intact
    expect(slideXml).toContain('<a:gradFill')
    expect(slideXml).toContain('val="4F46E5"')
    expect(slideXml).toContain('val="06B6D4"')
    expect(ppt.getShape('GradCard').text).toBe('New Header')
  })

  it('should update shape fill when fill parameter is explicitly provided', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape(1, 'rectangle', {
      id: 'MutableBox',
      x: 50,
      y: 50,
      width: 150,
      height: 80,
      fill: '#3B82F6',
    })

    expect(ppt.getShape('MutableBox').fill).toBe('#3B82F6')

    // 1. Explicitly update to green fill
    ppt.updateShape('MutableBox', {
      fill: '#10B981',
    })

    expect(ppt.getShape('MutableBox').fill).toBe('#10B981')
    let slideXml = ppt.slideManager.getSlideXml(1)
    expect(slideXml).toContain('val="10B981"')
    expect(slideXml).not.toContain('val="3B82F6"')

    // 2. Explicitly update to noFill ('none')
    ppt.updateShape('MutableBox', {
      fill: 'none',
    })

    slideXml = ppt.slideManager.getSlideXml(1)
    expect(slideXml).toMatch(/<a:noFill(\/>|><\/a:noFill>)/)
    expect(ppt.getShape('MutableBox').fill).toBeUndefined()
  })

  it('should preserve shape background fill when using updateStyle', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape(1, 'rectangle', {
      id: 'StyledBox',
      x: 60,
      y: 60,
      width: 200,
      height: 90,
      fill: '#8B5CF6',
      text: 'Original Label',
    })

    expect(ppt.getShape('StyledBox').fill).toBe('#8B5CF6')

    // Call updateStyle with typography overrides
    ppt.updateStyle('StyledBox', {
      fontSize: 22,
      bold: true,
      color: '#FFFFFF',
    })

    const shape = ppt.getShape('StyledBox')
    expect(shape.fill).toBe('#8B5CF6')

    const slideXml = ppt.slideManager.getSlideXml(1)
    expect(slideXml).toContain('val="8B5CF6"')
    expect(slideXml).toContain('sz="2200"')
  })

  it('should preserve shape background fill when using updateText', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape(1, 'rectangle', {
      id: 'TextBoxShape',
      x: 70,
      y: 70,
      width: 220,
      height: 100,
      fill: '#EC4899',
      text: 'Old Text Content',
    })

    expect(ppt.getShape('TextBoxShape').fill).toBe('#EC4899')

    // Call updateText
    ppt.updateText('TextBoxShape', 'Brand New Content', {
      fontSize: 18,
      color: '#FFFFFF',
    })

    const shape = ppt.getShape('TextBoxShape')
    expect(shape.text).toBe('Brand New Content')
    expect(shape.fill).toBe('#EC4899')

    const slideXml = ppt.slideManager.getSlideXml(1)
    expect(slideXml).toContain('val="EC4899"')
    expect(slideXml).toContain('Brand New Content')
  })

  it('should preserve table cell shape fill when updating cell shape properties', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Add a table first
    await ppt.addTable(
      [
        ['Header 1', 'Header 2'],
        ['Cell A', 'Cell B'],
      ],
      {
        id: 'SampleTable',
        x: 50,
        y: 50,
        width: 400,
        height: 200,
      }
    )

    // Add cell shape with fill #EF4444
    await ppt.addCellShape('SampleTable', 1, 0, {
      type: 'rectangle',
      fill: '#EF4444',
      text: 'Badge',
      width: 60,
      height: 30,
    })

    let cellShape = ppt.getCellShape('SampleTable', 1, 0, 0)
    expect(cellShape).not.toBeNull()
    expect(cellShape.fill).toBe('#EF4444')

    // Update cell shape text without providing fill
    await ppt.updateCellShape('SampleTable', 1, 0, 0, {
      text: 'Updated Badge',
    })

    cellShape = ppt.getCellShape('SampleTable', 1, 0, 0)
    expect(cellShape).not.toBeNull()
    expect(cellShape.text).toBe('Updated Badge')
    // Fill must be preserved!
    expect(cellShape.fill).toBe('#EF4444')

    // Update cell shape with explicit fill #10B981
    await ppt.updateCellShape('SampleTable', 1, 0, 0, {
      fill: '#10B981',
    })

    cellShape = ppt.getCellShape('SampleTable', 1, 0, 0)
    expect(cellShape.fill).toBe('#10B981')

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)
  })
})
