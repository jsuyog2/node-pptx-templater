import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Shape Operations Integration Tests (TASK-SHAPE-002)', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_shape_operations.pptx')

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

  it('should support addShape with various geometry presets, percentages, and fluent chaining', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Fluent method chaining
    const chained = ppt
      .addShape(1, 'rectangle', {
        id: 'HeaderCard',
        x: 50,
        y: 60,
        width: 300,
        height: 120,
        fill: '#0F2027',
        border: { color: '#203A43', width: 2 },
        text: 'Chained Header Card',
        textStyle: { bold: true, fontSize: 16, color: '#FFFFFF', align: 'center' },
      })
      .addShape(1, 'circle', {
        id: 'StatusPill',
        x: '75%',
        y: '20%',
        radius: 35,
        fill: '#00C853',
      })
      .addShape(1, 'triangle', {
        id: 'AccentTriangle',
        x: 100,
        y: 250,
        width: 80,
        height: 80,
        fill: '#FF6D00',
        rotation: 45,
        flipH: true,
      })
      .addShape(1, 'roundedRectangle', {
        id: 'RoundedBtn',
        x: 220,
        y: 250,
        width: 160,
        height: 60,
        borderRadius: 15,
        fill: '#2979FF',
        text: 'Action Button',
      })

    expect(chained).toBe(ppt)

    const shapes = ppt.getShapes(1)
    const header = shapes.find(s => s.name === 'HeaderCard' || s.id === 'HeaderCard')
    const pill = shapes.find(s => s.name === 'StatusPill' || s.id === 'StatusPill')
    const tri = shapes.find(s => s.name === 'AccentTriangle' || s.id === 'AccentTriangle')
    const btn = shapes.find(s => s.name === 'RoundedBtn' || s.id === 'RoundedBtn')

    expect(header).toBeDefined()
    expect(pill).toBeDefined()
    expect(tri).toBeDefined()
    expect(btn).toBeDefined()

    expect(header.text).toContain('Chained Header Card')
    expect(btn.text).toContain('Action Button')
    expect(tri.geometry).toBe('triangle')

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    const reloadedShapes = reloaded.getShapes(1)
    expect(reloadedShapes.some(s => s.name === 'HeaderCard')).toBe(true)
    expect(reloadedShapes.some(s => s.name === 'RoundedBtn')).toBe(true)
  })

  it('should update shape attributes and mutate them in-place with fluent return', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape(1, 'rectangle', {
      id: 'TargetToUpdate',
      x: 100,
      y: 100,
      width: 200,
      height: 100,
      fill: '#112233',
      text: 'Original Text',
    })

    const res = ppt.updateShape('TargetToUpdate', {
      x: 150,
      y: 180,
      width: 260,
      height: 140,
      fill: '#FF0055',
      text: 'Updated Text Label',
      rotation: 90,
    })

    expect(res).toBe(ppt)

    const shape = ppt.getShape('TargetToUpdate', 1)
    expect(shape).toBeDefined()
    expect(shape.position.x).toBe(150 * 9525)
    expect(shape.position.y).toBe(180 * 9525)
    expect(shape.position.cx).toBe(260 * 9525)
    expect(shape.position.cy).toBe(140 * 9525)
    expect(shape.text).toBe('Updated Text Label')

    // Error handling on non-existent shape
    expect(() => {
      ppt.updateShape('DoesNotExistShape', { x: 50 })
    }).toThrow(/not found/)
  })

  it('should clone an existing shape and position it with offsets', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape(1, 'star5', {
      id: 'OriginalStar',
      x: 100,
      y: 100,
      width: 120,
      height: 120,
      fill: '#FFD700',
      text: 'Gold Star',
    })

    const chained = ppt.cloneShape('OriginalStar', 'ClonedStarCopy', {
      x: 150,
      y: 180,
      text: 'Cloned Gold Star',
    })

    expect(chained).toBe(ppt)

    const cloned = ppt.getShape('ClonedStarCopy', 1)
    expect(cloned).toBeDefined()
    expect(cloned.name).toBe('ClonedStarCopy')
    expect(cloned.position.x).toBe(150 * 9525)
    expect(cloned.position.y).toBe(180 * 9525)
    expect(cloned.text).toBe('Cloned Gold Star')

    // Cloned shape should persist on save and reload
    await ppt.saveToFile(outputPath)
    const reloaded = await PPTXTemplater.load(outputPath)
    const starClone = reloaded.getShape('ClonedStarCopy', 1)
    expect(starClone).toBeDefined()
  })

  it('should remove shapes by ID or name and update z-order', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape(1, 'rectangle', {
      id: 'ShapeToDelete',
      x: 80,
      y: 80,
      width: 100,
      height: 100,
    })

    expect(ppt.getShape('ShapeToDelete', 1)).toBeDefined()

    const chained = ppt.removeShape('ShapeToDelete', 1)
    expect(chained).toBe(ppt)

    expect(ppt.getShape('ShapeToDelete', 1)).toBeNull()

    // Removing non-existent shape throws error
    expect(() => {
      ppt.removeShape('ShapeToDelete', 1)
    }).toThrow(/not found/)
  })

  it('should clear user shapes while preserving master and layout placeholders', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Check existing shapes on slide 1 (which may include Title placeholder)
    const initialShapes = ppt.getShapes(1)
    const initialPlaceholderCount = initialShapes.filter(s => s.isPlaceholder).length

    // Add 3 user shapes
    ppt
      .addShape(1, 'rectangle', { id: 'UserShape1', x: 20, y: 20, width: 80, height: 80 })
      .addShape(1, 'circle', { id: 'UserShape2', x: 120, y: 20, radius: 40 })
      .addShape(1, 'diamond', { id: 'UserShape3', x: 220, y: 20, width: 60, height: 60 })

    expect(ppt.getShapes(1).length).toBe(initialShapes.length + 3)

    // Clear shapes
    const chained = ppt.clearShapes(1)
    expect(chained).toBe(ppt)

    const remainingShapes = ppt.getShapes(1)
    // User shapes must be gone
    expect(remainingShapes.find(s => s.name === 'UserShape1')).toBeUndefined()
    expect(remainingShapes.find(s => s.name === 'UserShape2')).toBeUndefined()
    expect(remainingShapes.find(s => s.name === 'UserShape3')).toBeUndefined()

    // Placeholders must be preserved
    const remainingPlaceholders = remainingShapes.filter(s => s.isPlaceholder)
    expect(remainingPlaceholders.length).toBe(initialPlaceholderCount)
  })
})
