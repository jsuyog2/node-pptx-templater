import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Shape Inspection & Metadata Querying (TASK-SHAPE-004)', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_shape_inspection.pptx')

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

  it('should retrieve detailed metadata for individual shapes via getShape', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape(1, 'roundRect', {
      id: 'InspectableCard',
      x: 120,
      y: 140,
      width: 250,
      height: 150,
      borderRadius: 12,
      fill: '#4A90E2',
      border: { color: '#1B365D', width: 3 },
      rotation: 30,
      flipH: true,
      text: 'Inspected Shape Label',
      textStyle: { fontSize: 16, bold: true, color: '#FFFFFF' },
    })

    const descriptor = ppt.getShape('InspectableCard', 1)
    expect(descriptor).not.toBeNull()
    expect(descriptor.id).toBe('InspectableCard')
    expect(descriptor.name).toBe('InspectableCard')
    expect(descriptor.type).toBe('roundedRectangle')
    expect(descriptor.geometry).toBe('roundRect')

    // Coordinate & dimension checks in both pixels and EMUs
    expect(descriptor.x).toBe(120)
    expect(descriptor.y).toBe(140)
    expect(descriptor.width).toBe(250)
    expect(descriptor.height).toBe(150)

    expect(descriptor.bounds).toBeDefined()
    expect(descriptor.bounds.x).toBe(120 * 9525)
    expect(descriptor.bounds.y).toBe(140 * 9525)
    expect(descriptor.bounds.width).toBe(250 * 9525)
    expect(descriptor.bounds.height).toBe(150 * 9525)

    expect(descriptor.position.x).toBe(120 * 9525)
    expect(descriptor.position.y).toBe(140 * 9525)
    expect(descriptor.position.cx).toBe(250 * 9525)
    expect(descriptor.position.cy).toBe(150 * 9525)

    // Styling checks
    expect(descriptor.fill).toBe('#4A90E2')
    expect(descriptor.border).toBeDefined()
    expect(descriptor.border.color).toBe('#1B365D')
    expect(descriptor.border.width).toBe(3)
    expect(descriptor.rotation).toBe(30)
    expect(descriptor.flipH).toBe(true)
    expect(descriptor.flipV).toBe(false)

    // Inner text check
    expect(descriptor.text).toBe('Inspected Shape Label')
    expect(descriptor.isGroup).toBe(false)

    // Non-existent shape query returns null
    expect(ppt.getShape('NoSuchShapeExists', 1)).toBeNull()
  })

  it('should list all shapes across slide or presentation using getShapes', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt
      .addShape(1, 'rectangle', { id: 'BatchShapeA', x: 20, y: 30, width: 60, height: 40 })
      .addShape(1, 'ellipse', { id: 'BatchShapeB', x: 100, y: 30, width: 80, height: 60 })

    const slide1Shapes = ppt.getShapes(1)
    expect(slide1Shapes.length).toBeGreaterThanOrEqual(2)

    const shapeA = slide1Shapes.find(s => s.name === 'BatchShapeA')
    const shapeB = slide1Shapes.find(s => s.name === 'BatchShapeB')

    expect(shapeA).toBeDefined()
    expect(shapeB).toBeDefined()
    expect(shapeA.type).toBe('rectangle')
    expect(shapeB.type).toBe('ellipse')

    // Querying presentation-wide
    const allShapes = ppt.getShapes()
    expect(allShapes.some(s => s.name === 'BatchShapeA')).toBe(true)
    expect(allShapes.some(s => s.name === 'BatchShapeB')).toBe(true)
  })

  it('should inspect grouped shapes and reflect group metadata in getShapes', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt
      .addShape(1, 'rectangle', { id: 'ChildElem1', x: 50, y: 50, width: 100, height: 50 })
      .addShape(1, 'rectangle', { id: 'ChildElem2', x: 180, y: 50, width: 100, height: 50 })

    ppt.groupShapes(['ChildElem1', 'ChildElem2'], { name: 'InspectionGroup' })

    const groupMeta = ppt.getShape('InspectionGroup', 1)
    expect(groupMeta).toBeDefined()
    expect(groupMeta.isGroup).toBe(true)
    expect(groupMeta.type).toBe('group')

    const slideShapes = ppt.getShapes(1)
    const grpEntry = slideShapes.find(s => s.name === 'InspectionGroup')
    expect(grpEntry).toBeDefined()
    expect(grpEntry.isGroup).toBe(true)
  })
})
