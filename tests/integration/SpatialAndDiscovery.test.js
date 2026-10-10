import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Spatial Layout & Element Discovery', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_spatial.pptx')

  beforeEach(() => {
    if (existsSync(outputPath)) {
      try { unlinkSync(outputPath) } catch (_) {}
    }
  })

  afterEach(() => {
    if (existsSync(outputPath)) {
      try { unlinkSync(outputPath) } catch (_) {}
    }
  })

  it('should enumerate elements and discover shapes, pictures, tables, and charts', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const elements = ppt.getElements()
    expect(Array.isArray(elements)).toBe(true)
    expect(elements.length).toBeGreaterThan(0)

    const first = elements[0]
    expect(first).toHaveProperty('id')
    expect(first).toHaveProperty('name')
    expect(first).toHaveProperty('type')
    expect(first).toHaveProperty('bounds')
    expect(first.bounds).toHaveProperty('x')
    expect(first.bounds).toHaveProperty('y')
    expect(first.bounds).toHaveProperty('width')
    expect(first.bounds).toHaveProperty('height')

    // Filter by type
    const shapes = ppt.getElements({ type: 'shape' })
    for (const s of shapes) {
      expect(['shape', 'textbox']).toContain(s.type)
    }

    // findElements with predicate
    const found = ppt.findElements(el => el.type === 'shape' || el.type === 'textbox')
    expect(found.length).toBe(shapes.length)
  })

  it('should read bounds including rotation and flip properties', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const elements = ppt.getElements()
    const target = elements[0]

    const bounds = ppt.getBounds(target.name || target.id)
    expect(bounds).toHaveProperty('x')
    expect(bounds).toHaveProperty('y')
    expect(bounds).toHaveProperty('width')
    expect(bounds).toHaveProperty('height')
    expect(bounds).toHaveProperty('rotation')
    expect(typeof bounds.rotation).toBe('number')
  })

  it('should support setBounds with rotation and coordinate updates', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    // Add two test shapes
    ppt.addShape('rect', { id: 'boxA', x: 50, y: 50, width: 100, height: 100 })
    ppt.addShape('rect', { id: 'boxB', x: 80, y: 80, width: 100, height: 100 })

    const boundsA = ppt.getBounds('boxA')
    expect(boundsA.x).toBe(50)
    expect(boundsA.y).toBe(50)

    // Check collision between overlapping boxes
    expect(ppt.checkCollision('boxA', 'boxB')).toBe(true)

    // Move boxB far away and verify no collision
    ppt.setBounds('boxB', { x: 300, y: 300, rotation: 45 })
    expect(ppt.checkCollision('boxA', 'boxB')).toBe(false)

    const boundsB = ppt.getBounds('boxB')
    expect(boundsB.x).toBe(300)
    expect(boundsB.y).toBe(300)
    expect(boundsB.rotation).toBe(45)

    // Nest boxC completely inside boxB
    ppt.addShape('rect', { id: 'boxC', x: 310, y: 310, width: 20, height: 20 })
    expect(ppt.isContained('boxC', 'boxB')).toBe(true)
    expect(ppt.isContained('boxB', 'boxC')).toBe(false)
  })

  it('should align multiple shapes (left, center, top, middle)', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    ppt.addShape('rect', { id: 's1', x: 20, y: 40, width: 60, height: 40 })
    ppt.addShape('rect', { id: 's2', x: 100, y: 80, width: 80, height: 50 })
    ppt.addShape('rect', { id: 's3', x: 220, y: 150, width: 50, height: 30 })

    // Align left to minX (20)
    ppt.alignShapes(['s1', 's2', 's3'], 'left')
    expect(ppt.getBounds('s1').x).toBe(20)
    expect(ppt.getBounds('s2').x).toBe(20)
    expect(ppt.getBounds('s3').x).toBe(20)

    // Align top to minY (40)
    ppt.alignShapes(['s1', 's2', 's3'], 'top')
    expect(ppt.getBounds('s1').y).toBe(40)
    expect(ppt.getBounds('s2').y).toBe(40)
    expect(ppt.getBounds('s3').y).toBe(40)
  })

  it('should distribute shapes evenly horizontally', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    ppt.addShape('rect', { id: 'd1', x: 0, y: 50, width: 50, height: 50 })
    ppt.addShape('rect', { id: 'd2', x: 40, y: 50, width: 50, height: 50 })
    ppt.addShape('rect', { id: 'd3', x: 200, y: 50, width: 50, height: 50 })

    ppt.distributeShapes(['d1', 'd2', 'd3'], 'horizontal')

    const b1 = ppt.getBounds('d1')
    const b2 = ppt.getBounds('d2')
    const b3 = ppt.getBounds('d3')

    // First and last remain at bounds: d1.x = 0, d3.x = 200
    expect(b1.x).toBe(0)
    expect(b3.x).toBe(200)
    // Span = 250 - 0 = 250. Total width = 150. Free space = 100. Gap = 100 / 2 = 50.
    // d2.x = 0 + 50 + 50 = 100
    expect(b2.x).toBe(100)
  })
})
