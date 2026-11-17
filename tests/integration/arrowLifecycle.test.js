import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Arrow Lifecycle Integration Suite (TASK-ARROW-001 - 003)', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_arrow_lifecycle.pptx')

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

  it('should cover full arrow lifecycle: add, getShape, updateShape, and removeShape', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // 1. Add arrow with generic type 'arrow'
    ppt.addShape(1, 'arrow', {
      id: 'FlowArrow',
      x: 100,
      y: 150,
      width: 250,
      height: 80,
      fill: '#2563EB',
    })

    // 2. Query arrow via getShape() by ID and verify preset mapping to rightArrow
    const shape = ppt.getShape('FlowArrow', 1)
    expect(shape).not.toBeNull()
    expect(shape.id).toBe('FlowArrow')
    expect(shape.geometry).toBe('rightArrow')
    expect(shape.fill).toBe('#2563EB')
    expect(shape.width).toBe(250)
    expect(shape.height).toBe(80)

    // Verify raw DrawingML serialized AST
    const slideXml = ppt.slideManager.getSlideXml(1)
    expect(slideXml).toContain('prst="rightArrow"')
    expect(slideXml).toMatch(
      /<a:prstGeom prst="rightArrow">\s*(<a:avLst\/>|<a:avLst><\/a:avLst>)\s*<\/a:prstGeom>/
    )

    // 3. Update shape dimensions, position, and color
    ppt.updateShape('FlowArrow', {
      x: 150,
      y: 200,
      width: 300,
      height: 100,
      fill: '#10B981',
    })

    const updated = ppt.getShape('FlowArrow', 1)
    expect(updated).not.toBeNull()
    expect(updated.x).toBe(150)
    expect(updated.y).toBe(200)
    expect(updated.width).toBe(300)
    expect(updated.height).toBe(100)
    expect(updated.fill).toBe('#10B981')

    // 4. Verify removeShape deletes arrow
    ppt.removeShape('FlowArrow', 1)
    expect(ppt.getShape('FlowArrow', 1)).toBeNull()
  })

  it('should support directional aliases rightArrow, leftArrow, upArrow, downArrow', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt
      .addShape(1, 'right-arrow', {
        id: 'ArrowRight',
        x: 50,
        y: 50,
        width: 120,
        height: 60,
      })
      .addShape(1, 'left-arrow', {
        id: 'ArrowLeft',
        x: 200,
        y: 50,
        width: 120,
        height: 60,
      })
      .addShape(1, 'up-arrow', {
        id: 'ArrowUp',
        x: 350,
        y: 50,
        width: 60,
        height: 120,
      })
      .addShape(1, 'down-arrow', {
        id: 'ArrowDown',
        x: 450,
        y: 50,
        width: 60,
        height: 120,
      })

    expect(ppt.getShape('ArrowRight').geometry).toBe('rightArrow')
    expect(ppt.getShape('ArrowLeft').geometry).toBe('leftArrow')
    expect(ppt.getShape('ArrowUp').geometry).toBe('upArrow')
    expect(ppt.getShape('ArrowDown').geometry).toBe('downArrow')
  })

  it('should assign a fallback ID when none is provided and allow referencing it', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape(1, 'arrow', {
      x: 80,
      y: 80,
      width: 150,
      height: 70,
    })

    const shapes = ppt.getShapes(1)
    const arrow = shapes.find(s => s.geometry === 'rightArrow')
    expect(arrow).toBeDefined()
    expect(arrow.id).toBeDefined()

    // Can query by fallback ID
    const queried = ppt.getShape(arrow.id, 1)
    expect(queried).not.toBeNull()
    expect(queried.geometry).toBe('rightArrow')

    // Can update by fallback ID
    ppt.updateShape(arrow.id, {
      width: 220,
    })
    expect(ppt.getShape(arrow.id, 1).width).toBe(220)

    // Can remove by fallback ID
    ppt.removeShape(arrow.id, 1)
    expect(ppt.getShape(arrow.id, 1)).toBeNull()
  })

  it('should accurately query, mutate, and delete <p:cxnSp> connectors', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    const slideObj = ppt.slideManager.getSlideObj(1)
    const spTree = slideObj['p:sld']['p:cSld']['p:spTree']
    if (!spTree['p:cxnSp']) spTree['p:cxnSp'] = []
    spTree['p:cxnSp'].push({
      'p:nvCxnSpPr': {
        'p:cNvPr': {
          '@_id': '9901',
          '@_name': 'Connector9901',
        },
        'p:cNvCxnSpPr': {},
        'p:nvPr': {},
      },
      'p:spPr': {
        'a:xfrm': {
          'a:off': { '@_x': '952500', '@_y': '952500' },
          'a:ext': { '@_cx': '1905000', '@_cy': '0' },
        },
        'a:prstGeom': {
          '@_prst': 'line',
          'a:avLst': {},
        },
        'a:ln': {
          '@_w': '19050',
          'a:solidFill': { 'a:srgbClr': { '@_val': '000000' } },
          'a:headEnd': { '@_type': 'triangle' },
        },
      },
    })
    ppt.slideManager.markSlideObjDirty(1)

    // Query by ID and by Name
    expect(ppt.getShape('9901', 1)).not.toBeNull()
    expect(ppt.getShape('Connector9901', 1)).not.toBeNull()
    expect(ppt.getShape('9901', 1).isConnector).toBe(true)

    // Update connector
    ppt.updateShape('9901', {
      x: 150,
      width: 280,
    })
    expect(ppt.getShape('9901', 1).x).toBe(150)
    expect(ppt.getShape('9901', 1).width).toBe(280)

    // Remove connector
    ppt.removeShape('9901', 1)
    expect(ppt.getShape('9901', 1)).toBeNull()
  })
})
