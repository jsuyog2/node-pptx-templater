import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Arrow Shape Geometry & Preset Mapping (TASK-SHAPE-ARROW-001)', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_arrow_shape.pptx')

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

  it('should map generic "arrow" to "rightArrow" preset with valid DrawingML AST', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape(1, 'arrow', {
      id: 'DefaultArrow',
      x: 50,
      y: 100,
      width: 200,
      height: 80,
      fill: '#2563EB',
    })

    const shape = ppt.getShape('DefaultArrow')
    expect(shape).not.toBeNull()
    expect(shape.geometry).toBe('rightArrow')
    expect(shape.type).toBe('rightArrow')
    expect(shape.width).toBe(200)
    expect(shape.height).toBe(80)

    // Verify raw DrawingML slide XML
    const slideXml = ppt.slideManager.getSlideXml(1)
    expect(slideXml).toContain('prst="rightArrow"')
    expect(slideXml).toMatch(
      /<a:prstGeom prst="rightArrow">\s*(<a:avLst\/>|<a:avLst><\/a:avLst>)\s*<\/a:prstGeom>/
    )

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    // Re-load and verify
    const reloaded = await PPTXTemplater.load(outputPath)
    const reloadedShape = reloaded.getShape('DefaultArrow')
    expect(reloadedShape).not.toBeNull()
    expect(reloadedShape.geometry).toBe('rightArrow')
  })

  it('should support directional arrow presets (leftArrow, upArrow, downArrow, leftRightArrow)', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt
      .addShape(1, 'leftArrow', {
        id: 'LeftArrowShape',
        x: 50,
        y: 50,
        width: 150,
        height: 60,
        fill: '#10B981',
      })
      .addShape(1, 'upArrow', {
        id: 'UpArrowShape',
        x: 250,
        y: 50,
        width: 60,
        height: 150,
        fill: '#F59E0B',
      })
      .addShape(1, 'downArrow', {
        id: 'DownArrowShape',
        x: 350,
        y: 50,
        width: 60,
        height: 150,
        fill: '#EF4444',
      })
      .addShape(1, 'leftRightArrow', {
        id: 'BiDirectionalArrow',
        x: 450,
        y: 50,
        width: 200,
        height: 60,
        fill: '#8B5CF6',
      })

    expect(ppt.getShape('LeftArrowShape').geometry).toBe('leftArrow')
    expect(ppt.getShape('UpArrowShape').geometry).toBe('upArrow')
    expect(ppt.getShape('DownArrowShape').geometry).toBe('downArrow')
    expect(ppt.getShape('BiDirectionalArrow').geometry).toBe('leftRightArrow')

    const slideXml = ppt.slideManager.getSlideXml(1)
    expect(slideXml).toContain('prst="leftArrow"')
    expect(slideXml).toContain('prst="upArrow"')
    expect(slideXml).toContain('prst="downArrow"')
    expect(slideXml).toContain('prst="leftRightArrow"')
  })

  it('should update shape geometry preset to arrow dynamically', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape(1, 'rectangle', {
      id: 'MorphShape',
      x: 100,
      y: 100,
      width: 150,
      height: 75,
      fill: '#3B82F6',
    })

    expect(ppt.getShape('MorphShape').geometry).toBe('rect')

    // Update to arrow
    ppt.updateShape('MorphShape', {
      type: 'arrow',
    })

    expect(ppt.getShape('MorphShape').geometry).toBe('rightArrow')
    const slideXml = ppt.slideManager.getSlideXml(1)
    expect(slideXml).toContain('prst="rightArrow"')
  })

  it('should support line connectors with arrow heads and tail markers in <a:ln>', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape(1, 'lineArrow', {
      id: 'DirectionalConnector',
      x: 50,
      y: 200,
      width: 300,
      height: 0,
      strokeColor: '#DC2626',
      strokeWidth: 3,
      headEnd: 'triangle',
      tailEnd: 'oval',
    })

    const slideXml = ppt.slideManager.getSlideXml(1)
    expect(slideXml).toMatch(/<a:headEnd type="triangle"(\/>|><\/a:headEnd>)/)
    expect(slideXml).toMatch(/<a:tailEnd type="oval"(\/>|><\/a:tailEnd>)/)
    expect(slideXml).toContain('val="DC2626"')

    const shape = ppt.getShape('DirectionalConnector')
    expect(shape).not.toBeNull()
    expect(shape.geometry).toBe('line')
  })

  it('should calculate proper EMU bounding coordinates for arrow geometries', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.addShape(1, 'rightArrow', {
      id: 'EmuArrow',
      x: 100, // 100 * 9525 = 952500 EMUs
      y: 150, // 150 * 9525 = 1428750 EMUs
      width: 200, // 200 * 9525 = 1905000 EMUs
      height: 80, // 80 * 9525 = 762000 EMUs
    })

    const shape = ppt.getShape('EmuArrow')
    expect(shape.bounds.x).toBe(100 * 9525)
    expect(shape.bounds.y).toBe(150 * 9525)
    expect(shape.bounds.width).toBe(200 * 9525)
    expect(shape.bounds.height).toBe(80 * 9525)
  })

  it('should assign unique fallback ID when none provided and support getShape, updateShape, and removeShape', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Add arrow shape without specifying id or name
    ppt.addShape(1, 'arrow', {
      x: 120,
      y: 160,
      width: 180,
      height: 90,
      fill: '#10B981',
    })

    // Retrieve via getShapes
    const shapes = ppt.getShapes(1)
    const addedArrow = shapes.find(s => s.geometry === 'rightArrow' && s.x === 120)
    expect(addedArrow).toBeDefined()
    expect(addedArrow.id).toBeDefined()

    // Query by the assigned fallback ID
    const foundById = ppt.getShape(addedArrow.id, 1)
    expect(foundById).not.toBeNull()
    expect(foundById.geometry).toBe('rightArrow')
    expect(foundById.width).toBe(180)

    // Update using the assigned fallback ID
    ppt.updateShape(addedArrow.id, {
      width: 250,
      fill: '#F59E0B',
    })

    const updated = ppt.getShape(addedArrow.id, 1)
    expect(updated.width).toBe(250)
    expect(updated.fill).toBe('#F59E0B')

    // Remove using the assigned fallback ID
    ppt.removeShape(addedArrow.id, 1)
    expect(ppt.getShape(addedArrow.id, 1)).toBeNull()
  })

  it('should query, update, and remove <p:cxnSp> connector elements', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Directly inject a <p:cxnSp> connector into slide 1 spTree to simulate existing presentation connectors
    const slideObj = ppt.slideManager.getSlideObj(1)
    const spTree = slideObj['p:sld']['p:cSld']['p:spTree']
    if (!spTree['p:cxnSp']) spTree['p:cxnSp'] = []
    spTree['p:cxnSp'].push({
      'p:nvCxnSpPr': {
        'p:cNvPr': {
          '@_id': '8888',
          '@_name': 'ConnectorArrow8888',
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
          '@_w': '25400',
          'a:solidFill': { 'a:srgbClr': { '@_val': '2563EB' } },
          'a:headEnd': { '@_type': 'triangle' },
        },
      },
    })
    ppt.slideManager.markSlideObjDirty(1)

    // Query by ID
    const byId = ppt.getShape('8888', 1)
    expect(byId).not.toBeNull()
    expect(byId.name).toBe('ConnectorArrow8888')
    expect(byId.isConnector).toBe(true)
    expect(byId.geometry).toBe('line')
    expect(byId.border.color).toBe('#2563EB')
    expect(byId.border.headEnd).toBe('triangle')

    // Query by Name
    const byName = ppt.getShape('ConnectorArrow8888', 1)
    expect(byName).not.toBeNull()
    expect(byName.id).toBe('ConnectorArrow8888')

    // Discover in getShapes()
    const allShapes = ppt.getShapes(1)
    const foundInList = allShapes.find(
      s => s.id === 'ConnectorArrow8888' || s.name === 'ConnectorArrow8888'
    )
    expect(foundInList).toBeDefined()

    // Update connector
    ppt.updateShape('8888', {
      x: 200,
      width: 300,
      border: { color: '#DC2626', width: 4 },
    })

    const updatedCxn = ppt.getShape('8888', 1)
    expect(updatedCxn.x).toBe(200)
    expect(updatedCxn.width).toBe(300)
    expect(updatedCxn.border.color).toBe('#DC2626')

    // Remove connector
    ppt.removeShape('8888', 1)
    expect(ppt.getShape('8888', 1)).toBeNull()
  })
})
