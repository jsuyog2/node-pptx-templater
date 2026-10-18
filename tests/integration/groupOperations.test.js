import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Group and Ungroup Subsystem Integration Tests (TASK-SHAPE-003)', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_group_operations.pptx')

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

  it('should group multiple shapes, query group children, and verify composite bounds', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Add 3 shapes on slide 1
    ppt
      .addShape(1, 'rectangle', {
        id: 'CardBackground',
        x: 100,
        y: 100,
        width: 300,
        height: 200,
        fill: '#336699',
      })
      .addShape(1, 'circle', {
        id: 'CardIcon',
        x: 120,
        y: 120,
        radius: 30,
        fill: '#FFFFFF',
      })
      .addShape(1, 'rectangle', {
        id: 'CardTitle',
        x: 190,
        y: 125,
        width: 180,
        height: 40,
        text: 'Composite Card',
      })

    // Group the shapes with fluent return
    const chained = ppt.groupShapes(['CardBackground', 'CardIcon', 'CardTitle'], {
      name: 'CardGroupWidget',
    })

    expect(chained).toBe(ppt)

    // Query group shapes
    const children = ppt.getGroupShapes('CardGroupWidget')
    expect(children).toBeDefined()
    expect(children.length).toBe(3)

    const childNames = children.map(c => c.name)
    expect(childNames).toContain('CardBackground')
    expect(childNames).toContain('CardIcon')
    expect(childNames).toContain('CardTitle')

    // Top-level shapes on slide 1 should include the group and not the ungrouped children
    const shapes = ppt.getShapes(1)
    const groupDescriptor = shapes.find(s => s.name === 'CardGroupWidget')
    expect(groupDescriptor).toBeDefined()
    expect(groupDescriptor.isGroup).toBe(true)

    // Verify slide XML has <p:grpSp> container and composite transform
    const slideXml = ppt.slideManager.getSlideXml(1)
    expect(slideXml).toContain('CardGroupWidget')
    expect(slideXml).toContain('<p:grpSp>')
    expect(slideXml).toContain('<p:grpSpPr>')
    expect(slideXml).toContain('<a:chOff')
    expect(slideXml).toContain('<a:chExt')

    // Save and reload to verify package integrity
    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    const reloadedChildren = reloaded.getGroupShapes('CardGroupWidget', 1)
    expect(reloadedChildren.length).toBe(3)
  })

  it('should ungroup shapes, restore them to parent slide, and re-project coordinates', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    const origX1 = 150
    const origY1 = 150
    const origW1 = 200
    const origH1 = 100

    const origX2 = 250
    const origY2 = 200
    const origW2 = 120
    const origH2 = 80

    ppt
      .addShape(1, 'rectangle', {
        id: 'UngroupShapeA',
        x: origX1,
        y: origY1,
        width: origW1,
        height: origH1,
        fill: '#10B981',
      })
      .addShape(1, 'rectangle', {
        id: 'UngroupShapeB',
        x: origX2,
        y: origY2,
        width: origW2,
        height: origH2,
        fill: '#3B82F6',
      })

    ppt.groupShapes(['UngroupShapeA', 'UngroupShapeB'], { name: 'TemporaryGroup' })

    // Verify group exists
    expect(ppt.getShapes(1).some(s => s.name === 'TemporaryGroup')).toBe(true)

    // Dissolve group
    const chained = ppt.ungroupShapes('TemporaryGroup')
    expect(chained).toBe(ppt)

    // Group container should be gone
    expect(ppt.getShapes(1).some(s => s.name === 'TemporaryGroup')).toBe(false)

    // Children shapes must now exist directly on the slide with exact or reprojected coordinates
    const shapeA = ppt.getShape('UngroupShapeA', 1)
    const shapeB = ppt.getShape('UngroupShapeB', 1)

    expect(shapeA).toBeDefined()
    expect(shapeB).toBeDefined()

    expect(shapeA.position.x).toBe(origX1 * 9525)
    expect(shapeA.position.y).toBe(origY1 * 9525)
    expect(shapeA.position.cx).toBe(origW1 * 9525)
    expect(shapeA.position.cy).toBe(origH1 * 9525)

    expect(shapeB.position.x).toBe(origX2 * 9525)
    expect(shapeB.position.y).toBe(origY2 * 9525)
    expect(shapeB.position.cx).toBe(origW2 * 9525)
    expect(shapeB.position.cy).toBe(origH2 * 9525)

    // Validate presentation schema
    const report = await ppt.validatePresentation()
    expect(report.valid).toBe(true)
  })

  it('should validate inputs and throw informative errors on invalid grouping operations', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Attempting to group fewer than 2 shapes
    expect(() => {
      ppt.groupShapes(['SingleShape'])
    }).toThrow(/At least 2 shape IDs/)

    // Attempting to group non-existent shapes
    expect(() => {
      ppt.groupShapes(['GhostA', 'GhostB'])
    }).toThrow(/not found/)

    // Attempting to ungroup non-existent group
    expect(() => {
      ppt.ungroupShapes('NonExistentGroup')
    }).toThrow(/not found/)

    // Querying non-existent group returns empty array
    expect(ppt.getGroupShapes('NoGroupHere')).toEqual([])
  })
})
