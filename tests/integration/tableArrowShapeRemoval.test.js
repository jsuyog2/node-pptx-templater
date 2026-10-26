/**
 * @fileoverview tableArrowShapeRemoval.test.js - Comprehensive Integration Tests
 * for Arrow Shapes and Connector (<p:cxnSp>) Removal and Shifting across Table Operations.
 *
 * Verifies TASK-SHOWCASE-VERIFY-001, TASK-ARROW-REMOVAL-FIX-002, TASK-ARROW-STANDALONE-FIX-003, and TASK-TEST-VERIFY-004.
 */

'use strict'

import { describe, it, expect } from 'vitest'
import path from 'path'
import { PPTXTemplater } from '../../src/index.js'

describe('Arrow Shape and Connector Removal & Lifecycle Verification (Slides 8 and 9)', () => {
  const samplePath = path.resolve(__dirname, '../../showcase/templates/TablePowerPointSample.pptx')

  it('TASK-ARROW-REMOVAL-FIX-002 (Slide 8): purges connector Straight Arrow Connector 16 on removeTableRow and shifts Straight Arrow Connector 9', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Inspect initial Slide 8 state
    ppt.useSlide(8)
    const initialSlideObj = ppt.slideManager.getSlideObj(8)
    const initialSpTree = initialSlideObj['p:sld']['p:cSld']['p:spTree']

    const initialCxns = Array.isArray(initialSpTree['p:cxnSp'])
      ? initialSpTree['p:cxnSp']
      : [initialSpTree['p:cxnSp']]
    const initialCxnNames = initialCxns.map(c => c['p:nvCxnSpPr']?.['p:cNvPr']?.['@_name'])

    expect(initialCxnNames).toContain('Straight Arrow Connector 16')
    expect(initialCxnNames).toContain('Straight Arrow Connector 9')

    const connector9Initial = initialCxns.find(
      c => c['p:nvCxnSpPr']?.['p:cNvPr']?.['@_name'] === 'Straight Arrow Connector 9'
    )
    const connector9InitialY = parseInt(connector9Initial['p:spPr']['a:xfrm']['a:off']['@_y'], 10)
    expect(connector9InitialY).toBe(3062288)

    // Execute Slide 8 operations from showcase/table_advanced.js
    // Row 1 removal
    ppt.removeTableRow('UpdateTable', 1)

    // Verify Straight Arrow Connector 16 is purged immediately
    const midSlideObj = ppt.slideManager.getSlideObj(8)
    const midSpTree = midSlideObj['p:sld']['p:cSld']['p:spTree']
    const midCxns = Array.isArray(midSpTree['p:cxnSp'])
      ? midSpTree['p:cxnSp']
      : midSpTree['p:cxnSp']
        ? [midSpTree['p:cxnSp']]
        : []
    const midCxnNames = midCxns.map(c => c['p:nvCxnSpPr']?.['p:cNvPr']?.['@_name'])

    expect(midCxnNames).not.toContain('Straight Arrow Connector 16')
    expect(midCxnNames).toContain('Straight Arrow Connector 9')

    // Row 3 removal (which was originally row 4)
    ppt.removeTableRow('UpdateTable', 3)

    // Verify surviving Straight Arrow Connector 9 has shifted up by 2 row heights (2 * 370840 = 741680)
    const finalSlideObj = ppt.slideManager.getSlideObj(8)
    const finalSpTree = finalSlideObj['p:sld']['p:cSld']['p:spTree']
    const finalCxns = Array.isArray(finalSpTree['p:cxnSp'])
      ? finalSpTree['p:cxnSp']
      : [finalSpTree['p:cxnSp']]
    const connector9Final = finalCxns.find(
      c => c['p:nvCxnSpPr']?.['p:cNvPr']?.['@_name'] === 'Straight Arrow Connector 9'
    )

    expect(connector9Final).toBeDefined()
    const connector9FinalY = parseInt(connector9Final['p:spPr']['a:xfrm']['a:off']['@_y'], 10)
    expect(connector9FinalY).toBe(connector9InitialY - 741680)
    expect(connector9FinalY).toBe(2320608)
  })

  it('TASK-ARROW-REMOVAL-FIX-002 (Slide 9): purges Straight Arrow Connector 19 on removeTableColumn and shifts surviving column shapes', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Inspect initial Slide 9 state
    ppt.useSlide(9)
    const initialSlideObj = ppt.slideManager.getSlideObj(9)
    const initialSpTree = initialSlideObj['p:sld']['p:cSld']['p:spTree']

    const initialCxns = Array.isArray(initialSpTree['p:cxnSp'])
      ? initialSpTree['p:cxnSp']
      : [initialSpTree['p:cxnSp']]
    const initialCxnNames = initialCxns.map(c => c['p:nvCxnSpPr']?.['p:cNvPr']?.['@_name'])

    expect(initialCxnNames).toContain('Straight Arrow Connector 24') // In Column 1
    expect(initialCxnNames).toContain('Straight Arrow Connector 19') // In Column 4

    const initialSps = Array.isArray(initialSpTree['p:sp'])
      ? initialSpTree['p:sp']
      : [initialSpTree['p:sp']]
    const rect21 = initialSps.find(s => s['p:nvSpPr']?.['p:cNvPr']?.['@_name'] === 'Rectangle 21')
    expect(rect21).toBeDefined()
    const rect21InitialX = parseInt(rect21['p:spPr']['a:xfrm']['a:off']['@_x'], 10)
    expect(rect21InitialX).toBe(8223739)

    // Execute Slide 9 operations from showcase/table_advanced.js
    // 1. Remove column 2 (Rectangle 22 is in Col 2)
    ppt.removeTableColumn('UpdateTable', 2)

    // 2. Remove column 3 (originally Column 4 containing Straight Arrow Connector 19)
    ppt.removeTableColumn('UpdateTable', 3)

    // Verify Straight Arrow Connector 19 is purged, while Straight Arrow Connector 24 in Col 1 remains
    const finalSlideObj = ppt.slideManager.getSlideObj(9)
    const finalSpTree = finalSlideObj['p:sld']['p:cSld']['p:spTree']
    const finalCxns = Array.isArray(finalSpTree['p:cxnSp'])
      ? finalSpTree['p:cxnSp']
      : finalSpTree['p:cxnSp']
        ? [finalSpTree['p:cxnSp']]
        : []
    const finalCxnNames = finalCxns.map(c => c['p:nvCxnSpPr']?.['p:cNvPr']?.['@_name'])

    expect(finalCxnNames).not.toContain('Straight Arrow Connector 19')
    expect(finalCxnNames).toContain('Straight Arrow Connector 24')

    // Verify Rectangle 21 shifted left by 1 column width (2103120 EMUs)
    const finalSps = Array.isArray(finalSpTree['p:sp'])
      ? finalSpTree['p:sp']
      : [finalSpTree['p:sp']]
    const rect21Final = finalSps.find(
      s => s['p:nvSpPr']?.['p:cNvPr']?.['@_name'] === 'Rectangle 21'
    )
    expect(rect21Final).toBeDefined()
    const rect21FinalX = parseInt(rect21Final['p:spPr']['a:xfrm']['a:off']['@_x'], 10)
    expect(rect21FinalX).toBe(rect21InitialX - 2103120)
    expect(rect21FinalX).toBe(6120619)
  })

  it('TASK-ARROW-REMOVAL-FIX-002: handles polymorphic arrow cell shapes (inside, outside, shape-only) across row and column removals', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    // Add table with multiple arrow preset aliases
    ppt.addTable(
      1,
      3,
      3,
      [
        ['Header 1', 'Header 2', 'Header 3'],
        [
          {
            shape: { type: 'rightArrow', fill: '#3B82F6' },
            text: 'Inside Arrow',
            textPosition: 'inside',
          },
          {
            shape: { type: 'arrow', fill: '#EF4444' },
            text: 'Outside Arrow',
            textPosition: 'outside',
          },
          {
            shape: { type: 'right-arrow', fill: '#10B981' },
          },
        ],
        ['Bottom Left', { shape: { type: 'upArrow', fill: '#8B5CF6' } }, 'Bottom Right'],
      ],
      { id: 'ArrowTable', x: 100, y: 100, width: 600, height: 300 }
    )

    const slideObj = ppt.slideManager.getSlideObj(1)
    const spTree = slideObj['p:sld']['p:cSld']['p:spTree']
    const sps = Array.isArray(spTree['p:sp']) ? spTree['p:sp'] : [spTree['p:sp']]

    // Verify all 4 cell shapes exist initially
    const shapeNames = sps.map(s => s['p:nvSpPr']?.['p:cNvPr']?.['@_name'] || '')
    const row1Shapes = shapeNames.filter(n => n.includes('ArrowTable_r1'))
    const row2Shapes = shapeNames.filter(n => n.includes('ArrowTable_r2'))

    expect(row1Shapes.length).toBe(3)
    expect(row2Shapes.length).toBe(1)

    // Remove row 1 containing 3 arrow shapes
    ppt.removeTableRow('ArrowTable', 1)

    const afterRowSpTree = ppt.slideManager.getSlideObj(1)['p:sld']['p:cSld']['p:spTree']
    const afterRowSps = Array.isArray(afterRowSpTree['p:sp'])
      ? afterRowSpTree['p:sp']
      : [afterRowSpTree['p:sp']]
    const afterRowNames = afterRowSps.map(s => s['p:nvSpPr']?.['p:cNvPr']?.['@_name'] || '')

    // Assert original row 1 arrow shapes at c0 and c2 (which had no row 2 shape) are purged
    expect(afterRowNames.some(n => n.includes('ArrowTable_r1_c0'))).toBe(false)
    expect(afterRowNames.some(n => n.includes('ArrowTable_r1_c2'))).toBe(false)

    // Assert row 2 arrow at c1 shifted up to row 1 at c1
    expect(afterRowNames.some(n => n.includes('ArrowTable_r1_c1'))).toBe(true)
    expect(afterRowNames.some(n => n.includes('ArrowTable_r2_c1'))).toBe(false)

    // Remove table completely and ensure zero orphaned shapes remain
    ppt.removeTable('ArrowTable')
    const finalSpTree = ppt.slideManager.getSlideObj(1)['p:sld']['p:cSld']['p:spTree']
    const finalSps = Array.isArray(finalSpTree['p:sp'])
      ? finalSpTree['p:sp']
      : finalSpTree['p:sp']
        ? [finalSpTree['p:sp']]
        : []
    const remainingTableShapes = finalSps.filter(s =>
      (s['p:nvSpPr']?.['p:cNvPr']?.['@_name'] || '').includes('ArrowTable')
    )
    expect(remainingTableShapes.length).toBe(0)
  })

  it('TASK-ARROW-STANDALONE-FIX-003: supports removeShape and clearShapes on arrows and connectors with case-insensitive matching', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    // Add standalone arrow and connector
    const _arrowId = ppt.addShape(1, 'rightArrow', {
      name: 'CustomRightArrow',
      x: 100,
      y: 100,
      width: 200,
      height: 60,
      fill: '#0284C7',
    })

    const _connectorId = ppt.addShape(1, 'line', {
      name: 'CustomConnectorLine',
      x: 100,
      y: 200,
      width: 300,
      height: 0,
      border: { color: '#DC2626', width: 2 },
      headEnd: 'triangle',
    })

    // Verify lookup by exact name and case-insensitive name
    const foundArrow = ppt.getShape('CustomRightArrow')
    expect(foundArrow).toBeDefined()
    expect(foundArrow.name).toBe('CustomRightArrow')

    const foundArrowCaseInsensitive = ppt.getShape('customrightarrow')
    expect(foundArrowCaseInsensitive).toBeDefined()
    expect(foundArrowCaseInsensitive.name).toBe('CustomRightArrow')

    const foundConnector = ppt.getShape('CustomConnectorLine')
    expect(foundConnector).toBeDefined()
    expect(foundConnector.name).toBe('CustomConnectorLine')

    // Remove arrow using case-insensitive identifier
    ppt.removeShape('customrightarrow')
    expect(ppt.getShape('CustomRightArrow')).toBeNull()

    // Remove connector using exact identifier
    ppt.removeShape('CustomConnectorLine')
    expect(ppt.getShape('CustomConnectorLine')).toBeNull()

    // Save and validate OpenXML package
    const buffer = await ppt.toBuffer()
    const reloaded = await PPTXTemplater.load(buffer)
    const validation = reloaded.validate()
    expect(validation.valid).toBe(true)
    expect(validation.errors).toHaveLength(0)
  })
})
