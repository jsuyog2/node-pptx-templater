/**
 * @fileoverview Integration tests for TASK-TABLE-005: Complex Cell Merging & Splitting Subsystem.
 * Tests mergeCells and splitCell with ISO/IEC 29500 DrawingML compliance.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('TASK-TABLE-005: Complex Cell Merging & Splitting Subsystem', () => {
  let ppt

  beforeEach(async () => {
    ppt = await PPTXTemplater.createBlank()
    // Create a 4x4 matrix
    ppt.addTable(
      1,
      4,
      4,
      [
        ['A1', 'B1', 'C1', 'D1'],
        ['A2', 'B2', 'C2', 'D2'],
        ['A3', 'B3', 'C3', 'D3'],
        ['A4', 'B4', 'C4', 'D4'],
      ],
      { id: 'MergeTable', name: 'MergeTable' }
    )
  })

  it('should merge cells horizontally with mergeCells and return this', () => {
    const res = ppt.mergeCells('MergeTable', 0, 0, 0, 2)
    expect(res).toBe(ppt) // Fluent chaining

    const anchor = ppt.getCell('MergeTable', 0, 0)
    expect(anchor).not.toBeNull()
    expect(anchor.gridSpan).toBe(3)
    expect(anchor.rowSpan).toBe(1)
    expect(anchor.isMerged).toBe(true)

    const cell2 = ppt.getCell('MergeTable', 0, 1)
    expect(cell2.hMerge).toBe(true)
    expect(cell2.isMerged).toBe(true)

    const cell3 = ppt.getCell('MergeTable', 0, 2)
    expect(cell3.hMerge).toBe(true)
    expect(cell3.isMerged).toBe(true)
  })

  it('should merge cells vertically with mergeCells', () => {
    ppt.mergeCells('MergeTable', 1, 0, 3, 0)

    const anchor = ppt.getCell('MergeTable', 1, 0)
    expect(anchor.rowSpan).toBe(3)
    expect(anchor.gridSpan).toBe(1)
    expect(anchor.isMerged).toBe(true)

    const cell2 = ppt.getCell('MergeTable', 2, 0)
    expect(cell2.vMerge).toBe(true)
    expect(cell2.isMerged).toBe(true)

    const cell3 = ppt.getCell('MergeTable', 3, 0)
    expect(cell3.vMerge).toBe(true)
    expect(cell3.isMerged).toBe(true)
  })

  it('should merge a 2D block with both horizontal and vertical spans', () => {
    ppt.mergeCells('MergeTable', 1, 1, 2, 2)

    const anchor = ppt.getCell('MergeTable', 1, 1)
    expect(anchor.rowSpan).toBe(2)
    expect(anchor.gridSpan).toBe(2)
    expect(anchor.isMerged).toBe(true)

    const hNeighbor = ppt.getCell('MergeTable', 1, 2)
    expect(hNeighbor.hMerge).toBe(true)

    const vNeighbor = ppt.getCell('MergeTable', 2, 1)
    expect(vNeighbor.vMerge).toBe(true)

    const diagNeighbor = ppt.getCell('MergeTable', 2, 2)
    expect(diagNeighbor.vMerge).toBe(true)
    expect(diagNeighbor.hMerge).toBe(true)
  })

  it('should split merged cells back into independent cells with splitCell and return this', () => {
    ppt.mergeCells('MergeTable', 0, 0, 0, 2)

    const res = ppt.splitCell('MergeTable', 0, 0)
    expect(res).toBe(ppt) // Fluent chaining

    const c1 = ppt.getCell('MergeTable', 0, 0)
    const c2 = ppt.getCell('MergeTable', 0, 1)
    const c3 = ppt.getCell('MergeTable', 0, 2)

    expect(c1.gridSpan).toBe(1)
    expect(c1.rowSpan).toBe(1)
    expect(c1.isMerged).toBe(false)

    expect(c2.hMerge).toBe(false)
    expect(c2.isMerged).toBe(false)

    expect(c3.hMerge).toBe(false)
    expect(c3.isMerged).toBe(false)
  })

  it('should split merged region when splitCell is invoked on a secondary continuation cell', () => {
    ppt.mergeCells('MergeTable', 1, 1, 2, 2)

    // Call splitCell on secondary cell (2, 2)
    ppt.splitCell('MergeTable', 2, 2)

    const anchor = ppt.getCell('MergeTable', 1, 1)
    const diag = ppt.getCell('MergeTable', 2, 2)

    expect(anchor.rowSpan).toBe(1)
    expect(anchor.gridSpan).toBe(1)
    expect(anchor.isMerged).toBe(false)

    expect(diag.vMerge).toBe(false)
    expect(diag.hMerge).toBe(false)
    expect(diag.isMerged).toBe(false)
  })

  it('should export valid, repair-free PPTX package with merged cells', async () => {
    ppt.mergeCells('MergeTable', 0, 0, 0, 3)
    ppt.updateTableCell('MergeTable', 0, 0, 'Quarterly Summary Header')

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeInstanceOf(Buffer)

    const loaded = await PPTXTemplater.load(buffer)
    const anchor = loaded.getCell('MergeTable', 0, 0)
    expect(anchor.gridSpan).toBe(4)
    expect(anchor.text).toBe('Quarterly Summary Header')

    const validation = loaded.validate()
    expect(validation.valid).toBe(true)
    expect(validation.errors).toHaveLength(0)
  })
})
