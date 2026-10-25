/**
 * @fileoverview Integration tests for TASK-TABLE-006: Comprehensive Multi-Tier Styling Engine.
 * Tests styleTable, styleColumn, styleRow, and styleCell with ISO/IEC 29500 DrawingML compliance.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('TASK-TABLE-006: Comprehensive Multi-Tier Styling Engine', () => {
  let ppt

  beforeEach(async () => {
    ppt = await PPTXTemplater.createBlank()
    ppt.addTable(
      1,
      4,
      3,
      [
        ['Header 1', 'Header 2', 'Header 3'],
        ['Row 1 Col 1', 'Row 1 Col 2', 'Row 1 Col 3'],
        ['Row 2 Col 1', 'Row 2 Col 2', 'Row 2 Col 3'],
        ['Row 3 Col 1', 'Row 3 Col 2', 'Row 3 Col 3'],
      ],
      { id: 'StyleTable', name: 'StyleTable' }
    )
  })

  it('should apply table-wide styles with styleTable and return this', () => {
    const res = ppt.styleTable('StyleTable', {
      firstRow: true,
      bandRow: true,
      headerFill: '1E3A8A',
      evenFill: 'F1F5F9',
      oddFill: 'FFFFFF',
      border: { color: 'CBD5E1', width: 1 },
      fontSize: 12,
    })
    expect(res).toBe(ppt) // Fluent chaining

    const meta = ppt.getTable('StyleTable')
    expect(meta.properties.firstRow).toBe(true)
    expect(meta.properties.bandRow).toBe(true)

    // Header cell fill
    const hCell = ppt.getCell('StyleTable', 0, 0)
    expect(hCell.fill).toBe('1E3A8A')

    // Data rows
    const r1Cell = ppt.getCell('StyleTable', 1, 0)
    expect(r1Cell.fill).toBe('FFFFFF') // Row 1 is odd index

    const r2Cell = ppt.getCell('StyleTable', 2, 0)
    expect(r2Cell.fill).toBe('F1F5F9') // Row 2 is even index
  })

  it('should apply styles down a column with styleColumn and return this', () => {
    const res = ppt.styleColumn('StyleTable', 1, {
      fill: 'E0E7FF',
      color: '4338CA',
      bold: true,
      align: 'ctr',
    })
    expect(res).toBe(ppt)

    const c0 = ppt.getCell('StyleTable', 0, 1)
    const c1 = ppt.getCell('StyleTable', 1, 1)
    const c2 = ppt.getCell('StyleTable', 2, 1)

    expect(c0.fill).toBe('E0E7FF')
    expect(c1.fill).toBe('E0E7FF')
    expect(c2.fill).toBe('E0E7FF')
  })

  it('should apply styles across a row with styleRow and return this', () => {
    const res = ppt.styleRow('StyleTable', 0, {
      fill: '0F172A',
      color: 'F8FAFC',
      bold: true,
      height: 40,
    })
    expect(res).toBe(ppt)

    const c0 = ppt.getCell('StyleTable', 0, 0)
    const c1 = ppt.getCell('StyleTable', 0, 1)
    const c2 = ppt.getCell('StyleTable', 0, 2)

    expect(c0.fill).toBe('0F172A')
    expect(c1.fill).toBe('0F172A')
    expect(c2.fill).toBe('0F172A')

    const rowInfo = ppt.getRow('StyleTable', 0)
    expect(rowInfo.height).toBeGreaterThan(0)
  })

  it('should apply targeted cell styles with styleCell and return this', () => {
    const res = ppt.styleCell('StyleTable', 2, 2, {
      fill: '10B981',
      verticalAlign: 'ctr',
      align: 'r',
      border: { color: '047857', width: 2 },
      margins: 6,
    })
    expect(res).toBe(ppt)

    const cell = ppt.getCell('StyleTable', 2, 2)
    expect(cell.fill).toBe('10B981')
    expect(cell.anchor).toBe('ctr')
  })

  it('should export valid, repair-free PPTX package with multi-tier table styling', async () => {
    ppt
      .styleTable('StyleTable', { headerFill: '1E293B', oddFill: 'F8FAFC' })
      .styleRow('StyleTable', 0, { color: 'FFFFFF', bold: true })
      .styleColumn('StyleTable', 2, { align: 'r' })
      .styleCell('StyleTable', 3, 2, { fill: '22C55E', color: 'FFFFFF' })

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeInstanceOf(Buffer)

    const loaded = await PPTXTemplater.load(buffer)
    const cell = loaded.getCell('StyleTable', 3, 2)
    expect(cell.fill).toBe('22C55E')

    const validation = loaded.validate()
    expect(validation.valid).toBe(true)
    expect(validation.errors).toHaveLength(0)
  })
})
