/**
 * @fileoverview Integration tests for PowerPoint table cell merging and unmerging.
 */

import { describe, it, expect } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'
import JSZip from 'jszip'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURES_DIR = resolve(__dirname, '../fixtures')

describe('PPTXTemplater - Table Cell Merging & Unmerging', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')

  if (!existsSync(fixtureFile)) {
    it.skip('Fixture file sample.pptx not found — skipping merge tests.', () => {})
    return
  }

  it('should support merging and unmerging cells correctly', async () => {
    // 1. Load fresh templater
    const ppt = await PPTXTemplater.load(fixtureFile)
    ppt.useSlide(3) // Slide 3 contains our table
    const tableId = 'Table'

    // Verify initial state: no merged cells
    let merges = ppt.getMergedCells(tableId)
    expect(merges).toHaveLength(0)

    // 2. Horizontal Merge
    ppt.mergeCells({
      slide: 3,
      tableId,
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 2,
    })

    merges = ppt.getMergedCells(tableId)
    expect(merges).toHaveLength(1)
    expect(merges[0]).toEqual({ startRow: 0, startCol: 0, endRow: 0, endCol: 2 })

    // Verify cell status helpers
    expect(ppt.isMergedCell(tableId, 0, 1)).toBe(true)
    expect(ppt.getMergeParent(tableId, 0, 1)).toEqual({ row: 0, col: 0 })
    expect(ppt.getMergeRegion(tableId, 0, 1)).toEqual({
      startRow: 0,
      startCol: 0,
      endRow: 0,
      endCol: 2,
    })

    // 3. Unmerge Cell-Based
    ppt.unmergeCells({
      slide: 3,
      tableId,
      row: 0,
      col: 1,
    })

    merges = ppt.getMergedCells(tableId)
    expect(merges).toHaveLength(0)

    // 4. Vertical Merge
    ppt.mergeCells({
      slide: 3,
      tableId,
      startRow: 0,
      startCol: 1,
      endRow: 1,
      endCol: 1,
    })

    merges = ppt.getMergedCells(tableId)
    expect(merges).toHaveLength(1)
    expect(merges[0]).toEqual({ startRow: 0, startCol: 1, endRow: 1, endCol: 1 })

    // 5. Block Merge (Rectangular)
    // First split vertical
    ppt.unmergeCells({
      slide: 3,
      tableId,
      row: 0,
      col: 1,
    })

    // Add two rows to test larger rectangular block merges
    ppt.addTableRow(tableId, ['R3C1', 'R3C2', 'R3C3', 'R3C4', 'R3C5'])
    ppt.addTableRow(tableId, ['R4C1', 'R4C2', 'R4C3', 'R4C4', 'R4C5'])

    // Merge block (1, 1) to (3, 3)
    ppt.mergeCells({
      slide: 3,
      tableId,
      startRow: 1,
      startCol: 1,
      endRow: 3,
      endCol: 3,
    })

    merges = ppt.getMergedCells(tableId)
    expect(merges).toHaveLength(1)
    expect(merges[0]).toEqual({ startRow: 1, startCol: 1, endRow: 3, endCol: 3 })

    // 6. Overlap Validation
    expect(() => {
      ppt.mergeCells({
        slide: 3,
        tableId,
        startRow: 2,
        startCol: 2,
        endRow: 2,
        endCol: 4,
      })
    }).toThrow()

    // 7. Template-driven merging
    const freshPpt = await PPTXTemplater.load(fixtureFile)
    freshPpt.useSlide(3)

    // Update table with rowSpan/colSpan objects and merge configuration arrays
    freshPpt.updateTable(tableId, {
      rows: [
        ['Header 1', 'Header 2', 'Header 3', 'Header 4', 'Header 5'],
        ['Cell 1', { value: 'Cell Span H', colSpan: 2 }, 'Cell 4', 'Cell 5'],
        ['Cell A', 'Cell B', { value: 'Cell Span V', rowSpan: 2 }, 'Cell D', 'Cell E'],
        ['Cell X', 'Cell Y', '', 'Cell Z', 'Cell W'],
      ],
      merge: [{ startRow: 0, startCol: 0, endRow: 0, endCol: 2 }],
    })

    // Merges expected:
    // - (0,0) to (0,2) via merge array
    // - (1,1) to (1,2) via colSpan: 2
    // - (2,2) to (3,2) via rowSpan: 2
    merges = freshPpt.getMergedCells(tableId)
    expect(merges).toHaveLength(3)
    expect(merges).toContainEqual({ startRow: 0, startCol: 0, endRow: 0, endCol: 2 })
    expect(merges).toContainEqual({ startRow: 1, startCol: 1, endRow: 1, endCol: 2 })
    expect(merges).toContainEqual({ startRow: 2, startCol: 2, endRow: 3, endCol: 2 })

    // Verify formatting/structure is valid
    const report = await freshPpt.validatePresentation()
    expect(report.valid).toBe(true)
  })

  it('should correctly order a:r and a:endParaRPr when split cells inherit from a merged template cell, and support cell formatting', async () => {
    const templatePath = resolve(__dirname, '../../templates/sample.pptx')
    const ppt = await PPTXTemplater.load(templatePath)
    ppt.useSlide(3)

    ppt.updateTable('Table2', [
      ['Name', '', 'Role', 'Dept'],
      ['Alice', '', 'Engineer', 'Platform'],
      ['Bob', { value: 'Alice', align: 'r', fontSize: 14, fill: 'FF0000' }, 'Designer', 'Product'],
    ])

    ppt.mergeCells('Table2', 0, 0, 0, 1)
    ppt.mergeCells('Table2', 1, 0, 1, 1)

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const slideXml = await zip.file('ppt/slides/slide3.xml').async('text')

    const aliceColIndex = slideXml.indexOf('<a:t>Alice</a:t>')
    expect(aliceColIndex).toBeGreaterThan(-1)

    const endParaRPrIndex = slideXml.indexOf('<a:endParaRPr', aliceColIndex - 200)
    expect(endParaRPrIndex).toBeGreaterThan(aliceColIndex)

    // Verify cell styling applied from options
    expect(slideXml).toContain('algn="r"')
    expect(slideXml).toContain('sz="1400"')
    expect(slideXml).toContain('val="FF0000"')
  })
})
