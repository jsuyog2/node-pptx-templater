/**
 * @fileoverview Integration tests for TASK-TABLE-002: Table Lifecycle, Query & Cloning Engine.
 * Tests addTable, updateTable, updateTableData, removeTable, getTables, getTable, and cloneTable.
 */

import { describe, it, expect, beforeEach } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('TASK-TABLE-002: Table Lifecycle, Query & Cloning Engine', () => {
  let ppt

  beforeEach(async () => {
    ppt = await PPTXTemplater.createBlank()
  })

  it('should create a table with addTable, support percentage coordinates, and return this', async () => {
    const data = [
      ['Header 1', 'Header 2', 'Header 3'],
      ['Row 1 Col 1', 'Row 1 Col 2', 'Row 1 Col 3'],
      ['Row 2 Col 1', 'Row 2 Col 2', 'Row 2 Col 3'],
    ]

    const result = ppt.addTable(1, 3, 3, data, {
      id: 'MainTable',
      name: 'MainTable',
      x: '10%',
      y: '15%',
      width: '80%',
      height: '50%',
      firstRow: true,
      bandRow: true,
    })

    expect(result).toBe(ppt) // Fluent chaining

    const tables = ppt.getTables(1)
    expect(tables).toHaveLength(1)
    expect(tables[0].name).toBe('MainTable')
    expect(tables[0].rows).toBe(3)
    expect(tables[0].columns).toBe(3)
    expect(tables[0].x).toBeGreaterThan(0)
    expect(tables[0].y).toBeGreaterThan(0)
    expect(tables[0].width).toBeGreaterThan(0)
    expect(tables[0].height).toBeGreaterThan(0)
  })

  it('should retrieve a table via getTable and getTables', async () => {
    ppt.addTable(
      1,
      2,
      2,
      [
        ['A', 'B'],
        ['C', 'D'],
      ],
      {
        id: 'AlphaTable',
        name: 'AlphaTable',
      }
    )

    const found = ppt.getTable('AlphaTable')
    expect(found).not.toBeNull()
    expect(found.name).toBe('AlphaTable')
    expect(found.rows).toBe(2)
    expect(found.columns).toBe(2)
    expect(found.data).toEqual([
      ['A', 'B'],
      ['C', 'D'],
    ])

    const allTables = ppt.getTables()
    expect(allTables.some(t => t.name === 'AlphaTable')).toBe(true)

    const missing = ppt.getTable('NonExistentTable')
    expect(missing).toBeNull()
  })

  it('should update table geometry, position, and properties with updateTable', async () => {
    ppt.addTable(
      1,
      2,
      2,
      [
        ['A', 'B'],
        ['C', 'D'],
      ],
      {
        id: 'GeometryTable',
        name: 'GeometryTable',
        x: 100,
        y: 100,
        width: 400,
        height: 200,
      }
    )

    const chained = ppt.updateTable('GeometryTable', {
      x: 150,
      y: 200,
      width: 500,
      height: 250,
      bandCol: true,
    })

    expect(chained).toBe(ppt)

    const updated = ppt.getTable('GeometryTable')
    expect(updated).not.toBeNull()
    expect(updated.properties.bandCol).toBe(true)
  })

  it('should update matrix data with updateTableData while preserving table dimensions', async () => {
    ppt.addTable(
      1,
      2,
      2,
      [
        ['Old1', 'Old2'],
        ['Old3', 'Old4'],
      ],
      {
        id: 'DataUpdateTable',
        name: 'DataUpdateTable',
      }
    )

    const newData = [
      ['New1', 'New2'],
      ['New3', 'New4'],
    ]

    const chained = ppt.updateTableData('DataUpdateTable', newData)
    expect(chained).toBe(ppt)

    const updated = ppt.getTable('DataUpdateTable')
    expect(updated.data).toEqual(newData)
    expect(ppt.getTableData('DataUpdateTable')).toEqual(newData)
  })

  it('should deep clone a table with cloneTable assigning new ID and overrides', async () => {
    ppt.addTable(
      1,
      2,
      2,
      [
        ['Orig1', 'Orig2'],
        ['Orig3', 'Orig4'],
      ],
      {
        id: 'SourceTable',
        name: 'SourceTable',
        x: 100,
        y: 100,
        width: 400,
        height: 200,
      }
    )

    const chained = ppt.cloneTable('SourceTable', 'ClonedTable', {
      x: 200,
      y: 350,
    })

    expect(chained).toBe(ppt)

    const source = ppt.getTable('SourceTable')
    const cloned = ppt.getTable('ClonedTable')

    expect(source).not.toBeNull()
    expect(cloned).not.toBeNull()
    expect(cloned.name).toBe('ClonedTable')
    expect(cloned.data).toEqual(source.data)

    const tables = ppt.getTables(1)
    expect(tables).toHaveLength(2)
  })

  it('should remove a table with removeTable cleanly without corrupting slide', async () => {
    ppt.addTable(
      1,
      2,
      2,
      [
        ['R1', 'R2'],
        ['R3', 'R4'],
      ],
      {
        id: 'ToPurgeTable',
        name: 'ToPurgeTable',
      }
    )

    expect(ppt.getTable('ToPurgeTable')).not.toBeNull()

    const chained = ppt.removeTable('ToPurgeTable')
    expect(chained).toBe(ppt)

    expect(ppt.getTable('ToPurgeTable')).toBeNull()
    expect(ppt.getTables(1)).toHaveLength(0)

    expect(() => {
      ppt.removeTable('ToPurgeTable')
    }).toThrow(/TableNotFoundError|ToPurgeTable/)
  })

  it('should export valid, repair-free PPTX package with created, cloned, and updated tables', async () => {
    ppt.addTable(
      1,
      3,
      2,
      [
        ['Col1', 'Col2'],
        ['A', 'B'],
        ['C', 'D'],
      ],
      {
        id: 'ExportTable',
        name: 'ExportTable',
        x: '10%',
        y: '10%',
        width: '80%',
        height: '40%',
      }
    )

    ppt.cloneTable('ExportTable', 'ExportTableClone', {
      y: '55%',
    })

    ppt.updateTableData('ExportTableClone', [
      ['Col1', 'Col2'],
      ['Cloned A', 'Cloned B'],
      ['Cloned C', 'Cloned D'],
    ])

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeInstanceOf(Buffer)
    expect(buffer.length).toBeGreaterThan(1000)

    // Load back and verify structural integrity
    const loaded = await PPTXTemplater.load(buffer)
    const tables = loaded.getTables(1)
    expect(tables).toHaveLength(2)
    expect(tables.map(t => t.name).sort()).toEqual(['ExportTable', 'ExportTableClone'])

    const validation = loaded.validate()
    expect(validation.valid).toBe(true)
    expect(validation.errors).toHaveLength(0)
  })
})
