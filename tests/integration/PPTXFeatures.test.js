/**
 * @fileoverview Integration tests for extended PPTX APIs (Table, Chart, Text, Shape, Image, Slide, and Validation).
 */

import { describe, it, expect, beforeAll } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURES_DIR = resolve(__dirname, '../fixtures')

describe('PPTXTemplater - Extended Features & Safety', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')

  if (!existsSync(fixtureFile)) {
    it.skip('Fixture file sample.pptx not found — skipping extended feature tests.', () => {})
    return
  }

  let ppt

  beforeAll(async () => {
    ppt = await PPTXTemplater.load(fixtureFile)
  })

  describe('Table Operations & Safety', () => {
    it('should successfully update table content and check validation', async () => {
      // Slide 3 contains our table
      ppt.useSlide(3)

      const tableId = 'Table' // Name or ID from sample.pptx

      ppt.updateTable(tableId, [
        ['Header 1', 'Header 2', 'Header 3', 'Header 4', 'Header 5'],
        ['Val 1', 'Val 2', 'Val 3', 'Val 4', 'Val 5'],
        ['Val 6', 'Val 7', 'Val 8', 'Val 9', 'Val 10'],
      ])

      const report = await ppt.validatePresentation()
      expect(report.valid).toBe(true)
      expect(report.errors.length).toBe(0)
    })

    it('should insert, add, clone, and remove rows successfully', async () => {
      ppt.useSlide(3)
      const tableId = 'Table'

      // Verify row count before
      let tables = ppt.getTables()
      const t = tables.find(tbl => tbl.name === tableId || tbl.id === '5')
      const startRows = t ? t.rows : 0

      ppt.addTableRow(tableId, ['Added 1', 'Added 2', 'Added 3', 'Added 4', 'Added 5'])
      ppt.insertTableRow(tableId, 1, [
        'Inserted 1',
        'Inserted 2',
        'Inserted 3',
        'Inserted 4',
        'Inserted 5',
      ])
      ppt.cloneTableRow(tableId, 0, 2)
      ppt.removeTableRow(tableId, 3)

      tables = ppt.getTables()
      const t2 = tables.find(tbl => tbl.name === tableId || tbl.id === '5')
      expect(t2.rows).toBe(startRows + 2) // added 1, inserted 1, cloned 1, removed 1 => +2 net

      const report = await ppt.validatePresentation()
      expect(report.valid).toBe(true)
    })

    it('should merge, unmerge, update cells, autofit, and resize successfully', async () => {
      ppt.useSlide(3)
      const tableId = 'Table'

      ppt.updateCell(tableId, 0, 0, 'New Title', { fill: '00FF00', fontSize: 16, align: 'ctr' })
      ppt.mergeCells(tableId, 1, 0, 2, 1)
      ppt.unmergeCells(tableId, 1, 0, 2, 1)
      ppt.autoFitTable(tableId)
      ppt.resizeTable(tableId, 8.5, 4.2)

      const report = await ppt.validatePresentation()
      expect(report.valid).toBe(true)
    })
  })

  describe('Shape Operations', async () => {
    it('should update, clone, delete, and list shapes', async () => {
      ppt.useSlide(1)

      const shapes = ppt.getShapes()
      expect(shapes.length).toBeGreaterThan(0)

      const firstShapeId = shapes[0].name

      ppt.updateShapeText(firstShapeId, 'Updated Shape Header')
      ppt.cloneShape(firstShapeId, 'ClonedShapeCopy', { offsetX: 1.0, offsetY: 0.5 })
      ppt.deleteShape('ClonedShapeCopy')

      const report = await ppt.validatePresentation()
      expect(report.valid).toBe(true)
    })
  })

  describe('Text Search and Manipulation', () => {
    it('should replace text by tag and multiple tags, and list text elements', async () => {
      const freshPpt = await PPTXTemplater.load(fixtureFile)
      freshPpt.useSlide(1)

      freshPpt.replaceTextByTag('title', 'Awesome Presentation Title')
      freshPpt.replaceMultiple({
        title: 'Awesome Title Tag',
        company: 'Google DeepMind',
      })

      const elements = freshPpt.getTextElements()
      expect(elements.length).toBeGreaterThan(0)

      const matches = freshPpt.findText('Awesome')
      expect(matches.length).toBeGreaterThan(0)

      const report = await freshPpt.validatePresentation()
      expect(report.valid).toBe(true)
    })
  })

  describe('Slide Operations', () => {
    it('should duplicate, move, insert, and delete slides', async () => {
      const startCount = ppt.slideCount

      ppt.duplicateSlide(1, 2)
      expect(ppt.slideCount).toBe(startCount + 1)

      ppt.moveSlide(ppt.slideCount, 1)
      ppt.insertSlide(2, { title: 'New Layout Slide' })
      expect(ppt.slideCount).toBe(startCount + 2) // 1 duplicate + 1 insert

      ppt.deleteSlide(1)
      expect(ppt.slideCount).toBe(startCount + 1)

      const report = await ppt.validatePresentation()
      expect(report.valid).toBe(true)
    })
  })

  describe('Chart Custom Series and Properties', () => {
    it('should update title, categories, and replace series on charts', async () => {
      ppt.useSlide(2)

      const charts = ppt.getCharts()
      if (charts.length > 0) {
        const chartId = 'chart1' // Assuming chart1
        ppt.updateChartTitle(chartId, 'New Sales Summary')
        ppt.updateChartCategories(chartId, ['Q1', 'Q2', 'Q3', 'Q4'])
        ppt.replaceChartSeries(chartId, 0, { name: 'Direct Sales', values: [150, 180, 200, 250] })

        const report = await ppt.validatePresentation()
        expect(report.valid).toBe(true)
      }
    })
  })
})
