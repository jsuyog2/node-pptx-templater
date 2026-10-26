/**
 * @fileoverview TableManager - Strict 30-Method DrawingML Tables Subsystem.
 * ISO/IEC 29500 (ECMA-376) PresentationML & DrawingML compliant.
 * Pure JavaScript, zero native dependencies, zero repair warnings.
 */

'use strict'

const { createLogger } = require('../utils/logger.js')
const { TableNotFoundError, PPTXError } = require('../utils/errors.js')
const { toEmu, emuToPx } = require('../spatial/units.js')
const { resolveCoordinate, getCanvasDimensions } = require('../utils/coordinateUtils.js')
const { StyleBuilder } = require('../parser/StyleBuilder.js')
const { CellShapeManager } = require('./CellShapeManager.js')
const { Z_ORDER_SYMBOL } = require('../parsers/XMLParser.js')

const logger = createLogger('TableManager')

const DEFAULT_TABLE_STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'
const ROW_ID_URI = '{D42A27DB-BD31-4B8C-A354-60D93AC583D0}'
const ROW_ID_NS = 'http://schemas.microsoft.com/office/drawing/2014/main'

/**
 * Escapes raw strings for XML text nodes.
 * @param {*} str
 * @returns {string}
 */
function escapeXml(str) {
  if (str === null || str === undefined) return ''
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

/**
 * Generates an integer in range [1, 2147483647] for a16:rowId.
 * @returns {number}
 */
function generateRowId() {
  return Math.floor(Math.random() * 2147483640) + 1
}

/**
 * TableManager - DrawingML Table Engine adhering strictly to the 30-method whitelist boundary.
 */
class TableManager {
  /** @private @type {XMLParser} */
  #xmlParser

  /** @private @type {CellShapeManager} */
  #cellShapeManager

  /** @private @type {ShapeManager|null} */
  #shapeManager

  /**
   * @param {XMLParser} xmlParser
   * @param {ShapeManager} [shapeManager]
   */
  constructor(xmlParser, shapeManager = null) {
    this.#xmlParser = xmlParser
    this.#shapeManager = shapeManager
    this.#cellShapeManager = new CellShapeManager(xmlParser, shapeManager)
  }

  /**
   * Sets or updates the ShapeManager instance.
   * @param {ShapeManager} shapeManager
   */
  setShapeManager(shapeManager) {
    this.#shapeManager = shapeManager
    this.#cellShapeManager.setShapeManager(shapeManager)
  }

  /**
   * Returns the underlying CellShapeManager.
   * @returns {CellShapeManager}
   */
  getCellShapeManager() {
    return this.#cellShapeManager
  }

  // ==========================================================================
  // 1. Lifecycle & Creation
  // ==========================================================================

  /**
   * 1. addTable(slideId, rows, columns, data, options)
   * Constructs a valid <p:graphicFrame> enclosing <a:tbl> with <a:tblPr>, <a:tblGrid>, and initial <a:tr> elements.
   *
   * @param {number|string} slideId - Slide index or ID.
   * @param {number} rows - Number of rows.
   * @param {number} columns - Number of columns.
   * @param {Array<Array<*>>} [data] - Initial 2D table data.
   * @param {Object} [options] - Geometry, styling, and positioning options.
   * @param {SlideManager} slideManager - SlideManager instance.
   * @returns {string} Created table ID/name.
   */
  addTable(slideId, rows, columns, data = [], options = {}, slideManager) {
    const slideIndex = slideManager.resolveSlideIndex
      ? slideManager.resolveSlideIndex(slideId)
      : typeof slideId === 'number'
        ? slideId
        : 1
    const slideObj = slideManager.getSlideObj(slideIndex)
    if (!slideObj) {
      throw new PPTXError(`Slide not found for index/ID: ${slideId}`)
    }

    const cSld = slideObj['p:sld']?.['p:cSld'] || slideObj['p:sldLayout']?.['p:cSld']
    const spTree = cSld?.['p:spTree']
    if (!spTree) {
      throw new PPTXError(`Invalid slide structure: missing p:spTree on slide ${slideIndex}`)
    }

    const canvas = getCanvasDimensions(slideIndex, slideManager)
    const opt = typeof options === 'object' && options !== null ? options : {}
    if ((!data || data.length === 0) && Array.isArray(opt.data)) {
      data = opt.data
    }

    // Resolve coordinates (supports percentages, pt, px, in, emu)
    const x = opt.x !== undefined ? resolveCoordinate(opt.x, canvas.width, opt) : 1524000
    const y = opt.y !== undefined ? resolveCoordinate(opt.y, canvas.height, opt) : 1397000
    const width =
      opt.width !== undefined ? resolveCoordinate(opt.width, canvas.width, opt) : 6096000
    const height =
      opt.height !== undefined ? resolveCoordinate(opt.height, canvas.height, opt) : 1828800

    const numRows = Math.max(1, Number(rows) || (Array.isArray(data) ? data.length : 2))
    const numCols = Math.max(
      1,
      Number(columns) || (Array.isArray(data) && data[0] ? data[0].length : 2)
    )

    const shapeId = this.#generateUniqueShapeId(spTree)
    const tableId = opt.id || opt.name || `Table ${shapeId}`

    // Calculate grid column widths
    const gridCols = []
    const defaultColW = Math.round(width / numCols)
    for (let c = 0; c < numCols; c++) {
      let colW = defaultColW
      if (Array.isArray(opt.colWidths) && opt.colWidths[c] !== undefined) {
        colW = toEmu(opt.colWidths[c])
      }
      gridCols.push({ '@_w': String(colW) })
    }

    // Build row elements
    const defaultRowH = Math.round(height / numRows)
    const trList = []
    const deferredShapes = []

    for (let r = 0; r < numRows; r++) {
      let rowH = defaultRowH
      if (Array.isArray(opt.rowHeights) && opt.rowHeights[r] !== undefined) {
        rowH = toEmu(opt.rowHeights[r])
      }

      const tcList = []
      for (let c = 0; c < numCols; c++) {
        let raw = undefined
        if (Array.isArray(data) && data[r] && data[r][c] !== undefined) {
          raw = data[r][c]
        }

        const parsed = this.#cellShapeManager.parseCellData(raw)
        const cellText =
          parsed.kind === 'shape'
            ? parsed.shapeConfig?.textPosition === 'outside'
              ? parsed.text
              : ''
            : parsed.text
        const tcNode = this.#createCellNode(cellText, parsed.style)

        if (parsed.kind === 'shape') {
          deferredShapes.push({ r, c, shapeConfig: parsed.shapeConfig, tcNode })
        }

        tcList.push(tcNode)
      }

      const rowNode = {
        '@_h': String(rowH),
        'a:tc': tcList,
        'a:extLst': {
          'a:ext': {
            '@_uri': ROW_ID_URI,
            'a16:rowId': {
              '@_xmlns:a16': ROW_ID_NS,
              '@_val': String(generateRowId()),
            },
          },
        },
      }
      trList.push(rowNode)
    }

    // Table Properties <a:tblPr>
    const tblPr = StyleBuilder.buildTableProperties(
      {},
      {
        firstRow: opt.firstRow !== undefined ? opt.firstRow : true,
        bandRow: opt.bandRow !== undefined ? opt.bandRow : true,
        firstCol: opt.firstCol || false,
        bandCol: opt.bandCol || false,
        lastRow: opt.lastRow || false,
        lastCol: opt.lastCol || false,
        styleId: opt.styleId || DEFAULT_TABLE_STYLE_ID,
      }
    )

    // Construct full <p:graphicFrame>
    const graphicFrame = {
      'p:nvGraphicFramePr': {
        'p:cNvPr': {
          '@_id': String(shapeId),
          '@_name': tableId,
        },
        'p:cNvGraphicFramePr': {
          'a:graphicFrameLocks': { '@_noGrp': '1' },
        },
        'p:nvPr': {},
      },
      'p:xfrm': {
        'a:off': { '@_x': String(x), '@_y': String(y) },
        'a:ext': { '@_cx': String(width), '@_cy': String(height) },
      },
      'a:graphic': {
        'a:graphicData': {
          '@_uri': 'http://schemas.openxmlformats.org/drawingml/2006/table',
          'a:tbl': {
            'a:tblPr': tblPr,
            'a:tblGrid': {
              'a:gridCol': gridCols,
            },
            'a:tr': trList,
          },
        },
      },
    }

    if (!Array.isArray(spTree['p:graphicFrame'])) {
      if (spTree['p:graphicFrame']) {
        spTree['p:graphicFrame'] = [spTree['p:graphicFrame']]
      } else {
        spTree['p:graphicFrame'] = []
      }
    }
    spTree['p:graphicFrame'].push(graphicFrame)

    if (spTree[Z_ORDER_SYMBOL] && !spTree[Z_ORDER_SYMBOL].includes(String(shapeId))) {
      spTree[Z_ORDER_SYMBOL].push(String(shapeId))
    }

    // Apply any deferred cell shape overlays
    const tblObj = graphicFrame['a:graphic']['a:graphicData']['a:tbl']
    for (const item of deferredShapes) {
      this.#cellShapeManager.applyCellShape(
        slideIndex,
        tableId,
        item.r,
        item.c,
        item.shapeConfig,
        slideManager,
        graphicFrame,
        tblObj,
        item.tcNode,
        this.#shapeManager
      )
    }

    this.#markDirty(slideIndex, slideManager)
    logger.debug(`Created table "${tableId}" on slide ${slideIndex}`)
    return tableId
  }

  /**
   * 2. updateTable(tableId, options)
   * Updates table geometry, bounds, table properties, and styles.
   *
   * @param {string} tableId - Table identifier.
   * @param {Object} options - Updates for position, dimensions, properties, colWidths, rowHeights.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  updateTable(tableId, options = {}, slideIndex, slideManager) {
    if (Array.isArray(options)) {
      return this.updateTableData(tableId, options, slideIndex, slideManager)
    }

    const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )
    const canvas = getCanvasDimensions(resolvedSlideIndex, slideManager)

    const xfrm = frame['p:xfrm'] || (frame['p:xfrm'] = {})
    const off = xfrm['a:off'] || (xfrm['a:off'] = { '@_x': '0', '@_y': '0' })
    const ext = xfrm['a:ext'] || (xfrm['a:ext'] = { '@_cx': '0', '@_cy': '0' })

    if (options.x !== undefined) {
      off['@_x'] = String(resolveCoordinate(options.x, canvas.width, options))
    }
    if (options.y !== undefined) {
      off['@_y'] = String(resolveCoordinate(options.y, canvas.height, options))
    }
    if (options.width !== undefined) {
      ext['@_cx'] = String(resolveCoordinate(options.width, canvas.width, options))
    }
    if (options.height !== undefined) {
      ext['@_cy'] = String(resolveCoordinate(options.height, canvas.height, options))
    }

    // Table Properties
    if (tblObj['a:tblPr']) {
      tblObj['a:tblPr'] = StyleBuilder.buildTableProperties(tblObj['a:tblPr'], options)
    }

    // Column widths
    if (Array.isArray(options.colWidths)) {
      const grid = tblObj['a:tblGrid']?.['a:gridCol'] || []
      for (let i = 0; i < options.colWidths.length && i < grid.length; i++) {
        if (options.colWidths[i] !== undefined) {
          grid[i]['@_w'] = String(toEmu(options.colWidths[i]))
        }
      }
    }

    // Row heights
    if (Array.isArray(options.rowHeights)) {
      const trs = tblObj['a:tr'] || []
      for (let i = 0; i < options.rowHeights.length && i < trs.length; i++) {
        if (options.rowHeights[i] !== undefined) {
          trs[i]['@_h'] = String(toEmu(options.rowHeights[i]))
        }
      }
    }

    if (Array.isArray(options.data)) {
      this.updateTableData(tableId, options.data, resolvedSlideIndex, slideManager)
    }

    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * 3. updateTableData(tableId, data)
   * Replaces matrix cell data while strictly preserving existing column widths, cell styling, and row heights.
   *
   * @param {string} tableId - Table identifier.
   * @param {Array<Array<*>>} data - 2D matrix of cell values.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  updateTableData(tableId, data = [], slideIndex, slideManager) {
    const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )
    if (!Array.isArray(data) || data.length === 0) return this

    let trs = tblObj['a:tr']
    if (!Array.isArray(trs)) {
      trs = trs ? [trs] : []
      tblObj['a:tr'] = trs
    }

    const templateRow = trs[0] || null

    for (let r = 0; r < data.length; r++) {
      let tr = trs[r]
      if (!tr) {
        if (!templateRow) break
        tr = this.#xmlParser.deepClone(templateRow)
        this.#updateRowId(tr)
        trs.push(tr)
      }

      let tcs = tr['a:tc']
      if (!Array.isArray(tcs)) {
        tcs = tcs ? [tcs] : []
        tr['a:tc'] = tcs
      }

      const rowValues = Array.isArray(data[r]) ? data[r] : [data[r]]
      for (let c = 0; c < rowValues.length && c < tcs.length; c++) {
        const raw = rowValues[c]
        const parsed = this.#cellShapeManager.parseCellData(raw)

        if (parsed.kind === 'shape') {
          this.#cellShapeManager.applyCellShape(
            resolvedSlideIndex,
            tableId,
            r,
            c,
            parsed.shapeConfig,
            slideManager,
            frame,
            tblObj,
            tcs[c],
            this.#shapeManager
          )
        } else {
          this.#cellShapeManager.removeCellShape(resolvedSlideIndex, tableId, r, c, slideManager)
          if (tcs[c]['a:tcPr']) {
            delete tcs[c]['a:tcPr']['@_marL']
            delete tcs[c]['a:tcPr']['@_marR']
          }
          this.#setCellText(tcs[c], parsed.text, parsed.style)
          if (parsed.style && Object.keys(parsed.style).length > 0) {
            tcs[c]['a:tcPr'] = StyleBuilder.buildCellProperties(tcs[c]['a:tcPr'], parsed.style)
          }
        }
      }
    }

    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * 4. removeTable(tableId)
   * Purges the table <p:graphicFrame> from <p:spTree> cleanly without orphaned relationships.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  removeTable(tableId, slideIndex, slideManager) {
    const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )
    const slideObj = slideManager.getSlideObj(resolvedSlideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree']

    const targetFrameId = String(frame?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_id'] || '')

    // Strict geometric bounding-box containment: purge shapes strictly interior to the table
    const tableBounds = this.#cellShapeManager.getTableBounds(frame, tblObj)
    if (tableBounds) {
      this.#cellShapeManager.purgeShapesWithinBounds(
        resolvedSlideIndex,
        tableBounds,
        slideManager,
        tableId
      )
    }

    if (spTree?.['p:graphicFrame']) {
      if (Array.isArray(spTree['p:graphicFrame'])) {
        const idx = spTree['p:graphicFrame'].indexOf(frame)
        if (idx !== -1) {
          spTree['p:graphicFrame'].splice(idx, 1)
        }
      } else if (spTree['p:graphicFrame'] === frame) {
        delete spTree['p:graphicFrame']
      }
    }

    if (spTree?.[Z_ORDER_SYMBOL] && targetFrameId) {
      const zIdx = spTree[Z_ORDER_SYMBOL].indexOf(targetFrameId)
      if (zIdx !== -1) {
        spTree[Z_ORDER_SYMBOL].splice(zIdx, 1)
      }
    }

    this.#markDirty(resolvedSlideIndex, slideManager)
    logger.debug(`Removed table "${tableId}" from slide ${resolvedSlideIndex}`)
    return this
  }

  /**
   * 5. getTables(slideId)
   * Parses DrawingML nodes and returns normalized metadata for tables on the slide.
   *
   * @param {number|string} [slideId] - Slide identifier.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {Array<Object>} List of table descriptor objects.
   */
  getTables(slideId, slideManager) {
    const targetIndices =
      slideId !== undefined
        ? [
            slideManager.resolveSlideIndex
              ? slideManager.resolveSlideIndex(slideId)
              : Number(slideId),
          ]
        : slideManager.getAllSlideIndices()

    const results = []
    for (const sIdx of targetIndices) {
      const slideObj = slideManager.getSlideObj(sIdx)
      if (!slideObj) continue

      const spTree =
        slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
        slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
      if (!spTree) continue

      const frames = Array.isArray(spTree['p:graphicFrame'])
        ? spTree['p:graphicFrame']
        : spTree['p:graphicFrame']
          ? [spTree['p:graphicFrame']]
          : []

      for (const frame of frames) {
        const tbl = frame['a:graphic']?.['a:graphicData']?.['a:tbl']
        if (!tbl) continue

        const cNvPr = frame['p:nvGraphicFramePr']?.['p:cNvPr'] || {}
        const xfrm = frame['p:xfrm'] || {}
        const off = xfrm['a:off'] || {}
        const ext = xfrm['a:ext'] || {}

        const trs = Array.isArray(tbl['a:tr']) ? tbl['a:tr'] : tbl['a:tr'] ? [tbl['a:tr']] : []
        const cols = Array.isArray(tbl['a:tblGrid']?.['a:gridCol'])
          ? tbl['a:tblGrid']['a:gridCol']
          : tbl['a:tblGrid']?.['a:gridCol']
            ? [tbl['a:tblGrid']['a:gridCol']]
            : []

        const xVal = parseInt(off['@_x'] || '0', 10)
        const yVal = parseInt(off['@_y'] || '0', 10)
        const cxVal = parseInt(ext['@_cx'] || '0', 10)
        const cyVal = parseInt(ext['@_cy'] || '0', 10)

        results.push({
          id: String(cNvPr['@_id'] || ''),
          name: String(cNvPr['@_name'] || ''),
          slideIndex: sIdx,
          rows: trs.length,
          columns: cols.length,
          cols: cols.length,
          x: xVal,
          y: yVal,
          width: cxVal,
          height: cyVal,
          position: {
            x: xVal,
            y: yVal,
            width: cxVal,
            height: cyVal,
          },
          properties: {
            firstRow: tbl['a:tblPr']?.['@_firstRow'] === '1',
            bandRow: tbl['a:tblPr']?.['@_bandRow'] === '1',
            firstCol: tbl['a:tblPr']?.['@_firstCol'] === '1',
            bandCol: tbl['a:tblPr']?.['@_bandCol'] === '1',
            styleId: tbl['a:tblPr']?.['a:tableStyleId'] || null,
          },
          firstRow: tbl['a:tblPr']?.['@_firstRow'] === '1',
          bandRow: tbl['a:tblPr']?.['@_bandRow'] === '1',
          firstCol: tbl['a:tblPr']?.['@_firstCol'] === '1',
          bandCol: tbl['a:tblPr']?.['@_bandCol'] === '1',
          styleId: tbl['a:tblPr']?.['a:tableStyleId'] || null,
        })
      }
    }

    return results
  }

  /**
   * 6. getTable(tableId)
   * Queries and returns normalized table metadata, row/col counts, and position.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {Object|null} Normalized table metadata or null.
   */
  getTable(tableId, slideIndex, slideManager) {
    try {
      const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
        slideIndex,
        tableId,
        slideManager
      )
      const cNvPr = frame['p:nvGraphicFramePr']?.['p:cNvPr'] || {}
      const xfrm = frame['p:xfrm'] || {}
      const trs = Array.isArray(tblObj['a:tr'])
        ? tblObj['a:tr']
        : tblObj['a:tr']
          ? [tblObj['a:tr']]
          : []
      const cols = Array.isArray(tblObj['a:tblGrid']?.['a:gridCol'])
        ? tblObj['a:tblGrid']['a:gridCol']
        : tblObj['a:tblGrid']?.['a:gridCol']
          ? [tblObj['a:tblGrid']['a:gridCol']]
          : []

      const xVal = parseInt(xfrm['a:off']?.['@_x'] || '0', 10)
      const yVal = parseInt(xfrm['a:off']?.['@_y'] || '0', 10)
      const cxVal = parseInt(xfrm['a:ext']?.['@_cx'] || '0', 10)
      const cyVal = parseInt(xfrm['a:ext']?.['@_cy'] || '0', 10)
      const dataMatrix = this.getTableData(tableId, resolvedSlideIndex, slideManager)

      return {
        id: String(cNvPr['@_id'] || ''),
        name: String(cNvPr['@_name'] || ''),
        slideIndex: resolvedSlideIndex,
        rows: trs.length,
        columns: cols.length,
        cols: cols.length,
        x: xVal,
        y: yVal,
        width: cxVal,
        height: cyVal,
        position: {
          x: xVal,
          y: yVal,
          width: cxVal,
          height: cyVal,
        },
        properties: {
          firstRow: tblObj['a:tblPr']?.['@_firstRow'] === '1',
          bandRow: tblObj['a:tblPr']?.['@_bandRow'] === '1',
          firstCol: tblObj['a:tblPr']?.['@_firstCol'] === '1',
          bandCol: tblObj['a:tblPr']?.['@_bandCol'] === '1',
          styleId: tblObj['a:tblPr']?.['a:tableStyleId'] || null,
        },
        firstRow: tblObj['a:tblPr']?.['@_firstRow'] === '1',
        bandRow: tblObj['a:tblPr']?.['@_bandRow'] === '1',
        firstCol: tblObj['a:tblPr']?.['@_firstCol'] === '1',
        bandCol: tblObj['a:tblPr']?.['@_bandCol'] === '1',
        styleId: tblObj['a:tblPr']?.['a:tableStyleId'] || null,
        data: dataMatrix,
        tableData: dataMatrix,
      }
    } catch {
      return null
    }
  }

  /**
   * 7. cloneTable(tableId, clonedTableId, clonedOptions)
   * Deep-clones an existing table graphic frame, assigns unique shape ID and row IDs, and applies overrides.
   *
   * @param {string} tableId - Source table name or shape ID.
   * @param {string} clonedTableId - Cloned table identifier.
   * @param {Object} [clonedOptions] - Position and dimension overrides.
   * @param {number} [slideIndex] - Target slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  cloneTable(tableId, clonedTableId, clonedOptions = {}, slideIndex, slideManager) {
    const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )
    const slideObj = slideManager.getSlideObj(resolvedSlideIndex)
    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']

    const clonedFrame = this.#xmlParser.deepClone(frame)
    const newShapeId = this.#generateUniqueShapeId(spTree)
    const newName = clonedTableId || `Table ${newShapeId}`

    if (clonedFrame['p:nvGraphicFramePr']?.['p:cNvPr']) {
      clonedFrame['p:nvGraphicFramePr']['p:cNvPr']['@_id'] = String(newShapeId)
      clonedFrame['p:nvGraphicFramePr']['p:cNvPr']['@_name'] = newName
    }

    // Refresh row IDs for all rows in cloned table
    const clonedTbl = clonedFrame['a:graphic']?.['a:graphicData']?.['a:tbl']
    if (clonedTbl) {
      const trs = Array.isArray(clonedTbl['a:tr'])
        ? clonedTbl['a:tr']
        : clonedTbl['a:tr']
          ? [clonedTbl['a:tr']]
          : []
      for (const tr of trs) {
        this.#updateRowId(tr)
      }
    }

    // Apply coordinate overrides if provided
    const canvas = getCanvasDimensions(resolvedSlideIndex, slideManager)
    const xfrm = clonedFrame['p:xfrm'] || (clonedFrame['p:xfrm'] = {})
    const off = xfrm['a:off'] || (xfrm['a:off'] = { '@_x': '0', '@_y': '0' })
    const ext = xfrm['a:ext'] || (xfrm['a:ext'] = { '@_cx': '0', '@_cy': '0' })

    if (clonedOptions.x !== undefined) {
      off['@_x'] = String(resolveCoordinate(clonedOptions.x, canvas.width, clonedOptions))
    }
    if (clonedOptions.y !== undefined) {
      off['@_y'] = String(resolveCoordinate(clonedOptions.y, canvas.height, clonedOptions))
    }
    if (clonedOptions.width !== undefined) {
      ext['@_cx'] = String(resolveCoordinate(clonedOptions.width, canvas.width, clonedOptions))
    }
    if (clonedOptions.height !== undefined) {
      ext['@_cy'] = String(resolveCoordinate(clonedOptions.height, canvas.height, clonedOptions))
    }

    if (!Array.isArray(spTree['p:graphicFrame'])) {
      spTree['p:graphicFrame'] = [spTree['p:graphicFrame']]
    }
    spTree['p:graphicFrame'].push(clonedFrame)

    if (spTree[Z_ORDER_SYMBOL] && !spTree[Z_ORDER_SYMBOL].includes(String(newShapeId))) {
      spTree[Z_ORDER_SYMBOL].push(String(newShapeId))
    }

    const origOff = frame['p:xfrm']?.['a:off'] || {}
    const origX = parseInt(origOff['@_x'] || '0', 10)
    const origY = parseInt(origOff['@_y'] || '0', 10)
    const newX = parseInt(off['@_x'] || '0', 10)
    const newY = parseInt(off['@_y'] || '0', 10)

    this.#cellShapeManager.cloneCellShapes(
      resolvedSlideIndex,
      tableId,
      resolvedSlideIndex,
      newName,
      slideManager,
      newX - origX,
      newY - origY,
      frame,
      tblObj
    )

    this.#markDirty(resolvedSlideIndex, slideManager)
    logger.debug(`Cloned table "${tableId}" to "${newName}" on slide ${resolvedSlideIndex}`)
    return this
  }

  // ==========================================================================
  // 2. Row & Column Structural Operations
  // ==========================================================================

  /**
   * 8. addTableRow(tableId, rowIndex, data)
   * Inserts or appends a row at rowIndex with cell data, generating unique <a16:rowId>.
   *
   * @param {string} tableId - Table identifier.
   * @param {number|Array} rowIndex - 0-based insertion row index, or row data array if appending.
   * @param {Array<*>} [data] - Row cell data.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  addTableRow(tableId, rowIndex, data, slideIndex, slideManager) {
    const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )
    let trs = tblObj['a:tr']
    if (!Array.isArray(trs)) {
      trs = trs ? [trs] : []
      tblObj['a:tr'] = trs
    }

    let insertIdx = trs.length
    let rowValues = []

    if (typeof rowIndex === 'number') {
      insertIdx = Math.max(0, Math.min(rowIndex, trs.length))
      rowValues = Array.isArray(data) ? data : []
    } else if (Array.isArray(rowIndex)) {
      rowValues = rowIndex
    } else if (Array.isArray(data)) {
      rowValues = data
    }

    const templateRow = trs[Math.min(insertIdx, trs.length - 1)] || trs[0]
    let newRow
    const deferredShapes = []

    if (templateRow) {
      newRow = this.#xmlParser.deepClone(templateRow)
      this.#updateRowId(newRow)
      const tcs = Array.isArray(newRow['a:tc'])
        ? newRow['a:tc']
        : newRow['a:tc']
          ? [newRow['a:tc']]
          : []

      for (let c = 0; c < tcs.length; c++) {
        delete tcs[c]['@_rowSpan']
        delete tcs[c]['@_gridSpan']
        delete tcs[c]['@_hMerge']
        delete tcs[c]['@_vMerge']
        if (tcs[c]['a:tcPr']) {
          delete tcs[c]['a:tcPr']['@_marL']
          delete tcs[c]['a:tcPr']['@_marR']
        }
        const raw = rowValues[c] !== undefined ? rowValues[c] : ''
        const parsed = this.#cellShapeManager.parseCellData(raw)
        if (parsed.kind === 'shape') {
          if (parsed.shapeConfig?.textPosition === 'outside' && parsed.text) {
            this.#setCellText(tcs[c], parsed.text, parsed.shapeConfig.textStyle || parsed.style)
          } else {
            this.#setCellText(tcs[c], '')
          }
          deferredShapes.push({ c, shapeConfig: parsed.shapeConfig, tcNode: tcs[c] })
        } else {
          this.#setCellText(tcs[c], parsed.text, parsed.style)
          if (parsed.style && Object.keys(parsed.style).length > 0) {
            tcs[c]['a:tcPr'] = StyleBuilder.buildCellProperties(tcs[c]['a:tcPr'], parsed.style)
          }
        }
      }
    } else {
      const cols = tblObj['a:tblGrid']?.['a:gridCol'] || []
      const colCount = Array.isArray(cols) ? cols.length : 1
      const tcList = []
      for (let c = 0; c < colCount; c++) {
        const raw = rowValues[c] ?? ''
        const parsed = this.#cellShapeManager.parseCellData(raw)
        const cellText =
          parsed.kind === 'shape'
            ? parsed.shapeConfig?.textPosition === 'outside'
              ? parsed.text
              : ''
            : parsed.text
        const cellNode = this.#createCellNode(cellText, parsed.style)
        if (parsed.kind === 'shape') {
          deferredShapes.push({ c, shapeConfig: parsed.shapeConfig, tcNode: cellNode })
        }
        tcList.push(cellNode)
      }
      newRow = {
        '@_h': '370840',
        'a:tc': tcList,
        'a:extLst': {
          'a:ext': {
            '@_uri': ROW_ID_URI,
            'a16:rowId': {
              '@_xmlns:a16': ROW_ID_NS,
              '@_val': String(generateRowId()),
            },
          },
        },
      }
    }

    const insertedRowH = parseInt(newRow['@_h'] || '370840', 10)

    // Shift existing cell shapes at or below insertIdx downward BEFORE splicing
    this.#cellShapeManager.shiftRowShapes(
      resolvedSlideIndex,
      tableId,
      insertIdx,
      insertedRowH,
      slideManager
    )

    trs.splice(insertIdx, 0, newRow)

    for (const item of deferredShapes) {
      this.#cellShapeManager.applyCellShape(
        resolvedSlideIndex,
        tableId,
        insertIdx,
        item.c,
        item.shapeConfig,
        slideManager,
        frame,
        tblObj,
        item.tcNode,
        this.#shapeManager
      )
    }

    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * 9. addTableColumn(tableId, columnIndex, data)
   * Inserts or appends a column in <a:tblGrid> and creates corresponding <a:tc> cells across all rows.
   *
   * @param {string} tableId - Table identifier.
   * @param {number|Array} columnIndex - 0-based column index, or column data array if appending.
   * @param {Array<*>} [data] - Column data down each row.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  addTableColumn(tableId, columnIndex, data, slideIndex, slideManager) {
    const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )
    let gridCols = tblObj['a:tblGrid']?.['a:gridCol']
    if (!Array.isArray(gridCols)) {
      gridCols = gridCols ? [gridCols] : []
      if (!tblObj['a:tblGrid']) tblObj['a:tblGrid'] = {}
      tblObj['a:tblGrid']['a:gridCol'] = gridCols
    }

    let insertIdx = gridCols.length
    let colValues = []

    if (typeof columnIndex === 'number') {
      insertIdx = Math.max(0, Math.min(columnIndex, gridCols.length))
      colValues = Array.isArray(data) ? data : []
    } else if (Array.isArray(columnIndex)) {
      colValues = columnIndex
    } else if (Array.isArray(data)) {
      colValues = data
    }

    const defaultW = gridCols.length > 0 ? gridCols[0]['@_w'] || '2032000' : '2032000'
    const insertedColW = parseInt(defaultW, 10)

    // Shift existing cell shapes at or right of insertIdx rightward BEFORE splicing
    this.#cellShapeManager.shiftColShapes(
      resolvedSlideIndex,
      tableId,
      insertIdx,
      insertedColW,
      slideManager
    )

    gridCols.splice(insertIdx, 0, { '@_w': defaultW })

    const trs = Array.isArray(tblObj['a:tr'])
      ? tblObj['a:tr']
      : tblObj['a:tr']
        ? [tblObj['a:tr']]
        : []

    const deferredShapes = []

    for (let r = 0; r < trs.length; r++) {
      let tcs = trs[r]['a:tc']
      if (!Array.isArray(tcs)) {
        tcs = tcs ? [tcs] : []
        trs[r]['a:tc'] = tcs
      }
      const raw = colValues[r] !== undefined ? colValues[r] : ''
      const parsed = this.#cellShapeManager.parseCellData(raw)
      const templateCell = tcs[Math.min(insertIdx, tcs.length - 1)] || tcs[0]
      let newCell

      if (templateCell) {
        newCell = this.#xmlParser.deepClone(templateCell)
        delete newCell['@_rowSpan']
        delete newCell['@_gridSpan']
        delete newCell['@_hMerge']
        delete newCell['@_vMerge']
        if (newCell['a:tcPr']) {
          delete newCell['a:tcPr']['@_marL']
          delete newCell['a:tcPr']['@_marR']
        }
        if (parsed.kind === 'shape') {
          if (parsed.shapeConfig?.textPosition === 'outside' && parsed.text) {
            this.#setCellText(newCell, parsed.text, parsed.shapeConfig.textStyle || parsed.style)
          } else {
            this.#setCellText(newCell, '')
          }
          deferredShapes.push({ r, shapeConfig: parsed.shapeConfig, tcNode: newCell })
        } else {
          this.#setCellText(newCell, parsed.text, parsed.style)
          if (parsed.style && Object.keys(parsed.style).length > 0) {
            newCell['a:tcPr'] = StyleBuilder.buildCellProperties(newCell['a:tcPr'], parsed.style)
          }
        }
      } else {
        const cellText =
          parsed.kind === 'shape'
            ? parsed.shapeConfig?.textPosition === 'outside'
              ? parsed.text
              : ''
            : parsed.text
        newCell = this.#createCellNode(cellText, parsed.style)
        if (parsed.kind === 'shape') {
          deferredShapes.push({ r, shapeConfig: parsed.shapeConfig, tcNode: newCell })
        }
      }

      tcs.splice(insertIdx, 0, newCell)
    }

    for (const item of deferredShapes) {
      this.#cellShapeManager.applyCellShape(
        resolvedSlideIndex,
        tableId,
        item.r,
        insertIdx,
        item.shapeConfig,
        slideManager,
        frame,
        tblObj,
        item.tcNode,
        this.#shapeManager
      )
    }

    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * 10. removeTableRow(tableId, rowIndex)
   * Purges <a:tr> element at rowIndex.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  removeTableRow(tableId, rowIndex, slideIndex, slideManager) {
    const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )
    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]
    if (rowIndex < 0 || rowIndex >= trs.length) {
      throw new PPTXError(`Row index ${rowIndex} out of bounds (table has ${trs.length} rows)`)
    }

    const removedRowH = parseInt(trs[rowIndex]?.['@_h'] || '370840', 10)

    // Purge shapes strictly inside rowBounds and shift lower interior shapes upward
    this.#cellShapeManager.removeRowShapesAndShift(
      resolvedSlideIndex,
      tableId,
      rowIndex,
      removedRowH,
      slideManager,
      frame,
      tblObj
    )

    trs.splice(rowIndex, 1)

    // Update graphicFrame ext cy to reflect the row deletion
    if (frame?.['p:xfrm']?.['a:ext']?.['@_cy']) {
      const curCy = parseInt(frame['p:xfrm']['a:ext']['@_cy'] || '0', 10)
      if (curCy > removedRowH) {
        frame['p:xfrm']['a:ext']['@_cy'] = String(curCy - removedRowH)
      }
    }

    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * 11. removeTableColumn(tableId, columnIndex)
   * Purges column at columnIndex from <a:tblGrid> and <a:tc> from each row.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} columnIndex - 0-based column index.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  removeTableColumn(tableId, columnIndex, slideIndex, slideManager) {
    const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )
    const gridCols = tblObj['a:tblGrid']?.['a:gridCol']
    if (!Array.isArray(gridCols) || columnIndex < 0 || columnIndex >= gridCols.length) {
      throw new PPTXError(`Column index ${columnIndex} out of bounds`)
    }

    const removedColW = parseInt(gridCols[columnIndex]?.['@_w'] || '2032000', 10)

    // Purge shapes strictly inside colBounds and shift right interior shapes leftward
    this.#cellShapeManager.removeColShapesAndShift(
      resolvedSlideIndex,
      tableId,
      columnIndex,
      removedColW,
      slideManager,
      frame,
      tblObj
    )

    gridCols.splice(columnIndex, 1)

    // Update graphicFrame ext cx to reflect the column deletion
    if (frame?.['p:xfrm']?.['a:ext']?.['@_cx']) {
      const curCx = parseInt(frame['p:xfrm']['a:ext']['@_cx'] || '0', 10)
      if (curCx > removedColW) {
        frame['p:xfrm']['a:ext']['@_cx'] = String(curCx - removedColW)
      }
    }

    const trs = Array.isArray(tblObj['a:tr'])
      ? tblObj['a:tr']
      : tblObj['a:tr']
        ? [tblObj['a:tr']]
        : []
    for (let r = 0; r < trs.length; r++) {
      if (!trs[r] || !trs[r]['a:tc']) continue
      const tcs = Array.isArray(trs[r]['a:tc']) ? trs[r]['a:tc'] : [trs[r]['a:tc']]
      if (columnIndex >= 0 && columnIndex < tcs.length) {
        tcs.splice(columnIndex, 1)
      }
      trs[r]['a:tc'] = tcs
    }

    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * 12. updateRowData(tableId, rowIndex, data)
   * Updates cell text values along target row rowIndex.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {Array<*>} data - Row cell values.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  updateRowData(tableId, rowIndex, data, slideIndex, slideManager) {
    const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )
    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]
    if (rowIndex < 0 || rowIndex >= trs.length) {
      throw new PPTXError(`Row index ${rowIndex} out of bounds`)
    }

    const tcs = Array.isArray(trs[rowIndex]['a:tc'])
      ? trs[rowIndex]['a:tc']
      : [trs[rowIndex]['a:tc']]
    const vals = Array.isArray(data) ? data : [data]

    for (let c = 0; c < vals.length && c < tcs.length; c++) {
      const v = vals[c]
      const parsed = this.#cellShapeManager.parseCellData(v)
      if (parsed.kind === 'shape') {
        this.#cellShapeManager.applyCellShape(
          resolvedSlideIndex,
          tableId,
          rowIndex,
          c,
          parsed.shapeConfig,
          slideManager,
          frame,
          tblObj,
          tcs[c],
          this.#shapeManager
        )
      } else {
        this.#cellShapeManager.removeCellShape(
          resolvedSlideIndex,
          tableId,
          rowIndex,
          c,
          slideManager
        )
        if (tcs[c]['a:tcPr']) {
          delete tcs[c]['a:tcPr']['@_marL']
          delete tcs[c]['a:tcPr']['@_marR']
        }
        this.#setCellText(tcs[c], parsed.text, parsed.style)
        if (parsed.style && Object.keys(parsed.style).length > 0) {
          tcs[c]['a:tcPr'] = StyleBuilder.buildCellProperties(tcs[c]['a:tcPr'], parsed.style)
        }
      }
    }

    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * 13. updateColumnData(tableId, columnIndex, data)
   * Updates cell text values down target column columnIndex.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} columnIndex - 0-based column index.
   * @param {Array<*>} data - Column cell values.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  updateColumnData(tableId, columnIndex, data, slideIndex, slideManager) {
    const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )
    const trs = Array.isArray(tblObj['a:tr'])
      ? tblObj['a:tr']
      : tblObj['a:tr']
        ? [tblObj['a:tr']]
        : []
    const vals = Array.isArray(data) ? data : [data]

    for (let r = 0; r < vals.length && r < trs.length; r++) {
      const tcs = Array.isArray(trs[r]['a:tc']) ? trs[r]['a:tc'] : [trs[r]['a:tc']]
      if (columnIndex >= 0 && columnIndex < tcs.length) {
        const v = vals[r]
        const parsed = this.#cellShapeManager.parseCellData(v)
        if (parsed.kind === 'shape') {
          this.#cellShapeManager.applyCellShape(
            resolvedSlideIndex,
            tableId,
            r,
            columnIndex,
            parsed.shapeConfig,
            slideManager,
            frame,
            tblObj,
            tcs[columnIndex],
            this.#shapeManager
          )
        } else {
          this.#cellShapeManager.removeCellShape(
            resolvedSlideIndex,
            tableId,
            r,
            columnIndex,
            slideManager
          )
          if (tcs[columnIndex]['a:tcPr']) {
            delete tcs[columnIndex]['a:tcPr']['@_marL']
            delete tcs[columnIndex]['a:tcPr']['@_marR']
          }
          this.#setCellText(tcs[columnIndex], parsed.text, parsed.style)
          if (parsed.style && Object.keys(parsed.style).length > 0) {
            tcs[columnIndex]['a:tcPr'] = StyleBuilder.buildCellProperties(
              tcs[columnIndex]['a:tcPr'],
              parsed.style
            )
          }
        }
      }
    }

    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  // ==========================================================================
  // 3. Cell-Level Operations & Clear Handlers
  // ==========================================================================

  /**
   * 14. updateTableCell(tableId, rowIndex, columnIndex, options)
   * Updates cell text, typography, alignments, and background fills.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} columnIndex - 0-based col index.
   * @param {string|Object} options - Cell value or formatting object.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  updateTableCell(tableId, rowIndex, columnIndex, options = {}, slideIndex, slideManager) {
    const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )
    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]
    if (rowIndex < 0 || rowIndex >= trs.length) {
      throw new PPTXError(`Row index ${rowIndex} out of bounds`)
    }

    const tcs = Array.isArray(trs[rowIndex]['a:tc'])
      ? trs[rowIndex]['a:tc']
      : [trs[rowIndex]['a:tc']]
    if (columnIndex < 0 || columnIndex >= tcs.length) {
      throw new PPTXError(`Column index ${columnIndex} out of bounds`)
    }

    const cell = tcs[columnIndex]
    const parsed = this.#cellShapeManager.parseCellData(options)

    if (parsed.kind === 'shape') {
      this.#cellShapeManager.applyCellShape(
        resolvedSlideIndex,
        tableId,
        rowIndex,
        columnIndex,
        parsed.shapeConfig,
        slideManager,
        frame,
        tblObj,
        cell,
        this.#shapeManager
      )
    } else {
      this.#cellShapeManager.removeCellShape(
        resolvedSlideIndex,
        tableId,
        rowIndex,
        columnIndex,
        slideManager
      )
      if (cell['a:tcPr']) {
        delete cell['a:tcPr']['@_marL']
        delete cell['a:tcPr']['@_marR']
      }
      if (
        typeof options === 'string' ||
        typeof options === 'number' ||
        typeof options === 'boolean'
      ) {
        this.#setCellText(cell, options)
      } else if (typeof options === 'object' && options !== null) {
        if (parsed.text !== '' || options.value !== undefined || options.text !== undefined) {
          this.#setCellText(cell, parsed.text, parsed.style)
        } else if (
          Object.keys(parsed.style).some(k =>
            ['bold', 'italic', 'fontSize', 'color', 'fontFace', 'align', 'alignment'].includes(k)
          )
        ) {
          this.#applyTypography(cell, parsed.style)
        }

        // Cell properties (fill, border, margins, vertical alignment)
        cell['a:tcPr'] = StyleBuilder.buildCellProperties(cell['a:tcPr'], parsed.style)
      }
    }

    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * 15. clearRowData(tableId, rowIndex)
   * Purges text content from all cells across rowIndex while preserving formatting.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  clearRowData(tableId, rowIndex, slideIndex, slideManager) {
    const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )
    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]
    if (rowIndex >= 0 && rowIndex < trs.length) {
      const tcs = Array.isArray(trs[rowIndex]['a:tc'])
        ? trs[rowIndex]['a:tc']
        : [trs[rowIndex]['a:tc']]
      for (let c = 0; c < tcs.length; c++) {
        this.#setCellText(tcs[c], '')
        if (tcs[c]['a:tcPr']) {
          delete tcs[c]['a:tcPr']['@_marL']
          delete tcs[c]['a:tcPr']['@_marR']
        }
      }
    }
    // Strict geometric bounding-box containment: purge shapes strictly interior to row
    const rowBounds = this.#cellShapeManager.getTableRowBounds(frame, tblObj, rowIndex)
    if (rowBounds) {
      this.#cellShapeManager.purgeShapesWithinBounds(
        resolvedSlideIndex,
        rowBounds,
        slideManager,
        tableId
      )
    }
    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * 16. clearColumnData(tableId, columnIndex)
   * Purges text content from all cells down columnIndex while preserving formatting.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} columnIndex - 0-based col index.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  clearColumnData(tableId, columnIndex, slideIndex, slideManager) {
    const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )
    const trs = Array.isArray(tblObj['a:tr'])
      ? tblObj['a:tr']
      : tblObj['a:tr']
        ? [tblObj['a:tr']]
        : []
    for (let r = 0; r < trs.length; r++) {
      const tcs = Array.isArray(trs[r]['a:tc']) ? trs[r]['a:tc'] : [trs[r]['a:tc']]
      if (columnIndex >= 0 && columnIndex < tcs.length) {
        this.#setCellText(tcs[columnIndex], '')
        if (tcs[columnIndex]['a:tcPr']) {
          delete tcs[columnIndex]['a:tcPr']['@_marL']
          delete tcs[columnIndex]['a:tcPr']['@_marR']
        }
      }
    }
    // Strict geometric bounding-box containment: purge shapes strictly interior to column
    const colBounds = this.#cellShapeManager.getTableColumnBounds(frame, tblObj, columnIndex)
    if (colBounds) {
      this.#cellShapeManager.purgeShapesWithinBounds(
        resolvedSlideIndex,
        colBounds,
        slideManager,
        tableId
      )
    }
    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * 17. clearCellData(tableId, rowIndex, columnIndex)
   * Purges text content from a single cell while preserving cell properties <a:tcPr>.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} columnIndex - 0-based col index.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  clearCellData(tableId, rowIndex, columnIndex, slideIndex, slideManager) {
    const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )
    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]
    if (rowIndex >= 0 && rowIndex < trs.length) {
      const tcs = Array.isArray(trs[rowIndex]['a:tc'])
        ? trs[rowIndex]['a:tc']
        : [trs[rowIndex]['a:tc']]
      if (columnIndex >= 0 && columnIndex < tcs.length) {
        this.#setCellText(tcs[columnIndex], '')
        if (tcs[columnIndex]['a:tcPr']) {
          delete tcs[columnIndex]['a:tcPr']['@_marL']
          delete tcs[columnIndex]['a:tcPr']['@_marR']
        }
      }
    }
    // Strict geometric bounding-box containment: purge shapes strictly interior to cell
    const cellBounds = this.#cellShapeManager.getTableCellBounds(
      frame,
      tblObj,
      rowIndex,
      columnIndex
    )
    if (cellBounds) {
      this.#cellShapeManager.purgeShapesWithinBounds(
        resolvedSlideIndex,
        cellBounds,
        slideManager,
        tableId
      )
    }
    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * 18. getTableData(tableId)
   * Extracts raw string data from all cells as a 2D array.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {Array<Array<string>>} 2D array of cell values.
   */
  getTableData(tableId, slideIndex, slideManager) {
    const { tblObj, resolvedSlideIndex } = this.#getTableContext(slideIndex, tableId, slideManager)
    const trs = Array.isArray(tblObj['a:tr'])
      ? tblObj['a:tr']
      : tblObj['a:tr']
        ? [tblObj['a:tr']]
        : []

    const matrix = []
    for (let r = 0; r < trs.length; r++) {
      const tr = trs[r]
      const tcs = Array.isArray(tr['a:tc']) ? tr['a:tc'] : tr['a:tc'] ? [tr['a:tc']] : []
      const row = []
      for (let c = 0; c < tcs.length; c++) {
        let text = this.#getCellText(tcs[c])
        if (!text) {
          const shapeText = this.#cellShapeManager.getCellShapeText(
            resolvedSlideIndex,
            tableId,
            r,
            c,
            slideManager
          )
          if (shapeText) text = shapeText
        }
        row.push(text)
      }
      matrix.push(row)
    }
    return matrix
  }

  /**
   * 19. getColumnData(tableId, columnIndex)
   * Returns cell text array down columnIndex.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} columnIndex - 0-based col index.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {Array<string>} Column text values.
   */
  getColumnData(tableId, columnIndex, slideIndex, slideManager) {
    const matrix = this.getTableData(tableId, slideIndex, slideManager)
    return matrix.map(row => row[columnIndex] ?? '')
  }

  /**
   * 20. getRowData(tableId, rowIndex)
   * Returns cell text array along rowIndex.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {Array<string>} Row text values.
   */
  getRowData(tableId, rowIndex, slideIndex, slideManager) {
    const matrix = this.getTableData(tableId, slideIndex, slideManager)
    return matrix[rowIndex] || []
  }

  /**
   * 21. getColumn(tableId, columnIndex)
   * Returns column metadata: index, width, text data, and formatting.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} columnIndex - 0-based col index.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {Object} Column descriptor object.
   */
  getColumn(tableId, columnIndex, slideIndex, slideManager) {
    const { tblObj } = this.#getTableContext(slideIndex, tableId, slideManager)
    const gridCols = tblObj['a:tblGrid']?.['a:gridCol'] || []
    const cols = Array.isArray(gridCols) ? gridCols : [gridCols]
    const colNode = cols[columnIndex] || {}
    const width = parseInt(colNode['@_w'] || '0', 10)

    return {
      index: columnIndex,
      width,
      widthPx: Math.round(emuToPx(width)),
      data: this.getColumnData(tableId, columnIndex, slideIndex, slideManager),
    }
  }

  /**
   * 22. getRow(tableId, rowIndex)
   * Returns row metadata: index, height, text data, and formatting.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {Object} Row descriptor object.
   */
  getRow(tableId, rowIndex, slideIndex, slideManager) {
    const { tblObj } = this.#getTableContext(slideIndex, tableId, slideManager)
    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]
    const rowNode = trs[rowIndex] || {}
    const height = parseInt(rowNode['@_h'] || '0', 10)

    return {
      index: rowIndex,
      height,
      heightPx: Math.round(emuToPx(height)),
      data: this.getRowData(tableId, rowIndex, slideIndex, slideManager),
    }
  }

  /**
   * 23. getCell(tableId, rowIndex, columnIndex)
   * Returns detailed cell metadata, bounds, borders, fill, and merge state.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} columnIndex - 0-based col index.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {Object} Cell descriptor object.
   */
  getCell(tableId, rowIndex, columnIndex, slideIndex, slideManager) {
    const { frame, tblObj, resolvedSlideIndex } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )
    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]
    const tr = trs[rowIndex]
    if (!tr) return null

    const tcs = Array.isArray(tr['a:tc']) ? tr['a:tc'] : [tr['a:tc']]
    const tc = tcs[columnIndex]
    if (!tc) return null

    let text = this.#getCellText(tc)
    const cellShape = this.#cellShapeManager.getCellShape(
      resolvedSlideIndex,
      tableId,
      rowIndex,
      columnIndex,
      slideManager
    )
    if (!text && cellShape?.text) {
      text = cellShape.text
    }

    const tcPr = tc['a:tcPr'] || {}
    const fillClr = tcPr['a:solidFill']?.['a:srgbClr']?.['@_val'] || null

    // Compute cell bounding coordinates
    const bounds = this.#calculateCellBounds(frame, tblObj, rowIndex, columnIndex)

    return {
      row: rowIndex,
      column: columnIndex,
      text,
      fill: fillClr,
      anchor: tcPr['@_anchor'] || null,
      rowSpan: parseInt(tc['@_rowSpan'] || '1', 10),
      gridSpan: parseInt(tc['@_gridSpan'] || '1', 10),
      hMerge: tc['@_hMerge'] === '1',
      vMerge: tc['@_vMerge'] === '1',
      isMerged:
        tc['@_hMerge'] === '1' ||
        tc['@_vMerge'] === '1' ||
        parseInt(tc['@_rowSpan'] || '1', 10) > 1 ||
        parseInt(tc['@_gridSpan'] || '1', 10) > 1,
      bounds,
      shape: cellShape || null,
    }
  }

  /**
   * 24. getCellData(tableId, rowIndex, columnIndex)
   * Extracts raw string text from cell at (rowIndex, columnIndex).
   *
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} columnIndex - 0-based col index.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {string} Cell string value.
   */
  getCellData(tableId, rowIndex, columnIndex, slideIndex, slideManager) {
    const cell = this.getCell(tableId, rowIndex, columnIndex, slideIndex, slideManager)
    return cell ? cell.text : ''
  }

  // ==========================================================================
  // 4. Merging & Splitting
  // ==========================================================================

  /**
   * 25. mergeCells(tableId, startRow, startColumn, endRow, endColumn)
   * Applies rowSpan, gridSpan, hMerge, and vMerge to define a rectangular merged region.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} startRow - Start row index.
   * @param {number} startColumn - Start column index.
   * @param {number} endRow - End row index.
   * @param {number} endColumn - End column index.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  mergeCells(tableId, startRow, startColumn, endRow, endColumn, slideIndex, slideManager) {
    const { tblObj, resolvedSlideIndex } = this.#getTableContext(slideIndex, tableId, slideManager)

    const minR = Math.min(startRow, endRow)
    const maxR = Math.max(startRow, endRow)
    const minC = Math.min(startColumn, endColumn)
    const maxC = Math.max(startColumn, endColumn)

    const rowSpan = maxR - minR + 1
    const gridSpan = maxC - minC + 1

    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]

    for (let r = minR; r <= maxR; r++) {
      if (!trs[r]) continue
      const tcs = Array.isArray(trs[r]['a:tc']) ? trs[r]['a:tc'] : [trs[r]['a:tc']]

      for (let c = minC; c <= maxC; c++) {
        if (!tcs[c]) continue
        const cell = tcs[c]

        if (r === minR && c === minC) {
          // Primary anchor cell
          if (rowSpan > 1) cell['@_rowSpan'] = String(rowSpan)
          if (gridSpan > 1) cell['@_gridSpan'] = String(gridSpan)
          delete cell['@_hMerge']
          delete cell['@_vMerge']
        } else {
          // Secondary covered cells
          delete cell['@_rowSpan']
          delete cell['@_gridSpan']

          if (r === minR) {
            // Horizontal merge on the anchor row
            cell['@_hMerge'] = '1'
            delete cell['@_vMerge']
          } else {
            // Vertical merge continuation
            cell['@_vMerge'] = '1'
            if (c > minC) {
              cell['@_hMerge'] = '1'
            } else {
              delete cell['@_hMerge']
            }
          }

          // Clear text in covered cells to prevent duplicate content rendering
          this.#setCellText(cell, '')
        }
      }
    }

    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * 26. splitCell(tableId, rowIndex, columnIndex)
   * Restores independent cell boundaries by clearing rowSpan, gridSpan, hMerge, and vMerge.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - Row index of any cell in the merged region.
   * @param {number} columnIndex - Col index of any cell in the merged region.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  splitCell(tableId, rowIndex, columnIndex, slideIndex, slideManager) {
    const { tblObj, resolvedSlideIndex } = this.#getTableContext(slideIndex, tableId, slideManager)
    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]

    // Find anchor cell coordinates
    let anchorR = rowIndex
    let anchorC = columnIndex

    // Check if current cell is continuation
    const currentCell = trs[rowIndex]?.['a:tc']?.[columnIndex]
    if (currentCell) {
      if (currentCell['@_vMerge'] === '1' || currentCell['@_hMerge'] === '1') {
        // Trace back to anchor cell
        for (let r = rowIndex; r >= 0; r--) {
          const tc = trs[r]?.['a:tc']?.[columnIndex]
          if (tc && tc['@_vMerge'] !== '1') {
            anchorR = r
            break
          }
        }
        for (let c = columnIndex; c >= 0; c--) {
          const tc = trs[anchorR]?.['a:tc']?.[c]
          if (tc && tc['@_hMerge'] !== '1') {
            anchorC = c
            break
          }
        }
      }
    }

    const anchorCell = trs[anchorR]?.['a:tc']?.[anchorC]
    const rowSpan = anchorCell ? parseInt(anchorCell['@_rowSpan'] || '1', 10) : 1
    const gridSpan = anchorCell ? parseInt(anchorCell['@_gridSpan'] || '1', 10) : 1

    const maxR = anchorR + rowSpan - 1
    const maxC = anchorC + gridSpan - 1

    for (let r = anchorR; r <= maxR && r < trs.length; r++) {
      const tcs = Array.isArray(trs[r]['a:tc']) ? trs[r]['a:tc'] : [trs[r]['a:tc']]
      for (let c = anchorC; c <= maxC && c < tcs.length; c++) {
        delete tcs[c]['@_rowSpan']
        delete tcs[c]['@_gridSpan']
        delete tcs[c]['@_hMerge']
        delete tcs[c]['@_vMerge']
      }
    }

    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  // ==========================================================================
  // 5. Multi-Tier Styling Engine
  // ==========================================================================

  /**
   * 27. styleTable(tableId, options)
   * Applies table-wide properties, alternating banding fills, styles, and borders.
   *
   * @param {string} tableId - Table identifier.
   * @param {Object} options - Table styling options (bandRow, bandCol, fill, border, etc.).
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  styleTable(tableId, options = {}, slideIndex, slideManager) {
    const { tblObj, resolvedSlideIndex } = this.#getTableContext(slideIndex, tableId, slideManager)

    // Table properties
    if (tblObj['a:tblPr']) {
      tblObj['a:tblPr'] = StyleBuilder.buildTableProperties(tblObj['a:tblPr'], options)
    }

    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]

    for (let r = 0; r < trs.length; r++) {
      const isHeader = r === 0
      const isEven = r % 2 === 0
      const tcs = Array.isArray(trs[r]['a:tc']) ? trs[r]['a:tc'] : [trs[r]['a:tc']]

      for (let c = 0; c < tcs.length; c++) {
        const cell = tcs[c]
        let fill = options.fill

        if (isHeader && options.headerFill) {
          fill = options.headerFill
        } else if (!isHeader) {
          if (isEven && options.evenFill) fill = options.evenFill
          else if (!isEven && options.oddFill) fill = options.oddFill
        }

        const cellOpts = { ...options }
        if (fill !== undefined) cellOpts.fill = fill

        cell['a:tcPr'] = StyleBuilder.buildCellProperties(cell['a:tcPr'], cellOpts)
        this.#applyTypography(cell, options)
      }
    }

    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * 28. styleColumn(tableId, columnIndex, options)
   * Applies fills, typography, and borders down columnIndex.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} columnIndex - 0-based col index.
   * @param {Object} options - Column styling options.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  styleColumn(tableId, columnIndex, options = {}, slideIndex, slideManager) {
    const { tblObj, resolvedSlideIndex } = this.#getTableContext(slideIndex, tableId, slideManager)
    const trs = Array.isArray(tblObj['a:tr'])
      ? tblObj['a:tr']
      : tblObj['a:tr']
        ? [tblObj['a:tr']]
        : []

    for (const tr of trs) {
      const tcs = Array.isArray(tr['a:tc']) ? tr['a:tc'] : [tr['a:tc']]
      if (columnIndex >= 0 && columnIndex < tcs.length) {
        const cell = tcs[columnIndex]
        cell['a:tcPr'] = StyleBuilder.buildCellProperties(cell['a:tcPr'], options)
        this.#applyTypography(cell, options)
      }
    }

    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * 29. styleRow(tableId, rowIndex, options)
   * Applies fills, typography, height, and borders across rowIndex.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {Object} options - Row styling options.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  styleRow(tableId, rowIndex, options = {}, slideIndex, slideManager) {
    const { tblObj, resolvedSlideIndex } = this.#getTableContext(slideIndex, tableId, slideManager)
    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]
    if (rowIndex < 0 || rowIndex >= trs.length) {
      throw new PPTXError(`Row index ${rowIndex} out of bounds`)
    }

    const tr = trs[rowIndex]
    if (options.height !== undefined) {
      tr['@_h'] = String(toEmu(options.height))
    }

    const tcs = Array.isArray(tr['a:tc']) ? tr['a:tc'] : [tr['a:tc']]
    for (const cell of tcs) {
      cell['a:tcPr'] = StyleBuilder.buildCellProperties(cell['a:tcPr'], options)
      this.#applyTypography(cell, options)
    }

    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * 30. styleCell(tableId, rowIndex, columnIndex, options)
   * Applies fills, alignments, margins, and borders to cell at (rowIndex, columnIndex).
   *
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} columnIndex - 0-based col index.
   * @param {Object} options - Cell styling options.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager.
   * @returns {TableManager} this
   */
  styleCell(tableId, rowIndex, columnIndex, options = {}, slideIndex, slideManager) {
    const { tblObj, resolvedSlideIndex } = this.#getTableContext(slideIndex, tableId, slideManager)
    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]
    if (rowIndex < 0 || rowIndex >= trs.length) {
      throw new PPTXError(`Row index ${rowIndex} out of bounds`)
    }

    const tcs = Array.isArray(trs[rowIndex]['a:tc'])
      ? trs[rowIndex]['a:tc']
      : [trs[rowIndex]['a:tc']]
    if (columnIndex < 0 || columnIndex >= tcs.length) {
      throw new PPTXError(`Column index ${columnIndex} out of bounds`)
    }

    const cell = tcs[columnIndex]
    cell['a:tcPr'] = StyleBuilder.buildCellProperties(cell['a:tcPr'], options)
    this.#applyTypography(cell, options)

    this.#markDirty(resolvedSlideIndex, slideManager)
    return this
  }

  /**
   * Internal helper for slide inspection and validation.
   * @param {number} slideIndex
   * @param {SlideManager} slideManager
   * @returns {Array<Object>}
   */
  inspectTables(slideIndex, slideManager) {
    return this.getTables(slideIndex, slideManager)
  }

  // ==========================================================================
  // Private Helpers
  // ==========================================================================

  /**
   * Resolves table context across active slide or all slides.
   * @private
   */
  #getTableContext(slideIndex, tableId, slideManager) {
    let resolvedSlideIndex = slideIndex

    if (resolvedSlideIndex === undefined) {
      const allSlides = slideManager.getAllSlideIndices()
      for (const idx of allSlides) {
        const found = slideManager.getSlideTable(idx, tableId)
        if (found) {
          resolvedSlideIndex = idx
          break
        }
      }
      if (resolvedSlideIndex === undefined) {
        resolvedSlideIndex = allSlides[0] || 1
      }
    }

    const tableRes = slideManager.getSlideTable(resolvedSlideIndex, tableId)
    if (!tableRes) {
      throw new TableNotFoundError(tableId)
    }

    return {
      frame: tableRes.frame,
      tblObj: tableRes.table,
      resolvedSlideIndex,
    }
  }

  /**
   * Marks slide dirty and re-indexes tableMap.
   * @private
   */
  #markDirty(slideIndex, slideManager) {
    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Generates a unique numeric shape ID.
   * @private
   */
  #generateUniqueShapeId(container) {
    let maxId = 1
    const scan = obj => {
      if (!obj || typeof obj !== 'object') return
      for (const [k, v] of Object.entries(obj)) {
        if (k === 'p:cNvPr' && v && v['@_id']) {
          const id = parseInt(v['@_id'], 10)
          if (!isNaN(id) && id > maxId) maxId = id
        } else if (typeof v === 'object') {
          scan(v)
        }
      }
    }
    scan(container)
    return maxId + 1
  }

  /**
   * Generates a fresh a16:rowId on row AST node.
   * @private
   */
  #updateRowId(rowObj) {
    if (!rowObj['a:extLst']) {
      rowObj['a:extLst'] = {}
    }
    rowObj['a:extLst']['a:ext'] = {
      '@_uri': ROW_ID_URI,
      'a16:rowId': {
        '@_xmlns:a16': ROW_ID_NS,
        '@_val': String(generateRowId()),
      },
    }
  }

  /**
   * Extracts clean text content from a cell node.
   * @private
   */
  #getCellText(cellObj) {
    if (!cellObj || !cellObj['a:txBody']) return ''
    const paragraphs = Array.isArray(cellObj['a:txBody']['a:p'])
      ? cellObj['a:txBody']['a:p']
      : cellObj['a:txBody']['a:p']
        ? [cellObj['a:txBody']['a:p']]
        : []

    const textParts = []
    for (const p of paragraphs) {
      const runs = Array.isArray(p['a:r']) ? p['a:r'] : p['a:r'] ? [p['a:r']] : []
      for (const r of runs) {
        if (r['a:t'] !== undefined) {
          textParts.push(typeof r['a:t'] === 'object' ? r['a:t']['#text'] || '' : r['a:t'])
        }
      }
    }
    return textParts.join('')
  }

  /**
   * Sets text content in a cell node without destroying existing formatting nodes.
   * @private
   */
  #setCellText(cellObj, text, options = {}) {
    if (!cellObj['a:txBody']) {
      cellObj['a:txBody'] = {
        'a:bodyPr': {},
        'a:lstStyle': {},
        'a:p': {
          'a:r': {
            'a:rPr': {},
            'a:t': escapeXml(text),
          },
        },
      }
      return
    }

    const txBody = cellObj['a:txBody']
    let pList = txBody['a:p']
    if (!Array.isArray(pList)) {
      pList = pList ? [pList] : [{}]
      txBody['a:p'] = pList
    }

    const firstP = pList[0] || (pList[0] = {})
    let rList = firstP['a:r']
    if (!Array.isArray(rList)) {
      rList = rList ? [rList] : [{}]
      firstP['a:r'] = rList
    }

    const firstR = rList[0] || (rList[0] = {})
    firstR['a:t'] = escapeXml(text ?? '')

    // Remove secondary runs in the first paragraph to maintain clean single text run
    if (rList.length > 1) {
      rList.splice(1)
    }

    if (Object.keys(options).length > 0) {
      this.#applyTypography(cellObj, options)
    }
  }

  /**
   * Constructs an initial cell AST node.
   * @private
   */
  #createCellNode(text = '', options = {}) {
    const rPr = StyleBuilder.buildRunProperties({}, options)
    const pPr = StyleBuilder.buildParagraphProperties({}, options)
    const tcPr = StyleBuilder.buildCellProperties({}, options)

    return {
      'a:txBody': {
        'a:bodyPr': {},
        'a:lstStyle': {},
        'a:p': {
          ...(Object.keys(pPr).length > 0 ? { 'a:pPr': pPr } : {}),
          'a:r': {
            ...(Object.keys(rPr).length > 0 ? { 'a:rPr': rPr } : {}),
            'a:t': escapeXml(text),
          },
        },
      },
      'a:tcPr': tcPr,
    }
  }

  /**
   * Applies typography and alignment to cell AST paragraphs and runs.
   * @private
   */
  #applyTypography(cellObj, options = {}) {
    if (!cellObj['a:txBody']) return
    const paragraphs = Array.isArray(cellObj['a:txBody']['a:p'])
      ? cellObj['a:txBody']['a:p']
      : cellObj['a:txBody']['a:p']
        ? [cellObj['a:txBody']['a:p']]
        : []

    for (const p of paragraphs) {
      p['a:pPr'] = StyleBuilder.buildParagraphProperties(p['a:pPr'], options)
      const runs = Array.isArray(p['a:r']) ? p['a:r'] : p['a:r'] ? [p['a:r']] : []
      for (const r of runs) {
        r['a:rPr'] = StyleBuilder.buildRunProperties(r['a:rPr'], options)
      }
    }
  }

  /**
   * Calculates rendered cell bounding coordinates.
   * @private
   */
  #calculateCellBounds(frame, tblObj, targetRow, targetCol) {
    return this.#cellShapeManager.calculateCellBounds(frame, tblObj, targetRow, targetCol)
  }

  /**
   * Calculates the full bounding box of a table in EMUs and pixels.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager instance.
   * @returns {{ x: number, y: number, width: number, height: number, xMin: number, yMin: number, xMax: number, yMax: number, xPx: number, yPx: number, widthPx: number, heightPx: number }}
   */
  getTableBounds(tableId, slideIndex, slideManager) {
    const { frame, tblObj } = this.#getTableContext(slideIndex, tableId, slideManager)
    return this.#cellShapeManager.getTableBounds(frame, tblObj)
  }

  /**
   * Calculates the full horizontal row bounding box of a table row in EMUs and pixels.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager instance.
   * @returns {{ x: number, y: number, width: number, height: number, xMin: number, yMin: number, xMax: number, yMax: number, xPx: number, yPx: number, widthPx: number, heightPx: number }}
   */
  getTableRowBounds(tableId, rowIndex, slideIndex, slideManager) {
    const { frame, tblObj } = this.#getTableContext(slideIndex, tableId, slideManager)
    return this.#cellShapeManager.getTableRowBounds(frame, tblObj, rowIndex)
  }

  /**
   * Calculates the full vertical column bounding box of a table column in EMUs and pixels.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} columnIndex - 0-based column index.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager instance.
   * @returns {{ x: number, y: number, width: number, height: number, xMin: number, yMin: number, xMax: number, yMax: number, xPx: number, yPx: number, widthPx: number, heightPx: number }}
   */
  getTableColumnBounds(tableId, columnIndex, slideIndex, slideManager) {
    const { frame, tblObj } = this.#getTableContext(slideIndex, tableId, slideManager)
    return this.#cellShapeManager.getTableColumnBounds(frame, tblObj, columnIndex)
  }

  /**
   * Calculates the bounding box of a specific cell in EMUs and pixels.
   *
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} columnIndex - 0-based column index.
   * @param {number} [slideIndex] - Slide index.
   * @param {SlideManager} slideManager - SlideManager instance.
   * @returns {{ x: number, y: number, width: number, height: number, xMin: number, yMin: number, xMax: number, yMax: number, xPx: number, yPx: number, widthPx: number, heightPx: number }}
   */
  getTableCellBounds(tableId, rowIndex, columnIndex, slideIndex, slideManager) {
    const { frame, tblObj } = this.#getTableContext(slideIndex, tableId, slideManager)
    return this.#cellShapeManager.getTableCellBounds(frame, tblObj, rowIndex, columnIndex)
  }

  /**
   * Calculates rendered cell bounding coordinates.
   * Supports both (tableId, rowIndex, columnIndex, slideIndex, slideManager) and
   * (slideIndex, tableId, rowIndex, columnIndex, slideManager) signatures.
   *
   * @param {string|number} first
   * @param {string|number} second
   * @param {number} third
   * @param {number} fourth
   * @param {SlideManager} fifth
   * @returns {{ x: number, y: number, width: number, height: number, xMin: number, yMin: number, xMax: number, yMax: number, xPx: number, yPx: number, widthPx: number, heightPx: number }}
   */
  getCellBounds(first, second, third, fourth, fifth) {
    let slideIndex, tableId, rowIndex, columnIndex, slideManager
    if (typeof first === 'number' && typeof second === 'string') {
      slideIndex = first
      tableId = second
      rowIndex = third
      columnIndex = fourth
      slideManager = fifth
    } else {
      tableId = first
      rowIndex = second
      columnIndex = third
      slideIndex = fourth
      slideManager = fifth
    }
    const { frame, tblObj } = this.#getTableContext(slideIndex, tableId, slideManager)
    return this.#cellShapeManager.getTableCellBounds(frame, tblObj, rowIndex, columnIndex)
  }
}

module.exports = { TableManager }
