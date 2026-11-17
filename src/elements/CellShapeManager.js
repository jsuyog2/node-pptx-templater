/**
 * @fileoverview CellShapeManager - Manages table cell shape overlays, polymorphic cell data,
 * strict geometric bounding-box containment filtering, dynamic row/column shifting, cloning, and cascading deletions.
 * Pure JavaScript, zero native dependencies, 100% ECMA-376 PresentationML & DrawingML compliant.
 */

'use strict'

const { createLogger } = require('../utils/logger.js')
const { ptToEmu, emuToPx, toEmu } = require('../spatial/units.js')
const { StyleBuilder } = require('../parser/StyleBuilder.js')
const { Z_ORDER_SYMBOL } = require('../parsers/XMLParser.js')

const logger = createLogger('CellShapeManager')

/**
 * Supported preset shape geometries for cell badges/shapes.
 */
const SUPPORTED_SHAPE_TYPES = new Set([
  'roundRect',
  'roundedRectangle',
  'circle',
  'ellipse',
  'rect',
  'rectangle',
  'square',
  'triangle',
  'star4',
  'star5',
  'star6',
  'star',
  'rightArrow',
  'leftArrow',
  'upArrow',
  'downArrow',
  'diamond',
  'hexagon',
  'chevron',
  'pill',
  'pentagon',
  'octagon',
  'heart',
  // Arrow aliases & connectors
  'arrow',
  'right-arrow',
  'right_arrow',
  'left-arrow',
  'left_arrow',
  'up-arrow',
  'up_arrow',
  'down-arrow',
  'down_arrow',
  'leftRightArrow',
  'upDownArrow',
  'quadArrow',
  'bentArrow',
  'stripedRightArrow',
  'notchedRightArrow',
  'chevronArrow',
  'line',
  'lineArrow',
  'connector',
  'straightConnector1',
])

/**
 * Normalizes preset shape geometry aliases to canonical PresentationML presets.
 *
 * @param {string} type
 * @returns {string}
 */
function normalizeCellShapePreset(type) {
  if (!type || typeof type !== 'string') return 'roundRect'
  const clean = type.trim()
  const lower = clean.toLowerCase().replace(/[-_\s]/g, '')
  switch (lower) {
    case 'arrow':
    case 'rightarrow':
      return 'rightArrow'
    case 'leftarrow':
      return 'leftArrow'
    case 'uparrow':
      return 'upArrow'
    case 'downarrow':
      return 'downArrow'
    case 'leftrightarrow':
      return 'leftRightArrow'
    case 'updownarrow':
      return 'upDownArrow'
    case 'quadarrow':
      return 'quadArrow'
    case 'roundrect':
    case 'roundedrectangle':
    case 'pill':
      return 'roundRect'
    case 'circle':
    case 'ellipse':
      return 'ellipse'
    case 'rect':
    case 'rectangle':
    case 'square':
      return 'rect'
    case 'line':
    case 'connector':
    case 'straightconnector1':
      return 'straightConnector1'
    default:
      return clean
  }
}

/**
 * Determines whether a shape bounding box is strictly contained within a target container bounding box.
 * Enforces strict mathematical inequality:
 * shape.xMin > target.xMin AND shape.yMin > target.yMin AND shape.xMax < target.xMax AND shape.yMax < target.yMax
 *
 * If the shape touches the border, crosses an edge, or lies outside, it is NOT strictly contained.
 *
 * @param {Object} shapeBounds - Bounding box { x, y, width, height } or { xMin, yMin, xMax, yMax }
 * @param {Object} targetBounds - Target container bounding box { x, y, width, height } or { xMin, yMin, xMax, yMax }
 * @returns {boolean} True if and only if shape is strictly interior to target.
 */
function isStrictlyContainedWithin(shapeBounds, targetBounds) {
  if (!shapeBounds || !targetBounds) return false

  const sX1 = shapeBounds.x !== undefined ? shapeBounds.x : shapeBounds.xMin
  const sY1 = shapeBounds.y !== undefined ? shapeBounds.y : shapeBounds.yMin
  const sW =
    shapeBounds.width !== undefined ? shapeBounds.width : shapeBounds.xMax - shapeBounds.xMin
  const sH =
    shapeBounds.height !== undefined ? shapeBounds.height : shapeBounds.yMax - shapeBounds.yMin
  const sX2 = shapeBounds.xMax !== undefined ? shapeBounds.xMax : sX1 + sW
  const sY2 = shapeBounds.yMax !== undefined ? shapeBounds.yMax : sY1 + sH

  const tX1 = targetBounds.x !== undefined ? targetBounds.x : targetBounds.xMin
  const tY1 = targetBounds.y !== undefined ? targetBounds.y : targetBounds.yMin
  const tW =
    targetBounds.width !== undefined ? targetBounds.width : targetBounds.xMax - targetBounds.xMin
  const tH =
    targetBounds.height !== undefined ? targetBounds.height : targetBounds.yMax - targetBounds.yMin
  const tX2 = targetBounds.xMax !== undefined ? targetBounds.xMax : tX1 + tW
  const tY2 = targetBounds.yMax !== undefined ? targetBounds.yMax : tY1 + tH

  return sX1 > tX1 && sY1 > tY1 && sX2 < tX2 && sY2 < tY2
}

class CellShapeManager {
  /** @private @type {XMLParser} */
  #xmlParser

  /** @private @type {ShapeManager|null} */
  #shapeManager

  /**
   * Map tracking cell shape anchor metadata for repositioning or queries.
   * Key: `${slideIndex}_${tableId}_${rowIndex}_${colIndex}`
   * @private
   * @type {Map<string, Object>}
   */
  #anchors = new Map()

  /**
   * @param {XMLParser} xmlParser
   * @param {ShapeManager} [shapeManager]
   */
  constructor(xmlParser, shapeManager = null) {
    this.#xmlParser = xmlParser
    this.#shapeManager = shapeManager
  }

  /**
   * Sets or updates the ShapeManager instance.
   * @param {ShapeManager} shapeManager
   */
  setShapeManager(shapeManager) {
    this.#shapeManager = shapeManager
  }

  /**
   * Determines whether a shape bounding box is strictly contained within a target container bounding box.
   *
   * @param {Object} shapeBounds
   * @param {Object} targetBounds
   * @returns {boolean}
   */
  isStrictlyContainedWithin(shapeBounds, targetBounds) {
    return isStrictlyContainedWithin(shapeBounds, targetBounds)
  }

  /**
   * Checks whether a cell value specifies a shape overlay.
   *
   * @param {*} val - Cell value.
   * @returns {boolean}
   */
  isCellShape(val) {
    if (!val || typeof val !== 'object' || Array.isArray(val)) return false
    if (val.shape !== undefined) return true
    if (typeof val.type === 'string') {
      if (SUPPORTED_SHAPE_TYPES.has(val.type)) return true
      const normalized = val.type
        .trim()
        .toLowerCase()
        .replace(/[-_\s]/g, '')
      if (SUPPORTED_SHAPE_TYPES.has(normalized)) return true
    }
    return false
  }

  /**
   * Normalizes any polymorphic cell data into a standard descriptor:
   * (1) Primitive: string, number, boolean
   * (2) Styled: object with value/text and inline style
   * (3) Shape: object specifying an overlay shape (inside or outside text)
   *
   * @param {*} val - Incoming cell data.
   * @returns {{
   *   kind: 'primitive' | 'styled' | 'shape',
   *   text: string,
   *   style: Object,
   *   shapeConfig?: Object
   * }}
   */
  parseCellData(val) {
    if (val === null || val === undefined) {
      return { kind: 'primitive', text: '', style: {} }
    }

    if (typeof val !== 'object' || Array.isArray(val)) {
      return { kind: 'primitive', text: String(val), style: {} }
    }

    // Check if it's a shape configuration
    if (this.isCellShape(val)) {
      let shapeOpt
      if (typeof val.shape === 'object' && val.shape !== null) {
        shapeOpt = { ...val.shape }
      } else if (typeof val.shape === 'string') {
        shapeOpt = { type: val.shape, ...val }
        delete shapeOpt.shape
      } else {
        shapeOpt = { ...val }
      }

      const textVal =
        val.text !== undefined
          ? String(val.text)
          : shapeOpt.text !== undefined
            ? String(shapeOpt.text)
            : val.value !== undefined
              ? String(val.value)
              : shapeOpt.value !== undefined
                ? String(shapeOpt.value)
                : ''

      const textPos =
        val.textPosition ||
        shapeOpt.textPosition ||
        val.position ||
        shapeOpt.textPos ||
        (textVal ? 'inside' : 'none')

      const textStyle = val.textStyle || shapeOpt.textStyle || {}
      const shapePlacement =
        val.shapePlacement ||
        shapeOpt.shapePlacement ||
        shapeOpt.alignment ||
        shapeOpt.align ||
        'left'

      return {
        kind: 'shape',
        text: textVal,
        style: val.style || shapeOpt.style || {},
        shapeConfig: {
          shape: shapeOpt,
          text: textVal,
          textPosition: textPos,
          textStyle,
          shapePlacement,
        },
      }
    }

    // Otherwise, treat as styled value object
    const textVal =
      val.value !== undefined
        ? String(val.value ?? '')
        : val.text !== undefined
          ? String(val.text ?? '')
          : ''

    const combinedStyle = {
      ...(typeof val.style === 'object' && val.style !== null ? val.style : {}),
      ...val,
    }
    delete combinedStyle.value
    delete combinedStyle.text

    return {
      kind: 'styled',
      text: textVal,
      style: combinedStyle,
    }
  }

  /**
   * Resolves exact EMU bounding coordinates for a shape node (<p:sp>).
   *
   * @param {Object} sp - Shape AST node.
   * @returns {{ x: number, y: number, width: number, height: number, xMin: number, yMin: number, xMax: number, yMax: number, xPx: number, yPx: number, widthPx: number, heightPx: number } | null}
   */
  getShapeBounds(sp) {
    if (!sp) return null
    const xfrm = sp['p:spPr']?.['a:xfrm'] || {}
    const off = xfrm['a:off'] || {}
    const ext = xfrm['a:ext'] || {}
    const x = parseInt(off['@_x'] || '0', 10)
    const y = parseInt(off['@_y'] || '0', 10)
    const width = parseInt(ext['@_cx'] || '0', 10)
    const height = parseInt(ext['@_cy'] || '0', 10)

    let xMin = x
    let yMin = y
    let xMax = x + width
    let yMax = y + height

    const rotStr = xfrm['@_rot']
    if (rotStr) {
      const rotVal = parseInt(rotStr, 10)
      if (!isNaN(rotVal) && rotVal !== 0) {
        const rad = (rotVal / 60000) * (Math.PI / 180)
        const cx = x + width / 2
        const cy = y + height / 2
        const hw = width / 2
        const hh = height / 2
        const cos = Math.abs(Math.cos(rad))
        const sin = Math.abs(Math.sin(rad))
        const newHw = hw * cos + hh * sin
        const newHh = hw * sin + hh * cos
        xMin = Math.round(cx - newHw)
        xMax = Math.round(cx + newHw)
        yMin = Math.round(cy - newHh)
        yMax = Math.round(cy + newHh)
      }
    }

    return {
      x,
      y,
      width,
      height,
      xMin,
      yMin,
      xMax,
      yMax,
      xPx: Math.round(emuToPx(xMin)),
      yPx: Math.round(emuToPx(yMin)),
      widthPx: Math.round(emuToPx(xMax - xMin)),
      heightPx: Math.round(emuToPx(yMax - yMin)),
    }
  }

  /**
   * Calculates the full bounding box of a table in EMUs.
   *
   * @param {Object} frame - Table <p:graphicFrame> AST.
   * @param {Object} tblObj - Table <a:tbl> AST.
   * @returns {{ x: number, y: number, width: number, height: number, xMin: number, yMin: number, xMax: number, yMax: number, xPx: number, yPx: number, widthPx: number, heightPx: number } | null}
   */
  getTableBounds(frame, tblObj) {
    if (!frame || !tblObj) return null
    const xfrm = frame['p:xfrm'] || {}
    const off = xfrm['a:off'] || {}
    const ext = xfrm['a:ext'] || {}
    const startX = parseInt(off['@_x'] || '0', 10)
    const startY = parseInt(off['@_y'] || '0', 10)

    const gridCols = tblObj['a:tblGrid']?.['a:gridCol'] || []
    const cols = Array.isArray(gridCols) ? gridCols : [gridCols]
    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]

    let totalWidth = 0
    for (const c of cols) {
      totalWidth += parseInt(c?.['@_w'] || '0', 10)
    }
    if (totalWidth <= 0 && ext['@_cx']) {
      totalWidth = parseInt(ext['@_cx'] || '0', 10)
    }

    let totalHeight = 0
    for (const r of trs) {
      totalHeight += parseInt(r?.['@_h'] || '0', 10)
    }
    if (totalHeight <= 0 && ext['@_cy']) {
      totalHeight = parseInt(ext['@_cy'] || '0', 10)
    }

    return {
      x: startX,
      y: startY,
      width: totalWidth,
      height: totalHeight,
      xMin: startX,
      yMin: startY,
      xMax: startX + totalWidth,
      yMax: startY + totalHeight,
      xPx: Math.round(emuToPx(startX)),
      yPx: Math.round(emuToPx(startY)),
      widthPx: Math.round(emuToPx(totalWidth)),
      heightPx: Math.round(emuToPx(totalHeight)),
    }
  }

  /**
   * Calculates the full horizontal row bounding box of a table row in EMUs.
   *
   * @param {Object} frame - Table <p:graphicFrame> AST.
   * @param {Object} tblObj - Table <a:tbl> AST.
   * @param {number} rowIndex - 0-based row index.
   * @returns {{ x: number, y: number, width: number, height: number, xMin: number, yMin: number, xMax: number, yMax: number, xPx: number, yPx: number, widthPx: number, heightPx: number } | null}
   */
  getTableRowBounds(frame, tblObj, rowIndex) {
    if (!frame || !tblObj) return null
    const xfrm = frame['p:xfrm'] || {}
    const off = xfrm['a:off'] || {}
    const ext = xfrm['a:ext'] || {}
    const startX = parseInt(off['@_x'] || '0', 10)
    const startY = parseInt(off['@_y'] || '0', 10)

    const gridCols = tblObj['a:tblGrid']?.['a:gridCol'] || []
    const cols = Array.isArray(gridCols) ? gridCols : [gridCols]
    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]

    if (rowIndex < 0 || rowIndex >= trs.length) return null

    let totalWidth = 0
    for (const c of cols) {
      totalWidth += parseInt(c?.['@_w'] || '0', 10)
    }
    if (totalWidth <= 0 && ext['@_cx']) {
      totalWidth = parseInt(ext['@_cx'] || '0', 10)
    }

    let currentY = startY
    for (let r = 0; r < rowIndex && r < trs.length; r++) {
      currentY += parseInt(trs[r]?.['@_h'] || '0', 10)
    }
    const rowHeight = parseInt(trs[rowIndex]?.['@_h'] || '0', 10)

    return {
      x: startX,
      y: currentY,
      width: totalWidth,
      height: rowHeight,
      xMin: startX,
      yMin: currentY,
      xMax: startX + totalWidth,
      yMax: currentY + rowHeight,
      xPx: Math.round(emuToPx(startX)),
      yPx: Math.round(emuToPx(currentY)),
      widthPx: Math.round(emuToPx(totalWidth)),
      heightPx: Math.round(emuToPx(rowHeight)),
    }
  }

  /**
   * Calculates the full vertical column bounding box of a table column in EMUs.
   *
   * @param {Object} frame - Table <p:graphicFrame> AST.
   * @param {Object} tblObj - Table <a:tbl> AST.
   * @param {number} columnIndex - 0-based column index.
   * @returns {{ x: number, y: number, width: number, height: number, xMin: number, yMin: number, xMax: number, yMax: number, xPx: number, yPx: number, widthPx: number, heightPx: number } | null}
   */
  getTableColumnBounds(frame, tblObj, columnIndex) {
    if (!frame || !tblObj) return null
    const xfrm = frame['p:xfrm'] || {}
    const off = xfrm['a:off'] || {}
    const ext = xfrm['a:ext'] || {}
    const startX = parseInt(off['@_x'] || '0', 10)
    const startY = parseInt(off['@_y'] || '0', 10)

    const gridCols = tblObj['a:tblGrid']?.['a:gridCol'] || []
    const cols = Array.isArray(gridCols) ? gridCols : [gridCols]
    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]

    if (columnIndex < 0 || columnIndex >= cols.length) return null

    let totalHeight = 0
    for (const r of trs) {
      totalHeight += parseInt(r?.['@_h'] || '0', 10)
    }
    if (totalHeight <= 0 && ext['@_cy']) {
      totalHeight = parseInt(ext['@_cy'] || '0', 10)
    }

    let currentX = startX
    for (let c = 0; c < columnIndex && c < cols.length; c++) {
      currentX += parseInt(cols[c]?.['@_w'] || '0', 10)
    }
    const colWidth = parseInt(cols[columnIndex]?.['@_w'] || '0', 10)

    return {
      x: currentX,
      y: startY,
      width: colWidth,
      height: totalHeight,
      xMin: currentX,
      yMin: startY,
      xMax: currentX + colWidth,
      yMax: startY + totalHeight,
      xPx: Math.round(emuToPx(currentX)),
      yPx: Math.round(emuToPx(startY)),
      widthPx: Math.round(emuToPx(colWidth)),
      heightPx: Math.round(emuToPx(totalHeight)),
    }
  }

  /**
   * Calculates the bounding box of a specific cell in EMUs.
   *
   * @param {Object} frame - Table <p:graphicFrame> AST.
   * @param {Object} tblObj - Table <a:tbl> AST.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} columnIndex - 0-based column index.
   * @returns {{ x: number, y: number, width: number, height: number, xMin: number, yMin: number, xMax: number, yMax: number, xPx: number, yPx: number, widthPx: number, heightPx: number } | null}
   */
  getTableCellBounds(frame, tblObj, rowIndex, columnIndex) {
    return this.calculateCellBounds(frame, tblObj, rowIndex, columnIndex)
  }

  /**
   * Calculates rendered cell bounding coordinates in EMUs.
   *
   * @param {Object} frame - Table <p:graphicFrame> AST.
   * @param {Object} tblObj - Table <a:tbl> AST.
   * @param {number} targetRow - 0-based row index.
   * @param {number} targetCol - 0-based column index.
   * @returns {{ x: number, y: number, width: number, height: number, xMin: number, yMin: number, xMax: number, yMax: number, xPx: number, yPx: number, widthPx: number, heightPx: number } | null}
   */
  calculateCellBounds(frame, tblObj, targetRow, targetCol) {
    if (!frame || !tblObj) return null
    const xfrm = frame['p:xfrm'] || {}
    const off = xfrm['a:off'] || {}
    const startX = parseInt(off['@_x'] || '0', 10)
    const startY = parseInt(off['@_y'] || '0', 10)

    const gridCols = tblObj['a:tblGrid']?.['a:gridCol'] || []
    const cols = Array.isArray(gridCols) ? gridCols : [gridCols]
    const trs = Array.isArray(tblObj['a:tr']) ? tblObj['a:tr'] : [tblObj['a:tr']]

    if (targetRow >= trs.length || targetCol >= cols.length) return null

    let currentX = startX
    for (let c = 0; c < targetCol && c < cols.length; c++) {
      currentX += parseInt(cols[c]?.['@_w'] || '0', 10)
    }

    let currentY = startY
    for (let r = 0; r < targetRow && r < trs.length; r++) {
      currentY += parseInt(trs[r]?.['@_h'] || '0', 10)
    }

    const cellWidth = parseInt(cols[targetCol]?.['@_w'] || '0', 10)
    const cellHeight = parseInt(trs[targetRow]?.['@_h'] || '0', 10)

    return {
      x: currentX,
      y: currentY,
      width: cellWidth,
      height: cellHeight,
      xMin: currentX,
      yMin: currentY,
      xMax: currentX + cellWidth,
      yMax: currentY + cellHeight,
      xPx: Math.round(emuToPx(currentX)),
      yPx: Math.round(emuToPx(currentY)),
      widthPx: Math.round(emuToPx(cellWidth)),
      heightPx: Math.round(emuToPx(cellHeight)),
    }
  }

  /**
   * Scans <p:spTree> on the target slide and purges all shapes that are strictly contained
   * within targetBounds. Preserves shapes that touch borders, cross edges, or lie outside.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {Object} targetBounds - Target container bounding box.
   * @param {SlideManager} slideManager - SlideManager instance.
   * @param {string} [tableId] - Optional tableId to clean from anchor registry.
   * @returns {Array<Object>} List of purged shape nodes.
   */
  purgeShapesWithinBounds(slideIndex, targetBounds, slideManager, tableId = null) {
    if (!targetBounds || !slideManager) return []
    const slideObj = slideManager.getSlideObj(slideIndex)
    if (!slideObj) return []

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree) return []

    const shapeKeys = ['p:sp', 'p:cxnSp']
    const purgedIds = new Set()
    const purgedShapes = []

    for (const key of shapeKeys) {
      if (!spTree[key]) continue
      const spList = Array.isArray(spTree[key]) ? spTree[key] : [spTree[key]]
      const remaining = []

      for (const sp of spList) {
        const shapeBounds = this.getShapeBounds(sp)
        if (shapeBounds && this.isStrictlyContainedWithin(shapeBounds, targetBounds)) {
          const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr'] || sp?.['p:nvCxnSpPr']?.['p:cNvPr']
          const id = cNvPr?.['@_id'] ? String(cNvPr['@_id']) : null
          if (id) purgedIds.add(id)
          purgedShapes.push(sp)

          if (spTree[Z_ORDER_SYMBOL] && id) {
            const zIdx = spTree[Z_ORDER_SYMBOL].indexOf(id)
            if (zIdx !== -1) spTree[Z_ORDER_SYMBOL].splice(zIdx, 1)
          }
          continue
        }
        remaining.push(sp)
      }

      spTree[key] = remaining
    }

    // Clean matching anchors
    for (const [key, anchor] of Array.from(this.#anchors.entries())) {
      if (anchor.slideIndex === slideIndex && (tableId === null || anchor.tableId === tableId)) {
        if (purgedIds.has(String(anchor.shapeId))) {
          this.#anchors.delete(key)
        }
      }
    }

    if (purgedShapes.length > 0) {
      slideManager.markSlideObjDirty(slideIndex)
    }
    return purgedShapes
  }

  /**
   * Creates or updates a cell shape overlay positioned over the table cell.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string} tableId - Table identifier name.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} colIndex - 0-based column index.
   * @param {Object} shapeConfig - Shape overlay options.
   * @param {SlideManager} slideManager - SlideManager instance.
   * @param {Object} frame - Table <p:graphicFrame>.
   * @param {Object} tblObj - Table <a:tbl>.
   * @param {Object} tcNode - Table cell <a:tc>.
   * @param {ShapeManager} [shapeManager] - ShapeManager instance override.
   * @returns {string|null} Created shape ID.
   */
  applyCellShape(
    slideIndex,
    tableId,
    rowIndex,
    colIndex,
    shapeConfig,
    slideManager,
    frame,
    tblObj,
    tcNode,
    shapeManager = this.#shapeManager
  ) {
    if (!shapeManager) {
      logger.warn('ShapeManager not available in CellShapeManager. Skipping cell shape creation.')
      return null
    }

    // 1. Remove any previous overlay shape strictly within this specific cell
    const prevBounds = this.calculateCellBounds(frame, tblObj, rowIndex, colIndex)
    if (prevBounds) {
      this.purgeShapesWithinBounds(slideIndex, prevBounds, slideManager, tableId)
    }
    this.removeCellShape(slideIndex, tableId, rowIndex, colIndex, slideManager)

    // 2. Compute exact cell bounding box in slide EMUs
    const bounds = prevBounds || this.calculateCellBounds(frame, tblObj, rowIndex, colIndex)
    if (!bounds || bounds.width <= 0 || bounds.height <= 0) {
      logger.warn(`Could not compute cell bounds for ${tableId} cell (${rowIndex}, ${colIndex})`)
      return null
    }

    const shapeOpt = shapeConfig.shape || shapeConfig
    const shapeType = normalizeCellShapePreset(shapeOpt.type || shapeOpt.preset || 'roundRect')
    const text = shapeConfig.text || ''
    const textPosition = shapeConfig.textPosition || (text ? 'inside' : 'none')
    const textStyle = shapeConfig.textStyle || {}
    const shapePlacement =
      shapeConfig.shapePlacement ||
      shapeOpt.shapePlacement ||
      shapeOpt.alignment ||
      shapeOpt.align ||
      'left'

    // 3. Compute shape dimensions
    let shapeW
    let shapeH

    if (shapeType === 'circle' || shapeOpt.radius !== undefined) {
      const radiusVal =
        shapeOpt.radius !== undefined
          ? typeof shapeOpt.radius === 'number' && shapeOpt.radius < 50000
            ? ptToEmu(shapeOpt.radius)
            : toEmu(shapeOpt.radius)
          : Math.min(bounds.width, bounds.height) * 0.35
      shapeW = Math.round(radiusVal * 2)
      shapeH = Math.round(radiusVal * 2)
    } else {
      const rawW =
        shapeOpt.width !== undefined
          ? shapeOpt.width
          : shapeOpt.w !== undefined
            ? shapeOpt.w
            : shapeOpt.size
      const rawH =
        shapeOpt.height !== undefined
          ? shapeOpt.height
          : shapeOpt.h !== undefined
            ? shapeOpt.h
            : shapeOpt.size

      if (rawW !== undefined) {
        shapeW = typeof rawW === 'number' && rawW < 50000 ? ptToEmu(rawW) : toEmu(rawW)
      } else {
        shapeW = textPosition === 'inside' ? Math.round(bounds.width * 0.8) : ptToEmu(12)
      }

      if (rawH !== undefined) {
        shapeH = typeof rawH === 'number' && rawH < 50000 ? ptToEmu(rawH) : toEmu(rawH)
      } else {
        shapeH = textPosition === 'inside' ? Math.round(bounds.height * 0.65) : ptToEmu(12)
      }
    }

    // Clamp dimensions to cell bounds with a minimum margin so it is strictly interior
    const minMargin = ptToEmu(2)
    shapeW = Math.min(shapeW, Math.max(1, bounds.width - minMargin * 2))
    shapeH = Math.min(shapeH, Math.max(1, bounds.height - minMargin * 2))

    // 4. Compute shape position coordinates
    // Inset boundaries to guarantee strict inequality: shapeX > bounds.x && shapeX + shapeW < bounds.x + bounds.width
    const minX = bounds.x + 1
    const maxX = Math.max(minX, bounds.x + bounds.width - shapeW - 1)
    const minY = bounds.y + 1
    const maxY = Math.max(minY, bounds.y + bounds.height - shapeH - 1)

    // Vertical alignment / offsets
    const vAlign =
      shapeOpt.verticalAlignment ||
      shapeOpt.verticalAlign ||
      shapeOpt.valign ||
      shapeConfig.verticalAlignment ||
      'middle'

    let baseY
    if (vAlign === 'top') {
      baseY = bounds.y + ptToEmu(4)
    } else if (vAlign === 'bottom') {
      baseY = bounds.y + bounds.height - shapeH - ptToEmu(4)
    } else {
      // 'middle' / 'center'
      baseY = bounds.y + Math.round((bounds.height - shapeH) / 2)
    }

    if (shapeOpt.y !== undefined) {
      const customY =
        typeof shapeOpt.y === 'number' && shapeOpt.y < 50000
          ? ptToEmu(shapeOpt.y)
          : toEmu(shapeOpt.y)
      baseY = bounds.y + customY
    } else if (shapeOpt.offsetY !== undefined) {
      const customOffY =
        typeof shapeOpt.offsetY === 'number' && shapeOpt.offsetY < 50000
          ? ptToEmu(shapeOpt.offsetY)
          : toEmu(shapeOpt.offsetY)
      baseY += customOffY
    }
    const shapeY = Math.max(minY, Math.min(maxY, baseY))

    // Horizontal placement / alignment / offsets
    let baseX
    if (
      textPosition === 'inside' ||
      !text ||
      shapePlacement === 'center' ||
      shapeOpt.position === 'center' ||
      shapeOpt.align === 'center' ||
      shapeOpt.alignment === 'center'
    ) {
      baseX = bounds.x + Math.round((bounds.width - shapeW) / 2)
    } else if (
      shapePlacement === 'right' ||
      shapeOpt.position === 'right' ||
      shapeOpt.align === 'right' ||
      shapeOpt.alignment === 'right'
    ) {
      const pad = ptToEmu(6)
      baseX = bounds.x + bounds.width - shapeW - pad
    } else {
      // 'left'
      const pad = ptToEmu(6)
      baseX = bounds.x + pad
    }

    if (shapeOpt.x !== undefined) {
      const customX =
        typeof shapeOpt.x === 'number' && shapeOpt.x < 50000
          ? ptToEmu(shapeOpt.x)
          : toEmu(shapeOpt.x)
      baseX = bounds.x + customX
    } else if (shapeOpt.offsetX !== undefined) {
      const customOffX =
        typeof shapeOpt.offsetX === 'number' && shapeOpt.offsetX < 50000
          ? ptToEmu(shapeOpt.offsetX)
          : toEmu(shapeOpt.offsetX)
      baseX += customOffX
    }
    const shapeX = Math.max(minX, Math.min(maxX, baseX))

    // 5. Configure Border / Stroke
    let border = shapeOpt.border
    if (!border && (shapeOpt.stroke || shapeOpt.strokeColor)) {
      border = {
        color: shapeOpt.stroke || shapeOpt.strokeColor,
        width: shapeOpt.strokeWidth || shapeOpt.lineWidth || 1,
      }
    }

    const uniqueSuffix = `${Date.now()}_${Math.floor(Math.random() * 1000)}`
    const shapeId = `cellshape_${tableId}_r${rowIndex}_c${colIndex}_${uniqueSuffix}`
    const shapeName = `CellShape_${tableId}_r${rowIndex}_c${colIndex}`

    // 6. Handle Inside vs Outside Text
    const shapePayload = {
      id: shapeId,
      name: shapeName,
      type: shapeType,
      x: shapeX,
      y: shapeY,
      width: shapeW,
      height: shapeH,
      unit: 'emu',
      fill: shapeOpt.fill || shapeOpt.backgroundColor || shapeOpt.bgColor || '#10B981',
      border,
      opacity: shapeOpt.opacity !== undefined ? shapeOpt.opacity : shapeOpt.alpha,
      rotation: shapeOpt.rotation,
      borderRadius: shapeOpt.borderRadius,
    }

    if (textPosition === 'inside' && text) {
      shapePayload.text = text
      shapePayload.textStyle = textStyle
      shapePayload.textColor = textStyle.color || '#FFFFFF'
      shapePayload.fontSize = textStyle.fontSize || 10
      shapePayload.bold = textStyle.bold !== undefined ? textStyle.bold : true
      shapePayload.italic = textStyle.italic || false
      shapePayload.alignment = textStyle.alignment || textStyle.align || 'center'
      shapePayload.verticalAlignment =
        textStyle.verticalAlignment || textStyle.verticalAlign || 'middle'

      // Empty out text in underlying table cell
      if (tcNode && tcNode['a:txBody']) {
        this.#setCellText(tcNode, '')
      }
    } else if (textPosition === 'outside' && text) {
      // Shape has no text; table cell contains the outside text
      if (tcNode) {
        this.#setCellText(tcNode, text, textStyle)

        // Adjust cell margins so cell text does not collide with the overlay shape
        if (!tcNode['a:tcPr']) tcNode['a:tcPr'] = {}
        const marginPadding = shapeW + ptToEmu(10)
        if (shapePlacement === 'right') {
          tcNode['a:tcPr']['@_marR'] = String(marginPadding)
        } else {
          tcNode['a:tcPr']['@_marL'] = String(marginPadding)
        }
      }
    } else {
      // Standalone shape overlay (no text)
      if (tcNode && tcNode['a:txBody']) {
        this.#setCellText(tcNode, '')
      }
    }

    // 7. Add shape to slide
    const createdId = shapeManager.addShape(slideIndex, shapePayload, slideManager)

    // Register anchor
    const anchorKey = `${slideIndex}_${tableId}_${rowIndex}_${colIndex}`
    this.#anchors.set(anchorKey, {
      shapeId: String(createdId || shapePayload.id),
      shapeName,
      slideIndex,
      tableId,
      rowIndex,
      colIndex,
      text,
      config: shapeConfig,
    })

    logger.debug(
      `Applied cell shape "${shapeName}" at (${rowIndex}, ${colIndex}) on slide ${slideIndex}`
    )
    return String(createdId || shapePayload.id)
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
            'a:t': String(text ?? ''),
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
    firstR['a:t'] = String(text ?? '')

    if (rList.length > 1) {
      rList.splice(1)
    }

    if (options && Object.keys(options).length > 0) {
      firstP['a:pPr'] = StyleBuilder.buildParagraphProperties(firstP['a:pPr'], options)
      firstR['a:rPr'] = StyleBuilder.buildRunProperties(firstR['a:rPr'], options)
    }
  }

  /**
   * Queries all shape nodes belonging to a table on a slide.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {SlideManager} slideManager
   * @param {string} [tableName] - Optional table name from <p:cNvPr>.
   * @param {string} [tableShapeId] - Optional table numeric shape ID from <p:cNvPr>.
   * @returns {Array<{ sp: Object, name: string, id: string, r: number, c: number }>}
   */
  getAllCellShapesForTable(
    slideIndex,
    tableId,
    slideManager,
    tableName = null,
    tableShapeId = null
  ) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    if (!slideObj) return []

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree) return []

    const results = []

    const prefixes = new Set([`CellShape_${tableId}_`, `cellshape_${tableId}_`])
    if (tableName) {
      prefixes.add(`CellShape_${tableName}_`)
      prefixes.add(`cellshape_${tableName}_`)
    }
    if (tableShapeId) {
      prefixes.add(`CellShape_${tableShapeId}_`)
      prefixes.add(`cellshape_${tableShapeId}_`)
    }

    const shapeKeys = ['p:sp', 'p:cxnSp']
    for (const key of shapeKeys) {
      if (!spTree[key]) continue
      const spList = Array.isArray(spTree[key]) ? spTree[key] : [spTree[key]]

      for (const sp of spList) {
        const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr'] || sp?.['p:nvCxnSpPr']?.['p:cNvPr']
        const name = cNvPr?.['@_name'] || ''
        const id = cNvPr?.['@_id'] ? String(cNvPr['@_id']) : ''

        let matched = false
        for (const p of prefixes) {
          if (name.startsWith(p) || id.startsWith(p)) {
            matched = true
            break
          }
        }

        if (matched) {
          const match = name.match(/_r(\d+)_c(\d+)/)
          if (match) {
            const r = parseInt(match[1], 10)
            const c = parseInt(match[2], 10)
            results.push({ sp, name, id, r, c })
          }
        }
      }
    }

    return results
  }

  /**
   * Shifts existing cell shapes down by 1 row when a row is inserted.
   * Modifies logical row index (r >= fromRowIndex -> r + 1), XML shape name, and slide Y position.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} fromRowIndex - Row index at which the new row is inserted.
   * @param {number} deltaHeightEmu - Height of the inserted row in EMUs.
   * @param {SlideManager} slideManager
   */
  shiftRowShapes(slideIndex, tableId, fromRowIndex, deltaHeightEmu, slideManager) {
    const shapes = this.getAllCellShapesForTable(slideIndex, tableId, slideManager)
    const targetShapes = shapes.filter(s => s.r >= fromRowIndex)

    // Process from largest row index downward to prevent collision
    targetShapes.sort((a, b) => b.r - a.r)

    for (const item of targetShapes) {
      const oldR = item.r
      const newR = oldR + 1
      const col = item.c
      const newName = `CellShape_${tableId}_r${newR}_c${col}`

      // Update shape name in XML AST
      const cNvPr = item.sp?.['p:nvSpPr']?.['p:cNvPr'] || item.sp?.['p:nvCxnSpPr']?.['p:cNvPr']
      if (cNvPr) {
        cNvPr['@_name'] = newName
      }

      // Update vertical position in XML AST
      if (item.sp?.['p:spPr']?.['a:xfrm']?.['a:off']) {
        const off = item.sp['p:spPr']['a:xfrm']['a:off']
        const curY = parseInt(off['@_y'] || '0', 10)
        off['@_y'] = String(curY + deltaHeightEmu)
      }

      // Update internal anchor registry
      const oldKey = `${slideIndex}_${tableId}_${oldR}_${col}`
      const newKey = `${slideIndex}_${tableId}_${newR}_${col}`
      const existingAnchor = this.#anchors.get(oldKey)
      if (existingAnchor) {
        this.#anchors.delete(oldKey)
        this.#anchors.set(newKey, {
          ...existingAnchor,
          shapeName: newName,
          rowIndex: newR,
        })
      }
    }

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Removes shapes strictly within the removed row's bounding box and shifts remaining interior
   * shapes in lower rows upward without shifting external/edge shapes.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} removedRowIndex - 0-based index of removed row.
   * @param {number} removedHeightEmu - Height of removed row in EMUs.
   * @param {SlideManager} slideManager
   * @param {Object} [frame] - Table graphic frame AST.
   * @param {Object} [tblObj] - Table tbl AST.
   */
  removeRowShapesAndShift(
    slideIndex,
    tableId,
    removedRowIndex,
    removedHeightEmu,
    slideManager,
    frame = null,
    tblObj = null
  ) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    if (!slideObj) return

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree) return

    const tableName = frame?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_name'] || tableId
    const tableShapeId = String(frame?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_id'] || '')

    const tableBounds = frame && tblObj ? this.getTableBounds(frame, tblObj) : null
    const rowBounds =
      frame && tblObj ? this.getTableRowBounds(frame, tblObj, removedRowIndex) : null

    // 1. Purge shapes strictly inside the removed row's bounding box
    if (rowBounds) {
      this.purgeShapesWithinBounds(slideIndex, rowBounds, slideManager, tableId)
      if (tableName && tableName !== tableId) {
        this.purgeShapesWithinBounds(slideIndex, rowBounds, slideManager, tableName)
      }
    } else {
      this.removeRowShapes(slideIndex, tableId, removedRowIndex, slideManager)
      if (tableName && tableName !== tableId) {
        this.removeRowShapes(slideIndex, tableName, removedRowIndex, slideManager)
      }
    }

    // 2. Discover all shapes anchored or positioned in lower rows
    const shapesToShift = new Map()

    // 2a. First, collect named cell shapes for this table
    const cellShapes = this.getAllCellShapesForTable(
      slideIndex,
      tableId,
      slideManager,
      tableName,
      tableShapeId
    )
    for (const item of cellShapes) {
      if (item.r > removedRowIndex) {
        shapesToShift.set(item.sp, {
          sp: item.sp,
          r: item.r,
          c: item.c,
          name: item.name,
        })
      }
    }

    // 2b. Also inspect anchors for table
    for (const anchor of this.#anchors.values()) {
      if (
        anchor.slideIndex === slideIndex &&
        (anchor.tableId === tableId || (tableName && anchor.tableId === tableName)) &&
        anchor.rowIndex > removedRowIndex
      ) {
        const allShapes = [
          ...(Array.isArray(spTree['p:sp'])
            ? spTree['p:sp']
            : spTree['p:sp']
              ? [spTree['p:sp']]
              : []),
          ...(Array.isArray(spTree['p:cxnSp'])
            ? spTree['p:cxnSp']
            : spTree['p:cxnSp']
              ? [spTree['p:cxnSp']]
              : []),
        ]
        const foundSp = allShapes.find(
          s =>
            String(
              s?.['p:nvSpPr']?.['p:cNvPr']?.['@_id'] || s?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_id']
            ) === String(anchor.shapeId)
        )
        if (foundSp && !shapesToShift.has(foundSp)) {
          const cNvPr = foundSp['p:nvSpPr']?.['p:cNvPr'] || foundSp['p:nvCxnSpPr']?.['p:cNvPr']
          shapesToShift.set(foundSp, {
            sp: foundSp,
            r: anchor.rowIndex,
            c: anchor.colIndex,
            name: anchor.shapeName || cNvPr?.['@_name'],
          })
        }
      }
    }

    // 2c. Also check by coordinates: any shape whose bounds are inside table horizontally and below the removed row vertically
    if (rowBounds && tableBounds) {
      const allShapes = [
        ...(Array.isArray(spTree['p:sp'])
          ? spTree['p:sp']
          : spTree['p:sp']
            ? [spTree['p:sp']]
            : []),
        ...(Array.isArray(spTree['p:cxnSp'])
          ? spTree['p:cxnSp']
          : spTree['p:cxnSp']
            ? [spTree['p:cxnSp']]
            : []),
      ]
      for (const sp of allShapes) {
        if (shapesToShift.has(sp)) continue
        const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr'] || sp?.['p:nvCxnSpPr']?.['p:cNvPr']
        const name = cNvPr?.['@_name'] || ''

        const shapeBounds = this.getShapeBounds(sp)
        if (shapeBounds) {
          const withinTableX =
            shapeBounds.xMin >= tableBounds.xMin - 1000 &&
            shapeBounds.xMax <= tableBounds.xMax + 1000
          const belowRemovedRow =
            shapeBounds.yMin >= rowBounds.yMax - 1000 && shapeBounds.yMax <= tableBounds.yMax + 1000

          if (withinTableX && belowRemovedRow) {
            const match = name.match(/_r(\d+)_c(\d+)/)
            const r = match ? parseInt(match[1], 10) : removedRowIndex + 1
            const c = match ? parseInt(match[2], 10) : 0
            shapesToShift.set(sp, { sp, r, c, name })
          }
        }
      }
    }

    // Sort target shapes by row index ascending (smallest first)
    const targetShapes = Array.from(shapesToShift.values())
    targetShapes.sort((a, b) => a.r - b.r)

    for (const item of targetShapes) {
      const shapeBounds = this.getShapeBounds(item.sp)
      // Only skip if shape is completely outside the table horizontally or vertically
      if (
        tableBounds &&
        shapeBounds &&
        (shapeBounds.xMin > tableBounds.xMax ||
          shapeBounds.xMax < tableBounds.xMin ||
          shapeBounds.yMin > tableBounds.yMax + 1000)
      ) {
        continue
      }

      const oldR = item.r
      const newR = oldR - 1
      const col = item.c
      const baseName = tableName || tableId

      // Update name
      let newName = item.name
      if (item.name && item.name.includes('_r')) {
        newName = item.name.replace(/_r\d+/, `_r${newR}`)
      } else if (item.name && item.name.startsWith('CellShape_')) {
        newName = `CellShape_${baseName}_r${newR}_c${col}`
      }

      const cNvPr = item.sp?.['p:nvSpPr']?.['p:cNvPr'] || item.sp?.['p:nvCxnSpPr']?.['p:cNvPr']
      if (newName !== item.name && cNvPr) {
        cNvPr['@_name'] = newName
      }

      if (item.sp?.['p:spPr']?.['a:xfrm']?.['a:off']) {
        const off = item.sp['p:spPr']['a:xfrm']['a:off']
        const curY = parseInt(off['@_y'] || '0', 10)
        off['@_y'] = String(Math.max(0, curY - removedHeightEmu))
      }

      const oldKey = `${slideIndex}_${tableId}_${oldR}_${col}`
      const newKey = `${slideIndex}_${tableId}_${newR}_${col}`
      const existingAnchor =
        this.#anchors.get(oldKey) ||
        (tableName ? this.#anchors.get(`${slideIndex}_${tableName}_${oldR}_${col}`) : null)

      if (existingAnchor) {
        this.#anchors.delete(oldKey)
        if (tableName) this.#anchors.delete(`${slideIndex}_${tableName}_${oldR}_${col}`)
        this.#anchors.set(newKey, {
          ...existingAnchor,
          shapeName: newName,
          rowIndex: newR,
        })
      }
    }

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Shifts existing cell shapes right by 1 column when a column is inserted.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} fromColIndex - Column index at which the new column is inserted.
   * @param {number} deltaWidthEmu - Width of the inserted column in EMUs.
   * @param {SlideManager} slideManager
   */
  shiftColShapes(slideIndex, tableId, fromColIndex, deltaWidthEmu, slideManager) {
    const shapes = this.getAllCellShapesForTable(slideIndex, tableId, slideManager)
    const targetShapes = shapes.filter(s => s.c >= fromColIndex)

    // Process from largest column index downward
    targetShapes.sort((a, b) => b.c - a.c)

    for (const item of targetShapes) {
      const row = item.r
      const oldC = item.c
      const newC = oldC + 1
      const newName = `CellShape_${tableId}_r${row}_c${newC}`

      const cNvPr = item.sp?.['p:nvSpPr']?.['p:cNvPr'] || item.sp?.['p:nvCxnSpPr']?.['p:cNvPr']
      if (cNvPr) {
        cNvPr['@_name'] = newName
      }

      if (item.sp?.['p:spPr']?.['a:xfrm']?.['a:off']) {
        const off = item.sp['p:spPr']['a:xfrm']['a:off']
        const curX = parseInt(off['@_x'] || '0', 10)
        off['@_x'] = String(curX + deltaWidthEmu)
      }

      const oldKey = `${slideIndex}_${tableId}_${row}_${oldC}`
      const newKey = `${slideIndex}_${tableId}_${row}_${newC}`
      const existingAnchor = this.#anchors.get(oldKey)
      if (existingAnchor) {
        this.#anchors.delete(oldKey)
        this.#anchors.set(newKey, {
          ...existingAnchor,
          shapeName: newName,
          colIndex: newC,
        })
      }
    }

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Removes shapes strictly within the removed column's bounding box and shifts remaining interior
   * shapes in rightward columns leftward without shifting external/edge shapes.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} removedColIndex - 0-based index of removed column.
   * @param {number} removedWidthEmu - Width of removed column in EMUs.
   * @param {SlideManager} slideManager
   * @param {Object} [frame] - Table graphic frame AST.
   * @param {Object} [tblObj] - Table tbl AST.
   */
  removeColShapesAndShift(
    slideIndex,
    tableId,
    removedColIndex,
    removedWidthEmu,
    slideManager,
    frame = null,
    tblObj = null
  ) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    if (!slideObj) return

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree) return

    const tableName = frame?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_name'] || tableId
    const tableShapeId = String(frame?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_id'] || '')

    const tableBounds = frame && tblObj ? this.getTableBounds(frame, tblObj) : null
    const colBounds =
      frame && tblObj ? this.getTableColumnBounds(frame, tblObj, removedColIndex) : null

    // 1. Purge shapes strictly inside the removed column's bounding box
    if (colBounds) {
      this.purgeShapesWithinBounds(slideIndex, colBounds, slideManager, tableId)
      if (tableName && tableName !== tableId) {
        this.purgeShapesWithinBounds(slideIndex, colBounds, slideManager, tableName)
      }
    } else {
      this.removeColShapes(slideIndex, tableId, removedColIndex, slideManager)
      if (tableName && tableName !== tableId) {
        this.removeColShapes(slideIndex, tableName, removedColIndex, slideManager)
      }
    }

    // 2. Discover all shapes anchored or positioned in subsequent columns
    const shapesToShift = new Map()

    // 2a. First, collect named cell shapes for this table
    const cellShapes = this.getAllCellShapesForTable(
      slideIndex,
      tableId,
      slideManager,
      tableName,
      tableShapeId
    )
    for (const item of cellShapes) {
      if (item.c > removedColIndex) {
        shapesToShift.set(item.sp, {
          sp: item.sp,
          r: item.r,
          c: item.c,
          name: item.name,
        })
      }
    }

    // 2b. Also inspect anchors for table
    for (const anchor of this.#anchors.values()) {
      if (
        anchor.slideIndex === slideIndex &&
        (anchor.tableId === tableId || (tableName && anchor.tableId === tableName)) &&
        anchor.colIndex > removedColIndex
      ) {
        const allShapes = [
          ...(Array.isArray(spTree['p:sp'])
            ? spTree['p:sp']
            : spTree['p:sp']
              ? [spTree['p:sp']]
              : []),
          ...(Array.isArray(spTree['p:cxnSp'])
            ? spTree['p:cxnSp']
            : spTree['p:cxnSp']
              ? [spTree['p:cxnSp']]
              : []),
        ]
        const foundSp = allShapes.find(
          s =>
            String(
              s?.['p:nvSpPr']?.['p:cNvPr']?.['@_id'] || s?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_id']
            ) === String(anchor.shapeId)
        )
        if (foundSp && !shapesToShift.has(foundSp)) {
          const cNvPr = foundSp['p:nvSpPr']?.['p:cNvPr'] || foundSp['p:nvCxnSpPr']?.['p:cNvPr']
          shapesToShift.set(foundSp, {
            sp: foundSp,
            r: anchor.rowIndex,
            c: anchor.colIndex,
            name: anchor.shapeName || cNvPr?.['@_name'],
          })
        }
      }
    }

    // 2c. Also check by coordinates: if x >= colBounds.xMax and within table vertical span
    if (colBounds && tableBounds) {
      const allShapes = [
        ...(Array.isArray(spTree['p:sp'])
          ? spTree['p:sp']
          : spTree['p:sp']
            ? [spTree['p:sp']]
            : []),
        ...(Array.isArray(spTree['p:cxnSp'])
          ? spTree['p:cxnSp']
          : spTree['p:cxnSp']
            ? [spTree['p:cxnSp']]
            : []),
      ]
      for (const sp of allShapes) {
        if (shapesToShift.has(sp)) continue
        const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr'] || sp?.['p:nvCxnSpPr']?.['p:cNvPr']
        const name = cNvPr?.['@_name'] || ''

        const shapeBounds = this.getShapeBounds(sp)
        if (shapeBounds) {
          const withinTableY =
            shapeBounds.yMin >= tableBounds.yMin - 1000 &&
            shapeBounds.yMax <= tableBounds.yMax + 1000
          const toRightOfRemovedCol =
            shapeBounds.xMin >= colBounds.xMax - 1000 && shapeBounds.xMax <= tableBounds.xMax + 1000

          if (withinTableY && toRightOfRemovedCol) {
            const match = name.match(/_r(\d+)_c(\d+)/)
            const r = match ? parseInt(match[1], 10) : 0
            const c = match ? parseInt(match[2], 10) : removedColIndex + 1
            shapesToShift.set(sp, { sp, r, c, name })
          }
        }
      }
    }

    // Sort target shapes by column index ascending (smallest first)
    const targetShapes = Array.from(shapesToShift.values())
    targetShapes.sort((a, b) => a.c - b.c)

    for (const item of targetShapes) {
      const shapeBounds = this.getShapeBounds(item.sp)
      // Only skip if shape is completely outside the table vertically or past the table horizontally
      if (
        tableBounds &&
        shapeBounds &&
        (shapeBounds.yMin > tableBounds.yMax ||
          shapeBounds.yMax < tableBounds.yMin ||
          shapeBounds.xMin > tableBounds.xMax + 1000)
      ) {
        continue
      }

      const row = item.r
      const oldC = item.c
      const newC = oldC - 1
      const baseName = tableName || tableId

      // Update name
      let newName = item.name
      if (item.name && item.name.includes('_c')) {
        newName = item.name.replace(/_c\d+/, `_c${newC}`)
      } else if (item.name && item.name.startsWith('CellShape_')) {
        newName = `CellShape_${baseName}_r${row}_c${newC}`
      }

      const cNvPr = item.sp?.['p:nvSpPr']?.['p:cNvPr'] || item.sp?.['p:nvCxnSpPr']?.['p:cNvPr']
      if (newName !== item.name && cNvPr) {
        cNvPr['@_name'] = newName
      }

      if (item.sp?.['p:spPr']?.['a:xfrm']?.['a:off']) {
        const off = item.sp['p:spPr']['a:xfrm']['a:off']
        const curX = parseInt(off['@_x'] || '0', 10)
        off['@_x'] = String(Math.max(0, curX - removedWidthEmu))
      }

      const oldKey = `${slideIndex}_${tableId}_${row}_${oldC}`
      const newKey = `${slideIndex}_${tableId}_${row}_${newC}`
      const existingAnchor =
        this.#anchors.get(oldKey) ||
        (tableName ? this.#anchors.get(`${slideIndex}_${tableName}_${row}_${oldC}`) : null)

      if (existingAnchor) {
        this.#anchors.delete(oldKey)
        if (tableName) this.#anchors.delete(`${slideIndex}_${tableName}_${row}_${oldC}`)
        this.#anchors.set(newKey, {
          ...existingAnchor,
          shapeName: newName,
          colIndex: newC,
        })
      }
    }

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Removes overlay shape for a specific cell.
   *
   * @param {number} slideIndex - Slide index.
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} colIndex - 0-based column index.
   * @param {SlideManager} slideManager - SlideManager instance.
   */
  removeCellShape(
    slideIndex,
    tableId,
    rowIndex,
    colIndex,
    slideManager,
    frame = null,
    tblObj = null
  ) {
    if (frame && tblObj) {
      const cellBounds = this.getTableCellBounds(frame, tblObj, rowIndex, colIndex)
      if (cellBounds) {
        return this.purgeShapesWithinBounds(slideIndex, cellBounds, slideManager, tableId)
      }
    }

    const slideObj = slideManager.getSlideObj(slideIndex)
    if (!slideObj) return

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree) return

    const prefix = `CellShape_${tableId}_r${rowIndex}_c${colIndex}`
    const prefixOld = `cellshape_${tableId}_r${rowIndex}_c${colIndex}`
    const anchorKey = `${slideIndex}_${tableId}_${rowIndex}_${colIndex}`
    const anchor = this.#anchors.get(anchorKey)
    const targetId = anchor?.shapeId ? String(anchor.shapeId) : null

    for (const key of ['p:sp', 'p:cxnSp']) {
      if (!spTree[key]) continue
      const spList = Array.isArray(spTree[key]) ? spTree[key] : [spTree[key]]
      const remaining = []

      for (const sp of spList) {
        const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr'] || sp?.['p:nvCxnSpPr']?.['p:cNvPr']
        const name = cNvPr?.['@_name'] || ''
        const id = cNvPr?.['@_id'] ? String(cNvPr['@_id']) : ''

        if (
          name.startsWith(prefix) ||
          name.startsWith(prefixOld) ||
          id.startsWith(prefixOld) ||
          (targetId && id === targetId)
        ) {
          if (spTree[Z_ORDER_SYMBOL] && id) {
            const zIdx = spTree[Z_ORDER_SYMBOL].indexOf(id)
            if (zIdx !== -1) spTree[Z_ORDER_SYMBOL].splice(zIdx, 1)
          }
          continue
        }
        remaining.push(sp)
      }

      spTree[key] = remaining
    }

    this.#anchors.delete(anchorKey)
    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Removes all shapes across an entire row.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} rowIndex
   * @param {SlideManager} slideManager
   */
  removeRowShapes(slideIndex, tableId, rowIndex, slideManager, frame = null, tblObj = null) {
    if (frame && tblObj) {
      const rowBounds = this.getTableRowBounds(frame, tblObj, rowIndex)
      if (rowBounds) {
        return this.purgeShapesWithinBounds(slideIndex, rowBounds, slideManager, tableId)
      }
    }

    const slideObj = slideManager.getSlideObj(slideIndex)
    if (!slideObj) return

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree) return

    const rowPattern = new RegExp(`^(?:CellShape|cellshape)_${tableId}_r${rowIndex}_c\\d+`)

    for (const key of ['p:sp', 'p:cxnSp']) {
      if (!spTree[key]) continue
      const spList = Array.isArray(spTree[key]) ? spTree[key] : [spTree[key]]
      const remaining = []

      for (const sp of spList) {
        const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr'] || sp?.['p:nvCxnSpPr']?.['p:cNvPr']
        const name = cNvPr?.['@_name'] || ''
        const id = cNvPr?.['@_id'] ? String(cNvPr['@_id']) : ''

        if (rowPattern.test(name) || rowPattern.test(id)) {
          if (spTree[Z_ORDER_SYMBOL] && id) {
            const zIdx = spTree[Z_ORDER_SYMBOL].indexOf(id)
            if (zIdx !== -1) spTree[Z_ORDER_SYMBOL].splice(zIdx, 1)
          }
          continue
        }
        remaining.push(sp)
      }

      spTree[key] = remaining
    }

    // Clean matching keys from anchor map
    const keyPrefix = `${slideIndex}_${tableId}_${rowIndex}_`
    for (const key of Array.from(this.#anchors.keys())) {
      if (key.startsWith(keyPrefix)) {
        this.#anchors.delete(key)
      }
    }

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Removes all shapes down an entire column.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} colIndex
   * @param {SlideManager} slideManager
   */
  removeColShapes(slideIndex, tableId, colIndex, slideManager, frame = null, tblObj = null) {
    if (frame && tblObj) {
      const colBounds = this.getTableColumnBounds(frame, tblObj, colIndex)
      if (colBounds) {
        return this.purgeShapesWithinBounds(slideIndex, colBounds, slideManager, tableId)
      }
    }

    const slideObj = slideManager.getSlideObj(slideIndex)
    if (!slideObj) return

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree) return

    const colPattern = new RegExp(`^(?:CellShape|cellshape)_${tableId}_r\\d+_c${colIndex}$`)

    for (const key of ['p:sp', 'p:cxnSp']) {
      if (!spTree[key]) continue
      const spList = Array.isArray(spTree[key]) ? spTree[key] : [spTree[key]]
      const remaining = []

      for (const sp of spList) {
        const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr'] || sp?.['p:nvCxnSpPr']?.['p:cNvPr']
        const name = cNvPr?.['@_name'] || ''
        const id = cNvPr?.['@_id'] ? String(cNvPr['@_id']) : ''

        if (colPattern.test(name) || colPattern.test(id)) {
          if (spTree[Z_ORDER_SYMBOL] && id) {
            const zIdx = spTree[Z_ORDER_SYMBOL].indexOf(id)
            if (zIdx !== -1) spTree[Z_ORDER_SYMBOL].splice(zIdx, 1)
          }
          continue
        }
        remaining.push(sp)
      }

      spTree[key] = remaining
    }

    // Clean matching keys from anchor map
    for (const [key, anchor] of Array.from(this.#anchors.entries())) {
      if (
        anchor.slideIndex === slideIndex &&
        anchor.tableId === tableId &&
        anchor.colIndex === colIndex
      ) {
        this.#anchors.delete(key)
      }
    }

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Removes all overlay shapes associated with a table.
   *
   * @param {number} slideIndex - Slide index.
   * @param {string} tableId - Table identifier.
   * @param {SlideManager} slideManager - SlideManager instance.
   */
  removeAllCellShapesForTable(slideIndex, tableId, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    if (!slideObj) return

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree) return

    const prefix = `CellShape_${tableId}_`
    const prefixOld = `cellshape_${tableId}_`

    for (const key of ['p:sp', 'p:cxnSp']) {
      if (!spTree[key]) continue
      const spList = Array.isArray(spTree[key]) ? spTree[key] : [spTree[key]]
      const remaining = []

      for (const sp of spList) {
        const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr'] || sp?.['p:nvCxnSpPr']?.['p:cNvPr']
        const name = cNvPr?.['@_name'] || ''
        const id = cNvPr?.['@_id'] ? String(cNvPr['@_id']) : ''
        if (name.startsWith(prefix) || name.startsWith(prefixOld) || id.startsWith(prefixOld)) {
          if (spTree[Z_ORDER_SYMBOL] && id) {
            const zIdx = spTree[Z_ORDER_SYMBOL].indexOf(id)
            if (zIdx !== -1) spTree[Z_ORDER_SYMBOL].splice(zIdx, 1)
          }
          continue
        }
        remaining.push(sp)
      }

      spTree[key] = remaining
    }

    // Clean anchor map
    for (const key of Array.from(this.#anchors.keys())) {
      if (key.startsWith(`${slideIndex}_${tableId}_`)) {
        this.#anchors.delete(key)
      }
    }

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Clones all cell shapes from a source table to a cloned table with coordinate offset.
   *
   * @param {number} sourceSlideIndex
   * @param {string} sourceTableId
   * @param {number} targetSlideIndex
   * @param {string} targetTableId
   * @param {SlideManager} slideManager
   * @param {number} deltaX - EMU horizontal offset.
   * @param {number} deltaY - EMU vertical offset.
   * @param {Object} [frame] - Source table graphic frame AST.
   * @param {Object} [tblObj] - Source table tbl AST.
   */
  cloneCellShapes(
    sourceSlideIndex,
    sourceTableId,
    targetSlideIndex,
    targetTableId,
    slideManager,
    deltaX = 0,
    deltaY = 0,
    frame = null,
    tblObj = null
  ) {
    const slideObj = slideManager.getSlideObj(sourceSlideIndex)
    if (!slideObj) return

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree) return

    const sourceName = frame?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_name'] || sourceTableId
    const sourceId = String(frame?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_id'] || '')
    const sourceBounds = frame && tblObj ? this.getTableBounds(frame, tblObj) : null

    const prefixes = new Set([`CellShape_${sourceTableId}_`, `cellshape_${sourceTableId}_`])
    if (sourceName) {
      prefixes.add(`CellShape_${sourceName}_`)
      prefixes.add(`cellshape_${sourceName}_`)
    }
    if (sourceId) {
      prefixes.add(`CellShape_${sourceId}_`)
      prefixes.add(`cellshape_${sourceId}_`)
    }

    const targetSlideObj = slideManager.getSlideObj(targetSlideIndex)
    const targetSpTree =
      targetSlideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      targetSlideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree']

    if (!targetSpTree) return
    if (!targetSpTree['p:sp']) targetSpTree['p:sp'] = []
    if (!Array.isArray(targetSpTree['p:sp'])) targetSpTree['p:sp'] = [targetSpTree['p:sp']]
    if (!targetSpTree['p:cxnSp']) targetSpTree['p:cxnSp'] = []
    if (!Array.isArray(targetSpTree['p:cxnSp'])) targetSpTree['p:cxnSp'] = [targetSpTree['p:cxnSp']]

    // Find highest existing shape ID in target slide
    const existingIds = this.#getAllShapeIds(targetSpTree)
    let nextId = existingIds.length > 0 ? Math.max(...existingIds) + 1 : 2000
    if (nextId > 2147483640) nextId = 2000

    const matchedShapes = new Map()

    for (const key of ['p:sp', 'p:cxnSp']) {
      if (!spTree[key]) continue
      const spList = Array.isArray(spTree[key]) ? spTree[key] : [spTree[key]]

      for (const sp of spList) {
        const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr'] || sp?.['p:nvCxnSpPr']?.['p:cNvPr']
        const name = cNvPr?.['@_name'] || ''
        const id = cNvPr?.['@_id'] ? String(cNvPr['@_id']) : ''

        let isMatch = false
        for (const p of prefixes) {
          if (name.startsWith(p) || id.startsWith(p)) {
            isMatch = true
            break
          }
        }

        if (!isMatch) {
          for (const anchor of this.#anchors.values()) {
            if (
              anchor.slideIndex === sourceSlideIndex &&
              (anchor.tableId === sourceTableId || (sourceName && anchor.tableId === sourceName)) &&
              String(anchor.shapeId) === id
            ) {
              isMatch = true
              break
            }
          }
        }

        if (!isMatch && sourceBounds) {
          const shapeBounds = this.getShapeBounds(sp)
          if (shapeBounds && this.isStrictlyContainedWithin(shapeBounds, sourceBounds)) {
            isMatch = true
          }
        }

        if (isMatch && !matchedShapes.has(sp)) {
          matchedShapes.set(sp, { sp, name, id, tag: key })
        }
      }
    }

    for (const { sp, name, id, tag } of matchedShapes.values()) {
      const cloned = this.#xmlParser.deepClone(sp)

      while (existingIds.includes(nextId)) {
        nextId++
      }
      const newIdStr = String(nextId++)
      existingIds.push(parseInt(newIdStr, 10))

      let r, c
      const match = name.match(/_r(\d+)_c(\d+)/)
      if (match) {
        r = parseInt(match[1], 10)
        c = parseInt(match[2], 10)
      } else {
        for (const anchor of this.#anchors.values()) {
          if (
            anchor.slideIndex === sourceSlideIndex &&
            String(anchor.shapeId) === id &&
            anchor.rowIndex !== undefined &&
            anchor.colIndex !== undefined
          ) {
            r = anchor.rowIndex
            c = anchor.colIndex
            break
          }
        }
      }

      let newName
      if (r !== undefined && c !== undefined) {
        newName = `CellShape_${targetTableId}_r${r}_c${c}`
      } else {
        newName = name
          .replace(sourceTableId, targetTableId)
          .replace(sourceName || '', targetTableId)
        if (newName === name) {
          newName = `CellShape_${targetTableId}_${newIdStr}`
        }
      }

      const cNvPr = cloned?.['p:nvSpPr']?.['p:cNvPr'] || cloned?.['p:nvCxnSpPr']?.['p:cNvPr']
      if (cNvPr) {
        cNvPr['@_name'] = newName
        cNvPr['@_id'] = newIdStr
      }

      if (cloned['p:spPr']?.['a:xfrm']?.['a:off']) {
        const off = cloned['p:spPr']['a:xfrm']['a:off']
        const curX = parseInt(off['@_x'] || '0', 10)
        const curY = parseInt(off['@_y'] || '0', 10)
        off['@_x'] = String(curX + deltaX)
        off['@_y'] = String(curY + deltaY)
      }

      if (tag === 'p:cxnSp') {
        targetSpTree['p:cxnSp'].push(cloned)
      } else {
        targetSpTree['p:sp'].push(cloned)
      }

      if (targetSpTree[Z_ORDER_SYMBOL] && !targetSpTree[Z_ORDER_SYMBOL].includes(newIdStr)) {
        targetSpTree[Z_ORDER_SYMBOL].push(newIdStr)
      }

      let text = ''
      const txBody = cloned['p:txBody']
      if (txBody?.['a:p']) {
        const pList = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
        for (const p of pList) {
          const rList = Array.isArray(p['a:r']) ? p['a:r'] : p['a:r'] ? [p['a:r']] : []
          for (const rNode of rList) {
            if (rNode?.['a:t'] !== undefined) text += String(rNode['a:t'])
          }
        }
      }

      if (r !== undefined && c !== undefined) {
        const origAnchorKey = `${sourceSlideIndex}_${sourceTableId}_${r}_${c}`
        const origAnchor =
          this.#anchors.get(origAnchorKey) ||
          (sourceName ? this.#anchors.get(`${sourceSlideIndex}_${sourceName}_${r}_${c}`) : null)

        const anchorKey = `${targetSlideIndex}_${targetTableId}_${r}_${c}`
        this.#anchors.set(anchorKey, {
          shapeId: newIdStr,
          shapeName: newName,
          slideIndex: targetSlideIndex,
          tableId: targetTableId,
          rowIndex: r,
          colIndex: c,
          text,
          config: origAnchor?.config,
        })
      }
    }

    slideManager.markSlideObjDirty(targetSlideIndex)
  }

  /**
   * Discovers all shape IDs present in an spTree container.
   * @private
   */
  #getAllShapeIds(spTree) {
    const ids = []
    const tags = ['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp']
    for (const tag of tags) {
      const items = spTree[tag] || []
      const arr = Array.isArray(items) ? items : [items]
      for (const item of arr) {
        const pr =
          item?.['p:nvSpPr'] ||
          item?.['p:nvPicPr'] ||
          item?.['p:nvGraphicFramePr'] ||
          item?.['p:nvGrpSpPr'] ||
          item?.['p:nvCxnSpPr']
        const idVal = pr?.['p:cNvPr']?.['@_id']
        if (idVal !== undefined) {
          const parsed = parseInt(idVal, 10)
          if (!isNaN(parsed)) ids.push(parsed)
        }
      }
    }
    return ids
  }

  /**
   * Retrieves registered cell shape metadata or finds it in slide shapes.
   *
   * @param {number} slideIndex - Slide index.
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} colIndex - 0-based col index.
   * @param {SlideManager} [slideManager] - SlideManager instance.
   * @returns {Object|null}
   */
  getCellShape(slideIndex, tableId, rowIndex, colIndex, slideManager) {
    const key = `${slideIndex}_${tableId}_${rowIndex}_${colIndex}`
    if (this.#anchors.has(key)) {
      return this.#anchors.get(key)
    }

    if (!slideManager) return null
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree || !spTree['p:sp']) return null

    const prefix = `CellShape_${tableId}_r${rowIndex}_c${colIndex}`
    const prefixOld = `cellshape_${tableId}_r${rowIndex}_c${colIndex}`
    const spList = Array.isArray(spTree['p:sp']) ? spTree['p:sp'] : [spTree['p:sp']]

    for (const sp of spList) {
      const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr']
      const name = cNvPr?.['@_name'] || ''
      const id = cNvPr?.['@_id'] || ''
      if (name.startsWith(prefix) || name.startsWith(prefixOld) || id.startsWith(prefixOld)) {
        let text = ''
        const txBody = sp?.['p:txBody']
        if (txBody?.['a:p']) {
          const pList = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
          for (const p of pList) {
            const rList = Array.isArray(p['a:r']) ? p['a:r'] : p['a:r'] ? [p['a:r']] : []
            for (const r of rList) {
              if (r?.['a:t'] !== undefined) text += String(r['a:t'])
            }
          }
        }
        return {
          shapeId: id,
          shapeName: name,
          slideIndex,
          tableId,
          rowIndex,
          colIndex,
          text,
          shape: sp,
        }
      }
    }
    return null
  }

  /**
   * Retrieves text from inside an overlay shape if one exists.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} rowIndex
   * @param {number} colIndex
   * @param {SlideManager} [slideManager]
   * @returns {string|null}
   */
  getCellShapeText(slideIndex, tableId, rowIndex, colIndex, slideManager) {
    const shape = this.getCellShape(slideIndex, tableId, rowIndex, colIndex, slideManager)
    return shape?.text || null
  }
}

module.exports = {
  CellShapeManager,
  SUPPORTED_SHAPE_TYPES,
  isStrictlyContainedWithin,
}
