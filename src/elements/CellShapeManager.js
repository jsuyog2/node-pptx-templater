/**
 * @fileoverview CellShapeManager - Manages table cell shape overlays, polymorphic cell data,
 * dynamic row/column shifting, cloning, and cascading deletions.
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
])

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
   * Checks whether a cell value specifies a shape overlay.
   *
   * @param {*} val - Cell value.
   * @returns {boolean}
   */
  isCellShape(val) {
    if (!val || typeof val !== 'object' || Array.isArray(val)) return false
    if (val.shape !== undefined) return true
    if (typeof val.type === 'string' && SUPPORTED_SHAPE_TYPES.has(val.type)) return true
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
   * Calculates rendered cell bounding coordinates in EMUs.
   *
   * @param {Object} frame - Table <p:graphicFrame> AST.
   * @param {Object} tblObj - Table <a:tbl> AST.
   * @param {number} targetRow - 0-based row index.
   * @param {number} targetCol - 0-based column index.
   * @returns {{ x: number, y: number, width: number, height: number, xPx: number, yPx: number, widthPx: number, heightPx: number } | null}
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
      xPx: Math.round(emuToPx(currentX)),
      yPx: Math.round(emuToPx(currentY)),
      widthPx: Math.round(emuToPx(cellWidth)),
      heightPx: Math.round(emuToPx(cellHeight)),
    }
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

    // 1. Remove any previous overlay shape for this specific cell
    this.removeCellShape(slideIndex, tableId, rowIndex, colIndex, slideManager)

    // 2. Compute exact cell bounding box in slide EMUs
    const bounds = this.calculateCellBounds(frame, tblObj, rowIndex, colIndex)
    if (!bounds || bounds.width <= 0 || bounds.height <= 0) {
      logger.warn(`Could not compute cell bounds for ${tableId} cell (${rowIndex}, ${colIndex})`)
      return null
    }

    const shapeOpt = shapeConfig.shape || shapeConfig
    let shapeType = shapeOpt.type || shapeOpt.preset || 'roundRect'
    if (shapeType === 'pill') shapeType = 'roundRect'
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

    // Clamp dimensions to cell bounds with a margin
    shapeW = Math.min(shapeW, bounds.width)
    shapeH = Math.min(shapeH, bounds.height)

    // 4. Compute shape position coordinates
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
    const shapeY = baseY

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
    const shapeX = baseX

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
   * @returns {Array<{ sp: Object, name: string, id: string, r: number, c: number }>}
   */
  getAllCellShapesForTable(slideIndex, tableId, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    if (!slideObj) return []

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree || !spTree['p:sp']) return []

    const spList = Array.isArray(spTree['p:sp']) ? spTree['p:sp'] : [spTree['p:sp']]
    const results = []

    const prefix = `CellShape_${tableId}_`
    const prefixOld = `cellshape_${tableId}_`

    for (const sp of spList) {
      const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr']
      const name = cNvPr?.['@_name'] || ''
      const id = cNvPr?.['@_id'] ? String(cNvPr['@_id']) : ''

      if (name.startsWith(prefix) || name.startsWith(prefixOld) || id.startsWith(prefixOld)) {
        const match = name.match(/_r(\d+)_c(\d+)/)
        if (match) {
          const r = parseInt(match[1], 10)
          const c = parseInt(match[2], 10)
          results.push({ sp, name, id, r, c })
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
      if (item.sp?.['p:nvSpPr']?.['p:cNvPr']) {
        item.sp['p:nvSpPr']['p:cNvPr']['@_name'] = newName
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
   * Removes all shapes on a removed row and shifts lower shapes upward.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} removedRowIndex - 0-based index of removed row.
   * @param {number} removedHeightEmu - Height of removed row in EMUs.
   * @param {SlideManager} slideManager
   */
  removeRowShapesAndShift(slideIndex, tableId, removedRowIndex, removedHeightEmu, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    if (!slideObj) return

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree || !spTree['p:sp']) return

    // 1. Purge shapes at removedRowIndex
    this.removeRowShapes(slideIndex, tableId, removedRowIndex, slideManager)

    // 2. Shift shapes where r > removedRowIndex upward
    const shapes = this.getAllCellShapesForTable(slideIndex, tableId, slideManager)
    const targetShapes = shapes.filter(s => s.r > removedRowIndex)

    // Process from smallest row index upward
    targetShapes.sort((a, b) => a.r - b.r)

    for (const item of targetShapes) {
      const oldR = item.r
      const newR = oldR - 1
      const col = item.c
      const newName = `CellShape_${tableId}_r${newR}_c${col}`

      if (item.sp?.['p:nvSpPr']?.['p:cNvPr']) {
        item.sp['p:nvSpPr']['p:cNvPr']['@_name'] = newName
      }

      if (item.sp?.['p:spPr']?.['a:xfrm']?.['a:off']) {
        const off = item.sp['p:spPr']['a:xfrm']['a:off']
        const curY = parseInt(off['@_y'] || '0', 10)
        off['@_y'] = String(Math.max(0, curY - removedHeightEmu))
      }

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

      if (item.sp?.['p:nvSpPr']?.['p:cNvPr']) {
        item.sp['p:nvSpPr']['p:cNvPr']['@_name'] = newName
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
   * Removes all shapes on a removed column and shifts right-side shapes leftward.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} removedColIndex - 0-based index of removed column.
   * @param {number} removedWidthEmu - Width of removed column in EMUs.
   * @param {SlideManager} slideManager
   */
  removeColShapesAndShift(slideIndex, tableId, removedColIndex, removedWidthEmu, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    if (!slideObj) return

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree || !spTree['p:sp']) return

    // 1. Purge shapes at removedColIndex
    this.removeColShapes(slideIndex, tableId, removedColIndex, slideManager)

    // 2. Shift shapes where c > removedColIndex leftward
    const shapes = this.getAllCellShapesForTable(slideIndex, tableId, slideManager)
    const targetShapes = shapes.filter(s => s.c > removedColIndex)

    // Process from smallest column index upward
    targetShapes.sort((a, b) => a.c - b.c)

    for (const item of targetShapes) {
      const row = item.r
      const oldC = item.c
      const newC = oldC - 1
      const newName = `CellShape_${tableId}_r${row}_c${newC}`

      if (item.sp?.['p:nvSpPr']?.['p:cNvPr']) {
        item.sp['p:nvSpPr']['p:cNvPr']['@_name'] = newName
      }

      if (item.sp?.['p:spPr']?.['a:xfrm']?.['a:off']) {
        const off = item.sp['p:spPr']['a:xfrm']['a:off']
        const curX = parseInt(off['@_x'] || '0', 10)
        off['@_x'] = String(Math.max(0, curX - removedWidthEmu))
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
   * Removes overlay shape for a specific cell.
   *
   * @param {number} slideIndex - Slide index.
   * @param {string} tableId - Table identifier.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} colIndex - 0-based column index.
   * @param {SlideManager} slideManager - SlideManager instance.
   */
  removeCellShape(slideIndex, tableId, rowIndex, colIndex, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    if (!slideObj) return

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree || !spTree['p:sp']) return

    const prefix = `CellShape_${tableId}_r${rowIndex}_c${colIndex}`
    const prefixOld = `cellshape_${tableId}_r${rowIndex}_c${colIndex}`
    const anchorKey = `${slideIndex}_${tableId}_${rowIndex}_${colIndex}`
    const anchor = this.#anchors.get(anchorKey)
    const targetId = anchor?.shapeId ? String(anchor.shapeId) : null

    const spList = Array.isArray(spTree['p:sp']) ? spTree['p:sp'] : [spTree['p:sp']]
    const remaining = []

    for (const sp of spList) {
      const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr']
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

    spTree['p:sp'] = remaining
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
  removeRowShapes(slideIndex, tableId, rowIndex, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    if (!slideObj) return

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree || !spTree['p:sp']) return

    const rowPattern = new RegExp(`^(?:CellShape|cellshape)_${tableId}_r${rowIndex}_c\\d+`)
    const spList = Array.isArray(spTree['p:sp']) ? spTree['p:sp'] : [spTree['p:sp']]
    const remaining = []

    for (const sp of spList) {
      const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr']
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

    spTree['p:sp'] = remaining

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
  removeColShapes(slideIndex, tableId, colIndex, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    if (!slideObj) return

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree || !spTree['p:sp']) return

    const colPattern = new RegExp(`^(?:CellShape|cellshape)_${tableId}_r\\d+_c${colIndex}$`)
    const spList = Array.isArray(spTree['p:sp']) ? spTree['p:sp'] : [spTree['p:sp']]
    const remaining = []

    for (const sp of spList) {
      const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr']
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

    spTree['p:sp'] = remaining

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
    if (!spTree || !spTree['p:sp']) return

    const prefix = `CellShape_${tableId}_`
    const prefixOld = `cellshape_${tableId}_`

    const spList = Array.isArray(spTree['p:sp']) ? spTree['p:sp'] : [spTree['p:sp']]
    const remaining = []

    for (const sp of spList) {
      const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr']
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

    spTree['p:sp'] = remaining

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
   */
  cloneCellShapes(
    sourceSlideIndex,
    sourceTableId,
    targetSlideIndex,
    targetTableId,
    slideManager,
    deltaX = 0,
    deltaY = 0
  ) {
    const slideObj = slideManager.getSlideObj(sourceSlideIndex)
    if (!slideObj) return

    const spTree =
      slideObj['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj['p:sldLayout']?.['p:cSld']?.['p:spTree']
    if (!spTree || !spTree['p:sp']) return

    const prefix = `CellShape_${sourceTableId}_`
    const prefixOld = `cellshape_${sourceTableId}_`

    const spList = Array.isArray(spTree['p:sp']) ? spTree['p:sp'] : [spTree['p:sp']]
    const targetSlideObj = slideManager.getSlideObj(targetSlideIndex)
    const targetSpTree =
      targetSlideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      targetSlideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree']

    if (!targetSpTree) return
    if (!targetSpTree['p:sp']) targetSpTree['p:sp'] = []
    if (!Array.isArray(targetSpTree['p:sp'])) targetSpTree['p:sp'] = [targetSpTree['p:sp']]

    // Find highest existing shape ID in target slide
    const existingIds = this.#getAllShapeIds(targetSpTree)
    let nextId = existingIds.length > 0 ? Math.max(...existingIds) + 1 : 2000

    for (const sp of spList) {
      const cNvPr = sp?.['p:nvSpPr']?.['p:cNvPr']
      const name = cNvPr?.['@_name'] || ''
      if (name.startsWith(prefix) || name.startsWith(prefixOld)) {
        const cloned = this.#xmlParser.deepClone(sp)
        const newName = name.replace(sourceTableId, targetTableId)
        const newIdStr = String(nextId++)

        if (cloned['p:nvSpPr']?.['p:cNvPr']) {
          cloned['p:nvSpPr']['p:cNvPr']['@_name'] = newName
          cloned['p:nvSpPr']['p:cNvPr']['@_id'] = newIdStr
        }
        if (cloned['p:spPr']?.['a:xfrm']?.['a:off']) {
          const off = cloned['p:spPr']['a:xfrm']['a:off']
          const curX = parseInt(off['@_x'] || '0', 10)
          const curY = parseInt(off['@_y'] || '0', 10)
          off['@_x'] = String(curX + deltaX)
          off['@_y'] = String(curY + deltaY)
        }
        targetSpTree['p:sp'].push(cloned)

        if (targetSpTree[Z_ORDER_SYMBOL] && !targetSpTree[Z_ORDER_SYMBOL].includes(newIdStr)) {
          targetSpTree[Z_ORDER_SYMBOL].push(newIdStr)
        }

        // Parse row and col from name: CellShape_{targetTableId}_r{rowIndex}_c{colIndex}
        const match = newName.match(/_r(\d+)_c(\d+)/)
        if (match) {
          const r = parseInt(match[1], 10)
          const c = parseInt(match[2], 10)
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

          const origAnchorKey = `${sourceSlideIndex}_${sourceTableId}_${r}_${c}`
          const origAnchor = this.#anchors.get(origAnchorKey)

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
}
