/**
 * @fileoverview TableManager - Updates table data in PPTX slides.
 *
 * Tables in OpenXML PPTX:
 * ─────────────────────────────────────────────────────────────────
 * Tables are stored as a:tbl inside a p:graphicFrame shape.
 *
 *   <p:graphicFrame>
 *     <p:nvGraphicFramePr>
 *       <p:cNvPr id="6" name="employees-table"/>   ← table name
 *     </p:nvGraphicFramePr>
 *     <a:graphic>
 *       <a:graphicData uri="...table">
 *         <a:tbl>
 *           <a:tblGrid>
 *             <a:gridCol w="2286000"/>    ← column widths
 *           </a:tblGrid>
 *           <a:tr h="370840">             ← row (height in EMUs)
 *             <a:tc>                      ← table cell
 *               <a:txBody>
 *                 <a:p>
 *                   <a:r><a:t>Cell text</a:t></a:r>
 *                 </a:p>
 *               </a:txBody>
 *               <a:tcPr/>                 ← cell properties (borders, fills)
 *             </a:tc>
 *           </a:tr>
 *         </a:tbl>
 *       </a:graphicData>
 *     </a:graphic>
 *   </p:graphicFrame>
 *
 * Key challenge: Preserving cell formatting while replacing text.
 * We ONLY modify the <a:t> text nodes, keeping all <a:tcPr>, <a:rPr>, etc.
 * And critical to avoid PPT corruption: update the val of <a16:rowId> in each cloned row's <a:extLst>.
 */

const { createLogger } = require('../utils/logger.js')
const { TableNotFoundError, PPTXError } = require('../utils/errors.js')

const logger = createLogger('TableManager')

/**
 * @class TableManager
 * @description Handles table operations in PPTX templates.
 */
class TableManager {
  /** @private @type {XMLParser} */
  #xmlParser

  /**
   * @param {XMLParser} xmlParser
   */
  constructor(xmlParser) {
    this.#xmlParser = xmlParser
  }

  /**
   * Updates a table with new row data.
   * Finds the table by name/ID and replaces its row content.
   * Preserves all formatting properties.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string} tableId - Table name or shape ID.
   * @param {string[][]} rows - 2D array of cell values [row][col].
   * @param {SlideManager} slideManager
   * @throws {TableNotFoundError} If the table is not found.
   */
  updateTable(slideIndex, tableId, data, slideManager, shapeManager) {
    const { tblObj, frameObj, resolvedTableId } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )

    const trs = tblObj['a:tr'] || []
    if (trs.length === 0) {
      logger.warn('No rows found in table XML template')
      return
    }

    let rowsData = []
    let templateMerges = []
    let cellShapes = null

    if (Array.isArray(data)) {
      rowsData = data
    } else if (data && typeof data === 'object') {
      rowsData = data.rows || []
      templateMerges = data.merge || []
      cellShapes = data.cellShapes || null
    }

    const headerTemplate = trs[0]
    const dataTemplate = trs[1] || trs[0]

    const newRows = []
    const generatedMerges = []

    const headerNames = (trs[0]['a:tc'] || []).map(cell => this.#getCellText(cell).trim())
    const isObjectRows =
      rowsData.length > 0 && !Array.isArray(rowsData[0]) && typeof rowsData[0] === 'object'

    if (isObjectRows) {
      // 1. Keep/clone the header row
      const headerRow = this.#xmlParser.deepClone(headerTemplate)
      this.#updateRowId(headerRow)
      newRows.push(headerRow)

      // 2. Map objects to data rows
      for (let i = 0; i < rowsData.length; i++) {
        const newRow = this.#xmlParser.deepClone(dataTemplate)
        this.#updateRowId(newRow)

        const tcs = newRow['a:tc'] || []
        const rowObj = rowsData[i]

        for (let j = 0; j < tcs.length; j++) {
          const headerName = headerNames[j]
          let rawCell = undefined
          if (headerName) {
            if (rowObj[headerName] !== undefined) {
              rawCell = rowObj[headerName]
            } else if (rowObj[headerName.toLowerCase()] !== undefined) {
              rawCell = rowObj[headerName.toLowerCase()]
            }
          }
          if (rawCell === undefined && rowObj[j] !== undefined) {
            rawCell = rowObj[j]
          }
          if (rawCell === undefined) {
            rawCell = ''
          }

          let val = ''
          let cellOptions = {}

          if (tcs[j]['@_hMerge']) delete tcs[j]['@_hMerge']
          if (tcs[j]['@_vMerge']) delete tcs[j]['@_vMerge']
          if (tcs[j]['@_gridSpan']) delete tcs[j]['@_gridSpan']
          if (tcs[j]['@_rowSpan']) delete tcs[j]['@_rowSpan']

          if (rawCell && typeof rawCell === 'object') {
            val = rawCell.value !== undefined ? rawCell.value : ''
            const rowSpan = parseInt(rawCell.rowSpan || 1, 10)
            const colSpan = parseInt(rawCell.colSpan || rawCell.gridSpan || 1, 10)
            if (rowSpan > 1 || colSpan > 1) {
              generatedMerges.push({
                startRow: i + 1,
                startCol: j,
                endRow: i + 1 + rowSpan - 1,
                endCol: j + colSpan - 1,
              })
            }
            cellOptions = rawCell
          } else {
            val = String(rawCell)
          }

          this.#setCellTextObj(tcs[j], val)
          this.#applyCellOptions(tcs[j], cellOptions)
        }
        newRows.push(newRow)
      }
    } else {
      // 2D array mapping
      for (let i = 0; i < rowsData.length; i++) {
        const template = i === 0 ? headerTemplate : trs[i] || dataTemplate
        const newRow = this.#xmlParser.deepClone(template)
        this.#updateRowId(newRow)

        const tcs = newRow['a:tc'] || []
        const rowData = rowsData[i]

        for (let j = 0; j < tcs.length; j++) {
          const rawCell = rowData && rowData[j] !== undefined ? rowData[j] : ''
          let val = ''
          let cellOptions = {}

          if (tcs[j]['@_hMerge']) delete tcs[j]['@_hMerge']
          if (tcs[j]['@_vMerge']) delete tcs[j]['@_vMerge']
          if (tcs[j]['@_gridSpan']) delete tcs[j]['@_gridSpan']
          if (tcs[j]['@_rowSpan']) delete tcs[j]['@_rowSpan']

          if (rawCell && typeof rawCell === 'object') {
            val = rawCell.value !== undefined ? rawCell.value : ''
            const rowSpan = parseInt(rawCell.rowSpan || 1, 10)
            const colSpan = parseInt(rawCell.colSpan || rawCell.gridSpan || 1, 10)
            if (rowSpan > 1 || colSpan > 1) {
              generatedMerges.push({
                startRow: i,
                startCol: j,
                endRow: i + rowSpan - 1,
                endCol: j + colSpan - 1,
              })
            }
            cellOptions = rawCell
          } else {
            val = String(rawCell)
          }

          this.#setCellTextObj(tcs[j], val)
          this.#applyCellOptions(tcs[j], cellOptions)
        }
        newRows.push(newRow)
      }
    }

    tblObj['a:tr'] = newRows

    slideManager.markSlideObjDirty(slideIndex)

    const finalMerges = [...templateMerges, ...generatedMerges]
    for (const merge of finalMerges) {
      this.mergeCells(
        slideIndex,
        tableId,
        merge.startRow,
        merge.startCol,
        merge.endRow,
        merge.endCol,
        slideManager
      )
    }

    this.#calculateRowHeights(slideIndex, tableId, slideManager, tblObj)

    if (cellShapes) {
      this.#processCellShapes(
        slideIndex,
        tableId,
        resolvedTableId,
        rowsData,
        isObjectRows,
        cellShapes,
        slideManager,
        shapeManager,
        tblObj,
        frameObj
      )
    }

    logger.debug(
      `Updated table "${tableId}" with ${rowsData.length} rows and ${finalMerges.length} merges`
    )
  }

  /**
   * Adds a row at the end of the table.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {string[]} rowData
   * @param {SlideManager} slideManager
   */
  addTableRow(slideIndex, tableId, rowData, slideManager, options = {}) {
    const { tblObj } = this.#getTableContext(slideIndex, tableId, slideManager)

    const trs = tblObj['a:tr'] || []
    if (trs.length === 0) {
      throw new PPTXError('No rows to clone from')
    }

    const lastRow = trs[trs.length - 1]
    const numCols = lastRow['a:tc']?.length || 0

    // Compute target generated height
    const heights = []
    for (let c = 0; c < numCols; c++) {
      heights.push(this.#getNestedHeight(rowData[c]))
    }
    const targetHeight = Math.max(1, ...heights)

    // Expand each column value to targetHeight
    const expandedCols = []
    const strategy = options.mergeStrategy || 'auto'
    for (let c = 0; c < numCols; c++) {
      let colCells = this.#expandCellVal(rowData[c], targetHeight)
      if (strategy === 'none') {
        for (let i = 0; i < colCells.length; i++) {
          if (colCells[i].vMerge) {
            colCells[i] = { value: '', rowSpan: 1 }
          } else {
            colCells[i].rowSpan = 1
          }
        }
      } else if (strategy === 'auto') {
        colCells = this.#applyAutoMerge(colCells)
      }
      expandedCols.push(colCells)
    }

    // Clone and append rows
    for (let r = 0; r < targetHeight; r++) {
      const newRow = this.#xmlParser.deepClone(lastRow)
      this.#updateRowId(newRow)

      const tcs = newRow['a:tc'] || []
      for (let c = 0; c < numCols; c++) {
        const cellDef = expandedCols[c][r]
        const tcObj = tcs[c]

        // Clear any previous merge attributes
        if (tcObj['@_hMerge']) delete tcObj['@_hMerge']
        if (tcObj['@_vMerge']) delete tcObj['@_vMerge']
        if (tcObj['@_gridSpan']) delete tcObj['@_gridSpan']
        if (tcObj['@_rowSpan']) delete tcObj['@_rowSpan']

        if (cellDef.vMerge) {
          tcObj['@_vMerge'] = '1'
          this.#setCellTextObj(tcObj, '')
        } else {
          let text = cellDef.value
          let cellOpts = {}
          if (cellDef.value && typeof cellDef.value === 'object') {
            text = cellDef.value.value !== undefined ? cellDef.value.value : ''
            cellOpts = cellDef.value
          }
          this.#setCellTextObj(tcObj, text)
          if (cellDef.rowSpan && cellDef.rowSpan > 1 && strategy !== 'none') {
            tcObj['@_rowSpan'] = String(cellDef.rowSpan)
          }
          this.#applyCellOptions(tcObj, cellOpts)
        }
      }
      trs.push(newRow)
    }

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Removes a table row by index.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} rowIndex - 0-based row index.
   * @param {SlideManager} slideManager
   */
  removeTableRow(slideIndex, tableId, rowIndex, slideManager) {
    const { tblObj } = this.#getTableContext(slideIndex, tableId, slideManager)

    const trs = tblObj['a:tr'] || []
    if (rowIndex < 0 || rowIndex >= trs.length) {
      throw new PPTXError(`Row index ${rowIndex} out of bounds`)
    }

    trs.splice(rowIndex, 1)

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Inserts a table row at a specific index.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} rowIndex
   * @param {string[]} rowData
   * @param {SlideManager} slideManager
   */
  insertTableRow(slideIndex, tableId, rowIndex, rowData, slideManager) {
    const { tblObj } = this.#getTableContext(slideIndex, tableId, slideManager)

    const trs = tblObj['a:tr'] || []
    if (rowIndex < 0 || rowIndex > trs.length) {
      throw new PPTXError(`Row index ${rowIndex} out of bounds`)
    }

    // Use row above or first row as template
    const templateIndex = Math.max(0, rowIndex - 1)
    const template = trs[templateIndex] || trs[0]
    if (!template) {
      throw new PPTXError('No rows to insert and copy from')
    }

    const newRow = this.#xmlParser.deepClone(template)
    this.#updateRowId(newRow)

    const tcs = newRow['a:tc'] || []
    for (let j = 0; j < tcs.length; j++) {
      this.#setCellTextObj(tcs[j], rowData[j] !== undefined ? rowData[j] : '')
      if (tcs[j]['@_hMerge']) delete tcs[j]['@_hMerge']
      if (tcs[j]['@_vMerge']) delete tcs[j]['@_vMerge']
    }

    trs.splice(rowIndex, 0, newRow)

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Clones a row and inserts it at another index.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} sourceRowIndex
   * @param {number} targetRowIndex
   * @param {SlideManager} slideManager
   */
  cloneTableRow(slideIndex, tableId, sourceRowIndex, targetRowIndex, slideManager) {
    const { tblObj } = this.#getTableContext(slideIndex, tableId, slideManager)

    const trs = tblObj['a:tr'] || []
    if (sourceRowIndex < 0 || sourceRowIndex >= trs.length) {
      throw new PPTXError(`Source row index ${sourceRowIndex} out of bounds`)
    }
    if (targetRowIndex < 0 || targetRowIndex > trs.length) {
      throw new PPTXError(`Target row index ${targetRowIndex} out of bounds`)
    }

    const template = trs[sourceRowIndex]
    const newRow = this.#xmlParser.deepClone(template)
    this.#updateRowId(newRow)

    trs.splice(targetRowIndex, 0, newRow)

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Updates a single cell text and formatting.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} rowIndex
   * @param {number} colIndex
   * @param {string} value
   * @param {Object} options
   * @param {SlideManager} slideManager
   */
  updateCell(slideIndex, tableId, rowIndex, colIndex, value, options = {}, slideManager) {
    const { tblObj } = this.#getTableContext(slideIndex, tableId, slideManager)

    const row = tblObj['a:tr']?.[rowIndex]
    if (!row) {
      throw new PPTXError(`Row index ${rowIndex} out of bounds`)
    }

    const cell = row['a:tc']?.[colIndex]
    if (!cell) {
      throw new PPTXError(`Column index ${colIndex} out of bounds`)
    }

    this.#setCellTextObj(cell, value)

    if (options.fill) {
      if (!cell['a:tcPr']) cell['a:tcPr'] = {}
      cell['a:tcPr']['a:solidFill'] = {
        'a:srgbClr': { '@_val': options.fill },
      }
    }

    if (options.align) {
      const txBody = cell['a:txBody']
      if (txBody && txBody['a:p']) {
        const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
        for (const p of paras) {
          if (!p['a:pPr']) p['a:pPr'] = {}
          p['a:pPr']['@_algn'] = options.align
        }
      }
    }

    if (options.fontSize) {
      const sizeVal = options.fontSize * 100
      const txBody = cell['a:txBody']
      if (txBody && txBody['a:p']) {
        const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
        for (const p of paras) {
          if (p['a:r']) {
            const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
            for (const r of runs) {
              if (!r['a:rPr']) r['a:rPr'] = {}
              r['a:rPr']['@_sz'] = String(sizeVal)
            }
          }
        }
      }
    }

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Validates if a merge region can be applied to the table.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} startRow
   * @param {number} startCol
   * @param {number} endRow
   * @param {number} endCol
   * @param {SlideManager} slideManager
   * @returns {Object} Validation report { valid: boolean, errors: string[] }
   */
  validateMergeRegion(slideIndex, tableId, startRow, startCol, endRow, endCol, slideManager) {
    const errors = []
    const slideObj = slideManager.getSlideObj(slideIndex)
    const tblObj = this.#findTableObj(slideObj, tableId, slideManager, slideIndex)
    if (!tblObj) {
      errors.push(`Table "${tableId}" not found in slide ${slideIndex}`)
      return { valid: false, errors }
    }

    const trs = tblObj['a:tr'] || []
    const numRows = trs.length
    const numCols = trs[0]?.['a:tc']?.length || 0

    if (startRow < 0 || startRow >= numRows) {
      errors.push(`startRow ${startRow} is out of bounds (table has ${numRows} rows)`)
    }
    if (endRow < 0 || endRow >= numRows) {
      errors.push(`endRow ${endRow} is out of bounds (table has ${numRows} rows)`)
    }
    if (startCol < 0 || startCol >= numCols) {
      errors.push(`startCol ${startCol} is out of bounds (table has ${numCols} cols)`)
    }
    if (endCol < 0 || endCol >= numCols) {
      errors.push(`endCol ${endCol} is out of bounds (table has ${numCols} cols)`)
    }

    if (errors.length > 0) {
      return { valid: false, errors }
    }

    if (endRow < startRow) {
      errors.push(`endRow (${endRow}) cannot be less than startRow (${startRow})`)
    }
    if (endCol < startCol) {
      errors.push(`endCol (${endCol}) cannot be less than startCol (${startCol})`)
    }

    if (errors.length > 0) {
      return { valid: false, errors }
    }

    const existingMerges = this.getMergedCells(slideIndex, tableId, slideManager)
    for (const R of existingMerges) {
      const overlap =
        startRow <= R.endRow && endRow >= R.startRow && startCol <= R.endCol && endCol >= R.startCol
      if (overlap) {
        errors.push(
          `Requested merge region (${startRow}, ${startCol}) to (${endRow}, ${endCol}) overlaps with existing merged region (${R.startRow}, ${R.startCol}) to (${R.endRow}, ${R.endCol})`
        )
      }
    }

    return {
      valid: errors.length === 0,
      errors,
    }
  }

  /**
   * Merges cells in a table.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} startRow
   * @param {number} startCol
   * @param {number} endRow
   * @param {number} endCol
   * @param {SlideManager} slideManager
   */
  mergeCells(slideIndex, tableId, startRow, startCol, endRow, endCol, slideManager) {
    const validation = this.validateMergeRegion(
      slideIndex,
      tableId,
      startRow,
      startCol,
      endRow,
      endCol,
      slideManager
    )
    if (!validation.valid) {
      throw new PPTXError(`Invalid merge region: ${validation.errors.join('; ')}`)
    }

    const { tblObj } = this.#getTableContext(slideIndex, tableId, slideManager)
    const trs = tblObj['a:tr'] || []

    const allTexts = []

    for (let r = startRow; r <= endRow; r++) {
      const row = trs[r]
      if (!row) continue
      const tcs = row['a:tc'] || []

      for (let c = startCol; c <= endCol; c++) {
        const cell = tcs[c]
        if (!cell) continue

        if (r === startRow && c === startCol) {
          const text = this.#getCellText(cell)
          if (text) allTexts.push(text)
          if (cell['@_hMerge'] !== undefined) delete cell['@_hMerge']
          if (cell['@_vMerge'] !== undefined) delete cell['@_vMerge']
        } else {
          const text = this.#getCellText(cell)
          if (text) allTexts.push(text)

          if (cell['@_gridSpan'] !== undefined) delete cell['@_gridSpan']
          if (cell['@_rowSpan'] !== undefined) delete cell['@_rowSpan']

          if (r === startRow) {
            cell['@_hMerge'] = '1'
            if (cell['@_vMerge'] !== undefined) delete cell['@_vMerge']
          } else if (c === startCol) {
            cell['@_vMerge'] = '1'
            if (cell['@_hMerge'] !== undefined) delete cell['@_hMerge']
          } else {
            cell['@_hMerge'] = '1'
            cell['@_vMerge'] = '1'
          }
          this.#setCellTextObj(cell, '')
        }
      }
    }

    const originCell = trs[startRow]['a:tc'][startCol]
    if (endCol > startCol) {
      originCell['@_gridSpan'] = String(endCol - startCol + 1)
    }
    if (endRow > startRow) {
      originCell['@_rowSpan'] = String(endRow - startRow + 1)
    }

    const combinedText = allTexts.filter(t => t.trim() !== '').join('\n')
    this.#setCellTextObj(originCell, combinedText)

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Unmerges cells in a table.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} startRow
   * @param {number} startCol
   * @param {number} endRow
   * @param {number} endCol
   * @param {SlideManager} slideManager
   */
  unmergeCells(slideIndex, tableId, startRow, startCol, endRow, endCol, slideManager) {
    let actualSlideManager = slideManager
    let actualEndRow = endRow
    let actualEndCol = endCol

    if (typeof endRow === 'object' && endRow.getSlideXml) {
      actualSlideManager = endRow
      actualEndRow = undefined
      actualEndCol = undefined
    }

    const slideObj = actualSlideManager.getSlideObj(slideIndex)
    const tblObj = this.#findTableObj(slideObj, tableId, actualSlideManager, slideIndex)
    if (!tblObj) {
      throw new TableNotFoundError(`Table "${tableId}" not found in slide ${slideIndex}`)
    }

    const trs = tblObj['a:tr'] || []

    if (actualEndRow === undefined || actualEndCol === undefined) {
      const R = this.getMergeRegion(slideIndex, tableId, startRow, startCol, actualSlideManager)
      if (!R) return

      for (let r = R.startRow; r <= R.endRow; r++) {
        const rowObj = trs[r]
        if (!rowObj) continue
        const tcs = rowObj['a:tc'] || []
        for (let c = R.startCol; c <= R.endCol; c++) {
          const cell = tcs[c]
          if (!cell) continue
          if (cell['@_hMerge'] !== undefined) delete cell['@_hMerge']
          if (cell['@_vMerge'] !== undefined) delete cell['@_vMerge']
          if (cell['@_gridSpan'] !== undefined) delete cell['@_gridSpan']
          if (cell['@_rowSpan'] !== undefined) delete cell['@_rowSpan']
        }
      }
    } else {
      for (let r = startRow; r <= actualEndRow; r++) {
        const rowObj = trs[r]
        if (!rowObj) continue
        const tcs = rowObj['a:tc'] || []
        for (let c = startCol; c <= actualEndCol; c++) {
          const cell = tcs[c]
          if (!cell) continue
          if (cell['@_hMerge'] !== undefined) delete cell['@_hMerge']
          if (cell['@_vMerge'] !== undefined) delete cell['@_vMerge']
          if (cell['@_gridSpan'] !== undefined) delete cell['@_gridSpan']
          if (cell['@_rowSpan'] !== undefined) delete cell['@_rowSpan']
        }
      }
    }

    actualSlideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Scans the table grid and returns all merged regions.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {SlideManager} slideManager
   * @returns {Array<Object>} List of merged region coordinates
   */
  getMergedCells(slideIndex, tableId, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    const tblObj = this.#findTableObj(slideObj, tableId, slideManager, slideIndex)
    if (!tblObj) {
      throw new TableNotFoundError(`Table "${tableId}" not found in slide ${slideIndex}`)
    }

    const trs = tblObj['a:tr'] || []
    const merged = []

    for (let r = 0; r < trs.length; r++) {
      const row = trs[r]
      const tcs = row['a:tc'] || []
      for (let c = 0; c < tcs.length; c++) {
        const cell = tcs[c]
        if (!cell) continue

        const gridSpan = parseInt(cell['@_gridSpan'] || 1, 10)
        const rowSpan = parseInt(cell['@_rowSpan'] || 1, 10)

        if (gridSpan > 1 || rowSpan > 1) {
          merged.push({
            startRow: r,
            startCol: c,
            endRow: r + rowSpan - 1,
            endCol: c + gridSpan - 1,
          })
        }
      }
    }
    return merged
  }

  /**
   * Helper to get clean string content of a cell.
   */
  #getCellText(cellObj) {
    const txBody = cellObj?.['a:txBody']
    if (!txBody) return ''
    const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
    const text = []
    for (const p of paras) {
      let pText = ''
      if (p['a:r']) {
        const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
        for (const r of runs) {
          if (r['a:t']) {
            pText += String(r['a:t'])
          }
        }
      }
      text.push(pText)
    }
    return text.join('\n')
  }

  /**
   * Returns true if the cell is part of any merged region.
   */
  isMergedCell(slideIndex, tableId, row, col, slideManager) {
    return this.getMergeRegion(slideIndex, tableId, row, col, slideManager) !== null
  }

  /**
   * Returns the region coordinate containing cell (row, col).
   */
  getMergeRegion(slideIndex, tableId, row, col, slideManager) {
    const merges = this.getMergedCells(slideIndex, tableId, slideManager)
    for (const R of merges) {
      if (row >= R.startRow && row <= R.endRow && col >= R.startCol && col <= R.endCol) {
        return R
      }
    }
    return null
  }

  /**
   * Returns origin cell coordinates.
   */
  getMergeParent(slideIndex, tableId, row, col, slideManager) {
    const R = this.getMergeRegion(slideIndex, tableId, row, col, slideManager)
    if (R) {
      return { row: R.startRow, col: R.startCol }
    }
    return { row, col }
  }

  getCellBounds(slideIndex, tableId, rowIndex, colIndex, slideManager) {
    const { tblObj, frameObj } = this.#getTableContext(slideIndex, tableId, slideManager)

    this.#calculateRowHeights(slideIndex, tableId, slideManager, tblObj)

    const xfrm = frameObj['p:xfrm']
    const tableX = xfrm?.['a:off']?.['@_x'] ? parseInt(xfrm['a:off']['@_x'], 10) : 0
    const tableY = xfrm?.['a:off']?.['@_y'] ? parseInt(xfrm['a:off']['@_y'], 10) : 0

    const gridCols = tblObj['a:tblGrid']?.['a:gridCol'] || []
    const gridColsArr = Array.isArray(gridCols) ? gridCols : [gridCols]
    const colWidths = gridColsArr.map(col => parseInt(col['@_w'] || 0, 10))

    const trsArr = tblObj['a:tr'] || []
    const rowHeights = trsArr.map(row => parseInt(row['@_h'] || 0, 10))

    const parent = this.getMergeParent(slideIndex, tableId, rowIndex, colIndex, slideManager)
    const pr = parent.row
    const pc = parent.col

    let cellLeft = tableX
    for (let idx = 0; idx < pc; idx++) {
      cellLeft += colWidths[idx] || 0
    }

    let cellTop = tableY
    for (let idx = 0; idx < pr; idx++) {
      cellTop += rowHeights[idx] || 0
    }

    const parentCell = trsArr[pr]?.['a:tc']?.[pc]
    const gridSpan = parentCell?.['@_gridSpan'] ? parseInt(parentCell['@_gridSpan'], 10) : 1
    const rowSpan = parentCell?.['@_rowSpan'] ? parseInt(parentCell['@_rowSpan'], 10) : 1

    let cellWidth = 0
    for (let idx = 0; idx < gridSpan; idx++) {
      cellWidth += colWidths[pc + idx] || 0
    }

    let cellHeight = 0
    for (let idx = 0; idx < rowSpan; idx++) {
      cellHeight += rowHeights[pr + idx] || 0
    }

    return {
      x: Math.round(cellLeft / 9525),
      y: Math.round(cellTop / 9525),
      width: Math.round(cellWidth / 9525),
      height: Math.round(cellHeight / 9525),
    }
  }

  getCellPosition(slideIndex, tableId, rowIndex, colIndex, slideManager) {
    const bounds = this.getCellBounds(slideIndex, tableId, rowIndex, colIndex, slideManager)
    return {
      row: rowIndex,
      column: colIndex,
      x: bounds.x,
      y: bounds.y,
    }
  }

  /**
   * Splits a merged region containing cell (row, col).
   */
  splitMergedRegion(slideIndex, tableId, row, col, slideManager) {
    this.unmergeCells(slideIndex, tableId, row, col, slideManager)
  }

  /**
   * Clones a merged region.
   */
  cloneMergedRegion(slideIndex, tableId, row, col, targetRow, targetCol, slideManager) {
    const R = this.getMergeRegion(slideIndex, tableId, row, col, slideManager)
    if (!R) return

    const rowSpan = R.endRow - R.startRow + 1
    const colSpan = R.endCol - R.startCol + 1

    const targetEndRow = targetRow + rowSpan - 1
    const targetEndCol = targetCol + colSpan - 1

    this.mergeCells(
      slideIndex,
      tableId,
      targetRow,
      targetCol,
      targetEndRow,
      targetEndCol,
      slideManager
    )

    const slideObj = slideManager.getSlideObj(slideIndex)
    const tblObj = this.#findTableObj(slideObj, tableId, slideManager, slideIndex)
    if (!tblObj) return

    const trs = tblObj['a:tr'] || []
    const srcCell = trs[R.startRow]?.['a:tc']?.[R.startCol]
    const destCell = trs[targetRow]?.['a:tc']?.[targetCol]

    if (srcCell && destCell) {
      if (srcCell['a:tcPr']) {
        destCell['a:tcPr'] = this.#xmlParser.deepClone(srcCell['a:tcPr'])
      }
      if (srcCell['a:txBody']) {
        destCell['a:txBody'] = this.#xmlParser.deepClone(srcCell['a:txBody'])
      }
    }

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Auto-fits table columns based on text length.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {SlideManager} slideManager
   */
  autoFitTable(slideIndex, tableId, slideManager) {
    const { tblObj } = this.#getTableContext(slideIndex, tableId, slideManager)

    const trs = tblObj['a:tr'] || []
    const gridCols = tblObj['a:tblGrid']?.['a:gridCol']
    if (!gridCols || trs.length === 0) return

    const numCols = gridCols.length
    const maxLens = new Array(numCols).fill(0)

    for (const row of trs) {
      const tcs = row['a:tc'] || []
      for (let c = 0; c < numCols; c++) {
        const cell = tcs[c]
        if (!cell || cell['@_hMerge'] || cell['@_vMerge']) continue

        const txBody = cell['a:txBody']
        if (txBody) {
          const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
          let len = 0
          for (const p of paras) {
            if (p['a:r']) {
              const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
              for (const r of runs) {
                if (r['a:t']) {
                  len += String(r['a:t']).length
                }
              }
            }
          }
          if (len > maxLens[c]) {
            maxLens[c] = len
          }
        }
      }
    }

    for (let c = 0; c < numCols; c++) {
      const width = Math.max(1000000, maxLens[c] * 120000)
      gridCols[c]['@_w'] = String(width)
    }

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Resizes the table width and height.
   *
   * @param {number} slideIndex
   * @param {string} tableId
   * @param {number} width - New width in EMUs or inches.
   * @param {number} height - New height in EMUs or inches.
   * @param {SlideManager} slideManager
   */
  resizeTable(slideIndex, tableId, width, height, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)

    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']
    if (!spTree) return

    let frames = spTree['p:graphicFrame'] || []
    if (!Array.isArray(frames)) frames = [frames]

    let targetFrame = null
    let tbl = null

    for (const frame of frames) {
      tbl = frame?.['a:graphic']?.['a:graphicData']?.['a:tbl']
      if (!tbl) continue

      const cNvPr = frame?.['p:nvGraphicFramePr']?.['p:cNvPr']
      if (!cNvPr) continue

      const name = cNvPr['@_name']
      const id = String(cNvPr['@_id'])

      if (tableId === 'first' || name === tableId || id === tableId) {
        targetFrame = frame
        break
      }
    }

    if (!targetFrame || !tbl) {
      throw new TableNotFoundError(`Table "${tableId}" not found in slide ${slideIndex}`)
    }

    const emuWidth = width < 100 ? Math.round(width * 914400) : Math.round(width)
    const emuHeight = height < 100 ? Math.round(height * 914400) : Math.round(height)

    if (!targetFrame['p:xfrm']) targetFrame['p:xfrm'] = {}
    if (!targetFrame['p:xfrm']['a:ext']) targetFrame['p:xfrm']['a:ext'] = {}
    targetFrame['p:xfrm']['a:ext']['@_cx'] = String(emuWidth)
    targetFrame['p:xfrm']['a:ext']['@_cy'] = String(emuHeight)

    const gridCols = tbl['a:tblGrid']?.['a:gridCol']
    if (gridCols && gridCols.length > 0) {
      let currentTotalWidth = 0
      for (const col of gridCols) {
        currentTotalWidth += parseInt(col['@_w'] || 0, 10)
      }

      if (currentTotalWidth > 0) {
        const ratio = emuWidth / currentTotalWidth
        for (const col of gridCols) {
          const w = parseInt(col['@_w'] || 0, 10)
          col['@_w'] = String(Math.round(w * ratio))
        }
      } else {
        const evenWidth = Math.round(emuWidth / gridCols.length)
        for (const col of gridCols) {
          col['@_w'] = String(evenWidth)
        }
      }
    }

    const trs = tbl['a:tr'] || []
    if (trs.length > 0) {
      let currentTotalHeight = 0
      for (const row of trs) {
        currentTotalHeight += parseInt(row['@_h'] || 0, 10)
      }

      if (currentTotalHeight > 0) {
        const ratio = emuHeight / currentTotalHeight
        for (const row of trs) {
          const h = parseInt(row['@_h'] || 0, 10)
          row['@_h'] = String(Math.round(h * ratio))
        }
      } else {
        const evenHeight = Math.round(emuHeight / trs.length)
        for (const row of trs) {
          row['@_h'] = String(evenHeight)
        }
      }
    }

    slideManager.markSlideObjDirty(slideIndex)
  }

  /**
   * Inspects all tables in a slide.
   *
   * @param {number} slideIndex
   * @param {SlideManager} slideManager
   * @returns {Array<{name: string, id: string, rows: number, cols: number}>}
   */
  inspectTables(slideIndex, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    const tables = []

    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']
    if (!spTree) return []

    let frames = spTree['p:graphicFrame'] || []
    if (!Array.isArray(frames)) frames = [frames]

    for (const frame of frames) {
      const tbl = frame?.['a:graphic']?.['a:graphicData']?.['a:tbl']
      if (!tbl) continue

      const cNvPr = frame?.['p:nvGraphicFramePr']?.['p:cNvPr']
      const name = cNvPr ? cNvPr['@_name'] : 'unnamed'
      const id = cNvPr ? String(cNvPr['@_id']) : 'unknown'

      const trs = tbl['a:tr'] || []
      const cols = trs[0]?.['a:tc']?.length || 0

      tables.push({
        name,
        id,
        rows: trs.length,
        cols,
      })
    }

    return tables
  }

  /**
   * Helper to set cell text in parsed object structure.
   */
  #setCellTextObj(cellObj, text) {
    const val = text === undefined || text === null ? '' : String(text)

    if (!cellObj['a:txBody']) {
      cellObj['a:txBody'] = {
        'a:bodyPr': {},
        'a:lstStyle': {},
        'a:p': [],
      }
    }

    const txBody = cellObj['a:txBody']

    if (!txBody['a:p']) {
      txBody['a:p'] = []
    }
    if (!Array.isArray(txBody['a:p'])) {
      txBody['a:p'] = [txBody['a:p']]
    }
    if (txBody['a:p'].length === 0) {
      txBody['a:p'].push({})
    }

    const lines = val.split(/\r?\n/)
    const templatePara = txBody['a:p'][0]
    const newParas = []

    for (const line of lines) {
      const p = this.#xmlParser.deepClone(templatePara)

      if (!p['a:r']) {
        p['a:r'] = []
      }
      if (!Array.isArray(p['a:r'])) {
        p['a:r'] = [p['a:r']]
      }

      if (p['a:r'].length === 0) {
        p['a:r'].push({ 'a:t': line })
      } else {
        const firstRun = p['a:r'][0]
        firstRun['a:t'] = line
        p['a:r'] = [firstRun]
      }

      // Ensure strict schema ordering for DrawingML paragraph elements:
      // 1. a:pPr (paragraph properties)
      // 2. a:r / a:br / a:fld (text runs, breaks, fields)
      // 3. a:endParaRPr (end paragraph run properties)
      const pOrdered = {}
      if (p['a:pPr'] !== undefined) {
        pOrdered['a:pPr'] = p['a:pPr']
      }
      if (p['a:r'] !== undefined) {
        pOrdered['a:r'] = p['a:r']
      }
      if (p['a:br'] !== undefined) {
        pOrdered['a:br'] = p['a:br']
      }
      if (p['a:fld'] !== undefined) {
        pOrdered['a:fld'] = p['a:fld']
      }
      for (const key of Object.keys(p)) {
        if (
          key !== 'a:pPr' &&
          key !== 'a:r' &&
          key !== 'a:br' &&
          key !== 'a:fld' &&
          key !== 'a:endParaRPr'
        ) {
          pOrdered[key] = p[key]
        }
      }
      if (p['a:endParaRPr'] !== undefined) {
        pOrdered['a:endParaRPr'] = p['a:endParaRPr']
      }

      newParas.push(pOrdered)
    }

    txBody['a:p'] = newParas
  }

  /**
   * Helper to find a table element inside a slide parsed object.
   */
  #findTableObj(slideObj, tableId, slideManager, slideIndex) {
    if (slideManager && slideIndex !== undefined) {
      const res = slideManager.getSlideTable(slideIndex, tableId)
      return res ? res.table : null
    }

    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']
    if (!spTree) return null

    let frames = spTree['p:graphicFrame'] || []
    if (!Array.isArray(frames)) frames = [frames]

    for (const frame of frames) {
      const tbl = frame?.['a:graphic']?.['a:graphicData']?.['a:tbl']
      if (!tbl) continue

      const cNvPr = frame?.['p:nvGraphicFramePr']?.['p:cNvPr']
      if (!cNvPr) continue

      const name = cNvPr['@_name']
      const id = String(cNvPr['@_id'])

      if (tableId === 'first' || name === tableId || id === tableId) {
        return tbl
      }
    }

    return null
  }

  #getTableContext(slideIndex, tableId, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    const res = slideManager.getSlideTable(slideIndex, tableId)
    if (!res || !res.table) {
      throw new TableNotFoundError(`Table "${tableId}" not found in slide ${slideIndex}`)
    }
    const cNvPr = res.frame?.['p:nvGraphicFramePr']?.['p:cNvPr']
    const resolvedTableId = cNvPr ? cNvPr['@_name'] || String(cNvPr['@_id']) : tableId
    return { slideObj, tblObj: res.table, frameObj: res.frame, resolvedTableId }
  }

  #applyCellOptions(cellObj, cellOptions) {
    if (cellOptions.fill) {
      if (!cellObj['a:tcPr']) cellObj['a:tcPr'] = {}
      cellObj['a:tcPr']['a:solidFill'] = {
        'a:srgbClr': { '@_val': cellOptions.fill },
      }
    }

    if (cellOptions.align) {
      const txBody = cellObj['a:txBody']
      if (txBody && txBody['a:p']) {
        const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
        for (const p of paras) {
          if (!p['a:pPr']) p['a:pPr'] = {}
          p['a:pPr']['@_algn'] = cellOptions.align
        }
      }
    }

    if (cellOptions.fontSize) {
      const sizeVal = cellOptions.fontSize * 100
      const txBody = cellObj['a:txBody']
      if (txBody && txBody['a:p']) {
        const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
        for (const p of paras) {
          if (p['a:r']) {
            const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
            for (const r of runs) {
              if (!r['a:rPr']) r['a:rPr'] = {}
              r['a:rPr']['@_sz'] = String(sizeVal)
            }
          }
        }
      }
    }
  }

  #expandCellShape(config, cellBounds) {
    const cellLeft_px = Math.round(cellBounds.left / 9525)
    const cellTop_px = Math.round(cellBounds.top / 9525)
    const cellWidth_px = Math.round(cellBounds.width / 9525)
    const cellHeight_px = Math.round(cellBounds.height / 9525)

    const parseLength = (val, maxVal) => {
      if (typeof val === 'string' && val.endsWith('%')) {
        return (parseFloat(val) / 100) * maxVal
      }
      return val !== undefined ? parseFloat(val) : undefined
    }

    const isCellAnchored = config.anchor !== 'slide'

    // 1. Determine bounding box width and height
    let shapeWidth
    let shapeHeight

    if (config.type === 'progressBar') {
      shapeHeight = parseLength(config.height !== undefined ? config.height : 8, cellHeight_px)
      shapeWidth = parseLength(
        config.width !== undefined ? config.width : cellWidth_px - 10,
        cellWidth_px
      )
    } else if (config.type === 'badge') {
      const text = String(config.text !== undefined ? config.text : '')
      const fontSize = config.textStyle?.fontSize || 10
      const textWidth = text.length * fontSize * 0.6
      const paddingX = 12
      shapeWidth =
        parseLength(config.width, cellWidth_px) !== undefined
          ? parseLength(config.width, cellWidth_px)
          : textWidth + paddingX * 2
      shapeHeight =
        parseLength(config.height, cellHeight_px) !== undefined
          ? parseLength(config.height, cellHeight_px)
          : fontSize + 12
    } else if (config.type === 'icon') {
      const size = parseLength(
        config.size !== undefined ? config.size : 16,
        Math.min(cellWidth_px, cellHeight_px)
      )
      shapeWidth = size
      shapeHeight = size
    } else {
      shapeWidth = parseLength(config.width, cellWidth_px)
      if (shapeWidth === undefined) {
        const sizeVal = parseLength(config.size, Math.min(cellWidth_px, cellHeight_px))
        if (sizeVal !== undefined) {
          shapeWidth = sizeVal
        } else {
          const radiusVal = parseLength(config.radius, Math.min(cellWidth_px, cellHeight_px) / 2)
          if (radiusVal !== undefined) {
            shapeWidth = radiusVal * 2
          } else {
            shapeWidth = 12 // default
          }
        }
      }

      shapeHeight = parseLength(config.height, cellHeight_px)
      if (shapeHeight === undefined) {
        const sizeVal = parseLength(config.size, Math.min(cellWidth_px, cellHeight_px))
        if (sizeVal !== undefined) {
          shapeHeight = sizeVal
        } else {
          const radiusVal = parseLength(config.radius, Math.min(cellWidth_px, cellHeight_px) / 2)
          if (radiusVal !== undefined) {
            shapeHeight = radiusVal * 2
          } else {
            shapeHeight = 12 // default
          }
        }
      }
    }

    // 2. Determine alignment settings
    let alignX = config.alignX
    let alignY = config.alignY

    if (config.position) {
      switch (config.position) {
        case 'top-left':
          if (!alignX) alignX = 'left'
          if (!alignY) alignY = 'top'
          break
        case 'top-center':
        case 'top':
          if (!alignX) alignX = 'center'
          if (!alignY) alignY = 'top'
          break
        case 'top-right':
          if (!alignX) alignX = 'right'
          if (!alignY) alignY = 'top'
          break
        case 'middle-left':
        case 'left':
          if (!alignX) alignX = 'left'
          if (!alignY) alignY = 'middle'
          break
        case 'center':
        case 'middle-center':
          if (!alignX) alignX = 'center'
          if (!alignY) alignY = 'middle'
          break
        case 'middle-right':
        case 'right':
          if (!alignX) alignX = 'right'
          if (!alignY) alignY = 'middle'
          break
        case 'bottom-left':
          if (!alignX) alignX = 'left'
          if (!alignY) alignY = 'bottom'
          break
        case 'bottom-center':
        case 'bottom':
          if (!alignX) alignX = 'center'
          if (!alignY) alignY = 'bottom'
          break
        case 'bottom-right':
          if (!alignX) alignX = 'right'
          if (!alignY) alignY = 'bottom'
          break
      }
    }

    if (alignX && !alignY && config.y === undefined) {
      alignY = 'middle'
    }
    if (alignY && !alignX && config.x === undefined) {
      alignX = 'center'
    }

    if (!alignX && !alignY && config.x === undefined && config.y === undefined) {
      alignX = 'center'
      alignY = 'middle'
    }

    // 3. Compute coordinates
    let shapeLeft = cellLeft_px
    let shapeTop = cellTop_px

    if (isCellAnchored) {
      if (alignX === 'left') {
        shapeLeft = cellLeft_px + (config.x !== undefined ? config.x : 5)
      } else if (alignX === 'center') {
        shapeLeft = cellLeft_px + (cellWidth_px - shapeWidth) / 2 + (config.x || 0)
      } else if (alignX === 'right') {
        shapeLeft =
          cellLeft_px + cellWidth_px - shapeWidth - (config.x !== undefined ? config.x : 5)
      } else {
        shapeLeft =
          cellLeft_px + (config.x !== undefined ? config.x : (cellWidth_px - shapeWidth) / 2)
      }

      if (alignY === 'top') {
        shapeTop = cellTop_px + (config.y !== undefined ? config.y : 5)
      } else if (alignY === 'middle') {
        shapeTop = cellTop_px + (cellHeight_px - shapeHeight) / 2 + (config.y || 0)
      } else if (alignY === 'bottom') {
        shapeTop =
          cellTop_px + cellHeight_px - shapeHeight - (config.y !== undefined ? config.y : 5)
      } else {
        shapeTop =
          cellTop_px + (config.y !== undefined ? config.y : (cellHeight_px - shapeHeight) / 2)
      }

      // 4. Boundary Constraints Validation/Enforcement
      if (shapeWidth > cellWidth_px) {
        shapeLeft = cellLeft_px
      } else {
        shapeLeft = Math.max(
          cellLeft_px,
          Math.min(shapeLeft, cellLeft_px + cellWidth_px - shapeWidth)
        )
      }

      if (shapeHeight > cellHeight_px) {
        shapeTop = cellTop_px
      } else {
        shapeTop = Math.max(
          cellTop_px,
          Math.min(shapeTop, cellTop_px + cellHeight_px - shapeHeight)
        )
      }
    } else {
      shapeLeft = config.x || 0
      shapeTop = config.y || 0
    }

    // 5. Expand individual sub-elements / custom shapes
    if (config.type === 'progressBar') {
      const value = config.value !== undefined ? config.value : 0
      const max = config.max !== undefined ? config.max : 100
      const fill = config.fill || '#3B82F6'
      const bgFill = config.backgroundFill || '#E5E7EB'

      const shapes = []
      shapes.push({
        type: 'roundedRectangle',
        fill: bgFill,
        x: shapeLeft,
        y: shapeTop,
        width: shapeWidth,
        height: shapeHeight,
        borderRadius: shapeHeight / 2,
        zIndex: config.zIndex,
      })

      const pct = Math.min(1, Math.max(0, value / max))
      if (pct > 0) {
        const filledWidth = shapeWidth * pct
        shapes.push({
          type: 'roundedRectangle',
          fill: fill,
          x: shapeLeft,
          y: shapeTop,
          width: filledWidth,
          height: shapeHeight,
          borderRadius: shapeHeight / 2,
          zIndex: (config.zIndex || 0) + 1,
        })
      }
      return shapes
    }

    if (config.type === 'badge') {
      const text = String(config.text !== undefined ? config.text : '')
      const fontSize = config.textStyle?.fontSize || 10
      return [
        {
          type: 'roundedRectangle',
          fill: config.fill || '#10B981',
          borderRadius: shapeHeight / 2,
          x: shapeLeft,
          y: shapeTop,
          width: shapeWidth,
          height: shapeHeight,
          text: text,
          textStyle: {
            color: config.textStyle?.color || '#FFFFFF',
            fontSize: fontSize,
            bold: config.textStyle?.bold !== undefined ? config.textStyle.bold : true,
            align: 'center',
          },
          border: config.border,
          transparency: config.transparency,
          shadow: config.shadow,
          rotation: config.rotation,
          zIndex: config.zIndex,
        },
      ]
    }

    if (config.type === 'icon') {
      const iconFill = config.fill
      const fontSize = Math.round(shapeWidth * 0.8)

      let baseConfig = null
      switch (config.icon) {
        case 'check':
          baseConfig = {
            type: 'rectangle',
            fill: 'none',
            border: null,
            width: shapeWidth,
            height: shapeHeight,
            text: '✔',
            textStyle: {
              color: iconFill || '#10B981',
              bold: true,
              fontSize: fontSize,
              align: 'center',
            },
          }
          break
        case 'cross':
          baseConfig = {
            type: 'rectangle',
            fill: 'none',
            border: null,
            width: shapeWidth,
            height: shapeHeight,
            text: '✘',
            textStyle: {
              color: iconFill || '#EF4444',
              bold: true,
              fontSize: fontSize,
              align: 'center',
            },
          }
          break
        case 'warning':
          baseConfig = {
            type: 'triangle',
            fill: iconFill || '#F59E0B',
            border: null,
            width: shapeWidth,
            height: shapeHeight,
            text: '!',
            textStyle: {
              color: '#FFFFFF',
              bold: true,
              fontSize: Math.round(fontSize * 0.7),
              align: 'center',
            },
          }
          break
        case 'info':
          baseConfig = {
            type: 'circle',
            fill: iconFill || '#3B82F6',
            border: null,
            radius: shapeWidth / 2,
            text: 'i',
            textStyle: {
              color: '#FFFFFF',
              bold: true,
              fontSize: Math.round(fontSize * 0.7),
              align: 'center',
            },
          }
          break
        case 'star':
          baseConfig = {
            type: 'star5',
            fill: iconFill || '#FBBF24',
            border: null,
            width: shapeWidth,
            height: shapeHeight,
          }
          break
        case 'up':
          baseConfig = {
            type: 'upArrow',
            fill: iconFill || '#10B981',
            border: null,
            width: shapeWidth,
            height: shapeHeight,
          }
          break
        case 'down':
          baseConfig = {
            type: 'downArrow',
            fill: iconFill || '#EF4444',
            border: null,
            width: shapeWidth,
            height: shapeHeight,
          }
          break
        case 'arrow-right':
          baseConfig = {
            type: 'rightArrow',
            fill: iconFill || '#3B82F6',
            border: null,
            width: shapeWidth,
            height: shapeHeight,
          }
          break
        case 'arrow-left':
          baseConfig = {
            type: 'leftArrow',
            fill: iconFill || '#3B82F6',
            border: null,
            width: shapeWidth,
            height: shapeHeight,
          }
          break
        default:
          return []
      }

      baseConfig.x = shapeLeft
      baseConfig.y = shapeTop
      baseConfig.zIndex = config.zIndex
      if (config.border) baseConfig.border = config.border
      if (config.transparency !== undefined) baseConfig.transparency = config.transparency
      if (config.shadow !== undefined) baseConfig.shadow = config.shadow
      if (config.rotation !== undefined) baseConfig.rotation = config.rotation

      return [baseConfig]
    }

    const expanded = Object.assign({}, config, {
      x: shapeLeft,
      y: shapeTop,
      width: shapeWidth,
      height: shapeHeight,
    })

    if (expanded.type === 'circle' && expanded.radius === undefined) {
      expanded.radius = shapeWidth / 2
    }
    if (expanded.type === 'square' && expanded.size === undefined) {
      expanded.size = shapeWidth
    }

    return [expanded]
  }

  #processCellShapes(
    slideIndex,
    tableId,
    resolvedTableId,
    rowsData,
    isObjectRows,
    cellShapes,
    slideManager,
    shapeManager,
    tblObj,
    frameObj
  ) {
    if (!cellShapes || !shapeManager) return

    const shapes = shapeManager.getShapes(slideIndex, slideManager)
    const prefixToDelete = `cellshape_${resolvedTableId}_`
    const existingNames = shapes
      .map(s => s.name)
      .filter(name => name && name.startsWith(prefixToDelete))

    for (const name of existingNames) {
      try {
        shapeManager.deleteShape(slideIndex, name, slideManager)
      } catch (err) {
        logger.warn(`Failed to delete existing cell shape "${name}": ${err.message}`)
      }
    }

    this.#calculateRowHeights(slideIndex, tableId, slideManager, tblObj)

    const xfrm = frameObj['p:xfrm']
    const tableX = xfrm?.['a:off']?.['@_x'] ? parseInt(xfrm['a:off']['@_x'], 10) : 0
    const tableY = xfrm?.['a:off']?.['@_y'] ? parseInt(xfrm['a:off']['@_y'], 10) : 0

    const gridCols = tblObj['a:tblGrid']?.['a:gridCol'] || []
    const gridColsArr = Array.isArray(gridCols) ? gridCols : [gridCols]
    const colWidths = gridColsArr.map(col => parseInt(col['@_w'] || 0, 10))

    const trsArr = tblObj['a:tr'] || []
    const rowHeights = trsArr.map(row => parseInt(row['@_h'] || 0, 10))

    const getCellBounds = (r, c) => {
      const parent = this.getMergeParent(slideIndex, tableId, r, c, slideManager)
      const pr = parent.row
      const pc = parent.col

      let cellLeft = tableX
      for (let idx = 0; idx < pc; idx++) {
        cellLeft += colWidths[idx] || 0
      }

      let cellTop = tableY
      for (let idx = 0; idx < pr; idx++) {
        cellTop += rowHeights[idx] || 0
      }

      const parentCell = trsArr[pr]?.['a:tc']?.[pc]
      const gridSpan = parentCell?.['@_gridSpan'] ? parseInt(parentCell['@_gridSpan'], 10) : 1
      const rowSpan = parentCell?.['@_rowSpan'] ? parseInt(parentCell['@_rowSpan'], 10) : 1

      let cellWidth = 0
      for (let idx = 0; idx < gridSpan; idx++) {
        cellWidth += colWidths[pc + idx] || 0
      }

      let cellHeight = 0
      for (let idx = 0; idx < rowSpan; idx++) {
        cellHeight += rowHeights[pr + idx] || 0
      }

      return {
        left: cellLeft,
        top: cellTop,
        width: cellWidth,
        height: cellHeight,
      }
    }

    const shapesToCreate = []
    const headerNames = (trsArr[0]?.['a:tc'] || []).map(cell => this.#getCellText(cell).trim())

    for (let i = 0; i < rowsData.length; i++) {
      const rowData = rowsData[i]
      const finalRowIndex = isObjectRows ? i + 1 : i

      const numCols = trsArr[finalRowIndex]?.['a:tc']?.length || 0
      for (let j = 0; j < numCols; j++) {
        const headerName = headerNames[j]
        let shapeFn = null

        if (headerName) {
          shapeFn = cellShapes[headerName] || cellShapes[headerName.toLowerCase()]
        }
        if (!shapeFn) {
          shapeFn = cellShapes[j]
        }

        if (typeof shapeFn !== 'function') continue

        let configs = shapeFn(rowData, i)
        if (!configs) continue

        if (!Array.isArray(configs)) {
          configs = [configs]
        }

        configs.forEach((config, shapeIdx) => {
          shapesToCreate.push({
            config,
            rowIndex: finalRowIndex,
            colIndex: j,
            shapeIndex: shapeIdx,
          })
        })
      }
    }

    shapesToCreate.sort((a, b) => (a.config.zIndex || 0) - (b.config.zIndex || 0))

    shapesToCreate.forEach(item => {
      const bounds = getCellBounds(item.rowIndex, item.colIndex)
      const expandedConfigs = this.#expandCellShape(item.config, bounds)

      expandedConfigs.forEach((expandedConfig, expIdx) => {
        const finalShapeIndex =
          expandedConfigs.length > 1 ? `${item.shapeIndex}_${expIdx}` : item.shapeIndex
        expandedConfig.id = `cellshape_${resolvedTableId}_${item.rowIndex}_${item.colIndex}_${finalShapeIndex}`

        shapeManager.addShape(slideIndex, expandedConfig, slideManager)
      })
    })
  }

  getTableRows(slideIndex, tableId, options = {}, slideManager) {
    const { tblObj } = this.#getTableContext(slideIndex, tableId, slideManager)
    const trs = tblObj['a:tr'] || []
    if (trs.length === 0) {
      return options.includeMetadata
        ? { rows: [], rowCount: 0, columnCount: 0, mergedCells: [] }
        : []
    }

    const numRows = trs.length
    const gridCols = tblObj['a:tblGrid']?.['a:gridCol'] || []
    const gridColsArr = Array.isArray(gridCols) ? gridCols : [gridCols]
    const numCols = gridColsArr.length

    // Extract all raw cell text, resolving merges to their parent's text
    const matrix = []
    for (let r = 0; r < numRows; r++) {
      const rowCells = []
      for (let c = 0; c < numCols; c++) {
        const parent = this.getMergeParent(slideIndex, tableId, r, c, slideManager)
        const cell = trs[parent.row]?.['a:tc']?.[parent.col]
        const text = cell ? this.#getCellText(cell) : ''
        rowCells.push(text)
      }
      matrix.push(rowCells)
    }

    // Header names are extracted from the first row (index 0)
    const headerNames = matrix[0].map((hText, cIdx) => {
      const cleaned = hText.trim()
      return cleaned || `column${cIdx + 1}`
    })

    // Compute the data rows (excluding the header row at index 0)
    const dataRows = matrix.slice(1)

    let rowsResult = []
    if (options.raw) {
      rowsResult = dataRows
    } else {
      for (const rowCells of dataRows) {
        const rowObj = {}
        for (let c = 0; c < numCols; c++) {
          const key = headerNames[c]
          rowObj[key] = rowCells[c] || ''
        }
        rowsResult.push(rowObj)
      }
    }

    if (options.includeMetadata) {
      const mergedCells = this.getMergedCells(slideIndex, tableId, slideManager)
      return {
        rows: rowsResult,
        rowCount: numRows,
        columnCount: numCols,
        mergedCells,
      }
    }

    return rowsResult
  }

  addCellShape(slideIndex, tableId, rowIndex, colIndex, options, slideManager, shapeManager) {
    const { tblObj, frameObj, resolvedTableId } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )

    this.#calculateRowHeights(slideIndex, tableId, slideManager, tblObj)

    const xfrm = frameObj['p:xfrm']
    const tableX = xfrm?.['a:off']?.['@_x'] ? parseInt(xfrm['a:off']['@_x'], 10) : 0
    const tableY = xfrm?.['a:off']?.['@_y'] ? parseInt(xfrm['a:off']['@_y'], 10) : 0

    const gridCols = tblObj['a:tblGrid']?.['a:gridCol'] || []
    const gridColsArr = Array.isArray(gridCols) ? gridCols : [gridCols]
    const colWidths = gridColsArr.map(col => parseInt(col['@_w'] || 0, 10))

    const trsArr = tblObj['a:tr'] || []
    const rowHeights = trsArr.map(row => parseInt(row['@_h'] || 0, 10))

    const parent = this.getMergeParent(slideIndex, tableId, rowIndex, colIndex, slideManager)
    const pr = parent.row
    const pc = parent.col

    let cellLeft = tableX
    for (let idx = 0; idx < pc; idx++) {
      cellLeft += colWidths[idx] || 0
    }

    let cellTop = tableY
    for (let idx = 0; idx < pr; idx++) {
      cellTop += rowHeights[idx] || 0
    }

    const parentCell = trsArr[pr]?.['a:tc']?.[pc]
    const gridSpan = parentCell?.['@_gridSpan'] ? parseInt(parentCell['@_gridSpan'], 10) : 1
    const rowSpan = parentCell?.['@_rowSpan'] ? parseInt(parentCell['@_rowSpan'], 10) : 1

    let cellWidth = 0
    for (let idx = 0; idx < gridSpan; idx++) {
      cellWidth += colWidths[pc + idx] || 0
    }

    let cellHeight = 0
    for (let idx = 0; idx < rowSpan; idx++) {
      cellHeight += rowHeights[pr + idx] || 0
    }

    const bounds = { left: cellLeft, top: cellTop, width: cellWidth, height: cellHeight }

    const shapes = shapeManager.getShapes(slideIndex, slideManager)
    const prefix = `cellshape_${resolvedTableId}_${rowIndex}_${colIndex}_`
    let maxShapeIndex = -1
    for (const s of shapes) {
      if (s.name && s.name.startsWith(prefix)) {
        const remaining = s.name.slice(prefix.length)
        const parts = remaining.split('_')
        const idxVal = parseInt(parts[0], 10)
        if (!isNaN(idxVal) && idxVal > maxShapeIndex) {
          maxShapeIndex = idxVal
        }
      }
    }
    const nextShapeIndex = maxShapeIndex + 1

    const expandedConfigs = this.#expandCellShape(options, bounds)

    expandedConfigs.forEach((expandedConfig, expIdx) => {
      const finalShapeIndex =
        expandedConfigs.length > 1 ? `${nextShapeIndex}_${expIdx}` : nextShapeIndex
      expandedConfig.id = `cellshape_${resolvedTableId}_${rowIndex}_${colIndex}_${finalShapeIndex}`

      shapeManager.addShape(slideIndex, expandedConfig, slideManager)
    })
  }

  updateCellShape(
    slideIndex,
    tableId,
    rowIndex,
    colIndex,
    shapeIndex,
    options,
    slideManager,
    shapeManager
  ) {
    const { tblObj, frameObj, resolvedTableId } = this.#getTableContext(
      slideIndex,
      tableId,
      slideManager
    )

    this.#calculateRowHeights(slideIndex, tableId, slideManager, tblObj)

    const shapes = shapeManager.getShapes(slideIndex, slideManager)
    const prefix = `cellshape_${resolvedTableId}_${rowIndex}_${colIndex}_${shapeIndex}`
    const matchingShapes = shapes.filter(
      s => s.name && (s.name === prefix || s.name.startsWith(prefix + '_'))
    )

    if (matchingShapes.length === 0) {
      throw new PPTXError(`Cell shape "${shapeIndex}" not found in cell (${rowIndex}, ${colIndex})`)
    }

    for (const s of matchingShapes) {
      shapeManager.deleteShape(slideIndex, s.name, slideManager)
    }

    const xfrm = frameObj['p:xfrm']
    const tableX = xfrm?.['a:off']?.['@_x'] ? parseInt(xfrm['a:off']['@_x'], 10) : 0
    const tableY = xfrm?.['a:off']?.['@_y'] ? parseInt(xfrm['a:off']['@_y'], 10) : 0

    const gridCols = tblObj['a:tblGrid']?.['a:gridCol'] || []
    const gridColsArr = Array.isArray(gridCols) ? gridCols : [gridCols]
    const colWidths = gridColsArr.map(col => parseInt(col['@_w'] || 0, 10))

    const trsArr = tblObj['a:tr'] || []
    const rowHeights = trsArr.map(row => parseInt(row['@_h'] || 0, 10))

    const parent = this.getMergeParent(slideIndex, tableId, rowIndex, colIndex, slideManager)
    const pr = parent.row
    const pc = parent.col

    let cellLeft = tableX
    for (let idx = 0; idx < pc; idx++) {
      cellLeft += colWidths[idx] || 0
    }

    let cellTop = tableY
    for (let idx = 0; idx < pr; idx++) {
      cellTop += rowHeights[idx] || 0
    }

    const parentCell = trsArr[pr]?.['a:tc']?.[pc]
    const gridSpan = parentCell?.['@_gridSpan'] ? parseInt(parentCell['@_gridSpan'], 10) : 1
    const rowSpan = parentCell?.['@_rowSpan'] ? parseInt(parentCell['@_rowSpan'], 10) : 1

    let cellWidth = 0
    for (let idx = 0; idx < gridSpan; idx++) {
      cellWidth += colWidths[pc + idx] || 0
    }

    let cellHeight = 0
    for (let idx = 0; idx < rowSpan; idx++) {
      cellHeight += rowHeights[pr + idx] || 0
    }

    const bounds = { left: cellLeft, top: cellTop, width: cellWidth, height: cellHeight }

    const expandedConfigs = this.#expandCellShape(options, bounds)

    expandedConfigs.forEach((expandedConfig, expIdx) => {
      const finalShapeIndex = expandedConfigs.length > 1 ? `${shapeIndex}_${expIdx}` : shapeIndex
      expandedConfig.id = `cellshape_${resolvedTableId}_${rowIndex}_${colIndex}_${finalShapeIndex}`

      shapeManager.addShape(slideIndex, expandedConfig, slideManager)
    })
  }

  removeCellShape(slideIndex, tableId, rowIndex, colIndex, shapeIndex, slideManager, shapeManager) {
    const { resolvedTableId } = this.#getTableContext(slideIndex, tableId, slideManager)

    const shapes = shapeManager.getShapes(slideIndex, slideManager)
    const prefix = `cellshape_${resolvedTableId}_${rowIndex}_${colIndex}_${shapeIndex}`
    const matchingShapes = shapes.filter(
      s => s.name && (s.name === prefix || s.name.startsWith(prefix + '_'))
    )

    if (matchingShapes.length === 0) {
      throw new PPTXError(`Cell shape "${shapeIndex}" not found in cell (${rowIndex}, ${colIndex})`)
    }

    for (const s of matchingShapes) {
      shapeManager.deleteShape(slideIndex, s.name, slideManager)
    }
  }

  getCellShape(slideIndex, tableId, rowIndex, colIndex, shapeIndex, slideManager, shapeManager) {
    const { resolvedTableId } = this.#getTableContext(slideIndex, tableId, slideManager)

    const prefix = `cellshape_${resolvedTableId}_${rowIndex}_${colIndex}_${shapeIndex}`
    const shapes = shapeManager.getShapes(slideIndex, slideManager)
    const primaryShape = shapes.find(
      s => s.name === prefix || s.name === `${prefix}_0` || s.name === `${prefix}_1`
    )

    if (!primaryShape) return null

    return shapeManager.getShape(slideIndex, primaryShape.name, slideManager)
  }

  /**
   * Generates a new rowId for the given row object.
   */
  #updateRowId(rowObj) {
    const randomVal = String(this.#generateRandomUint32())
    if (rowObj['a:extLst']?.['a:ext']) {
      const ext = rowObj['a:extLst']['a:ext']
      const exts = Array.isArray(ext) ? ext : [ext]
      let updated = false
      for (const e of exts) {
        if (e['a16:rowId']) {
          e['a16:rowId']['@_val'] = randomVal
          updated = true
        }
      }
      if (!updated) {
        exts.push({
          '@_uri': '{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}',
          'a16:rowId': {
            '@_xmlns:a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
            '@_val': randomVal,
          },
        })
        rowObj['a:extLst']['a:ext'] = exts
      }
    } else {
      rowObj['a:extLst'] = {
        'a:ext': {
          '@_uri': '{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}',
          'a16:rowId': {
            '@_xmlns:a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
            '@_val': randomVal,
          },
        },
      }
    }
  }

  #calculateRowHeights(slideIndex, tableId, slideManager, tblObj) {
    const trsArr = tblObj['a:tr'] || []
    if (trsArr.length === 0) return []

    const gridCols = tblObj['a:tblGrid']?.['a:gridCol'] || []
    const gridColsArr = Array.isArray(gridCols) ? gridCols : [gridCols]
    const colWidths = gridColsArr.map(col => parseInt(col['@_w'] || 0, 10))

    const numRows = trsArr.length
    const numCols = colWidths.length

    // Initialize rowHeights with original height or 0
    const rowHeights = trsArr.map(row => parseInt(row['@_h'] || 0, 10))

    // Helper to get paragraph font size
    const getParagraphFontSize = p => {
      let maxSz = 14 // default 14pt
      if (p['a:pPr']?.['a:defRPr']?.['@_sz']) {
        maxSz = parseInt(p['a:pPr']['a:defRPr']['@_sz'], 10) / 100
      }
      if (p['a:r']) {
        const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
        for (const r of runs) {
          if (r['a:rPr']?.['@_sz']) {
            const szVal = parseInt(r['a:rPr']['@_sz'], 10) / 100
            if (szVal > maxSz) {
              maxSz = szVal
            }
          }
        }
      }
      return maxSz
    }

    // Helper to wrap text
    const wrapText = (text, availWidth_px, fontSize) => {
      const charWidth = fontSize * 0.65
      const words = text.split(/(\s+)/)
      let linesCount = 0
      let currentLineLen = 0

      for (const word of words) {
        if (!word) continue
        const wordWidth = word.length * charWidth
        if (wordWidth > availWidth_px) {
          if (currentLineLen > 0) {
            linesCount++
            currentLineLen = 0
          }
          let remainingWidth = wordWidth
          while (remainingWidth > 0) {
            linesCount++
            remainingWidth -= availWidth_px
          }
        } else {
          if (currentLineLen + wordWidth > availWidth_px) {
            linesCount++
            currentLineLen = word.trim() ? wordWidth : 0
          } else {
            currentLineLen += wordWidth
          }
        }
      }
      if (currentLineLen > 0 || linesCount === 0) {
        linesCount++
      }
      return linesCount
    }

    // Helper to get cell margins
    const getCellMargins = cell => {
      const tcPr = cell['a:tcPr']
      const marL = tcPr?.['@_marL'] !== undefined ? parseInt(tcPr['@_marL'], 10) : 91440
      const marR = tcPr?.['@_marR'] !== undefined ? parseInt(tcPr['@_marR'], 10) : 91440
      const marT = tcPr?.['@_marT'] !== undefined ? parseInt(tcPr['@_marT'], 10) : 45720
      const marB = tcPr?.['@_marB'] !== undefined ? parseInt(tcPr['@_marB'], 10) : 45720
      return { marL, marR, marT, marB }
    }

    // Calculate required height for each cell
    const cellHeights = Array.from({ length: numRows }, () => new Array(numCols).fill(0))

    for (let r = 0; r < numRows; r++) {
      const row = trsArr[r]
      const tcs = row['a:tc'] || []
      for (let c = 0; c < numCols; c++) {
        const cell = tcs[c]
        if (!cell || cell['@_hMerge'] || cell['@_vMerge']) continue

        const parent = this.getMergeParent(slideIndex, tableId, r, c, slideManager)
        const gridSpan = cell['@_gridSpan'] ? parseInt(cell['@_gridSpan'], 10) : 1

        // Calculate cell width
        let cellWidth = 0
        for (let idx = 0; idx < gridSpan; idx++) {
          cellWidth += colWidths[parent.col + idx] || 0
        }

        const { marL, marR, marT, marB } = getCellMargins(cell)
        const availWidth = cellWidth - marL - marR
        const availWidth_px = Math.max(1, availWidth / 9525)

        // Calculate text height
        const txBody = cell['a:txBody']
        let textHeight_emu = 0
        if (txBody) {
          const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
          for (const p of paras) {
            const fontSize = getParagraphFontSize(p)
            let pText = ''
            if (p['a:r']) {
              const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
              for (const r of runs) {
                if (r['a:t']) {
                  pText += String(r['a:t'])
                }
              }
            }

            const linesCount = wrapText(pText, availWidth_px, fontSize)
            const lineHeight_emu = fontSize * 20780 // 1.4 line height multiplier

            let pHeight_emu = linesCount * lineHeight_emu
            if (p['a:pPr']?.['a:spcBef']?.['a:spcPts']?.['@_val']) {
              pHeight_emu += parseInt(p['a:pPr']['a:spcBef']['a:spcPts']['@_val'], 10) * 127
            }
            if (p['a:pPr']?.['a:spcAft']?.['a:spcPts']?.['@_val']) {
              pHeight_emu += parseInt(p['a:pPr']['a:spcAft']['a:spcPts']['@_val'], 10) * 127
            }
            textHeight_emu += pHeight_emu
          }
        }

        const totalCellHeight_emu = marT + marB + textHeight_emu
        cellHeights[r][c] = totalCellHeight_emu
      }
    }

    // Now resolve row heights based on required cell heights
    // First, non-vertically-merged cells define row heights directly
    for (let r = 0; r < numRows; r++) {
      let maxCellHeight = rowHeights[r] // Start with original template height as floor
      const row = trsArr[r]
      const tcs = row['a:tc'] || []
      for (let c = 0; c < numCols; c++) {
        const cell = tcs[c]
        if (!cell || cell['@_vMerge'] || cell['@_hMerge']) continue
        const rowSpan = cell['@_rowSpan'] ? parseInt(cell['@_rowSpan'], 10) : 1
        if (rowSpan === 1) {
          if (cellHeights[r][c] > maxCellHeight) {
            maxCellHeight = cellHeights[r][c]
          }
        }
      }
      rowHeights[r] = maxCellHeight
    }

    // Next, adjust for vertically merged cells (rowSpan > 1)
    for (let r = 0; r < numRows; r++) {
      const row = trsArr[r]
      const tcs = row['a:tc'] || []
      for (let c = 0; c < numCols; c++) {
        const cell = tcs[c]
        if (!cell || cell['@_vMerge'] || cell['@_hMerge']) continue
        const rowSpan = cell['@_rowSpan'] ? parseInt(cell['@_rowSpan'], 10) : 1
        if (rowSpan > 1) {
          const reqHeight = cellHeights[r][c]
          // Sum currently allocated row heights for spanned rows
          let currentSpanHeight = 0
          for (let idx = 0; idx < rowSpan; idx++) {
            currentSpanHeight += rowHeights[r + idx] || 0
          }
          if (reqHeight > currentSpanHeight) {
            // Distribute the extra required height equally across all spanned rows
            const diff = reqHeight - currentSpanHeight
            const extraPerRow = Math.ceil(diff / rowSpan)
            for (let idx = 0; idx < rowSpan; idx++) {
              rowHeights[r + idx] += extraPerRow
            }
          }
        }
      }
    }

    // Update row heights in XML
    for (let r = 0; r < numRows; r++) {
      trsArr[r]['@_h'] = String(rowHeights[r])
    }

    return rowHeights
  }

  #getNestedHeight(val) {
    if (Array.isArray(val)) {
      if (val.length === 0) return 1
      return val.reduce((sum, item) => sum + this.#getNestedHeight(item), 0)
    }
    return 1
  }

  #expandCellVal(val, targetHeight) {
    if (!Array.isArray(val)) {
      const res = []
      res.push({ value: val !== undefined ? val : '', rowSpan: targetHeight })
      for (let i = 1; i < targetHeight; i++) {
        res.push({ vMerge: true })
      }
      return res
    }

    if (val.length === 0) {
      const res = []
      res.push({ value: '', rowSpan: targetHeight })
      for (let i = 1; i < targetHeight; i++) {
        res.push({ vMerge: true })
      }
      return res
    }

    const itemHeights = val.map(item => this.#getNestedHeight(item))
    const currentSum = itemHeights.reduce((a, b) => a + b, 0)

    const allocatedHeights = []
    let remaining = targetHeight
    for (let i = 0; i < val.length; i++) {
      const share = Math.round((itemHeights[i] / currentSum) * targetHeight)
      allocatedHeights.push(share)
      remaining -= share
    }

    if (remaining !== 0) {
      let idx = 0
      while (remaining > 0) {
        allocatedHeights[idx % allocatedHeights.length]++
        remaining--
        idx++
      }
      while (remaining < 0) {
        let reduced = false
        for (let i = 0; i < allocatedHeights.length; i++) {
          const actualIdx = (idx + i) % allocatedHeights.length
          if (allocatedHeights[actualIdx] > 1) {
            allocatedHeights[actualIdx]--
            remaining++
            reduced = true
            break
          }
        }
        if (!reduced) break
        idx++
      }
    }

    const result = []
    for (let i = 0; i < val.length; i++) {
      result.push(...this.#expandCellVal(val[i], allocatedHeights[i]))
    }
    return result
  }

  #applyAutoMerge(cells) {
    const result = [...cells]
    let i = 0
    while (i < result.length) {
      const cell = result[i]
      if (cell.vMerge) {
        i++
        continue
      }
      let count = 1
      let j = i + 1
      while (
        j < result.length &&
        !result[j].vMerge &&
        result[j].value === cell.value &&
        cell.value !== ''
      ) {
        count++
        j++
      }
      if (count > 1) {
        cell.rowSpan = count
        for (let k = i + 1; k < j; k++) {
          result[k] = { vMerge: true }
        }
      }
      i = j
    }
    return result
  }

  #generateRandomUint32() {
    return Math.floor(Math.random() * 4294967296)
  }
}

module.exports = { TableManager }
