/**
 * @fileoverview ShapeManager - Handles shape text replacement, cloning, deletion, and enumeration.
 */

const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')

const logger = createLogger('ShapeManager')

/**
 * @class ShapeManager
 * @description Manages shape elements inside PPTX slides.
 */
class ShapeManager {
  /** @private @type {XMLParser} */
  #xmlParser

  /**
   * @param {XMLParser} xmlParser
   */
  constructor(xmlParser) {
    this.#xmlParser = xmlParser
  }

  /**
   * Updates shape text content.
   *
   * @param {number} slideIndex
   * @param {string} shapeId
   * @param {string} text
   * @param {SlideManager} slideManager
   */
  updateShapeText(slideIndex, shapeId, text, slideManager) {
    const res = slideManager.getSlideShape(slideIndex, shapeId)

    if (!res) {
      throw new PPTXError(`Shape "${shapeId}" not found in slide ${slideIndex}`)
    }

    this.#setShapeTextObj(res.shape, text)
    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Updated text for shape "${shapeId}" on slide ${slideIndex}`)
  }

  /**
   * Updates an existing shape's position and/or dimensions.
   *
   * @param {number} slideIndex
   * @param {string} shapeId
   * @param {Object} options Position and dimensions configuration.
   * @param {number} [options.x] Absolute X offset coordinate (in EMUs).
   * @param {number} [options.y] Absolute Y offset coordinate (in EMUs).
   * @param {number} [options.width] Bounding box width (in EMUs).
   * @param {number} [options.height] Bounding box height (in EMUs).
   * @param {SlideManager} slideManager
   */
  updateShapePosition(slideIndex, shapeId, options = {}, slideManager) {
    const res = slideManager.getSlideShape(slideIndex, shapeId)

    if (!res) {
      throw new PPTXError(`Shape "${shapeId}" not found in slide ${slideIndex}`)
    }

    if (!res.shape['p:spPr']) {
      res.shape['p:spPr'] = {}
    }
    if (!res.shape['p:spPr']['a:xfrm']) {
      res.shape['p:spPr']['a:xfrm'] = {}
    }
    const xfrm = res.shape['p:spPr']['a:xfrm']

    if (!xfrm['a:off']) {
      xfrm['a:off'] = {}
    }
    if (!xfrm['a:ext']) {
      xfrm['a:ext'] = {}
    }

    if (options.x !== undefined) {
      xfrm['a:off']['@_x'] = String(Math.round(options.x))
    } else if (xfrm['a:off']['@_x'] === undefined) {
      xfrm['a:off']['@_x'] = '0'
    }

    if (options.y !== undefined) {
      xfrm['a:off']['@_y'] = String(Math.round(options.y))
    } else if (xfrm['a:off']['@_y'] === undefined) {
      xfrm['a:off']['@_y'] = '0'
    }

    if (options.width !== undefined) {
      xfrm['a:ext']['@_cx'] = String(Math.round(options.width))
    } else if (xfrm['a:ext']['@_cx'] === undefined) {
      xfrm['a:ext']['@_cx'] = '0'
    }

    if (options.height !== undefined) {
      xfrm['a:ext']['@_cy'] = String(Math.round(options.height))
    } else if (xfrm['a:ext']['@_cy'] === undefined) {
      xfrm['a:ext']['@_cy'] = '0'
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Updated position/dimensions for shape "${shapeId}" on slide ${slideIndex}`)
  }

  /**
   * Updates an existing textbox shape's position and/or dimensions.
   *
   * @param {number} slideIndex
   * @param {string} textBoxId
   * @param {Object} options Position and dimensions configuration.
   * @param {number} [options.x] Absolute X offset coordinate (in EMUs).
   * @param {number} [options.y] Absolute Y offset coordinate (in EMUs).
   * @param {number} [options.width] Bounding box width (in EMUs).
   * @param {number} [options.height] Bounding box height (in EMUs).
   * @param {SlideManager} slideManager
   */
  updateTextBoxPosition(slideIndex, textBoxId, options = {}, slideManager) {
    try {
      this.updateShapePosition(slideIndex, textBoxId, options, slideManager)
    } catch (err) {
      if (err.message.includes('not found')) {
        throw new PPTXError(`Textbox "${textBoxId}" not found in slide ${slideIndex}`)
      }
      throw err
    }
  }

  /**
   * Clones a shape and adds it with offsets.
   *
   * @param {number} slideIndex
   * @param {string} shapeId
   * @param {string} newShapeId
   * @param {Object} options
   * @param {SlideManager} slideManager
   */
  cloneShape(slideIndex, shapeId, newShapeId, options = {}, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    const res = slideManager.getSlideShape(slideIndex, shapeId)

    if (!res) {
      throw new PPTXError(`Shape "${shapeId}" not found in slide ${slideIndex}`)
    }

    const newShape = this.#xmlParser.deepClone(res.shape)
    const cNvPr = newShape['p:nvSpPr']?.['p:cNvPr']

    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    const existingIds = this.#getAllShapeIds(spTree)
    const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 1000
    const newId = maxId + 1

    if (cNvPr) {
      cNvPr['@_id'] = String(newId)
      cNvPr['@_name'] = newShapeId
    }

    const xfrm = newShape['p:spPr']?.['a:xfrm']
    if (xfrm) {
      if (options.offsetX !== undefined) {
        const dx =
          options.offsetX < 100 ? Math.round(options.offsetX * 914400) : Math.round(options.offsetX)
        const x = parseInt(xfrm['a:off']?.['@_x'] || 0, 10) + dx
        if (!xfrm['a:off']) xfrm['a:off'] = {}
        xfrm['a:off']['@_x'] = String(x)
      }
      if (options.offsetY !== undefined) {
        const dy =
          options.offsetY < 100 ? Math.round(options.offsetY * 914400) : Math.round(options.offsetY)
        const y = parseInt(xfrm['a:off']?.['@_y'] || 0, 10) + dy
        if (!xfrm['a:off']) xfrm['a:off'] = {}
        xfrm['a:off']['@_y'] = String(y)
      }
      if (options.width !== undefined) {
        const cx =
          options.width < 100 ? Math.round(options.width * 914400) : Math.round(options.width)
        if (!xfrm['a:ext']) xfrm['a:ext'] = {}
        xfrm['a:ext']['@_cx'] = String(cx)
      }
      if (options.height !== undefined) {
        const cy =
          options.height < 100 ? Math.round(options.height * 914400) : Math.round(options.height)
        if (!xfrm['a:ext']) xfrm['a:ext'] = {}
        xfrm['a:ext']['@_cy'] = String(cy)
      }
    }

    const parent = res.parent
    if (!parent['p:sp']) parent['p:sp'] = []
    if (!Array.isArray(parent['p:sp'])) {
      parent['p:sp'] = [parent['p:sp']]
    }
    parent['p:sp'].push(newShape)

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Cloned shape "${shapeId}" as "${newShapeId}"`)
  }

  /**
   * Deletes a shape from the slide.
   *
   * @param {number} slideIndex
   * @param {string} shapeId
   * @param {SlideManager} slideManager
   */
  deleteShape(slideIndex, shapeId, slideManager) {
    const res = slideManager.getSlideShape(slideIndex, shapeId)

    if (!res) {
      throw new PPTXError(`Shape "${shapeId}" not found in slide ${slideIndex}`)
    }

    const parent = res.parent
    if (parent['p:sp']) {
      if (Array.isArray(parent['p:sp'])) {
        parent['p:sp'] = parent['p:sp'].filter(s => s !== res.shape)
      } else if (parent['p:sp'] === res.shape) {
        delete parent['p:sp']
      }
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Deleted shape "${shapeId}" from slide ${slideIndex}`)
  }

  /**
   * Gets all shapes inside a slide.
   *
   * @param {number} slideIndex
   * @param {SlideManager} slideManager
   * @returns {Array<Object>}
   */
  getShapes(slideIndex, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    const shapesInfo = []
    this.#collectShapesInfo(spTree, shapesInfo)
    return shapesInfo
  }

  /**
   * Helper to recursively scan a container for shapes.
   */
  findShapeRecursive(container, shapeId) {
    if (!container) return null

    let shapes = container['p:sp'] || []
    if (!Array.isArray(shapes)) shapes = [shapes]

    for (const shape of shapes) {
      const cNvPr = shape?.['p:nvSpPr']?.['p:cNvPr']
      if (cNvPr) {
        if (cNvPr['@_name'] === shapeId || String(cNvPr['@_id']) === shapeId) {
          return { shape, parent: container, type: 'sp' }
        }
      }
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]

    for (const group of groups) {
      const res = this.findShapeRecursive(group, shapeId)
      if (res) return res
    }

    return null
  }

  #setShapeTextObj(shape, text) {
    const val = text === undefined || text === null ? '' : String(text)

    if (!shape['p:txBody']) {
      shape['p:txBody'] = {
        'a:bodyPr': {},
        'a:lstStyle': {},
        'a:p': [],
      }
    }

    const txBody = shape['p:txBody']
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
      newParas.push(p)
    }

    txBody['a:p'] = newParas
  }

  #getAllShapeIds(container) {
    const ids = []
    if (!container) return ids

    let shapes = container['p:sp'] || []
    if (!Array.isArray(shapes)) shapes = [shapes]
    for (const s of shapes) {
      const id = parseInt(s?.['p:nvSpPr']?.['p:cNvPr']?.['@_id'], 10)
      if (!isNaN(id)) ids.push(id)
    }

    let pics = container['p:pic'] || []
    if (!Array.isArray(pics)) pics = [pics]
    for (const p of pics) {
      const id = parseInt(p?.['p:nvPicPr']?.['p:cNvPr']?.['@_id'], 10)
      if (!isNaN(id)) ids.push(id)
    }

    let frames = container['p:graphicFrame'] || []
    if (!Array.isArray(frames)) frames = [frames]
    for (const f of frames) {
      const id = parseInt(f?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_id'], 10)
      if (!isNaN(id)) ids.push(id)
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const g of groups) {
      const id = parseInt(g?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_id'], 10)
      if (!isNaN(id)) ids.push(id)
      ids.push(...this.#getAllShapeIds(g))
    }

    return ids
  }

  #collectShapesInfo(container, results) {
    if (!container) return

    let shapes = container['p:sp'] || []
    if (!Array.isArray(shapes)) shapes = [shapes]

    for (const shape of shapes) {
      const cNvPr = shape?.['p:nvSpPr']?.['p:cNvPr']
      if (!cNvPr) continue

      const name = cNvPr['@_name']
      const id = String(cNvPr['@_id'])

      let text = ''
      const txBody = shape['p:txBody']
      if (txBody && txBody['a:p']) {
        const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
        const textParts = []
        for (const p of paras) {
          if (p['a:r']) {
            const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
            for (const r of runs) {
              if (r['a:t']) textParts.push(String(r['a:t']))
            }
          }
        }
        text = textParts.join('\n')
      }

      const xfrm = shape['p:spPr']?.['a:xfrm']
      const position = xfrm
        ? {
            x: parseInt(xfrm['a:off']?.['@_x'] || 0, 10),
            y: parseInt(xfrm['a:off']?.['@_y'] || 0, 10),
            cx: parseInt(xfrm['a:ext']?.['@_cx'] || 0, 10),
            cy: parseInt(xfrm['a:ext']?.['@_cy'] || 0, 10),
          }
        : null

      results.push({
        type: 'shape',
        id,
        name,
        text,
        position,
      })
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const g of groups) {
      this.#collectShapesInfo(g, results)
    }
  }
}

module.exports = { ShapeManager }
