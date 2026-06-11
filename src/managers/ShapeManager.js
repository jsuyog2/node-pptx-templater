/**
 * @fileoverview ShapeManager - Handles shape text replacement, cloning, deletion, and enumeration.
 */

const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')
const { Z_ORDER_SYMBOL } = require('../parsers/XMLParser.js')

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

    const zOrder = parent[Z_ORDER_SYMBOL]
    if (zOrder && res.shape['p:nvSpPr']?.['p:cNvPr']?.['@_id']) {
      const elementId = String(res.shape['p:nvSpPr']['p:cNvPr']['@_id'])
      parent[Z_ORDER_SYMBOL] = zOrder.filter(id => String(id) !== elementId)
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

  /**
   * Validates shape configuration options.
   *
   * @param {Object} options Shape configuration options.
   * @returns {Array<string>} Array of validation error messages.
   */
  validateShape(options) {
    const errors = []

    if (!options) {
      errors.push('Shape options object is missing.')
      return errors
    }

    // Missing IDs
    if (options.id === undefined || options.id === null || String(options.id).trim() === '') {
      errors.push('Shape ID is missing or empty.')
    }

    // Unsupported shape types
    const supportedTypes = [
      'rectangle',
      'square',
      'circle',
      'ellipse',
      'roundedRectangle',
      'triangle',
      'star5',
      'upArrow',
      'downArrow',
      'leftArrow',
      'rightArrow'
    ]
    if (!options.type) {
      errors.push('Shape type is missing.')
    } else if (!supportedTypes.includes(options.type)) {
      errors.push(
        `Unsupported shape type: "${options.type}". Supported types are: ${supportedTypes.join(
          ', '
        )}.`
      )
    }

    // Invalid dimensions
    if (options.type === 'square') {
      if (options.size === undefined || typeof options.size !== 'number' || options.size <= 0) {
        errors.push('Square "size" must be a positive number.')
      }
    } else if (options.type === 'circle') {
      if (options.radius === undefined || typeof options.radius !== 'number' || options.radius <= 0) {
        errors.push('Circle "radius" must be a positive number.')
      }
    } else if (options.type) {
      if (options.width === undefined || typeof options.width !== 'number' || options.width <= 0) {
        errors.push(`${options.type} "width" must be a positive number.`)
      }
      if (options.height === undefined || typeof options.height !== 'number' || options.height <= 0) {
        errors.push(`${options.type} "height" must be a positive number.`)
      }
    }

    if (options.x !== undefined && typeof options.x !== 'number') {
      errors.push('Shape coordinate "x" must be a number.')
    }
    if (options.y !== undefined && typeof options.y !== 'number') {
      errors.push('Shape coordinate "y" must be a number.')
    }

    const isValidHexColor = (color) => {
      if (typeof color !== 'string') return false
      return /^#?[0-9a-fA-F]{3}([0-9a-fA-F]{3})?$/.test(color)
    }

    // Invalid colors
    if (options.fill) {
      if (typeof options.fill === 'string') {
        if (!isValidHexColor(options.fill)) {
          errors.push(`Invalid fill color: "${options.fill}". Must be a valid hex color.`)
        }
      } else if (typeof options.fill === 'object') {
        if (options.fill.type === 'gradient') {
          if (!Array.isArray(options.fill.colors) || options.fill.colors.length < 2) {
            errors.push('Gradient fill must contain an array of at least 2 colors.')
          } else {
            options.fill.colors.forEach((color, idx) => {
              if (!isValidHexColor(color)) {
                errors.push(
                  `Invalid gradient fill color at index ${idx}: "${color}". Must be a valid hex color.`
                )
              }
            })
          }
        } else {
          errors.push('Invalid fill configuration.')
        }
      }
    }

    if (options.border) {
      if (typeof options.border !== 'object') {
        errors.push('Border option must be an object.')
      } else {
        if (options.border.color && !isValidHexColor(options.border.color)) {
          errors.push(`Invalid border color: "${options.border.color}". Must be a valid hex color.`)
        }
        if (
          options.border.width !== undefined &&
          (typeof options.border.width !== 'number' || options.border.width <= 0)
        ) {
          errors.push('Border width must be a positive number.')
        }
      }
    }

    if (options.textStyle) {
      if (typeof options.textStyle !== 'object') {
        errors.push('TextStyle option must be an object.')
      } else {
        if (options.textStyle.color && !isValidHexColor(options.textStyle.color)) {
          errors.push(
            `Invalid text style color: "${options.textStyle.color}". Must be a valid hex color.`
          )
        }
        if (
          options.textStyle.fontSize !== undefined &&
          (typeof options.textStyle.fontSize !== 'number' || options.textStyle.fontSize <= 0)
        ) {
          errors.push('Text style fontSize must be a positive number.')
        }
      }
    }

    // Invalid border radius
    if (options.borderRadius !== undefined) {
      if (options.type !== 'roundedRectangle') {
        errors.push(`Shape type "${options.type}" does not support borderRadius.`)
      } else if (typeof options.borderRadius !== 'number' || options.borderRadius <= 0) {
        errors.push('Border radius must be a positive number.')
      }
    }

    return errors
  }

  /**
   * Adds a new shape to a slide.
   */
  addShape(slideIndex, options, slideManager) {
    const errors = this.validateShape(options)
    if (errors.length > 0) {
      throw new PPTXError(`Shape validation failed:\n- ${errors.join('\n- ')}`)
    }

    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    if (!spTree) {
      throw new PPTXError(`Invalid slide structure for slide ${slideIndex}`)
    }

    // Generate unique shape ID
    const existingIds = this.#getAllShapeIds(spTree)
    const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 1000
    const newId = maxId + 1

    // Build shape XML
    const type = options.type
    let preset = 'rect'
    let width = options.width || 100
    let height = options.height || 100

    if (type === 'square') {
      preset = 'rect'
      width = options.size || 100
      height = options.size || 100
    } else if (type === 'circle') {
      preset = 'ellipse'
      width = (options.radius || 50) * 2
      height = (options.radius || 50) * 2
    } else if (type === 'ellipse') {
      preset = 'ellipse'
    } else if (type === 'roundedRectangle') {
      preset = 'roundRect'
    } else if (
      [
        'triangle',
        'star5',
        'upArrow',
        'downArrow',
        'leftArrow',
        'rightArrow'
      ].includes(type)
    ) {
      preset = type
    }

    const xEmu = Math.round((options.x || 0) * 9525)
    const yEmu = Math.round((options.y || 0) * 9525)
    const wEmu = Math.round(width * 9525)
    const hEmu = Math.round(height * 9525)

    const name = options.id || `${type.charAt(0).toUpperCase() + type.slice(1)} ${newId}`

    // Fill properties
    let fillXml = '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>' // default
    if (options.fill) {
      if (typeof options.fill === 'string') {
        const hex = options.fill.replace('#', '')
        const alphaXml =
          options.transparency !== undefined
            ? `<a:alpha val="${Math.round((100 - options.transparency) * 1000)}"/>`
            : ''
        fillXml = `<a:solidFill><a:srgbClr val="${hex}">${alphaXml}</a:srgbClr></a:solidFill>`
      } else if (typeof options.fill === 'object' && options.fill.type === 'gradient') {
        const colors = options.fill.colors || []
        const transparency = options.transparency
        fillXml = `<a:gradFill flip="none" rotWithShape="1">
          <a:gsLst>
            ${colors
              .map((color, i) => {
                const pos = Math.round((i / (colors.length - 1)) * 100000)
                const hexColor = color.replace('#', '')
                const alphaXml =
                  transparency !== undefined
                    ? `<a:alpha val="${Math.round((100 - transparency) * 1000)}"/>`
                    : ''
                return `<a:gs pos="${pos}"><a:srgbClr val="${hexColor}">${alphaXml}</a:srgbClr></a:gs>`
              })
              .join('')}
          </a:gsLst>
          <a:lin ang="5400000" scaled="1"/>
        </a:gradFill>`
      }
    }

    // Border properties
    let borderXml = ''
    if (options.border) {
      const bColor = (options.border.color || '#000000').replace('#', '')
      const bWidth = Math.round((options.border.width || 1) * 9525)
      borderXml = `<a:ln w="${bWidth}">
        <a:solidFill><a:srgbClr val="${bColor}"/></a:solidFill>
      </a:ln>`
    }

    // Adjustments (border radius)
    let avLstXml = '<a:avLst/>'
    if (preset === 'roundRect') {
      let adjVal = 16667 // PPT default
      if (options.borderRadius !== undefined) {
        const shorterSide = Math.min(width, height)
        adjVal = Math.min(
          50000,
          Math.max(0, Math.round((options.borderRadius / shorterSide) * 100000))
        )
      }
      avLstXml = `<a:avLst><a:gd name="adj" fmla="val ${adjVal}"/></a:avLst>`
    }

    // Shadow properties
    let shadowXml = ''
    if (options.shadow) {
      let blur = 5
      let distance = 3
      let opacity = 50
      if (typeof options.shadow === 'object') {
        if (options.shadow.blur !== undefined) blur = options.shadow.blur
        if (options.shadow.distance !== undefined) distance = options.shadow.distance
        if (options.shadow.opacity !== undefined) opacity = options.shadow.opacity
      }
      const blurEmu = Math.round(blur * 9525)
      const distEmu = Math.round(distance * 9525)
      const alphaVal = Math.round(opacity * 1000)
      shadowXml = `<a:effectLst>
        <a:outerShdw blurRad="${blurEmu}" dist="${distEmu}" dir="5400000" algn="tl" rotWithShape="0">
          <a:srgbClr val="000000">
            <a:alpha val="${alphaVal}"/>
          </a:srgbClr>
        </a:outerShdw>
      </a:effectLst>`
    }

    // Rotation attribute
    const rotAttr =
      options.rotation !== undefined ? ` rot="${Math.round(options.rotation * 60000)}"` : ''

    // Text box body properties
    let txBodyXml = ''
    if (options.text !== undefined && options.text !== null) {
      const textStyle = options.textStyle || {}
      const fontSizeVal = (textStyle.fontSize || 14) * 100
      const boldAttr = textStyle.bold ? ' b="1"' : ''
      const italicAttr = textStyle.italic ? ' i="1"' : ''

      let alignAttr = ''
      if (textStyle.align) {
        const alignMap = { center: 'ctr', right: 'r', left: 'l', justify: 'just' }
        const algn = alignMap[textStyle.align] || 'l'
        alignAttr = `<a:pPr algn="${algn}"/>`
      }

      let colorFill = ''
      if (textStyle.color) {
        const colorHex = textStyle.color.replace('#', '')
        colorFill = `<a:solidFill><a:srgbClr val="${colorHex}"/></a:solidFill>`
      }

      const lines = String(options.text).split(/\r?\n/)
      const paragraphsXml = lines
        .map((line) => {
          return `<a:p>
          ${alignAttr}
          <a:r>
            <a:rPr lang="en-US" sz="${fontSizeVal}"${boldAttr}${italicAttr}>
              ${colorFill}
            </a:rPr>
            <a:t>${escapeXml(line)}</a:t>
          </a:r>
        </a:p>`
        })
        .join('')

      txBodyXml = `<p:txBody>
        <a:bodyPr wrap="square" rtlCol="0">
          <a:normAutofit/>
        </a:bodyPr>
        <a:lstStyle/>
        ${paragraphsXml}
      </p:txBody>`
    }

    // Build shape XML block
    const shapeXml = `<p:sp>
      <p:nvSpPr>
        <p:cNvPr id="${newId}" name="${escapeXml(name)}"/>
        <p:cNvSpPr/>
        <p:nvPr/>
      </p:nvSpPr>
      <p:spPr>
        <a:xfrm${rotAttr}>
          <a:off x="${xEmu}" y="${yEmu}"/>
          <a:ext cx="${wEmu}" cy="${hEmu}"/>
        </a:xfrm>
        <a:prstGeom prst="${preset}">${avLstXml}</a:prstGeom>
        ${fillXml}
        ${borderXml}
        ${shadowXml}
      </p:spPr>
      ${txBodyXml}
    </p:sp>`

    const parsed = this.#xmlParser.parse(shapeXml, 'shape.xml')['p:sp']
    const shapeObj = Array.isArray(parsed) ? parsed[0] : parsed

    if (!spTree['p:sp']) {
      spTree['p:sp'] = []
    }
    if (!Array.isArray(spTree['p:sp'])) {
      spTree['p:sp'] = [spTree['p:sp']]
    }
    spTree['p:sp'].push(shapeObj)

    if (spTree[Z_ORDER_SYMBOL] && !spTree[Z_ORDER_SYMBOL].includes(String(newId))) {
      spTree[Z_ORDER_SYMBOL].push(String(newId))
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Added shape "${name}" with ID ${newId} to slide ${slideIndex}`)
  }

  /**
   * Updates an existing shape in-place.
   */
  updateShape(slideIndex, shapeId, options, slideManager) {
    const res = slideManager.getSlideShape(slideIndex, shapeId)

    if (!res) {
      throw new PPTXError(`Shape "${shapeId}" not found in slide ${slideIndex}`)
    }

    const shape = res.shape
    const spPr = shape['p:spPr']

    if (!spPr) {
      throw new PPTXError(`Invalid shape structure for "${shapeId}" on slide ${slideIndex}`)
    }

    // Update coordinates & dimensions
    let w = options.width
    let h = options.height
    if (options.size !== undefined) {
      w = options.size
      h = options.size
    } else if (options.radius !== undefined) {
      w = options.radius * 2
      h = options.radius * 2
    }

    const xfrm = spPr['a:xfrm']
    if (xfrm) {
      if (options.x !== undefined) xfrm['a:off']['@_x'] = String(Math.round(options.x * 9525))
      if (options.y !== undefined) xfrm['a:off']['@_y'] = String(Math.round(options.y * 9525))
      if (w !== undefined) xfrm['a:ext']['@_cx'] = String(Math.round(w * 9525))
      if (h !== undefined) xfrm['a:ext']['@_cy'] = String(Math.round(h * 9525))

      // Update rotation
      if (options.rotation !== undefined) {
        if (options.rotation === null) {
          delete xfrm['@_rot']
        } else {
          xfrm['@_rot'] = String(Math.round(options.rotation * 60000))
        }
      }
    }

    // Update preset geometry type if specified
    if (options.type) {
      let preset = 'rect'
      if (options.type === 'square') preset = 'rect'
      else if (options.type === 'circle' || options.type === 'ellipse') preset = 'ellipse'
      else if (options.type === 'roundedRectangle') preset = 'roundRect'
      else if (
        [
          'triangle',
          'star5',
          'upArrow',
          'downArrow',
          'leftArrow',
          'rightArrow'
        ].includes(options.type)
      ) {
        preset = options.type
      }
      if (spPr['a:prstGeom']) {
        spPr['a:prstGeom']['@_prst'] = preset
      }
    }

    // Update fill properties
    let fillXml = ''
    if (options.fill) {
      if (typeof options.fill === 'string') {
        const hex = options.fill.replace('#', '')
        const alphaXml =
          options.transparency !== undefined
            ? `<a:alpha val="${Math.round((100 - options.transparency) * 1000)}"/>`
            : ''
        fillXml = `<a:solidFill><a:srgbClr val="${hex}">${alphaXml}</a:srgbClr></a:solidFill>`
      } else if (typeof options.fill === 'object' && options.fill.type === 'gradient') {
        const colors = options.fill.colors || []
        const transparency = options.transparency
        fillXml = `<a:gradFill flip="none" rotWithShape="1">
          <a:gsLst>
            ${colors
              .map((color, i) => {
                const pos = Math.round((i / (colors.length - 1)) * 100000)
                const hexColor = color.replace('#', '')
                const alphaXml =
                  transparency !== undefined
                    ? `<a:alpha val="${Math.round((100 - transparency) * 1000)}"/>`
                    : ''
                return `<a:gs pos="${pos}"><a:srgbClr val="${hexColor}">${alphaXml}</a:srgbClr></a:gs>`
              })
              .join('')}
          </a:gsLst>
          <a:lin ang="5400000" scaled="1"/>
        </a:gradFill>`
      }

      if (fillXml) {
        delete spPr['a:noFill']
        delete spPr['a:solidFill']
        delete spPr['a:gradFill']
        delete spPr['a:pattFill']
        delete spPr['a:grpFill']
        const parsedFill = this.#xmlParser.parse(fillXml, 'fill.xml')
        const fillKey = Object.keys(parsedFill)[0]
        const fillVal = parsedFill[fillKey]
        spPr[fillKey] = Array.isArray(fillVal) ? fillVal[0] : fillVal
      }
    } else if (options.transparency !== undefined) {
      // Transparency only
      const solidFill = spPr['a:solidFill']
      if (solidFill && solidFill['a:srgbClr']) {
        const srgbClr = solidFill['a:srgbClr']
        srgbClr['a:alpha'] = { '@_val': String(Math.round((100 - options.transparency) * 1000)) }
      }
    }

    // Update border properties
    if (options.border) {
      delete spPr['a:ln']
      const bColor = (options.border.color || '#000000').replace('#', '')
      const bWidth = Math.round((options.border.width || 1) * 9525)
      const borderXml = `<a:ln w="${bWidth}">
        <a:solidFill><a:srgbClr val="${bColor}"/></a:solidFill>
      </a:ln>`
      const parsedBorder = this.#xmlParser.parse(borderXml, 'border.xml')['a:ln']
      spPr['a:ln'] = Array.isArray(parsedBorder) ? parsedBorder[0] : parsedBorder
    }

    // Update shadow properties
    if (options.shadow !== undefined) {
      delete spPr['a:effectLst']
      if (options.shadow) {
        let blur = 5
        let distance = 3
        let opacity = 50
        if (typeof options.shadow === 'object') {
          if (options.shadow.blur !== undefined) blur = options.shadow.blur
          if (options.shadow.distance !== undefined) distance = options.shadow.distance
          if (options.shadow.opacity !== undefined) opacity = options.shadow.opacity
        }
        const blurEmu = Math.round(blur * 9525)
        const distEmu = Math.round(distance * 9525)
        const alphaVal = Math.round(opacity * 1000)
        const shadowXml = `<a:effectLst>
          <a:outerShdw blurRad="${blurEmu}" dist="${distEmu}" dir="5400000" algn="tl" rotWithShape="0">
            <a:srgbClr val="000000">
              <a:alpha val="${alphaVal}"/>
            </a:srgbClr>
          </a:outerShdw>
        </a:effectLst>`
        const parsedShadow = this.#xmlParser.parse(shadowXml, 'shadow.xml')['a:effectLst']
        spPr['a:effectLst'] = Array.isArray(parsedShadow) ? parsedShadow[0] : parsedShadow
      }
    }

    // Update border radius
    if (options.borderRadius !== undefined) {
      const prstGeom = spPr['a:prstGeom']
      if (prstGeom && prstGeom['@_prst'] === 'roundRect') {
        let curW = 100
        let curH = 100
        if (xfrm) {
          curW = parseInt(xfrm['a:ext']?.['@_cx'] || 0, 10) / 9525 || 100
          curH = parseInt(xfrm['a:ext']?.['@_cy'] || 0, 10) / 9525 || 100
        }
        const shorterSide = Math.min(curW, curH)
        const adjVal = Math.min(
          50000,
          Math.max(0, Math.round((options.borderRadius / shorterSide) * 100000))
        )
        const avLstXml = `<a:avLst><a:gd name="adj" fmla="val ${adjVal}"/></a:avLst>`
        const parsedAvLst = this.#xmlParser.parse(avLstXml, 'avLst.xml')['a:avLst']
        prstGeom['a:avLst'] = Array.isArray(parsedAvLst) ? parsedAvLst[0] : parsedAvLst
      }
    }

    // Update text
    if (options.text !== undefined || options.textStyle !== undefined) {
      let textVal = options.text
      if (textVal === undefined) {
        textVal = this.getShapeText(shape) || ''
      }

      const textStyle = options.textStyle || {}
      const fontSizeVal = (textStyle.fontSize || 14) * 100
      const boldAttr = textStyle.bold ? ' b="1"' : ''
      const italicAttr = textStyle.italic ? ' i="1"' : ''

      let alignAttr = ''
      if (textStyle.align) {
        const alignMap = { center: 'ctr', right: 'r', left: 'l', justify: 'just' }
        const algn = alignMap[textStyle.align] || 'l'
        alignAttr = `<a:pPr algn="${algn}"/>`
      }

      let colorFill = ''
      if (textStyle.color) {
        const colorHex = textStyle.color.replace('#', '')
        colorFill = `<a:solidFill><a:srgbClr val="${colorHex}"/></a:solidFill>`
      }

      const lines = String(textVal).split(/\r?\n/)
      const paragraphsXml = lines
        .map((line) => {
          return `<a:p>
          ${alignAttr}
          <a:r>
            <a:rPr lang="en-US" sz="${fontSizeVal}"${boldAttr}${italicAttr}>
              ${colorFill}
            </a:rPr>
            <a:t>${escapeXml(line)}</a:t>
          </a:r>
        </a:p>`
        })
        .join('')

      const txBodyXml = `<p:txBody>
        <a:bodyPr wrap="square" rtlCol="0">
          <a:normAutofit/>
        </a:bodyPr>
        <a:lstStyle/>
        ${paragraphsXml}
      </p:txBody>`

      const parsedTxBody = this.#xmlParser.parse(txBodyXml, 'txBody.xml')['p:txBody']
      shape['p:txBody'] = Array.isArray(parsedTxBody) ? parsedTxBody[0] : parsedTxBody
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Updated shape "${shapeId}" on slide ${slideIndex}`)
  }

  /**
   * Removes a shape from a slide (alias for deleteShape).
   */
  removeShape(slideIndex, shapeId, slideManager) {
    this.deleteShape(slideIndex, shapeId, slideManager)
  }

  /**
   * Helper to get existing text of a shape.
   */
  getShapeText(shape) {
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
      return textParts.join('\n')
    }
    return ''
  }

  /**
   * Gets details of a shape.
   */
  getShape(slideIndex, shapeId, slideManager) {
    const res = slideManager.getSlideShape(slideIndex, shapeId)
    if (!res) return null

    const cNvPr = res.shape['p:nvSpPr']?.['p:cNvPr']
    const xfrm = res.shape['p:spPr']?.['a:xfrm']
    const prstGeom = res.shape['p:spPr']?.['a:prstGeom']

    const id = cNvPr ? cNvPr['@_name'] || String(cNvPr['@_id']) : shapeId
    const preset = prstGeom ? prstGeom['@_prst'] : 'rect'

    const xEmu = xfrm?.['a:off']?.['@_x'] ? parseInt(xfrm['a:off']['@_x'], 10) : 0
    const yEmu = xfrm?.['a:off']?.['@_y'] ? parseInt(xfrm['a:off']['@_y'], 10) : 0
    const wEmu = xfrm?.['a:ext']?.['@_cx'] ? parseInt(xfrm['a:ext']['@_cx'], 10) : 0
    const hEmu = xfrm?.['a:ext']?.['@_cy'] ? parseInt(xfrm['a:ext']['@_cy'], 10) : 0

    const x = Math.round(xEmu / 9525)
    const y = Math.round(yEmu / 9525)
    const width = Math.round(wEmu / 9525)
    const height = Math.round(hEmu / 9525)

    const type = mapPresetToType(preset, width, height)

    return {
      id,
      type,
      x,
      y,
      width,
      height,
    }
  }
}

function escapeXml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

function mapPresetToType(preset, w, h) {
  if (preset === 'rect') {
    return w === h ? 'square' : 'rectangle'
  }
  if (preset === 'ellipse') {
    return w === h ? 'circle' : 'ellipse'
  }
  if (preset === 'roundRect') {
    return 'roundedRectangle'
  }
  if (
    [
      'triangle',
      'star5',
      'upArrow',
      'downArrow',
      'leftArrow',
      'rightArrow'
    ].includes(preset)
  ) {
    return preset
  }
  return 'rectangle'
}

module.exports = { ShapeManager }
