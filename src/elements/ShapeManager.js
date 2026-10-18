/**
 * @fileoverview ShapeManager - Strict 10-Method Shapes & Geometry Manager.
 *
 * Whitelisted public methods:
 * 1. addShape
 * 2. updateShape
 * 3. removeShape
 * 4. clearShapes
 * 5. getShapes
 * 6. getShape
 * 7. cloneShape
 * 8. getGroupShapes
 * 9. groupShapes
 * 10. ungroupShapes
 */

const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')
const { Z_ORDER_SYMBOL } = require('../parsers/XMLParser.js')
const { resolveElementBounds, resolveCoordinate } = require('../utils/coordinateUtils.js')

const logger = createLogger('ShapeManager')

/**
 * @class ShapeManager
 * @description Manages DrawingML shape elements inside PPTX presentation slides.
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
   * Adds a new shape to a slide.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string|Object} typeOrOptions - Preset shape geometry or options object.
   * @param {Object} [maybeOptions={}] - Configuration options if type specified as string.
   * @param {SlideManager} [maybeSlideManager] - SlideManager instance.
   * @returns {string} Unique shape identifier.
   */
  addShape(slideIndex, typeOrOptions, maybeOptions = {}, maybeSlideManager) {
    let options
    let slideManager
    if (typeof typeOrOptions === 'string') {
      options = { ...maybeOptions, type: typeOrOptions }
      slideManager = maybeSlideManager
    } else {
      options = { ...typeOrOptions }
      slideManager = maybeOptions
    }

    const validationErrors = this.#validateShapeOptions(options)
    if (validationErrors.length > 0) {
      throw new PPTXError(`Invalid shape options: ${validationErrors.join('; ')}`)
    }

    const normalized = this.#normalizeShapeOptions(options)
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

    let assignedId = newId
    if (normalized.id !== undefined && !isNaN(parseInt(normalized.id, 10))) {
      const parsedNum = parseInt(normalized.id, 10)
      if (parsedNum > 0 && !existingIds.includes(parsedNum)) {
        assignedId = parsedNum
      }
    }

    const type = normalized.type || 'rectangle'
    const preset = normalizePresetGeometry(type)
    let width = normalized.width || 100
    let height = normalized.height || 100

    if (preset === 'rect') {
      if (type === 'square' || normalized.size !== undefined) {
        width = normalized.size || width
        height = normalized.size || height
      }
    } else if (preset === 'ellipse') {
      if (type === 'circle' || normalized.radius !== undefined) {
        width = (normalized.radius || width / 2) * 2
        height = (normalized.radius || height / 2) * 2
      }
    }

    const bounds = resolveElementBounds(
      {
        ...normalized,
        width,
        height,
      },
      slideManager
    )
    const xEmu = bounds.x
    const yEmu = bounds.y
    const wEmu = bounds.width
    const hEmu = bounds.height

    const fallbackName = `${preset.charAt(0).toUpperCase() + preset.slice(1)} ${assignedId}`
    const name =
      normalized.name ||
      (normalized.id && isNaN(parseInt(normalized.id, 10)) ? normalized.id : fallbackName)

    const isLine = preset === 'line'
    let headEndXml = ''
    let tailEndXml = ''
    if (
      normalized.headEnd ||
      normalized.arrow ||
      normalized.endArrow ||
      type === 'lineArrow' ||
      type === 'arrowLine'
    ) {
      const headType = typeof normalized.headEnd === 'string' ? normalized.headEnd : 'triangle'
      headEndXml = `<a:headEnd type="${headType}"/>`
    }
    if (normalized.tailEnd || normalized.startArrow) {
      const tailType = typeof normalized.tailEnd === 'string' ? normalized.tailEnd : 'triangle'
      tailEndXml = `<a:tailEnd type="${tailType}"/>`
    }

    // Fill properties
    let fillXml =
      isLine && !normalized.fill
        ? '<a:noFill/>'
        : '<a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>'
    if (normalized.fill) {
      if (typeof normalized.fill === 'string') {
        const fillStr = normalized.fill.toLowerCase()
        if (fillStr === 'none' || fillStr === 'transparent') {
          fillXml = '<a:noFill/>'
        } else {
          const hex = normalized.fill.replace('#', '').toUpperCase()
          const alphaXml =
            normalized.transparency !== undefined
              ? `<a:alpha val="${Math.round((100 - normalized.transparency) * 1000)}"/>`
              : ''
          fillXml = `<a:solidFill><a:srgbClr val="${hex}">${alphaXml}</a:srgbClr></a:solidFill>`
        }
      } else if (typeof normalized.fill === 'object') {
        if (normalized.fill.type === 'none') {
          fillXml = '<a:noFill/>'
        } else if (normalized.fill.type === 'gradient') {
          const colors = normalized.fill.colors || []
          const transparency = normalized.transparency
          fillXml = `<a:gradFill flip="none" rotWithShape="1">
            <a:gsLst>
              ${colors
                .map((color, i) => {
                  const pos = Math.round((i / (colors.length - 1)) * 100000)
                  const hexColor = color.replace('#', '').toUpperCase()
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
    }

    // Border properties
    let borderXml = ''
    if (normalized.border) {
      if (normalized.border === 'none' || normalized.border === false) {
        borderXml = '<a:ln><a:noFill/></a:ln>'
      } else {
        const bColor = (normalized.border.color || '#000000').replace('#', '').toUpperCase()
        const bWidth = Math.round((normalized.border.width || 1) * 9525)
        borderXml = `<a:ln w="${bWidth}">
          <a:solidFill><a:srgbClr val="${bColor}"/></a:solidFill>${headEndXml}${tailEndXml}
        </a:ln>`
      }
    } else if (isLine || headEndXml || tailEndXml) {
      const strokeColor = (normalized.strokeColor || normalized.color || '#000000')
        .replace('#', '')
        .toUpperCase()
      const strokeWidth = Math.round(
        (normalized.strokeWidth || normalized.lineWidth || (isLine ? 2 : 1)) * 9525
      )
      borderXml = `<a:ln w="${strokeWidth}">
        <a:solidFill><a:srgbClr val="${strokeColor}"/></a:solidFill>${headEndXml}${tailEndXml}
      </a:ln>`
    }

    // Adjustments (border radius)
    let avLstXml = '<a:avLst/>'
    if (preset === 'roundRect') {
      let adjVal = 16667
      if (normalized.borderRadius !== undefined) {
        const shorterSide = Math.min(width, height)
        adjVal = Math.min(
          50000,
          Math.max(0, Math.round((normalized.borderRadius / shorterSide) * 100000))
        )
      }
      avLstXml = `<a:avLst><a:gd name="adj" fmla="val ${adjVal}"/></a:avLst>`
    }

    // Shadow properties
    let shadowXml = ''
    if (normalized.shadow) {
      let blur = 5
      let distance = 3
      let opacity = 50
      if (typeof normalized.shadow === 'object') {
        if (normalized.shadow.blur !== undefined) blur = normalized.shadow.blur
        if (normalized.shadow.distance !== undefined) distance = normalized.shadow.distance
        if (normalized.shadow.opacity !== undefined) opacity = normalized.shadow.opacity
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

    // Rotation & Flip attributes
    let xfrmAttrs = ''
    if (normalized.rotation !== undefined) {
      xfrmAttrs += ` rot="${Math.round(normalized.rotation * 60000)}"`
    }
    if (normalized.flipH) {
      xfrmAttrs += ' flipH="1"'
    }
    if (normalized.flipV) {
      xfrmAttrs += ' flipV="1"'
    }

    // Text box body
    let txBodyXml = ''
    if (normalized.text !== undefined && normalized.text !== null) {
      txBodyXml = buildTextBodyXml(normalized.text, normalized.textStyle, normalized)
    } else {
      txBodyXml = `<p:txBody><a:bodyPr rtlCol="0"/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody>`
    }

    const shapeXml = `<p:sp><p:nvSpPr><p:cNvPr id="${assignedId}" name="${escapeXml(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm${xfrmAttrs}><a:off x="${xEmu}" y="${yEmu}"/><a:ext cx="${wEmu}" cy="${hEmu}"/></a:xfrm><a:prstGeom prst="${preset}">${avLstXml}</a:prstGeom>${fillXml}${borderXml}${shadowXml}</p:spPr>${txBodyXml}</p:sp>`

    const parsed = this.#xmlParser.parse(shapeXml, 'shape.xml')['p:sp']
    const shapeObj = Array.isArray(parsed) ? parsed[0] : parsed

    this.#reorderSpPrKeys(shapeObj['p:spPr'])

    if (!spTree['p:sp']) {
      spTree['p:sp'] = []
    }
    if (!Array.isArray(spTree['p:sp'])) {
      spTree['p:sp'] = [spTree['p:sp']]
    }
    spTree['p:sp'].push(shapeObj)

    if (spTree[Z_ORDER_SYMBOL] && !spTree[Z_ORDER_SYMBOL].includes(String(assignedId))) {
      spTree[Z_ORDER_SYMBOL].push(String(assignedId))
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Added shape "${name}" with ID ${assignedId} to slide ${slideIndex}`)
    return String(assignedId)
  }

  /**
   * Updates an existing shape in-place.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string} shapeId - Shape ID or name.
   * @param {Object} options - Updates to apply.
   * @param {SlideManager} slideManager - SlideManager instance.
   */
  updateShape(slideIndex, shapeId, options = {}, slideManager) {
    const res = this.#findShape(slideIndex, shapeId, slideManager)
    if (!res) {
      throw new PPTXError(`Shape "${shapeId}" not found in slide ${slideIndex}`)
    }

    const shape = res.shape
    if (!shape['p:spPr']) {
      shape['p:spPr'] = {}
    }
    const spPr = shape['p:spPr']

    const normalized = this.#normalizeShapeOptions(options, true)

    let w = normalized.width
    let h = normalized.height
    if (normalized.size !== undefined) {
      w = normalized.size
      h = normalized.size
    } else if (normalized.radius !== undefined) {
      w = normalized.radius * 2
      h = normalized.radius * 2
    }

    if (!spPr['a:xfrm']) {
      spPr['a:xfrm'] = {}
    }
    const xfrm = spPr['a:xfrm']
    if (!xfrm['a:off']) xfrm['a:off'] = {}
    if (!xfrm['a:ext']) xfrm['a:ext'] = {}

    const dims =
      slideManager && typeof slideManager.getSlideSize === 'function'
        ? slideManager.getSlideSize()
        : { width: 12192000, height: 6858000 }

    if (normalized.x !== undefined) {
      xfrm['a:off']['@_x'] = String(resolveCoordinate(normalized.x, dims.width, normalized))
    }
    if (normalized.y !== undefined) {
      xfrm['a:off']['@_y'] = String(resolveCoordinate(normalized.y, dims.height, normalized))
    }
    if (w !== undefined) {
      xfrm['a:ext']['@_cx'] = String(resolveCoordinate(w, dims.width, normalized))
    }
    if (h !== undefined) {
      xfrm['a:ext']['@_cy'] = String(resolveCoordinate(h, dims.height, normalized))
    }

    if (normalized.rotation !== undefined) {
      if (normalized.rotation === null) {
        delete xfrm['@_rot']
      } else {
        xfrm['@_rot'] = String(Math.round(normalized.rotation * 60000))
      }
    }

    if (normalized.flipH !== undefined) {
      xfrm['@_flipH'] = normalized.flipH ? '1' : '0'
    }
    if (normalized.flipV !== undefined) {
      xfrm['@_flipV'] = normalized.flipV ? '1' : '0'
    }

    // Update preset geometry type if specified
    if (normalized.type) {
      const preset = normalizePresetGeometry(normalized.type)
      if (!spPr['a:prstGeom']) {
        spPr['a:prstGeom'] = { 'a:avLst': {} }
      }
      spPr['a:prstGeom']['@_prst'] = preset
      if (!spPr['a:prstGeom']['a:avLst']) {
        spPr['a:prstGeom']['a:avLst'] = {}
      }
    }

    // Update fill properties ONLY if fill was explicitly specified
    if (normalized.fill !== undefined) {
      let fillXml = ''
      if (typeof normalized.fill === 'string') {
        const fillStr = normalized.fill.toLowerCase()
        if (fillStr === 'none' || fillStr === 'transparent') {
          fillXml = '<a:noFill/>'
        } else {
          const hex = normalized.fill.replace('#', '').toUpperCase()
          const alphaXml =
            normalized.transparency !== undefined
              ? `<a:alpha val="${Math.round((100 - normalized.transparency) * 1000)}"/>`
              : ''
          fillXml = `<a:solidFill><a:srgbClr val="${hex}">${alphaXml}</a:srgbClr></a:solidFill>`
        }
      } else if (typeof normalized.fill === 'object' && normalized.fill !== null) {
        if (normalized.fill.type === 'none') {
          fillXml = '<a:noFill/>'
        } else if (normalized.fill.type === 'gradient') {
          const colors = normalized.fill.colors || []
          const transparency = normalized.transparency
          fillXml = `<a:gradFill flip="none" rotWithShape="1">
            <a:gsLst>
              ${colors
                .map((color, i) => {
                  const pos = Math.round((i / (colors.length - 1)) * 100000)
                  const hexColor = color.replace('#', '').toUpperCase()
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

      if (fillXml) {
        delete spPr['a:noFill']
        delete spPr['a:solidFill']
        delete spPr['a:gradFill']
        delete spPr['a:pattFill']
        delete spPr['a:grpFill']
        delete spPr['a:blipFill']
        const parsedFill = this.#xmlParser.parse(fillXml, 'fill.xml')
        const fillKey = Object.keys(parsedFill)[0]
        const fillVal = parsedFill[fillKey]
        spPr[fillKey] = Array.isArray(fillVal) ? fillVal[0] : fillVal
      }
    } else if (normalized.transparency !== undefined) {
      const solidFill = spPr['a:solidFill']
      if (solidFill && solidFill['a:srgbClr']) {
        const srgbClr = solidFill['a:srgbClr']
        srgbClr['a:alpha'] = { '@_val': String(Math.round((100 - normalized.transparency) * 1000)) }
      }
    }

    // Update border properties
    if (normalized.border !== undefined) {
      delete spPr['a:ln']
      if (normalized.border === 'none' || normalized.border === false) {
        spPr['a:ln'] = { 'a:noFill': {} }
      } else {
        const bColor = (normalized.border.color || '#000000').replace('#', '').toUpperCase()
        const bWidth = Math.round((normalized.border.width || 1) * 9525)
        let endMarkers = ''
        if (normalized.headEnd || normalized.arrow || normalized.endArrow) {
          const hType = typeof normalized.headEnd === 'string' ? normalized.headEnd : 'triangle'
          endMarkers += `<a:headEnd type="${hType}"/>`
        }
        if (normalized.tailEnd || normalized.startArrow) {
          const tType = typeof normalized.tailEnd === 'string' ? normalized.tailEnd : 'triangle'
          endMarkers += `<a:tailEnd type="${tType}"/>`
        }
        const borderXml = `<a:ln w="${bWidth}">
          <a:solidFill><a:srgbClr val="${bColor}"/></a:solidFill>${endMarkers}
        </a:ln>`
        const parsedBorder = this.#xmlParser.parse(borderXml, 'border.xml')['a:ln']
        spPr['a:ln'] = Array.isArray(parsedBorder) ? parsedBorder[0] : parsedBorder
      }
    } else if (
      normalized.headEnd !== undefined ||
      normalized.tailEnd !== undefined ||
      normalized.arrow !== undefined
    ) {
      if (!spPr['a:ln']) {
        spPr['a:ln'] = {
          '@_w': '19050',
          'a:solidFill': { 'a:srgbClr': { '@_val': '000000' } },
        }
      }
      if (normalized.headEnd || normalized.arrow) {
        const hType = typeof normalized.headEnd === 'string' ? normalized.headEnd : 'triangle'
        spPr['a:ln']['a:headEnd'] = { '@_type': hType }
      }
      if (normalized.tailEnd) {
        const tType = typeof normalized.tailEnd === 'string' ? normalized.tailEnd : 'triangle'
        spPr['a:ln']['a:tailEnd'] = { '@_type': tType }
      }
    }

    // Update shadow properties
    if (normalized.shadow !== undefined) {
      delete spPr['a:effectLst']
      if (normalized.shadow) {
        let blur = 5
        let distance = 3
        let opacity = 50
        if (typeof normalized.shadow === 'object') {
          if (normalized.shadow.blur !== undefined) blur = normalized.shadow.blur
          if (normalized.shadow.distance !== undefined) distance = normalized.shadow.distance
          if (normalized.shadow.opacity !== undefined) opacity = normalized.shadow.opacity
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
    if (normalized.borderRadius !== undefined) {
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
          Math.max(0, Math.round((normalized.borderRadius / shorterSide) * 100000))
        )
        const avLstXml = `<a:avLst><a:gd name="adj" fmla="val ${adjVal}"/></a:avLst>`
        const parsedAvLst = this.#xmlParser.parse(avLstXml, 'avLst.xml')['a:avLst']
        prstGeom['a:avLst'] = Array.isArray(parsedAvLst) ? parsedAvLst[0] : parsedAvLst
      }
    }

    this.#reorderSpPrKeys(spPr)

    // Update text & typography non-destructively
    const textPropKeys = [
      'text',
      'textStyle',
      'fontSize',
      'fontFamily',
      'color',
      'textColor',
      'bold',
      'italic',
      'underline',
      'strike',
      'strikethrough',
      'align',
      'textAlign',
      'alignment',
      'verticalAlign',
      'verticalAlignment',
      'lineSpacing',
      'lineSpacingPts',
      'spaceBefore',
      'spaceAfter',
      'characterSpacing',
      'tracking',
      'superscript',
      'subscript',
    ]
    const hasTextChange = textPropKeys.some(
      k => normalized[k] !== undefined || options[k] !== undefined
    )

    if (hasTextChange) {
      const textVal =
        normalized.text !== undefined ? normalized.text : this.#getShapeText(shape) || ''

      const existingTxBody = shape['p:txBody']
      const existingParas = existingTxBody?.['a:p']
        ? Array.isArray(existingTxBody['a:p'])
          ? existingTxBody['a:p']
          : [existingTxBody['a:p']]
        : []
      const existingStyle = {}
      for (const p of existingParas) {
        if (p && p['a:r']) {
          const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
          for (const r of runs) {
            if (r && r['a:rPr']) {
              const rPr = r['a:rPr']
              if (rPr['@_sz']) existingStyle.fontSize = parseInt(rPr['@_sz'], 10) / 100
              if (rPr['@_b'] === '1' || rPr['@_b'] === true) existingStyle.bold = true
              if (rPr['@_i'] === '1' || rPr['@_i'] === true) existingStyle.italic = true
              if (rPr['@_u']) existingStyle.underline = rPr['@_u']
              if (rPr['a:solidFill']?.['a:srgbClr']?.['@_val']) {
                existingStyle.color = '#' + rPr['a:solidFill']['a:srgbClr']['@_val']
              }
              if (rPr['a:latin']?.['@_typeface']) {
                existingStyle.fontFamily = rPr['a:latin']['@_typeface']
              }
              break
            }
          }
          if (Object.keys(existingStyle).length > 0) break
        }
      }

      const mergedTextStyle = {
        ...existingStyle,
        ...(normalized.textStyle || {}),
      }
      if (normalized.fontSize !== undefined) mergedTextStyle.fontSize = normalized.fontSize
      if (normalized.fontFamily !== undefined) mergedTextStyle.fontFamily = normalized.fontFamily
      if (normalized.color !== undefined) mergedTextStyle.color = normalized.color
      if (normalized.textColor !== undefined) mergedTextStyle.color = normalized.textColor
      if (normalized.bold !== undefined) mergedTextStyle.bold = normalized.bold
      if (normalized.italic !== undefined) mergedTextStyle.italic = normalized.italic
      if (normalized.underline !== undefined) mergedTextStyle.underline = normalized.underline
      if (normalized.align !== undefined || normalized.textAlign !== undefined) {
        mergedTextStyle.align = normalized.align || normalized.textAlign
      }

      const txBodyXml = buildTextBodyXml(textVal, mergedTextStyle, {
        ...normalized,
        ...mergedTextStyle,
      })
      const parsedTxBody = this.#xmlParser.parse(txBodyXml, 'txBody.xml')['p:txBody']
      shape['p:txBody'] = Array.isArray(parsedTxBody) ? parsedTxBody[0] : parsedTxBody
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Updated shape "${shapeId}" on slide ${slideIndex}`)
  }

  /**
   * Removes a shape from the slide.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string} shapeId - Shape ID or name.
   * @param {SlideManager} slideManager - SlideManager instance.
   */
  removeShape(slideIndex, shapeId, slideManager) {
    const res = this.#findShape(slideIndex, shapeId, slideManager)
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
    if (parent['p:cxnSp']) {
      if (Array.isArray(parent['p:cxnSp'])) {
        parent['p:cxnSp'] = parent['p:cxnSp'].filter(s => s !== res.shape)
      } else if (parent['p:cxnSp'] === res.shape) {
        delete parent['p:cxnSp']
      }
    }
    if (parent['p:grpSp']) {
      if (Array.isArray(parent['p:grpSp'])) {
        parent['p:grpSp'] = parent['p:grpSp'].filter(s => s !== res.shape)
      } else if (parent['p:grpSp'] === res.shape) {
        delete parent['p:grpSp']
      }
    }

    const zOrder = parent[Z_ORDER_SYMBOL]
    const numericId =
      res.shape?.['p:nvSpPr']?.['p:cNvPr']?.['@_id'] ||
      res.shape?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_id'] ||
      res.shape?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_id']
    if (zOrder && numericId !== undefined) {
      const elementId = String(numericId)
      parent[Z_ORDER_SYMBOL] = zOrder.filter(id => String(id) !== elementId)
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Removed shape "${shapeId}" from slide ${slideIndex}`)
  }

  /**
   * Clears all user-created shapes from a slide while preserving placeholders and layout elements.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {SlideManager} slideManager - SlideManager instance.
   */
  clearShapes(slideIndex, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    if (!spTree) return

    const removedIds = new Set()

    if (spTree['p:sp']) {
      let shapes = spTree['p:sp']
      if (!Array.isArray(shapes)) shapes = [shapes]

      const preserved = []
      for (const shape of shapes) {
        const isPlaceholder = Boolean(shape?.['p:nvSpPr']?.['p:nvPr']?.['p:ph'])
        if (isPlaceholder) {
          preserved.push(shape)
        } else {
          const id = shape?.['p:nvSpPr']?.['p:cNvPr']?.['@_id']
          if (id !== undefined && id !== null) removedIds.add(String(id))
        }
      }
      spTree['p:sp'] = preserved
    }

    if (spTree['p:cxnSp']) {
      let cxns = spTree['p:cxnSp']
      if (!Array.isArray(cxns)) cxns = [cxns]
      for (const cxn of cxns) {
        const id = cxn?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_id']
        if (id !== undefined && id !== null) removedIds.add(String(id))
      }
      spTree['p:cxnSp'] = []
    }

    if (spTree['p:grpSp']) {
      let groups = spTree['p:grpSp']
      if (!Array.isArray(groups)) groups = [groups]
      for (const group of groups) {
        const id = group?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_id']
        if (id !== undefined && id !== null) removedIds.add(String(id))
      }
      spTree['p:grpSp'] = []
    }

    const zOrder = spTree[Z_ORDER_SYMBOL]
    if (zOrder && Array.isArray(zOrder)) {
      spTree[Z_ORDER_SYMBOL] = zOrder.filter(id => !removedIds.has(String(id)))
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Cleared user shapes from slide ${slideIndex}`)
  }

  /**
   * Retrieves all shape descriptors on a slide.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {SlideManager} slideManager - SlideManager instance.
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
   * Retrieves a specific shape descriptor by ID or name.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string} shapeId - Shape ID or name.
   * @param {SlideManager} slideManager - SlideManager instance.
   * @returns {Object|null}
   */
  getShape(slideIndex, shapeId, slideManager) {
    const res = this.#findShape(slideIndex, shapeId, slideManager)
    if (!res) return null
    return this.#normalizeShapeDescriptor(res.shape)
  }

  /**
   * Clones an existing shape with a new ID and optional overrides.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string} shapeId - Source shape ID or name.
   * @param {string} clonedShapeId - Identifier for cloned shape.
   * @param {Object} [clonedOptions={}] - Positioning and styling overrides.
   * @param {SlideManager} slideManager - SlideManager instance.
   */
  cloneShape(slideIndex, shapeId, clonedShapeId, clonedOptions = {}, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    const res = this.#findShape(slideIndex, shapeId, slideManager)

    if (!res) {
      throw new PPTXError(`Shape "${shapeId}" not found in slide ${slideIndex}`)
    }

    const newShape = this.#xmlParser.deepClone(res.shape)
    const cNvPr =
      newShape['p:nvSpPr']?.['p:cNvPr'] ||
      newShape['p:nvCxnSpPr']?.['p:cNvPr'] ||
      newShape['p:nvGrpSpPr']?.['p:cNvPr']

    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    const existingIds = this.#getAllShapeIds(spTree)
    const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 1000
    const newId = maxId + 1

    if (cNvPr) {
      cNvPr['@_id'] = String(newId)
      cNvPr['@_name'] = clonedShapeId
    }

    const xfrm = newShape['p:spPr']?.['a:xfrm']
    if (xfrm) {
      const offsetX = clonedOptions.offsetX !== undefined ? clonedOptions.offsetX : clonedOptions.dx
      const offsetY = clonedOptions.offsetY !== undefined ? clonedOptions.offsetY : clonedOptions.dy

      if (offsetX !== undefined) {
        const dx = Math.abs(offsetX) < 100 ? Math.round(offsetX * 914400) : Math.round(offsetX)
        const x = parseInt(xfrm['a:off']?.['@_x'] || 0, 10) + dx
        if (!xfrm['a:off']) xfrm['a:off'] = {}
        xfrm['a:off']['@_x'] = String(x)
      } else if (clonedOptions.x !== undefined) {
        const dims = slideManager.getSlideSize()
        if (!xfrm['a:off']) xfrm['a:off'] = {}
        xfrm['a:off']['@_x'] = String(resolveCoordinate(clonedOptions.x, dims.width, clonedOptions))
      }

      if (offsetY !== undefined) {
        const dy = Math.abs(offsetY) < 100 ? Math.round(offsetY * 914400) : Math.round(offsetY)
        const y = parseInt(xfrm['a:off']?.['@_y'] || 0, 10) + dy
        if (!xfrm['a:off']) xfrm['a:off'] = {}
        xfrm['a:off']['@_y'] = String(y)
      } else if (clonedOptions.y !== undefined) {
        const dims = slideManager.getSlideSize()
        if (!xfrm['a:off']) xfrm['a:off'] = {}
        xfrm['a:off']['@_y'] = String(
          resolveCoordinate(clonedOptions.y, dims.height, clonedOptions)
        )
      }

      if (clonedOptions.width !== undefined) {
        const dims = slideManager.getSlideSize()
        if (!xfrm['a:ext']) xfrm['a:ext'] = {}
        xfrm['a:ext']['@_cx'] = String(
          resolveCoordinate(clonedOptions.width, dims.width, clonedOptions)
        )
      }

      if (clonedOptions.height !== undefined) {
        const dims = slideManager.getSlideSize()
        if (!xfrm['a:ext']) xfrm['a:ext'] = {}
        xfrm['a:ext']['@_cy'] = String(
          resolveCoordinate(clonedOptions.height, dims.height, clonedOptions)
        )
      }
    }

    if (clonedOptions.text !== undefined) {
      const txBodyXml = buildTextBodyXml(
        clonedOptions.text,
        clonedOptions.textStyle || {},
        clonedOptions
      )
      const parsedTxBody = this.#xmlParser.parse(txBodyXml, 'txBody.xml')['p:txBody']
      newShape['p:txBody'] = Array.isArray(parsedTxBody) ? parsedTxBody[0] : parsedTxBody
    }

    const parent = res.parent
    const targetKey = res.type === 'cxnSp' || newShape['p:nvCxnSpPr'] ? 'p:cxnSp' : 'p:sp'
    if (!parent[targetKey]) parent[targetKey] = []
    if (!Array.isArray(parent[targetKey])) {
      parent[targetKey] = [parent[targetKey]]
    }
    parent[targetKey].push(newShape)

    if (spTree[Z_ORDER_SYMBOL] && !spTree[Z_ORDER_SYMBOL].includes(String(newId))) {
      spTree[Z_ORDER_SYMBOL].push(String(newId))
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Cloned shape "${shapeId}" as "${clonedShapeId}"`)
    return String(newId)
  }

  /**
   * Inspects a group shape and returns normalized child shape descriptors.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string} groupId - Group shape ID or name.
   * @param {SlideManager} slideManager - SlideManager instance.
   * @returns {Array<Object>}
   */
  getGroupShapes(slideIndex, groupId, slideManager) {
    const res = this.#findShape(slideIndex, groupId, slideManager)
    if (!res || res.type !== 'grpSp') return []

    const groupObj = res.shape
    const childShapes = []
    if (groupObj['p:sp']) {
      const sps = Array.isArray(groupObj['p:sp']) ? groupObj['p:sp'] : [groupObj['p:sp']]
      childShapes.push(...sps)
    }
    if (groupObj['p:cxnSp']) {
      const cxns = Array.isArray(groupObj['p:cxnSp']) ? groupObj['p:cxnSp'] : [groupObj['p:cxnSp']]
      childShapes.push(...cxns)
    }

    return childShapes.map(shape => this.#normalizeShapeDescriptor(shape))
  }

  /**
   * Groups multiple shapes into a composite group container (<p:grpSp>).
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string[]} shapeIds - Array of shape IDs or names to group.
   * @param {Object} [groupOptions={}] - Configuration options (e.g. name, id).
   * @param {SlideManager} slideManager - SlideManager instance.
   * @returns {{ id: string, name: string }}
   */
  groupShapes(slideIndex, shapeIds, groupOptions = {}, slideManager) {
    if (!Array.isArray(shapeIds) || shapeIds.length < 2) {
      throw new PPTXError('At least 2 shape IDs are required to create a group.')
    }

    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    if (!spTree) {
      throw new PPTXError(`Invalid slide structure for slide ${slideIndex}`)
    }

    const targetShapes = []
    for (const shapeId of shapeIds) {
      const res = this.#findShape(slideIndex, shapeId, slideManager)
      if (!res) {
        throw new PPTXError(`Shape "${shapeId}" not found in slide ${slideIndex}`)
      }
      targetShapes.push(res)
    }

    let minX = Infinity
    let minY = Infinity
    let maxX = -Infinity
    let maxY = -Infinity

    for (const { shape } of targetShapes) {
      const xfrm = shape?.['p:spPr']?.['a:xfrm']
      const x = parseInt(xfrm?.['a:off']?.['@_x'] || 0, 10)
      const y = parseInt(xfrm?.['a:off']?.['@_y'] || 0, 10)
      const cx = parseInt(xfrm?.['a:ext']?.['@_cx'] || 0, 10)
      const cy = parseInt(xfrm?.['a:ext']?.['@_cy'] || 0, 10)

      if (x < minX) minX = x
      if (y < minY) minY = y
      if (x + cx > maxX) maxX = x + cx
      if (y + cy > maxY) maxY = y + cy
    }

    const groupW = Math.max(1, maxX - minX)
    const groupH = Math.max(1, maxY - minY)

    const existingIds = this.#getAllShapeIds(spTree)
    const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 1000
    const newId = maxId + 1

    const groupName = groupOptions.name || groupOptions.id || `Group ${newId}`

    const groupXml = `<p:grpSp>
      <p:nvGrpSpPr>
        <p:cNvPr id="${newId}" name="${escapeXml(groupName)}"/>
        <p:cNvGrpSpPr/>
        <p:nvPr/>
      </p:nvGrpSpPr>
      <p:grpSpPr>
        <a:xfrm>
          <a:off x="${minX}" y="${minY}"/>
          <a:ext cx="${groupW}" cy="${groupH}"/>
          <a:chOff x="${minX}" y="${minY}"/>
          <a:chExt cx="${groupW}" cy="${groupH}"/>
        </a:xfrm>
      </p:grpSpPr>
    </p:grpSp>`

    const parsedGroup = this.#xmlParser.parse(groupXml, 'group.xml')['p:grpSp']
    const groupObj = Array.isArray(parsedGroup) ? parsedGroup[0] : parsedGroup

    groupObj['p:sp'] = []
    const removedNumericIds = new Set(
      targetShapes.map(t =>
        String(
          t.shape?.['p:nvSpPr']?.['p:cNvPr']?.['@_id'] ||
            t.shape?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_id'] ||
            t.shape?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_id']
        )
      )
    )

    for (const { shape, parent, type } of targetShapes) {
      if (parent['p:sp']) {
        if (Array.isArray(parent['p:sp'])) {
          parent['p:sp'] = parent['p:sp'].filter(s => s !== shape)
        } else if (parent['p:sp'] === shape) {
          delete parent['p:sp']
        }
      }
      if (parent['p:cxnSp']) {
        if (Array.isArray(parent['p:cxnSp'])) {
          parent['p:cxnSp'] = parent['p:cxnSp'].filter(s => s !== shape)
        } else if (parent['p:cxnSp'] === shape) {
          delete parent['p:cxnSp']
        }
      }
      if (type === 'cxnSp' || shape['p:nvCxnSpPr']) {
        if (!groupObj['p:cxnSp']) groupObj['p:cxnSp'] = []
        if (!Array.isArray(groupObj['p:cxnSp'])) groupObj['p:cxnSp'] = [groupObj['p:cxnSp']]
        groupObj['p:cxnSp'].push(shape)
      } else {
        groupObj['p:sp'].push(shape)
      }
    }

    if (!spTree['p:grpSp']) {
      spTree['p:grpSp'] = []
    }
    if (!Array.isArray(spTree['p:grpSp'])) {
      spTree['p:grpSp'] = [spTree['p:grpSp']]
    }
    spTree['p:grpSp'].push(groupObj)

    const zOrder = spTree[Z_ORDER_SYMBOL]
    if (zOrder && Array.isArray(zOrder)) {
      let replaced = false
      const newZOrder = []
      for (const id of zOrder) {
        if (removedNumericIds.has(String(id))) {
          if (!replaced) {
            newZOrder.push(String(newId))
            replaced = true
          }
        } else {
          newZOrder.push(id)
        }
      }
      if (!replaced) newZOrder.push(String(newId))
      spTree[Z_ORDER_SYMBOL] = newZOrder
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(
      `Grouped ${targetShapes.length} shapes into "${groupName}" (ID ${newId}) on slide ${slideIndex}`
    )
    return { id: String(newId), name: groupName }
  }

  /**
   * Ungroups a composite group container, reprojecting child transforms back to absolute slide coordinates.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string} groupId - Group shape ID or name.
   * @param {SlideManager} slideManager - SlideManager instance.
   */
  ungroupShapes(slideIndex, groupId, slideManager) {
    const res = this.#findShape(slideIndex, groupId, slideManager)
    if (!res || res.type !== 'grpSp') {
      throw new PPTXError(`Group shape "${groupId}" not found in slide ${slideIndex}`)
    }

    const groupObj = res.shape
    const parentContainer = res.parent
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    const grpXfrm = groupObj['p:grpSpPr']?.['a:xfrm']
    const gX = parseInt(grpXfrm?.['a:off']?.['@_x'] || 0, 10)
    const gY = parseInt(grpXfrm?.['a:off']?.['@_y'] || 0, 10)
    const gW = parseInt(grpXfrm?.['a:ext']?.['@_cx'] || 0, 10)
    const gH = parseInt(grpXfrm?.['a:ext']?.['@_cy'] || 0, 10)
    const chX = parseInt(
      grpXfrm?.['a:chOff']?.['@_x'] !== undefined ? grpXfrm['a:chOff']['@_x'] : gX,
      10
    )
    const chY = parseInt(
      grpXfrm?.['a:chOff']?.['@_y'] !== undefined ? grpXfrm['a:chOff']['@_y'] : gY,
      10
    )
    const chW = parseInt(
      grpXfrm?.['a:chExt']?.['@_cx'] !== undefined ? grpXfrm['a:chExt']['@_cx'] : gW,
      10
    )
    const chH = parseInt(
      grpXfrm?.['a:chExt']?.['@_cy'] !== undefined ? grpXfrm['a:chExt']['@_cy'] : gH,
      10
    )

    const scaleX = chW > 0 ? gW / chW : 1
    const scaleY = chH > 0 ? gH / chH : 1

    const childShapes = []
    if (groupObj['p:sp']) {
      const sps = Array.isArray(groupObj['p:sp']) ? groupObj['p:sp'] : [groupObj['p:sp']]
      childShapes.push(...sps)
    }
    if (groupObj['p:cxnSp']) {
      const cxns = Array.isArray(groupObj['p:cxnSp']) ? groupObj['p:cxnSp'] : [groupObj['p:cxnSp']]
      childShapes.push(...cxns)
    }

    if (!spTree['p:sp']) spTree['p:sp'] = []
    if (!Array.isArray(spTree['p:sp'])) spTree['p:sp'] = [spTree['p:sp']]

    const childIds = []

    for (const child of childShapes) {
      if (!child['p:spPr']) child['p:spPr'] = {}
      if (!child['p:spPr']['a:xfrm']) child['p:spPr']['a:xfrm'] = {}
      const xfrm = child['p:spPr']['a:xfrm']
      if (!xfrm['a:off']) xfrm['a:off'] = {}
      if (!xfrm['a:ext']) xfrm['a:ext'] = {}

      const cX = parseInt(xfrm['a:off']['@_x'] || 0, 10)
      const cY = parseInt(xfrm['a:off']['@_y'] || 0, 10)
      const cW = parseInt(xfrm['a:ext']['@_cx'] || 0, 10)
      const cH = parseInt(xfrm['a:ext']['@_cy'] || 0, 10)

      const absX = Math.round(gX + (cX - chX) * scaleX)
      const absY = Math.round(gY + (cY - chY) * scaleY)
      const absW = Math.round(cW * scaleX)
      const absH = Math.round(cH * scaleY)

      xfrm['a:off']['@_x'] = String(absX)
      xfrm['a:off']['@_y'] = String(absY)
      xfrm['a:ext']['@_cx'] = String(absW)
      xfrm['a:ext']['@_cy'] = String(absH)

      const cid =
        child?.['p:nvSpPr']?.['p:cNvPr']?.['@_id'] || child?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_id']
      if (cid !== undefined && cid !== null) childIds.push(String(cid))

      if (child['p:nvCxnSpPr']) {
        if (!spTree['p:cxnSp']) spTree['p:cxnSp'] = []
        if (!Array.isArray(spTree['p:cxnSp'])) spTree['p:cxnSp'] = [spTree['p:cxnSp']]
        spTree['p:cxnSp'].push(child)
      } else {
        spTree['p:sp'].push(child)
      }
    }

    if (parentContainer['p:grpSp']) {
      if (Array.isArray(parentContainer['p:grpSp'])) {
        parentContainer['p:grpSp'] = parentContainer['p:grpSp'].filter(g => g !== groupObj)
      } else if (parentContainer['p:grpSp'] === groupObj) {
        delete parentContainer['p:grpSp']
      }
    }

    const grpId = String(groupObj?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_id'])
    const zOrder = spTree[Z_ORDER_SYMBOL]
    if (zOrder && Array.isArray(zOrder)) {
      const newZOrder = []
      for (const id of zOrder) {
        if (String(id) === grpId) {
          newZOrder.push(...childIds)
        } else {
          newZOrder.push(id)
        }
      }
      spTree[Z_ORDER_SYMBOL] = newZOrder
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Ungrouped "${groupId}" into ${childShapes.length} shapes on slide ${slideIndex}`)
  }

  // --- Private Helpers ---

  #findShape(slideIndex, shapeId, slideManager) {
    const res = slideManager.getSlideShape(slideIndex, shapeId)
    if (res) return res

    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    if (!spTree) return null

    const targetIdStr = String(shapeId).trim()

    const searchContainer = container => {
      if (!container) return null

      let shapes = container['p:sp'] || []
      if (!Array.isArray(shapes)) shapes = [shapes]
      for (const s of shapes) {
        const cNvPr = s?.['p:nvSpPr']?.['p:cNvPr']
        if (
          cNvPr &&
          (String(cNvPr['@_id']) === targetIdStr || String(cNvPr['@_name']) === targetIdStr)
        ) {
          return { shape: s, parent: container, type: 'sp' }
        }
      }

      let cxns = container['p:cxnSp'] || []
      if (!Array.isArray(cxns)) cxns = [cxns]
      for (const c of cxns) {
        const cNvPr = c?.['p:nvCxnSpPr']?.['p:cNvPr']
        if (
          cNvPr &&
          (String(cNvPr['@_id']) === targetIdStr || String(cNvPr['@_name']) === targetIdStr)
        ) {
          return { shape: c, parent: container, type: 'cxnSp' }
        }
      }

      let groups = container['p:grpSp'] || []
      if (!Array.isArray(groups)) groups = [groups]
      for (const g of groups) {
        const cNvPr = g?.['p:nvGrpSpPr']?.['p:cNvPr']
        if (
          cNvPr &&
          (String(cNvPr['@_id']) === targetIdStr || String(cNvPr['@_name']) === targetIdStr)
        ) {
          return { shape: g, parent: container, type: 'grpSp' }
        }
        const inner = searchContainer(g)
        if (inner) return inner
      }

      return null
    }

    return searchContainer(spTree)
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

    let cxns = container['p:cxnSp'] || []
    if (!Array.isArray(cxns)) cxns = [cxns]
    for (const c of cxns) {
      const id = parseInt(c?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_id'], 10)
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
      results.push(this.#normalizeShapeDescriptor(shape))
    }

    let cxnShapes = container['p:cxnSp'] || []
    if (!Array.isArray(cxnShapes)) cxnShapes = [cxnShapes]

    for (const cxn of cxnShapes) {
      results.push(this.#normalizeShapeDescriptor(cxn))
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const g of groups) {
      results.push(this.#normalizeShapeDescriptor(g))
      this.#collectShapesInfo(g, results)
    }
  }

  #normalizeShapeDescriptor(shape) {
    const isGroup = Boolean(shape?.['p:nvGrpSpPr'] || shape?.['p:grpSpPr'])
    const isConnector = Boolean(shape?.['p:nvCxnSpPr'])
    const isPlaceholder = Boolean(shape?.['p:nvSpPr']?.['p:nvPr']?.['p:ph'])
    const cNvPr =
      shape?.['p:nvSpPr']?.['p:cNvPr'] ||
      shape?.['p:nvCxnSpPr']?.['p:cNvPr'] ||
      shape?.['p:nvGrpSpPr']?.['p:cNvPr']
    const xfrm = shape?.['p:spPr']?.['a:xfrm'] || shape?.['p:grpSpPr']?.['a:xfrm']
    const prstGeom = shape?.['p:spPr']?.['a:prstGeom']

    const rawId = cNvPr?.['@_id'] !== undefined ? String(cNvPr['@_id']) : ''
    const rawName = cNvPr?.['@_name'] || ''
    const id = rawName || rawId
    const name = rawName || id
    const preset = isGroup ? 'group' : prstGeom ? prstGeom['@_prst'] : isConnector ? 'line' : 'rect'

    const xEmu = xfrm?.['a:off']?.['@_x'] !== undefined ? parseInt(xfrm['a:off']['@_x'], 10) : 0
    const yEmu = xfrm?.['a:off']?.['@_y'] !== undefined ? parseInt(xfrm['a:off']['@_y'], 10) : 0
    const wEmu = xfrm?.['a:ext']?.['@_cx'] !== undefined ? parseInt(xfrm['a:ext']['@_cx'], 10) : 0
    const hEmu = xfrm?.['a:ext']?.['@_cy'] !== undefined ? parseInt(xfrm['a:ext']['@_cy'], 10) : 0

    const x = Math.round(xEmu / 9525)
    const y = Math.round(yEmu / 9525)
    const width = Math.round(wEmu / 9525)
    const height = Math.round(hEmu / 9525)

    const type = isGroup ? 'group' : mapPresetToType(preset, width, height)

    let fill = undefined
    const solidFill = shape?.['p:spPr']?.['a:solidFill']
    if (solidFill && solidFill['a:srgbClr']) {
      fill = '#' + (solidFill['a:srgbClr']['@_val'] || '').toUpperCase()
    }

    let transparency = undefined
    if (solidFill && solidFill['a:srgbClr']?.['a:alpha']?.['@_val']) {
      transparency = Math.round(
        (100000 - parseInt(solidFill['a:srgbClr']['a:alpha']['@_val'], 10)) / 1000
      )
    }

    let border = undefined
    const ln = shape?.['p:spPr']?.['a:ln']
    if (ln) {
      const bColor = ln['a:solidFill']?.['a:srgbClr']?.['@_val']
      const bWidth = ln['@_w'] ? parseInt(ln['@_w'], 10) / 9525 : 1
      border = {
        color: bColor ? '#' + bColor.toUpperCase() : '#000000',
        width: bWidth,
      }
      if (ln['a:headEnd']?.['@_type']) {
        border.headEnd = ln['a:headEnd']['@_type']
      }
      if (ln['a:tailEnd']?.['@_type']) {
        border.tailEnd = ln['a:tailEnd']['@_type']
      }
    }

    const text = this.#getShapeText(shape) || undefined

    const rotEmu = xfrm?.['@_rot'] ? parseInt(xfrm['@_rot'], 10) : 0
    const rotation = rotEmu ? Math.round(rotEmu / 60000) : 0
    const flipH = xfrm?.['@_flipH'] === '1' || xfrm?.['@_flipH'] === true
    const flipV = xfrm?.['@_flipV'] === '1' || xfrm?.['@_flipV'] === true

    const position = xfrm
      ? {
          x: xEmu,
          y: yEmu,
          cx: wEmu,
          cy: hEmu,
        }
      : null

    return {
      id,
      name,
      type,
      geometry: preset,
      isGroup,
      isConnector,
      isPlaceholder,
      x,
      y,
      width,
      height,
      bounds: xfrm ? { x: xEmu, y: yEmu, width: wEmu, height: hEmu } : null,
      position,
      rotation,
      flipH,
      flipV,
      fill,
      transparency,
      border,
      stroke: border,
      text,
    }
  }

  #getShapeText(shape) {
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

  #reorderSpPrKeys(spPr) {
    if (!spPr) return
    const order = [
      '@_bwMode',
      'a:xfrm',
      'a:custGeom',
      'a:prstGeom',
      'a:noFill',
      'a:solidFill',
      'a:gradFill',
      'a:blipFill',
      'a:pattFill',
      'a:grpFill',
      'a:ln',
      'a:effectLst',
      'a:effectDag',
      'a:scene3d',
      'a:sp3d',
      'a:extLst',
    ]

    const newSpPr = {}
    for (const key of order) {
      if (spPr[key] !== undefined) {
        newSpPr[key] = spPr[key]
      }
    }
    for (const key of Object.keys(spPr)) {
      if (!order.includes(key)) {
        newSpPr[key] = spPr[key]
      }
    }

    for (const key of Object.keys(spPr)) {
      delete spPr[key]
    }
    for (const key of Object.keys(newSpPr)) {
      spPr[key] = newSpPr[key]
    }
  }

  #normalizeShapeOptions(options, isUpdate = false) {
    if (!options) return options
    const normalized = { ...options }

    if (normalized.type === 'rect') {
      normalized.type = 'rectangle'
    }

    if (normalized.id === undefined && normalized.name !== undefined) {
      normalized.id = normalized.name
    }
    if (normalized.name === undefined && normalized.id !== undefined) {
      normalized.name = normalized.id
    }

    let fillVal =
      options.fill !== undefined
        ? options.fill
        : options.fillColor !== undefined
          ? options.fillColor
          : options.backgroundColor !== undefined
            ? options.backgroundColor
            : options.bg

    if (
      !isUpdate &&
      fillVal === undefined &&
      options.color !== undefined &&
      options.text === undefined &&
      !options.textColor
    ) {
      fillVal = options.color
    }

    if (fillVal !== undefined) {
      normalized.fill = fillVal
    }

    if (options.opacity !== undefined) {
      normalized.transparency = Math.round((1 - parseFloat(options.opacity)) * 100)
    }

    const strokeVal =
      options.stroke !== undefined
        ? options.stroke
        : options.border !== undefined
          ? options.border
          : options.strokeColor !== undefined
            ? options.strokeColor
            : options.borderColor
    if (strokeVal !== undefined && strokeVal !== null) {
      const strokeWidth =
        options.strokeWidth !== undefined
          ? options.strokeWidth
          : options.borderWidth !== undefined
            ? options.borderWidth
            : 1
      if (typeof strokeVal === 'string') {
        normalized.border = {
          color: strokeVal,
          width: parseFloat(strokeWidth),
        }
      } else if (typeof strokeVal === 'object') {
        normalized.border = {
          color: strokeVal.color || strokeVal.fill || '#000000',
          width: parseFloat(strokeVal.width !== undefined ? strokeVal.width : strokeWidth),
        }
      }
    } else if (options.strokeWidth !== undefined || options.borderWidth !== undefined) {
      const strokeWidth =
        options.strokeWidth !== undefined ? options.strokeWidth : options.borderWidth
      normalized.border = {
        color: '#000000',
        width: parseFloat(strokeWidth),
      }
    }

    const cleanHexColor = colorStr => {
      if (typeof colorStr !== 'string') return colorStr
      let clean = colorStr.replace('#', '').trim()
      if (clean.length === 3) {
        clean = clean[0] + clean[0] + clean[1] + clean[1] + clean[2] + clean[2]
      } else if (clean.length === 8) {
        const rgbPart = clean.slice(0, 6)
        const alphaPart = clean.slice(6, 8)
        if (options.opacity === undefined && options.transparency === undefined) {
          const alphaVal = parseInt(alphaPart, 16)
          normalized.transparency = Math.round((1 - alphaVal / 255) * 100)
        }
        clean = rgbPart
      }
      return '#' + clean.toUpperCase()
    }

    if (
      typeof normalized.fill === 'string' &&
      normalized.fill !== 'none' &&
      normalized.fill !== 'transparent'
    ) {
      normalized.fill = cleanHexColor(normalized.fill)
    } else if (typeof normalized.fill === 'object' && normalized.fill?.type === 'gradient') {
      if (Array.isArray(normalized.fill.colors)) {
        normalized.fill.colors = normalized.fill.colors.map(c => cleanHexColor(c))
      }
    }

    if (normalized.border && typeof normalized.border.color === 'string') {
      normalized.border.color = cleanHexColor(normalized.border.color)
    }

    return normalized
  }

  #validateShapeOptions(options) {
    const errors = []
    if (!options) return errors

    const supportedTypes = [
      'rectangle',
      'rect',
      'square',
      'circle',
      'ellipse',
      'roundedRectangle',
      'roundRect',
      'triangle',
      'star5',
      'arrow',
      'upArrow',
      'downArrow',
      'leftArrow',
      'rightArrow',
      'leftRightArrow',
      'upDownArrow',
      'quadArrow',
      'leftUpArrow',
      'bentArrow',
      'curvedRightArrow',
      'curvedLeftArrow',
      'curvedUpArrow',
      'curvedDownArrow',
      'stripedRightArrow',
      'notchedRightArrow',
      'homePlate',
      'chevron',
      'diamond',
      'hexagon',
      'line',
      'lineArrow',
      'connector',
    ]

    if (options.type && !supportedTypes.includes(options.type)) {
      // allow custom openxml presets or report
    }

    return errors
  }
}

// --- Module-scoped Utility Functions ---

function normalizePresetGeometry(type) {
  if (!type || typeof type !== 'string') return 'rect'
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
    case 'leftuparrow':
      return 'leftUpArrow'
    case 'bentarrow':
      return 'bentArrow'
    case 'curvedrightarrow':
      return 'curvedRightArrow'
    case 'curvedleftarrow':
      return 'curvedLeftArrow'
    case 'curveduparrow':
      return 'curvedUpArrow'
    case 'curveddownarrow':
      return 'curvedDownArrow'
    case 'stripedrightarrow':
      return 'stripedRightArrow'
    case 'notchedrightarrow':
      return 'notchedRightArrow'
    case 'homeplate':
      return 'homePlate'
    case 'chevron':
      return 'chevron'
    case 'rect':
    case 'rectangle':
    case 'square':
      return 'rect'
    case 'ellipse':
    case 'circle':
      return 'ellipse'
    case 'roundrect':
    case 'roundedrectangle':
      return 'roundRect'
    case 'line':
    case 'connector':
    case 'linearrow':
    case 'arrowline':
      return 'line'
    case 'triangle':
      return 'triangle'
    case 'diamond':
      return 'diamond'
    case 'hexagon':
      return 'hexagon'
    case 'star5':
    case 'star':
      return 'star5'
    default:
      return clean
  }
}

function escapeXml(str) {
  if (typeof str !== 'string') return String(str ?? '')
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&apos;')
}

function mapPresetToType(preset, w, h) {
  if (preset === 'roundRect') return 'roundedRectangle'
  if (preset === 'rect') {
    if (w === h && w > 0) return 'square'
    return 'rectangle'
  }
  if (preset === 'ellipse') {
    if (w === h && w > 0) return 'circle'
    return 'ellipse'
  }
  if (preset === 'rightArrow') return 'rightArrow'
  if (preset === 'leftArrow') return 'leftArrow'
  if (preset === 'upArrow') return 'upArrow'
  if (preset === 'downArrow') return 'downArrow'
  return preset
}

function buildTextBodyXml(text, textStyle = {}, options = {}) {
  const { toEmu } = require('../spatial/units.js')
  const fontSize = textStyle.fontSize || options.fontSize || 14
  const fontSizeVal = Math.round(fontSize * 100)
  const boldAttr = textStyle.bold || options.bold ? ' b="1"' : ''
  const italicAttr = textStyle.italic || options.italic ? ' i="1"' : ''

  const underline = textStyle.underline !== undefined ? textStyle.underline : options.underline
  let uAttr = ''
  if (underline) {
    uAttr = ` u="${typeof underline === 'string' ? underline : 'sng'}"`
  }

  const strike = textStyle.strike !== undefined ? textStyle.strike : options.strike
  let strikeAttr = ''
  if (strike) {
    strikeAttr = ` strike="${strike === 'double' || strike === 'dblStrike' ? 'dblStrike' : 'sngStrike'}"`
  }

  let baselineAttr = ''
  if (textStyle.superscript || options.superscript) {
    baselineAttr = ' baseline="30000"'
  } else if (textStyle.subscript || options.subscript) {
    baselineAttr = ' baseline="-25000"'
  } else if (textStyle.baseline !== undefined) {
    baselineAttr = ` baseline="${Math.round(textStyle.baseline)}"`
  } else if (options.baseline !== undefined) {
    baselineAttr = ` baseline="${Math.round(options.baseline)}"`
  }

  const charSpacing =
    textStyle.characterSpacing !== undefined
      ? textStyle.characterSpacing
      : textStyle.tracking !== undefined
        ? textStyle.tracking
        : options.characterSpacing !== undefined
          ? options.characterSpacing
          : options.tracking
  const spcAttr = charSpacing !== undefined ? ` spc="${Math.round(charSpacing * 100)}"` : ''

  const fontFamily = textStyle.fontFamily || options.fontFamily || 'Arial'
  const ff = escapeXml(fontFamily)
  const fontXml = `<a:latin typeface="${ff}"/><a:ea typeface="${ff}"/><a:cs typeface="${ff}"/>`

  const fontColor = textStyle.color || options.textColor || options.color || '#000000'
  const cleanHex = fontColor.replace('#', '').toUpperCase()
  const solidFillXml = `<a:solidFill><a:srgbClr val="${cleanHex}"/></a:solidFill>`

  // Paragraph properties (alignment, line spacing, spacing before/after)
  let pPrChildren = ''
  const alignMap = {
    center: 'ctr',
    ctr: 'ctr',
    right: 'r',
    r: 'r',
    left: 'l',
    l: 'l',
    justify: 'just',
    just: 'just',
  }
  const align = textStyle.align || textStyle.alignment || options.textAlign || options.align || 'l'
  const algn = alignMap[align.toLowerCase()] || 'l'
  const pPrAttrs = ` algn="${algn}"`

  const lineSpacing =
    textStyle.lineSpacing !== undefined ? textStyle.lineSpacing : options.lineSpacing
  const lineSpacingPts =
    textStyle.lineSpacingPts !== undefined ? textStyle.lineSpacingPts : options.lineSpacingPts
  if (lineSpacing !== undefined) {
    const pctVal =
      typeof lineSpacing === 'number' && lineSpacing < 10
        ? Math.round(lineSpacing * 100000)
        : Math.round(lineSpacing * 1000)
    pPrChildren += `<a:lnSpc><a:spcPct val="${pctVal}"/></a:lnSpc>`
  } else if (lineSpacingPts !== undefined) {
    pPrChildren += `<a:lnSpc><a:spcPts val="${Math.round(lineSpacingPts * 100)}"/></a:lnSpc>`
  }

  const spaceBefore =
    textStyle.spaceBefore !== undefined ? textStyle.spaceBefore : options.spaceBefore
  if (spaceBefore !== undefined) {
    pPrChildren += `<a:spcBfr><a:spcPts val="${Math.round(spaceBefore * 100)}"/></a:spcBfr>`
  }
  const spaceAfter = textStyle.spaceAfter !== undefined ? textStyle.spaceAfter : options.spaceAfter
  if (spaceAfter !== undefined) {
    pPrChildren += `<a:spcAft><a:spcPts val="${Math.round(spaceAfter * 100)}"/></a:spcAft>`
  }

  const pPrXml = `<a:pPr${pPrAttrs}>${pPrChildren}</a:pPr>`

  const lines = String(text ?? '').split(/\r?\n/)
  const paragraphsXml = lines
    .map(line => {
      if (!line) {
        return `<a:p>${pPrXml}<a:endParaRPr lang="en-US" sz="${fontSizeVal}"${boldAttr}${italicAttr}${uAttr}${strikeAttr}${baselineAttr}${spcAttr}>${solidFillXml}${fontXml}</a:endParaRPr></a:p>`
      }
      return `<a:p>${pPrXml}<a:r><a:rPr lang="en-US" sz="${fontSizeVal}"${boldAttr}${italicAttr}${uAttr}${strikeAttr}${baselineAttr}${spcAttr}>${solidFillXml}${fontXml}</a:rPr><a:t>${escapeXml(line)}</a:t></a:r></a:p>`
    })
    .join('')

  // BodyPr properties (vertical alignment, padding, wrap)
  const anchorMap = {
    top: 't',
    t: 't',
    middle: 'ctr',
    center: 'ctr',
    ctr: 'ctr',
    bottom: 'b',
    b: 'b',
    just: 'just',
  }
  const vAlign = textStyle.verticalAlign || options.verticalAlign || 'ctr'
  const anchorAttr =
    vAlign && anchorMap[String(vAlign).toLowerCase()]
      ? ` anchor="${anchorMap[String(vAlign).toLowerCase()]}"`
      : ' anchor="ctr"'
  const wrapAttr =
    options.wrap === false || options.wordWrap === false ? ' wrap="none"' : ' wrap="square"'

  let padAttrs = ''
  const pad = options.padding || textStyle.padding
  if (pad) {
    if (typeof pad === 'number') {
      const emu = toEmu(pad)
      padAttrs = ` lIns="${emu}" tIns="${emu}" rIns="${emu}" bIns="${emu}"`
    } else if (typeof pad === 'object') {
      if (pad.left !== undefined) padAttrs += ` lIns="${toEmu(pad.left)}"`
      if (pad.top !== undefined) padAttrs += ` tIns="${toEmu(pad.top)}"`
      if (pad.right !== undefined) padAttrs += ` rIns="${toEmu(pad.right)}"`
      if (pad.bottom !== undefined) padAttrs += ` bIns="${toEmu(pad.bottom)}"`
    }
  }

  return `<p:txBody><a:bodyPr${wrapAttr} rtlCol="0"${anchorAttr}${padAttrs}><a:normAutofit/></a:bodyPr><a:lstStyle/>${paragraphsXml}</p:txBody>`
}

module.exports = { ShapeManager }
