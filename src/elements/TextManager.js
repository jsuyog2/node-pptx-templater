/**
 * @fileoverview TextManager - Handles slide text manipulation, search, lists, and DrawingML styling.
 * Strictly adheres to the 18 approved methods under the Text & Placeholders category.
 */

'use strict'

const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')
const { toEmu } = require('../spatial/units.js')
const { StyleBuilder, applyTextCasing } = require('../parser/StyleBuilder.js')
const { ValidationEngine } = require('../core/ValidationEngine.js')

const logger = createLogger('TextManager')

class TextManager {
  /** @private @type {XMLParser} */
  #xmlParser

  /**
   * @param {XMLParser} xmlParser
   */
  constructor(xmlParser) {
    this.#xmlParser = xmlParser
  }

  // ==========================================================================
  // 1. replaceText(slideIndex, replacements, slideManager, templateEngine)
  // ==========================================================================
  /**
   * Replaces placeholders across text runs in a slide XML with AST run healing.
   *
   * @param {number} slideIndex
   * @param {Object.<string, string|number>} replacements
   * @param {SlideManager} slideManager
   * @param {TemplateEngine} templateEngine
   */
  replaceText(slideIndex, replacements, slideManager, templateEngine) {
    if (!replacements || typeof replacements !== 'object') return
    const slideXml = slideManager.getSlideXml(slideIndex)

    // Normalize keys: ensure both 'key' and '{{key}}' are matched
    const normalized = {}
    for (const [key, val] of Object.entries(replacements)) {
      normalized[key] = String(val)
      if (!key.startsWith('{{') || !key.endsWith('}}')) {
        normalized[`{{${key.replace(/^\{\{|\}\}$/g, '')}}}`] = String(val)
      }
    }

    const updatedXml = templateEngine.replaceTextInXml(slideXml, normalized)
    slideManager.setSlideXml(slideIndex, updatedXml)
    logger.debug(`replaceText: Replaced placeholders in slide ${slideIndex}`)
  }

  // ==========================================================================
  // 2. addText(slideIndex, textContent, options, slideManager)
  // ==========================================================================
  /**
   * Adds a new formatted text box shape to a slide or adds text to an existing shape.
   *
   * @param {number} slideIndex
   * @param {string} textContent
   * @param {Object} [options]
   * @param {SlideManager} slideManager
   * @returns {Object} Shape descriptor
   */
  addText(slideIndex, textContent, options = {}, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    if (!spTree) {
      throw new PPTXError(`Cannot locate spTree in slide ${slideIndex}`)
    }

    // Determine target shape if target/shapeId specified
    const targetId = options.target || options.shapeId || options.id
    if (targetId && !options.x && !options.y) {
      const found = this.findShapeRecursive(spTree, targetId)
      if (found && found.shape) {
        this.#setTextBody(found.shape, textContent, options)
        slideManager.markSlideObjDirty(slideIndex)
        return { shapeId: targetId, type: 'text' }
      }
    }

    // Generate unique shape ID and name
    const existingIds = this.#collectShapeIds(spTree)
    const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 1000
    const newId = maxId + 1
    const shapeName = options.name || options.id || `TextBox_${newId}`

    // Coordinate resolution: supports numbers, percentages ('10%'), pixels ('100px'), and EMUs
    // Standard slide dimensions: 12192000 x 6858000
    const SLIDE_WIDTH = 12192000
    const SLIDE_HEIGHT = 6858000
    const x = toEmu(options.x !== undefined ? options.x : 100, SLIDE_WIDTH)
    const y = toEmu(options.y !== undefined ? options.y : 100, SLIDE_HEIGHT)
    const width = toEmu(options.width !== undefined ? options.width : 300, SLIDE_WIDTH)
    const height = toEmu(options.height !== undefined ? options.height : 60, SLIDE_HEIGHT)

    // Build style and casing
    const style = options.style || options
    const styledText = applyTextCasing(textContent, style)

    const rPr = StyleBuilder.buildRunProperties({ '@_lang': 'en-US' }, style)
    const pPr = StyleBuilder.buildParagraphProperties({}, style)
    const bodyPr = StyleBuilder.buildBodyProperties({ '@_wrap': 'square', '@_rtlCol': '0' }, style)

    const shapeNode = {
      'p:nvSpPr': {
        'p:cNvPr': {
          '@_id': String(newId),
          '@_name': shapeName,
        },
        'p:cNvSpPr': {
          '@_txBox': '1',
        },
        'p:nvPr': {},
      },
      'p:spPr': {
        'a:xfrm': {
          'a:off': { '@_x': String(x), '@_y': String(y) },
          'a:ext': { '@_cx': String(width), '@_cy': String(height) },
        },
        'a:prstGeom': {
          '@_prst': 'rect',
          'a:avLst': {},
        },
        'a:noFill': {},
      },
      'p:txBody': {
        'a:bodyPr': bodyPr,
        'a:lstStyle': {},
        'a:p': [
          {
            'a:pPr': pPr,
            'a:r': [
              {
                'a:rPr': rPr,
                'a:t': styledText,
              },
            ],
          },
        ],
      },
    }

    if (!spTree['p:sp']) {
      spTree['p:sp'] = []
    } else if (!Array.isArray(spTree['p:sp'])) {
      spTree['p:sp'] = [spTree['p:sp']]
    }
    spTree['p:sp'].push(shapeNode)

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`addText: Added text box "${shapeName}" (ID ${newId}) on slide ${slideIndex}`)
    return { shapeId: String(newId), shapeName }
  }

  // ==========================================================================
  // 3. replaceTextByTag(slideIndex, tag, replacement, options, slideManager, templateEngine)
  // ==========================================================================
  /**
   * Replaces the entire text content of a text box or shape identified by ID, name, or tag.
   * Clears out all existing text runs across all paragraphs within <p:txBody> while
   * preserving root paragraph properties (<a:pPr>) and base run formatting (<a:rPr>).
   * Supports multi-line replacements (split on \n) and optional style overrides.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string|number} tagOrId - Shape ID, shape name, or semantic tag.
   * @param {string|number} newText - New replacement text.
   * @param {Object} [options={}] - Optional formatting options or overrides.
   * @param {SlideManager} slideManager - Slide manager instance.
   * @param {TemplateEngine} [_templateEngine] - Optional template engine reference for backward compatibility.
   */
  replaceTextByTag(slideIndex, tagOrId, newText, options = {}, slideManager, _templateEngine) {
    if (tagOrId === undefined || tagOrId === null) {
      logger.warn(`replaceTextByTag: Invalid tagOrId provided on slide ${slideIndex}`)
      return
    }

    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    if (!spTree) {
      logger.warn(`replaceTextByTag: Cannot locate spTree on slide ${slideIndex}`)
      return
    }

    const targetKey = String(tagOrId).trim()
    let found = this.findShapeRecursive(spTree, targetKey)
    if (!found) {
      // Try semantic placeholder match
      const normalizedTag =
        targetKey.startsWith('{{') && targetKey.endsWith('}}') ? targetKey : `{{${targetKey}}}`
      found =
        this.#collectMatchingShapeOrCell(spTree, targetKey) ||
        this.#collectMatchingShapeOrCell(spTree, normalizedTag)
    }

    if (!found || !found.shape) {
      logger.warn(
        `replaceTextByTag: Text box/shape with ID or tag "${tagOrId}" not found on slide ${slideIndex}`
      )
      return
    }

    const shape = found.shape
    let txBody = shape['p:txBody'] || shape['a:txBody']
    if (!txBody) {
      txBody = {
        'a:bodyPr': { '@_wrap': 'square', '@_rtlCol': '0' },
        'a:lstStyle': {},
        'a:p': [],
      }
      if (found.type === 'cell') {
        shape['a:txBody'] = txBody
      } else {
        shape['p:txBody'] = txBody
      }
    }

    // Preserve root paragraph properties and base run typography
    const existingParas = Array.isArray(txBody['a:p'])
      ? txBody['a:p']
      : txBody['a:p']
        ? [txBody['a:p']]
        : []

    let basePPr = {}
    if (existingParas.length > 0 && existingParas[0]?.['a:pPr']) {
      basePPr = this.#xmlParser.deepClone(existingParas[0]['a:pPr'])
    }

    let baseRPr = { '@_lang': 'en-US' }
    let foundRPr = false
    for (const p of existingParas) {
      if (p && p['a:r']) {
        const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
        for (const r of runs) {
          if (r && r['a:rPr']) {
            baseRPr = this.#xmlParser.deepClone(r['a:rPr'])
            foundRPr = true
            break
          }
        }
        if (foundRPr) break
      }
    }

    if (!foundRPr && existingParas[0]?.['a:endParaRPr']) {
      baseRPr = this.#xmlParser.deepClone(existingParas[0]['a:endParaRPr'])
    }

    // Apply optional style overrides from options / options.style
    const style = options.style || options
    if (style && typeof style === 'object') {
      basePPr = StyleBuilder.buildParagraphProperties(basePPr, style)
      baseRPr = StyleBuilder.buildRunProperties(baseRPr, style)
      if (txBody['a:bodyPr']) {
        txBody['a:bodyPr'] = StyleBuilder.buildBodyProperties(txBody['a:bodyPr'], style)
      }
    }

    // Text casing and multi-line handling (splitting on \n)
    const rawText = newText !== undefined && newText !== null ? String(newText) : ''
    const styledText = applyTextCasing(rawText, style)
    const lines = styledText.split(/\r?\n/)

    const newParas = lines.map(line => ({
      'a:pPr': this.#xmlParser.deepClone(basePPr),
      'a:r': [
        {
          'a:rPr': this.#xmlParser.deepClone(baseRPr),
          'a:t': line,
        },
      ],
    }))

    txBody['a:p'] = newParas
    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(
      `replaceTextByTag: Replaced full text in shape "${tagOrId}" on slide ${slideIndex}`
    )
  }

  // ==========================================================================
  // 4. addLists(slideIndex, items, options, slideManager)
  // ==========================================================================
  /**
   * Adds a new bulleted or numbered list shape to a slide.
   *
   * @param {number} slideIndex
   * @param {Array<string|Object>} items
   * @param {Object} [options]
   * @param {SlideManager} slideManager
   * @returns {Object} Shape descriptor
   */
  addLists(slideIndex, items, options = {}, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    if (!spTree) {
      throw new PPTXError(`Cannot locate spTree in slide ${slideIndex}`)
    }

    const existingIds = this.#collectShapeIds(spTree)
    const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 1000
    const newId = maxId + 1
    const shapeName = options.name || options.id || `ListBox_${newId}`

    const SLIDE_WIDTH = 12192000
    const SLIDE_HEIGHT = 6858000
    const x = toEmu(options.x !== undefined ? options.x : 100, SLIDE_WIDTH)
    const y = toEmu(options.y !== undefined ? options.y : 100, SLIDE_HEIGHT)
    const width = toEmu(options.width !== undefined ? options.width : 400, SLIDE_WIDTH)
    const height = toEmu(options.height !== undefined ? options.height : 200, SLIDE_HEIGHT)

    const paras = this.#buildListParagraphs(items, options)
    const bodyPr = StyleBuilder.buildBodyProperties(
      { '@_wrap': 'square', '@_rtlCol': '0' },
      options.style || options
    )

    const shapeNode = {
      'p:nvSpPr': {
        'p:cNvPr': {
          '@_id': String(newId),
          '@_name': shapeName,
        },
        'p:cNvSpPr': {
          '@_txBox': '1',
        },
        'p:nvPr': {},
      },
      'p:spPr': {
        'a:xfrm': {
          'a:off': { '@_x': String(x), '@_y': String(y) },
          'a:ext': { '@_cx': String(width), '@_cy': String(height) },
        },
        'a:prstGeom': {
          '@_prst': 'rect',
          'a:avLst': {},
        },
        'a:noFill': {},
      },
      'p:txBody': {
        'a:bodyPr': bodyPr,
        'a:lstStyle': {},
        'a:p': paras,
      },
    }

    if (!spTree['p:sp']) {
      spTree['p:sp'] = []
    } else if (!Array.isArray(spTree['p:sp'])) {
      spTree['p:sp'] = [spTree['p:sp']]
    }
    spTree['p:sp'].push(shapeNode)

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`addLists: Created list box "${shapeName}" (ID ${newId}) on slide ${slideIndex}`)
    return { shapeId: String(newId), shapeName }
  }

  // ==========================================================================
  // 5. getList(slideIndex, tag, slideManager)
  // ==========================================================================
  /**
   * Retrieves list items from a shape or table cell by tag or shape name/ID.
   *
   * @param {number} slideIndex
   * @param {string} tag
   * @param {SlideManager} slideManager
   * @returns {Array<string|Object>}
   */
  getList(slideIndex, tag, slideManager) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']

    let res = this.findShapeRecursive(spTree, tag)
    if (!res) {
      res = this.#collectMatchingShapeOrCell(spTree, tag)
    }

    if (!res || !res.shape) return []

    const txBody = res.shape['p:txBody'] || res.shape['a:txBody']
    if (!txBody) return []

    const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
    const flatItems = []

    for (const p of paras) {
      if (!p) continue
      let lvl = 0
      if (p['a:pPr'] && p['a:pPr']['@_lvl'] !== undefined) {
        lvl = parseInt(p['a:pPr']['@_lvl'], 10) || 0
      }

      let text = ''
      if (p['a:r']) {
        const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
        const textParts = []
        for (const r of runs) {
          if (r && r['a:t'] !== undefined) textParts.push(String(r['a:t']))
        }
        text = textParts.join('')
      }

      if (text.trim() !== '') {
        flatItems.push({ text: text.trim(), lvl })
      }
    }

    if (flatItems.length === 0) return []

    // Construct hierarchy based on lvl
    const result = []
    const stack = []

    for (const item of flatItems) {
      const node = { text: item.text, children: [] }
      while (stack.length > 0 && stack[stack.length - 1].lvl >= item.lvl) {
        stack.pop()
      }
      if (stack.length === 0) {
        result.push(node)
      } else {
        const parent = stack[stack.length - 1].node
        parent.children.push(node)
      }
      stack.push({ lvl: item.lvl, node })
    }

    const cleanNode = n => {
      if (n.children.length === 0) return n.text
      return { text: n.text, children: n.children.map(cleanNode) }
    }

    return result.map(cleanNode)
  }

  // ==========================================================================
  // 6. findText(slideIndex, query, slideManager)
  // ==========================================================================
  /**
   * Searches for a text string inside all text runs across a slide.
   * Performs AST run healing to detect matches split across separate <a:r>/<a:t> nodes.
   *
   * @param {number} slideIndex
   * @param {string} query
   * @param {SlideManager} slideManager
   * @param {Object} [options]
   * @returns {Array<Object>}
   */
  findText(slideIndex, query, slideManager, options = {}) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
    const results = []

    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    if (!spTree) return results

    const caseSensitive = options.caseSensitive === true
    const searchTarget = caseSensitive ? String(query) : String(query).toLowerCase()

    this.#searchShapesForText(spTree, query, searchTarget, caseSensitive, results, slideIndex)
    return results
  }

  // ==========================================================================
  // 7. addStyle(slideIndex, target, style, slideManager)
  // ==========================================================================
  /**
   * Merges style properties into an existing shape or text element.
   *
   * @param {number} slideIndex
   * @param {string} target
   * @param {Object} style
   * @param {SlideManager} slideManager
   */
  addStyle(slideIndex, target, style, slideManager) {
    this.updateStyle(slideIndex, target, style, slideManager)
  }

  // ==========================================================================
  // 8. removeStyle(slideIndex, target, styleProperties, slideManager)
  // ==========================================================================
  /**
   * Removes specific style properties from a target shape or text element.
   *
   * @param {number} slideIndex
   * @param {string} target
   * @param {Array<string>|Object} styleProperties
   * @param {SlideManager} slideManager
   */
  removeStyle(slideIndex, target, styleProperties, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    const found =
      this.findShapeRecursive(spTree, target) || this.#collectMatchingShapeOrCell(spTree, target)
    if (!found || !found.shape) {
      logger.warn(`removeStyle: Target "${target}" not found on slide ${slideIndex}`)
      return
    }

    const txBody = found.shape['p:txBody'] || found.shape['a:txBody']
    if (!txBody) return

    const bodyPr = txBody['a:bodyPr']
    const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]

    for (const p of paras) {
      if (!p) continue
      const pPr = p['a:pPr']
      const runs = p['a:r'] ? (Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]) : []
      for (const r of runs) {
        if (r && r['a:rPr']) {
          StyleBuilder.removeStyleProperties(r['a:rPr'], pPr, bodyPr, styleProperties)
        }
      }
      if (pPr) {
        StyleBuilder.removeStyleProperties(null, pPr, bodyPr, styleProperties)
      }
    }

    if (bodyPr) {
      StyleBuilder.removeStyleProperties(null, null, bodyPr, styleProperties)
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`removeStyle: Removed styles from "${target}" on slide ${slideIndex}`)
  }

  // ==========================================================================
  // 9. updateStyle(slideIndex, target, style, slideManager)
  // ==========================================================================
  /**
   * Updates/merges style properties on a target shape or text element.
   *
   * @param {number} slideIndex
   * @param {string} target
   * @param {Object} style
   * @param {SlideManager} slideManager
   */
  updateStyle(slideIndex, target, style, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    const found =
      this.findShapeRecursive(spTree, target) || this.#collectMatchingShapeOrCell(spTree, target)
    if (!found || !found.shape) {
      logger.warn(`updateStyle: Target "${target}" not found on slide ${slideIndex}`)
      return
    }

    const txBody = found.shape['p:txBody'] || found.shape['a:txBody']
    if (!txBody) return

    if (!txBody['a:bodyPr']) txBody['a:bodyPr'] = {}
    txBody['a:bodyPr'] = StyleBuilder.buildBodyProperties(txBody['a:bodyPr'], style)

    const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
    for (const p of paras) {
      if (!p) continue
      if (!p['a:pPr']) p['a:pPr'] = {}
      p['a:pPr'] = StyleBuilder.buildParagraphProperties(p['a:pPr'], style)

      const runs = p['a:r'] ? (Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]) : []
      for (const r of runs) {
        if (!r) continue
        if (!r['a:rPr']) r['a:rPr'] = {}
        r['a:rPr'] = StyleBuilder.buildRunProperties(r['a:rPr'], style)

        if (r['a:t'] !== undefined) {
          r['a:t'] = applyTextCasing(String(r['a:t']), style)
        }
      }
    }

    const fillValue =
      style?.fill !== undefined
        ? style.fill
        : style?.backgroundColor !== undefined
          ? style.backgroundColor
          : style?.fillColor
    if (fillValue !== undefined && found.shape['p:spPr']) {
      const spPr = found.shape['p:spPr']
      delete spPr['a:noFill']
      delete spPr['a:solidFill']
      delete spPr['a:gradFill']
      delete spPr['a:blipFill']
      if (fillValue === 'none' || fillValue === 'transparent') {
        spPr['a:noFill'] = {}
      } else if (typeof fillValue === 'string') {
        const hex = fillValue.replace('#', '').toUpperCase()
        spPr['a:solidFill'] = { 'a:srgbClr': { '@_val': hex } }
      }
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`updateStyle: Applied styles to "${target}" on slide ${slideIndex}`)
  }

  /**
   * Updates/replaces text in an existing shape or placeholder non-destructively.
   *
   * @param {number} slideIndex
   * @param {string} target
   * @param {string|Object} newText
   * @param {Object} [options]
   * @param {SlideManager} slideManager
   * @param {TemplateEngine} [templateEngine]
   */
  updateText(slideIndex, target, newText, options = {}, slideManager, templateEngine) {
    let actualText = newText
    let actualOptions = options
    if (
      typeof newText === 'object' &&
      newText !== null &&
      (options === undefined || Object.keys(options).length === 0)
    ) {
      actualOptions = newText
      actualText = actualOptions.text
    }
    return this.replaceTextByTag(
      slideIndex,
      target,
      actualText,
      actualOptions,
      slideManager,
      templateEngine
    )
  }

  // ==========================================================================
  // 10. getText(slideIndex, tag, slideManager)
  // ==========================================================================
  /**
   * Retrieves full concatenated plain text from a shape or tag.
   *
   * @param {number} slideIndex
   * @param {string} tag
   * @param {SlideManager} slideManager
   * @returns {string}
   */
  getText(slideIndex, tag, slideManager) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']

    const found =
      this.findShapeRecursive(spTree, tag) || this.#collectMatchingShapeOrCell(spTree, tag)
    if (!found || !found.shape) return ''

    const txBody = found.shape['p:txBody'] || found.shape['a:txBody']
    if (!txBody) return ''

    const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
    const parts = []

    for (const p of paras) {
      if (!p) continue
      const runs = p['a:r'] ? (Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]) : []
      for (const r of runs) {
        if (r && r['a:t'] !== undefined) {
          parts.push(String(r['a:t']))
        }
      }
    }
    return parts.join('')
  }

  // ==========================================================================
  // 11. getLists(slideIndex, slideManager)
  // ==========================================================================
  /**
   * Retrieves all list structures found across shapes on a slide.
   *
   * @param {number} slideIndex
   * @param {SlideManager} slideManager
   * @returns {Array<{shapeId: string, shapeName: string, items: Array}>}
   */
  getLists(slideIndex, slideManager) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    const lists = []

    if (!spTree) return lists

    const shapes = this.#collectAllShapes(spTree)
    for (const shape of shapes) {
      const cNvPr = shape?.['p:nvSpPr']?.['p:cNvPr']
      const shapeName = cNvPr ? cNvPr['@_name'] : 'unnamed'
      const shapeId = cNvPr ? String(cNvPr['@_id']) : 'unknown'

      const txBody = shape['p:txBody']
      if (!txBody || !txBody['a:p']) continue

      const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
      const hasBullet = paras.some(
        p =>
          p['a:pPr'] && (p['a:pPr']['a:buChar'] || p['a:pPr']['a:buAutoNum'] || p['a:pPr']['@_lvl'])
      )

      if (hasBullet) {
        const items = this.getList(slideIndex, shapeName, slideManager)
        if (items && items.length > 0) {
          lists.push({ shapeId, shapeName, items })
        }
      }
    }

    return lists
  }

  // ==========================================================================
  // 12. findLists(slideIndex, query, slideManager)
  // ==========================================================================
  /**
   * Finds lists on a slide matching a query (by shape name, ID, or item contents).
   *
   * @param {number} slideIndex
   * @param {string} query
   * @param {SlideManager} slideManager
   * @returns {Array<Object>}
   */
  findLists(slideIndex, query, slideManager) {
    const all = this.getLists(slideIndex, slideManager)
    const q = String(query).toLowerCase()

    return all.filter(entry => {
      if (entry.shapeName.toLowerCase().includes(q)) return true
      if (entry.shapeId.includes(q)) return true
      const hasMatch = items => {
        return items.some(item => {
          if (typeof item === 'string') return item.toLowerCase().includes(q)
          if (item && item.text) {
            if (item.text.toLowerCase().includes(q)) return true
            if (Array.isArray(item.children)) return hasMatch(item.children)
          }
          return false
        })
      }
      return hasMatch(entry.items)
    })
  }

  // ==========================================================================
  // 13. removeText(slideIndex, tag, slideManager)
  // ==========================================================================
  /**
   * Removes a text shape or text body matching the tag.
   *
   * @param {number} slideIndex
   * @param {string} tag
   * @param {SlideManager} slideManager
   */
  removeText(slideIndex, tag, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    const found =
      this.findShapeRecursive(spTree, tag) || this.#collectMatchingShapeOrCell(spTree, tag)
    if (!found || !found.shape || !found.parent) {
      logger.warn(`removeText: Shape "${tag}" not found on slide ${slideIndex}`)
      return
    }

    if (found.parent['p:sp']) {
      const shapes = Array.isArray(found.parent['p:sp'])
        ? found.parent['p:sp']
        : [found.parent['p:sp']]
      found.parent['p:sp'] = shapes.filter(s => s !== found.shape)
      slideManager.markSlideObjDirty(slideIndex)
      logger.debug(`removeText: Removed shape "${tag}" on slide ${slideIndex}`)
    }
  }

  // ==========================================================================
  // 14. removeList(slideIndex, tag, slideManager)
  // ==========================================================================
  /**
   * Removes a list shape matching tag from the slide.
   *
   * @param {number} slideIndex
   * @param {string} tag
   * @param {SlideManager} slideManager
   */
  removeList(slideIndex, tag, slideManager) {
    this.removeText(slideIndex, tag, slideManager)
  }

  // ==========================================================================
  // 15. clearText(slideIndex, tag, slideManager)
  // ==========================================================================
  /**
   * Clears text content inside a shape while preserving container geometry and position.
   *
   * @param {number} slideIndex
   * @param {string} tag
   * @param {SlideManager} slideManager
   */
  clearText(slideIndex, tag, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    const found =
      this.findShapeRecursive(spTree, tag) || this.#collectMatchingShapeOrCell(spTree, tag)
    if (!found || !found.shape) {
      logger.warn(`clearText: Shape "${tag}" not found on slide ${slideIndex}`)
      return
    }

    const txBody = found.shape['p:txBody'] || found.shape['a:txBody']
    if (txBody) {
      // Empty paragraph preserving paragraph container
      txBody['a:p'] = [{ 'a:pPr': {}, 'a:r': [] }]
      slideManager.markSlideObjDirty(slideIndex)
      logger.debug(`clearText: Cleared text in shape "${tag}" on slide ${slideIndex}`)
    }
  }

  // ==========================================================================
  // 16. clearList(slideIndex, tag, slideManager)
  // ==========================================================================
  /**
   * Clears list items from a shape while retaining shape position and bullet properties.
   *
   * @param {number} slideIndex
   * @param {string} tag
   * @param {SlideManager} slideManager
   */
  clearList(slideIndex, tag, slideManager) {
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    const found =
      this.findShapeRecursive(spTree, tag) || this.#collectMatchingShapeOrCell(spTree, tag)
    if (!found || !found.shape) return

    const txBody = found.shape['p:txBody'] || found.shape['a:txBody']
    if (txBody) {
      const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
      const firstPPr = paras.length > 0 && paras[0]['a:pPr'] ? paras[0]['a:pPr'] : {}
      txBody['a:p'] = [{ 'a:pPr': firstPPr, 'a:r': [] }]
      slideManager.markSlideObjDirty(slideIndex)
      logger.debug(`clearList: Cleared list in shape "${tag}" on slide ${slideIndex}`)
    }
  }

  // ==========================================================================
  // 17. replaceList(slideIndex, tag, newItems, options, slideManager, templateEngine)
  // ==========================================================================
  /**
   * Replaces list items inside a shape or table cell matching tag.
   *
   * @param {number} slideIndex
   * @param {string} tag
   * @param {Array<string|Object>|Object} newItems
   * @param {Object} [options]
   * @param {SlideManager} slideManager
   * @param {TemplateEngine} templateEngine
   */
  replaceList(slideIndex, tag, newItems, options = {}, slideManager, templateEngine) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const normalizedTag = tag.startsWith('{{') && tag.endsWith('}}') ? tag : `{{${tag}}}`

    // Option A: Tag exists as a placeholder in slide XML
    const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']

    let res = this.findShapeRecursive(spTree, tag)
    if (!res) {
      res =
        this.#collectMatchingShapeOrCell(spTree, normalizedTag) ||
        this.#collectMatchingShapeOrCell(spTree, tag)
    }

    if (!res || !res.shape) {
      throw new PPTXError(`Text placeholder or shape "${tag}" not found in slide ${slideIndex}`)
    }

    const shape = res.shape
    const listConfig = Array.isArray(newItems)
      ? { list: newItems, ...options }
      : typeof newItems === 'object' && newItems !== null
        ? newItems.list !== undefined || newItems.items !== undefined
          ? { ...newItems, ...options }
          : { list: [newItems], ...options }
        : { list: [String(newItems)], ...options }

    const validation = ValidationEngine.validateList(listConfig)
    if (!validation.valid) {
      throw new PPTXError(`List validation failed: ${validation.errors.join(', ')}`)
    }

    const txBody = shape['p:txBody'] || shape['a:txBody'] || {}
    if (!shape['p:txBody'] && !shape['a:txBody']) {
      shape['p:txBody'] = txBody
      txBody['a:bodyPr'] = { '@_wrap': 'square', '@_rtlCol': '0' }
      txBody['a:lstStyle'] = {}
    }

    const originalParas = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
    const templatePara = originalParas.length > 0 && originalParas[0] ? originalParas[0] : {}
    const templateRuns = templatePara['a:r']
      ? Array.isArray(templatePara['a:r'])
        ? templatePara['a:r']
        : [templatePara['a:r']]
      : []
    const firstRun = templateRuns.length > 0 ? templateRuns[0] : { 'a:rPr': {} }

    const firstRunXml = this.#xmlParser.build({ 'a:r': firstRun })
    const dummyParaXml = `<a:p>${firstRunXml}</a:p>`
    const generatedXml = templateEngine.generateListParagraphs(
      dummyParaXml,
      { xml: firstRunXml },
      listConfig
    )

    const wrappedXml = `<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${generatedXml}</root>`
    const parsedObj = this.#xmlParser.parse(wrappedXml)
    let newParas = parsedObj?.root?.['a:p'] || []
    if (!Array.isArray(newParas)) {
      newParas = [newParas]
    }

    txBody['a:p'] = newParas

    const decl = this.#xmlParser.extractDeclaration(slideXml)
    slideManager.setSlideXml(slideIndex, this.#xmlParser.build(slideObj, decl))
    logger.debug(`replaceList: Replaced list items for "${tag}" on slide ${slideIndex}`)
  }

  // ==========================================================================
  // 18. getTextElements(slideIndex, slideManager)
  // ==========================================================================
  /**
   * Extracts and returns all text elements on a slide.
   *
   * @param {number} slideIndex
   * @param {SlideManager} slideManager
   * @returns {Array<{shapeId: string, shapeName: string, paragraphIndex: number, text: string}>}
   */
  getTextElements(slideIndex, slideManager) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
    const results = []

    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    if (!spTree) return results

    this.#collectTextElements(spTree, results, slideIndex)
    return results
  }

  // ==========================================================================
  // Internal Helpers
  // ==========================================================================

  #setTextBody(shape, textContent, options = {}) {
    if (!shape['p:txBody']) {
      shape['p:txBody'] = {
        'a:bodyPr': { '@_wrap': 'square', '@_rtlCol': '0' },
        'a:lstStyle': {},
        'a:p': [],
      }
    }
    const txBody = shape['p:txBody']
    const style = options.style || options
    const styledText = applyTextCasing(textContent, style)

    const rPr = StyleBuilder.buildRunProperties({ '@_lang': 'en-US' }, style)
    const pPr = StyleBuilder.buildParagraphProperties({}, style)
    txBody['a:bodyPr'] = StyleBuilder.buildBodyProperties(txBody['a:bodyPr'] || {}, style)

    txBody['a:p'] = [
      {
        'a:pPr': pPr,
        'a:r': [
          {
            'a:rPr': rPr,
            'a:t': styledText,
          },
        ],
      },
    ]
  }

  #buildListParagraphs(items, options = {}) {
    const paragraphs = []
    const baseStyle = options.style || options

    const processItems = (itemList, level = 0) => {
      for (const item of itemList) {
        if (typeof item === 'string') {
          const itemStyle = { ...baseStyle, indentLevel: level }
          if (options.ordered) {
            itemStyle.bulletStyle = 'ordered'
          } else if (!itemStyle.bulletChar && !itemStyle.bulletStyle) {
            itemStyle.bulletChar = '•'
          }
          const pPr = StyleBuilder.buildParagraphProperties({}, itemStyle)
          const rPr = StyleBuilder.buildRunProperties({ '@_lang': 'en-US' }, itemStyle)
          paragraphs.push({
            'a:pPr': pPr,
            'a:r': [{ 'a:rPr': rPr, 'a:t': applyTextCasing(item, itemStyle) }],
          })
        } else if (item && typeof item === 'object') {
          const text = item.text || item.title || ''
          const itemStyle = {
            ...baseStyle,
            ...(item.style || {}),
            indentLevel: item.level !== undefined ? item.level : level,
          }
          if (options.ordered || item.ordered) {
            itemStyle.bulletStyle = 'ordered'
          } else if (!itemStyle.bulletChar && !itemStyle.bulletStyle) {
            itemStyle.bulletChar = '•'
          }
          const pPr = StyleBuilder.buildParagraphProperties({}, itemStyle)
          const rPr = StyleBuilder.buildRunProperties({ '@_lang': 'en-US' }, itemStyle)
          paragraphs.push({
            'a:pPr': pPr,
            'a:r': [{ 'a:rPr': rPr, 'a:t': applyTextCasing(text, itemStyle) }],
          })
          if (Array.isArray(item.children) && item.children.length > 0) {
            processItems(item.children, level + 1)
          }
        }
      }
    }

    const initialLevel =
      options.indentLevel !== undefined
        ? options.indentLevel
        : options.bulletLevel !== undefined
          ? options.bulletLevel
          : options.style?.indentLevel !== undefined
            ? options.style.indentLevel
            : options.style?.bulletLevel !== undefined
              ? options.style.bulletLevel
              : 0

    processItems(items, initialLevel)
    return paragraphs
  }

  /**
   * Helper to recursively scan a container for shapes matching shapeId or shapeName.
   */
  findShapeRecursive(container, shapeId) {
    if (!container) return null

    const targetStr = String(shapeId).trim()
    const targetLower = targetStr.toLowerCase()

    let shapes = container['p:sp'] || []
    if (!Array.isArray(shapes)) shapes = [shapes]

    for (const shape of shapes) {
      const cNvPr = shape?.['p:nvSpPr']?.['p:cNvPr']
      if (cNvPr) {
        const id = String(cNvPr['@_id'] || '').trim()
        const name = String(cNvPr['@_name'] || '').trim()
        const tag = String(cNvPr['@_tag'] || '').trim()
        const descr = String(cNvPr['@_descr'] || '').trim()

        if (
          id === targetStr ||
          name === targetStr ||
          name.toLowerCase() === targetLower ||
          tag === targetStr ||
          tag.toLowerCase() === targetLower ||
          descr === targetStr ||
          descr.toLowerCase() === targetLower
        ) {
          return { shape, parent: container, type: 'sp' }
        }
      }

      // Check placeholder type (e.g. title, subTitle, body, ctrTitle)
      const ph = shape?.['p:nvSpPr']?.['p:nvPr']?.['p:ph']
      if (ph) {
        const phType = String(ph['@_type'] || '')
          .trim()
          .toLowerCase()
        const phIdx = String(ph['@_idx'] || '').trim()
        if ((phType && phType === targetLower) || (phIdx && phIdx === targetStr)) {
          return { shape, parent: container, type: 'sp' }
        }
      }
    }

    let frames = container['p:graphicFrame'] || []
    if (!Array.isArray(frames)) frames = [frames]

    for (const frame of frames) {
      const cNvPr = frame?.['p:nvGraphicFramePr']?.['p:cNvPr']
      if (cNvPr) {
        const id = String(cNvPr['@_id'] || '').trim()
        const name = String(cNvPr['@_name'] || '').trim()
        if (id === targetStr || name === targetStr || name.toLowerCase() === targetLower) {
          return { shape: frame, parent: container, type: 'graphicFrame' }
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

  #collectMatchingShapeOrCell(container, tag) {
    if (!container) return null

    let shapes = container['p:sp'] || []
    if (!Array.isArray(shapes)) shapes = [shapes]

    for (const shape of shapes) {
      const txBody = shape['p:txBody']
      if (txBody && txBody['a:p']) {
        const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
        for (const p of paras) {
          let pText = ''
          if (p['a:r']) {
            const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
            for (const r of runs) {
              if (r['a:t'] !== undefined) pText += String(r['a:t'])
            }
          }
          if (pText.includes(tag)) {
            return { shape, parent: container, type: 'sp' }
          }
        }
      }
    }

    let frames = container['p:graphicFrame'] || []
    if (!Array.isArray(frames)) frames = [frames]

    for (const frame of frames) {
      const tbl = frame?.['a:graphic']?.['a:graphicData']?.['a:tbl']
      if (tbl && tbl['a:tr']) {
        const rows = Array.isArray(tbl['a:tr']) ? tbl['a:tr'] : [tbl['a:tr']]
        for (const row of rows) {
          const cells = Array.isArray(row['a:tc']) ? row['a:tc'] : [row['a:tc']]
          for (const cell of cells) {
            const txBody = cell['a:txBody']
            if (txBody && txBody['a:p']) {
              const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
              for (const p of paras) {
                let pText = ''
                if (p['a:r']) {
                  const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
                  for (const r of runs) {
                    if (r['a:t'] !== undefined) pText += String(r['a:t'])
                  }
                }
                if (pText.includes(tag)) {
                  return { shape: cell, parent: row, type: 'cell' }
                }
              }
            }
          }
        }
      }
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]

    for (const group of groups) {
      const matched = this.#collectMatchingShapeOrCell(group, tag)
      if (matched) return matched
    }

    return null
  }

  #collectShapeIds(container) {
    const ids = []
    if (!container) return ids

    let shapes = container['p:sp'] || []
    if (!Array.isArray(shapes)) shapes = [shapes]

    for (const shape of shapes) {
      const id = shape?.['p:nvSpPr']?.['p:cNvPr']?.['@_id']
      if (id !== undefined) {
        const parsed = parseInt(id, 10)
        if (!isNaN(parsed)) ids.push(parsed)
      }
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const g of groups) {
      ids.push(...this.#collectShapeIds(g))
    }

    return ids
  }

  #collectAllShapes(container) {
    const shapes = []
    if (!container) return shapes

    let directShapes = container['p:sp'] || []
    if (!Array.isArray(directShapes)) directShapes = [directShapes]
    shapes.push(...directShapes)

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const g of groups) {
      shapes.push(...this.#collectAllShapes(g))
    }

    return shapes
  }

  #searchShapesForText(container, rawSearchText, searchTarget, caseSensitive, results, slideIndex) {
    if (!container) return

    let shapes = container['p:sp'] || []
    if (!Array.isArray(shapes)) shapes = [shapes]

    for (const shape of shapes) {
      const cNvPr = shape?.['p:nvSpPr']?.['p:cNvPr']
      const shapeName = cNvPr ? cNvPr['@_name'] : 'unnamed'
      const shapeId = cNvPr ? String(cNvPr['@_id']) : 'unknown'

      const txBody = shape['p:txBody']
      if (txBody && txBody['a:p']) {
        const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
        paras.forEach((p, pIdx) => {
          let pText = ''
          if (p['a:r']) {
            const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
            runs.forEach(r => {
              if (r && r['a:t'] !== undefined) pText += String(r['a:t'])
            })
          }

          const haystack = caseSensitive ? pText : pText.toLowerCase()
          if (haystack.includes(searchTarget)) {
            results.push({
              slideIndex,
              shapeId,
              shapeName,
              paragraphIndex: pIdx,
              text: pText,
              match: rawSearchText,
            })
          }
        })
      }
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const g of groups) {
      this.#searchShapesForText(g, rawSearchText, searchTarget, caseSensitive, results, slideIndex)
    }
  }

  #collectTextElements(container, results, slideIndex) {
    if (!container) return

    let shapes = container['p:sp'] || []
    if (!Array.isArray(shapes)) shapes = [shapes]

    for (const shape of shapes) {
      const cNvPr = shape?.['p:nvSpPr']?.['p:cNvPr']
      const shapeName = cNvPr ? cNvPr['@_name'] : 'unnamed'
      const shapeId = cNvPr ? String(cNvPr['@_id']) : 'unknown'

      const txBody = shape['p:txBody']
      if (txBody && txBody['a:p']) {
        const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
        paras.forEach((p, pIdx) => {
          let pText = ''
          if (p['a:r']) {
            const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
            runs.forEach(r => {
              if (r && r['a:t'] !== undefined) pText += String(r['a:t'])
            })
          }

          if (pText.trim()) {
            results.push({
              slideIndex,
              shapeId,
              shapeName,
              paragraphIndex: pIdx,
              text: pText,
            })
          }
        })
      }
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const g of groups) {
      this.#collectTextElements(g, results, slideIndex)
    }
  }
}

module.exports = { TextManager }
