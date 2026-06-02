/**
 * @fileoverview TextManager - Handles slide text search, retrieval, and replacement.
 */

const { createLogger } = require('../utils/logger.js')

const logger = createLogger('TextManager')

/**
 * @class TextManager
 * @description Manages text elements, replacements, and search inside slide XML.
 */
class TextManager {
  /** @private @type {XMLParser} */
  #xmlParser

  /**
   * @param {XMLParser} xmlParser
   */
  constructor(xmlParser) {
    this.#xmlParser = xmlParser
  }

  /**
   * Replaces a specific tag/placeholder with a value.
   *
   * @param {number} slideIndex
   * @param {string} tag - E.g. '{{name}}' or 'name' (auto-wraps if simple).
   * @param {string} value
   * @param {Object} options
   * @param {SlideManager} slideManager
   * @param {TemplateEngine} templateEngine
   */
  replaceTextByTag(slideIndex, tag, value, _options = {}, slideManager, templateEngine) {
    const slideXml = slideManager.getSlideXml(slideIndex)

    // Auto-wrap tag in {{}} if not already present
    const normalizedTag = tag.startsWith('{{') && tag.endsWith('}}') ? tag : `{{${tag}}}`

    const replacements = { [normalizedTag]: value }
    const updatedXml = templateEngine.replaceTextInXml(slideXml, replacements)

    slideManager.setSlideXml(slideIndex, updatedXml)
    logger.debug(`Replaced text tag "${normalizedTag}" with value on slide ${slideIndex}`)
  }

  /**
   * Performs multiple text replacements at once.
   *
   * @param {number} slideIndex
   * @param {Object.<string, string>} replacements - Map of key -> value.
   * @param {Object} options
   * @param {SlideManager} slideManager
   * @param {TemplateEngine} templateEngine
   */
  replaceMultiple(slideIndex, replacements, _options = {}, slideManager, templateEngine) {
    const slideXml = slideManager.getSlideXml(slideIndex)

    // Normalize keys in the replacements map to ensure they are wrapped in placeholders
    const normalized = {}
    for (const [key, val] of Object.entries(replacements)) {
      const normalizedKey = key.startsWith('{{') && key.endsWith('}}') ? key : `{{${key}}}`
      normalized[normalizedKey] = val
    }

    const updatedXml = templateEngine.replaceTextInXml(slideXml, normalized)
    slideManager.setSlideXml(slideIndex, updatedXml)
    logger.debug(`Replaced multiple tags on slide ${slideIndex}`)
  }

  /**
   * Searches for a text string inside all text runs on a slide.
   *
   * @param {number} slideIndex
   * @param {string} searchText
   * @param {SlideManager} slideManager
   * @returns {Array<Object>} List of match details.
   */
  findText(slideIndex, searchText, slideManager) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
    const results = []

    // Find in shapes (p:sp)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    if (!spTree) return results

    this.#searchShapesForText(spTree, searchText, results)
    return results
  }

  /**
   * Extracts and returns all text elements on a slide.
   *
   * @param {number} slideIndex
   * @param {SlideManager} slideManager
   * @returns {Array<Object>} Elements with text contents.
   */
  getTextElements(slideIndex, slideManager) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
    const results = []

    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    if (!spTree) return results

    this.#collectTextElements(spTree, results)
    return results
  }

  #searchShapesForText(container, searchText, results) {
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
              if (r['a:t']) pText += String(r['a:t'])
            })
          }

          if (pText.toLowerCase().includes(searchText.toLowerCase())) {
            results.push({
              shapeId,
              shapeName,
              paragraphIndex: pIdx,
              text: pText,
              match: searchText,
            })
          }
        })
      }
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const g of groups) {
      this.#searchShapesForText(g, searchText, results)
    }
  }

  #collectTextElements(container, results) {
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
              if (r['a:t']) pText += String(r['a:t'])
            })
          }

          if (pText.trim()) {
            results.push({
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
      this.#collectTextElements(g, results)
    }
  }
}

module.exports = { TextManager }
