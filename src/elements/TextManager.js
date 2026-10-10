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

  /**
   * Updates shape text or list content by placeholder tag or shape name/ID.
   *
   * @param {number} slideIndex
   * @param {string} tag - Placeholder tag (e.g. '{{name}}') or shape name/ID.
   * @param {string|Object} data - String value or list configuration object.
   * @param {SlideManager} slideManager
   * @param {TemplateEngine} templateEngine
   */
  updateText(slideIndex, tag, data, slideManager, templateEngine) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const normalizedTag = tag.startsWith('{{') && tag.endsWith('}}') ? tag : `{{${tag}}}`

    // Option A: Tag exists as a placeholder in the slide XML
    if (slideXml.includes(normalizedTag)) {
      const replacements = { [normalizedTag]: data }
      const updatedXml = templateEngine.replaceTextInXml(slideXml, replacements)
      slideManager.setSlideXml(slideIndex, updatedXml)
      logger.debug(`Updated text tag "${normalizedTag}" on slide ${slideIndex}`)
      return
    }

    // Option B: Search for a shape whose name or ID matches the tag
    const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    const res = this.findShapeRecursive(spTree, tag)

    if (!res) {
      const { PPTXError } = require('../utils/errors.js')
      throw new PPTXError(`Text placeholder or shape "${tag}" not found in slide ${slideIndex}`)
    }

    // Replace the text body of the shape
    const shape = res.shape
    const listConfig =
      typeof data === 'object' && data !== null
        ? data.list !== undefined
          ? data
          : { list: [data] }
        : { list: [String(data)] }

    const { ValidationEngine } = require('../core/ValidationEngine.js')
    const validation = ValidationEngine.validateList(listConfig)
    if (!validation.valid) {
      throw new Error(`List validation failed: ${validation.errors.join(', ')}`)
    }

    if (!shape['p:txBody']) {
      shape['p:txBody'] = {
        'a:bodyPr': {},
        'a:lstStyle': {},
        'a:p': [],
      }
    }

    const txBody = shape['p:txBody']
    const originalParas = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]
    const templatePara = originalParas.length > 0 ? originalParas[0] : {}
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

    // Parse the generated XML paragraphs back to objects
    const wrappedXml = `<root xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${generatedXml}</root>`
    const parsedObj = this.#xmlParser.parse(wrappedXml)
    let newParas = parsedObj?.root?.['a:p'] || []
    if (!Array.isArray(newParas)) {
      newParas = [newParas]
    }

    txBody['a:p'] = newParas

    const decl = this.#xmlParser.extractDeclaration(slideXml)
    slideManager.setSlideXml(slideIndex, this.#xmlParser.build(slideObj, decl))
    logger.debug(`Updated text content for shape "${tag}" on slide ${slideIndex}`)
  }

  /**
   * Retrieves list items from a shape or text box by name or placeholder tag.
   *
   * @param {number} slideIndex
   * @param {string} tag - Shape name/ID or placeholder tag.
   * @param {SlideManager} slideManager
   * @returns {Array} Nested list structure of items.
   */
  getList(slideIndex, tag, slideManager) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']

    // Step 1: Find shape by name or ID matching tag
    let res = this.findShapeRecursive(spTree, tag)

    // Step 2: If not found, look for any shape containing the placeholder string
    if (!res) {
      const collectMatchingShape = container => {
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
                  if (r['a:t']) pText += String(r['a:t'])
                }
              }
              if (pText.includes(tag)) {
                return { shape, parent: container, type: 'sp' }
              }
            }
          }
        }

        // Search in tables inside graphicFrames
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
                        if (r['a:t']) pText += String(r['a:t'])
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
          const matched = collectMatchingShape(group)
          if (matched) return matched
        }

        return null
      }
      res = collectMatchingShape(spTree)
    }

    if (!res || !res.shape || !(res.shape['p:txBody'] || res.shape['a:txBody'])) {
      return []
    }

    const txBody = res.shape['p:txBody'] || res.shape['a:txBody']
    const paras = Array.isArray(txBody['a:p']) ? txBody['a:p'] : [txBody['a:p']]

    const flatItems = []
    for (const p of paras) {
      let lvl = 0
      if (p['a:pPr'] && p['a:pPr']['@_lvl'] !== undefined) {
        lvl = parseInt(p['a:pPr']['@_lvl'], 10) || 0
      }

      let text = ''
      if (p['a:r']) {
        const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
        const textParts = []
        for (const r of runs) {
          if (r['a:t']) textParts.push(String(r['a:t']))
        }
        text = textParts.join('')
      }

      if (text.trim() !== '') {
        flatItems.push({ text: text.trim(), lvl })
      }
    }

    if (flatItems.length === 0) {
      return []
    }

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
      if (n.children.length === 0) {
        return n.text
      }
      return {
        text: n.text,
        children: n.children.map(cleanNode),
      }
    }

    return result.map(cleanNode)
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
}

module.exports = { TextManager }
