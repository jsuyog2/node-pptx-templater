/**
 * @fileoverview XMLParser - Unified XML parsing and serialization layer.
 *
 * Uses fast-xml-parser for high-performance XML → JS object conversion.
 * Provides consistent, reusable parse/build methods used by all managers.
 *
 * Key design decisions:
 *  - Attributes are stored with '@_' prefix (fast-xml-parser convention)
 *  - Arrays are preserved for elements that can repeat (e.g., slides, rows)
 *  - Text content uses '#text' key
 *  - CDATA sections are preserved
 *  - XML namespaces (a:, p:, r:, etc.) are preserved as-is
 *
 * OpenXML namespace prefixes you'll see:
 *  a:  — DrawingML (shapes, fonts, colors)
 *  p:  — PresentationML (slides, layouts, masters)
 *  r:  — Relationships (rId references)
 *  c:  — ChartML (chart data)
 *  w:  — WordprocessingML (not typically in PPTX)
 *  mc: — Markup Compatibility
 */

const { XMLParser: FastXMLParser, XMLBuilder } = require('fast-xml-parser')
const { PPTXError } = require('../utils/errors.js')

const Z_ORDER_SYMBOL = Symbol('zOrder')

const drawingTags = new Set(['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp'])

function findcNvPr(node) {
  const tagName = Object.keys(node).find(k => k !== ':@')
  if (!tagName) return null
  const children = node[tagName]
  if (!Array.isArray(children)) return null

  const nvPrNode = children.find(child => {
    const k = Object.keys(child).find(key => key !== ':@')
    return (
      k &&
      (k === 'p:nvSpPr' ||
        k === 'p:nvPicPr' ||
        k === 'p:nvGraphicFramePr' ||
        k === 'p:nvGrpSpPr' ||
        k === 'p:nvCxnSpPr')
    )
  })

  if (nvPrNode) {
    const nvPrTagName = Object.keys(nvPrNode).find(k => k !== ':@')
    const nvPrChildren = nvPrNode[nvPrTagName]
    if (Array.isArray(nvPrChildren)) {
      const cNvPrNode = nvPrChildren.find(child => Object.keys(child).includes('p:cNvPr'))
      if (cNvPrNode) {
        return cNvPrNode[':@']
      }
    }
  }
  return null
}

function buildZOrderMap(orderedNode, containerMap = new Map()) {
  if (!orderedNode) return containerMap
  const tagName = Object.keys(orderedNode).find(k => k !== ':@')
  if (!tagName) return containerMap
  const children = orderedNode[tagName]
  if (!Array.isArray(children)) return containerMap

  if (tagName === 'p:spTree') {
    const order = []
    for (const child of children) {
      const childTagName = Object.keys(child).find(k => k !== ':@')
      if (drawingTags.has(childTagName)) {
        const attrs = findcNvPr(child)
        if (attrs && attrs['@_id']) {
          order.push(String(attrs['@_id']))
        }
        buildZOrderMap(child, containerMap)
      }
    }
    containerMap.set('root', order)
  } else if (tagName === 'p:grpSp') {
    const attrs = findcNvPr(orderedNode)
    const grpId = attrs ? String(attrs['@_id']) : null
    const order = []
    for (const child of children) {
      const childTagName = Object.keys(child).find(k => k !== ':@')
      if (drawingTags.has(childTagName)) {
        const attrs = findcNvPr(child)
        if (attrs && attrs['@_id']) {
          order.push(String(attrs['@_id']))
        }
        buildZOrderMap(child, containerMap)
      }
    }
    if (grpId) {
      containerMap.set(grpId, order)
    }
  } else {
    for (const child of children) {
      buildZOrderMap(child, containerMap)
    }
  }

  return containerMap
}

function attachZOrder(normalContainer, containerId, containerMap) {
  if (!normalContainer) return

  const order = containerMap.get(containerId)
  if (order) {
    normalContainer[Z_ORDER_SYMBOL] = order
  }

  let grpSps = normalContainer['p:grpSp'] || []
  if (!Array.isArray(grpSps)) grpSps = [grpSps]

  for (const grpSp of grpSps) {
    const cNvPr = grpSp?.['p:nvGrpSpPr']?.['p:cNvPr']
    const grpId = cNvPr ? String(cNvPr['@_id']) : null
    if (grpId) {
      attachZOrder(grpSp, grpId, containerMap)
    }
  }
}

/**
 * Parser configuration for fast-xml-parser.
 * These settings ensure lossless round-trip XML parsing.
 */
const PARSER_OPTIONS = {
  ignoreAttributes: false,
  ignoreDeclaration: true,
  attributeNamePrefix: '@_',
  allowBooleanAttributes: true,
  parseAttributeValue: false, // Keep all values as strings to avoid type coercion
  parseTagValue: false,
  cdataPropName: '__cdata',
  commentPropName: '__comment',
  preserveOrder: false,
  trimValues: false,
  processEntities: true,
  htmlEntities: false,
  isArray: (name, jpath) => {
    // Elements that should ALWAYS be arrays (even when there's only one)
    const alwaysArrayPaths = [
      'p:sld.p:cSld.p:spTree.p:sp',
      'p:sld.p:cSld.p:spTree.p:pic',
      'p:sld.p:cSld.p:spTree.p:graphicFrame',
      'p:sld.p:cSld.p:spTree.p:grpSp',
      'p:sldMaster.p:sldLayoutIdLst.p:sldLayoutId',
      'p:presentation.p:sldMasterIdLst.p:sldMasterId',
      'p:presentation.p:sldIdLst.p:sldId',
      'a:tbl.a:tr',
      'a:tr.a:tc',
      'c:ser',
      'c:pt',
      'c:cat.c:strRef.c:strCache.c:pt',
      'c:val.c:numRef.c:numCache.c:pt',
      'p:sp',
      'p:pic',
      'a:r', // text runs
      'Relationship',
      'Override',
      'Default',
      'p14:sldId',
      'p14:section',
    ]
    return alwaysArrayPaths.some(path => jpath.endsWith(path) || name === path.split('.').pop())
  },
}

/**
 * Builder configuration for XMLBuilder.
 * Must match the parser configuration for correct round-trip.
 */
const BUILDER_OPTIONS = {
  ignoreAttributes: false,
  ignoreDeclaration: true,
  attributeNamePrefix: '@_',
  cdataPropName: '__cdata',
  commentPropName: '__comment',
  suppressEmptyNode: false,
  format: false, // No extra whitespace — PPTX is sensitive to whitespace in some cases
  processEntities: true,
}

/**
 * @class XMLParser
 * @description Provides XML parsing and serialization with OpenXML-aware configuration.
 */
class XMLParser {
  /**
   * @private
   * @type {FastXMLParser}
   */
  #parser

  /**
   * @private
   * @type {XMLBuilder}
   */
  #builder

  /**
   * @private
   * @type {FastXMLParser}
   */
  #orderedParser

  constructor() {
    this.#parser = new FastXMLParser(PARSER_OPTIONS)
    this.#builder = new XMLBuilder(BUILDER_OPTIONS)
    this.#orderedParser = new FastXMLParser({
      ignoreAttributes: false,
      preserveOrder: true,
      attributeNamePrefix: '@_',
      parseAttributeValue: false,
      parseTagValue: false,
    })
  }

  /**
   * Parses an XML string into a JavaScript object.
   * The resulting object preserves all attributes, namespaces, and structure.
   *
   * @param {string} xmlString - Raw XML content.
   * @param {string} [context] - Optional context description for error messages.
   * @returns {Object} Parsed JavaScript object.
   * @throws {PPTXError} If XML is malformed.
   *
   * @example
   * const obj = parser.parse('<p:sp><p:nvSpPr>...</p:nvSpPr></p:sp>');
   */
  parse(xmlString, context = '') {
    if (!xmlString || typeof xmlString !== 'string') {
      throw new PPTXError(`Invalid XML input${context ? ` (${context})` : ''}`)
    }

    try {
      const normalObj = this.#parser.parse(xmlString)

      // Automatically inspect for spTree to build and attach Z-order
      const spTree =
        normalObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
        normalObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
        normalObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

      if (spTree) {
        try {
          const orderedObj = this.#orderedParser.parse(xmlString)
          const orderedRootNode = orderedObj.find(n => {
            const keys = Object.keys(n)
            return (
              keys.includes('p:sld') ||
              keys.includes('p:sldLayout') ||
              keys.includes('p:sldMaster')
            )
          })
          const containerMap = buildZOrderMap(orderedRootNode)
          attachZOrder(spTree, 'root', containerMap)
        } catch (err) {
          // Fallback gracefully if ordered parsing fails
        }
      }

      return normalObj
    } catch (err) {
      throw new PPTXError(`XML parse error${context ? ` in ${context}` : ''}: ${err.message}`, err)
    }
  }

  /**
   * Serializes a JavaScript object back to an XML string.
   *
   * @param {Object} obj - JavaScript object (from parse() or manually constructed).
   * @param {string} [xmlDeclaration] - Optional XML declaration to prepend.
   * @returns {string} XML string.
   *
   * @example
   * const xml = parser.build(obj, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
   */
  build(obj, xmlDeclaration = '') {
    try {
      const spTreeObj =
        obj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
        obj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
        obj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

      let xml = this.#builder.build(obj)

      if (spTreeObj) {
        const correctSpTreeXml = this.serializeContainer(spTreeObj, 'p:spTree')
        if (xml.includes('<p:spTree/>')) {
          xml = xml.replace('<p:spTree/>', correctSpTreeXml)
        } else {
          xml = xml.replace(/<p:spTree>[\s\S]*<\/p:spTree>/, correctSpTreeXml)
        }
      }

      return xmlDeclaration ? `${xmlDeclaration}\n${xml}` : xml
    } catch (err) {
      throw new PPTXError(`XML build error: ${err.message}`, err)
    }
  }

  /**
   * Helper to serialize drawing containers (p:spTree, p:grpSp) recursively in Z-order.
   */
  serializeContainer(container, containerTagName) {
    const zOrder = container[Z_ORDER_SYMBOL] || []

    let headerXml = ''
    if (container['p:nvGrpSpPr']) {
      headerXml += this.#builder.build({ 'p:nvGrpSpPr': container['p:nvGrpSpPr'] })
    }
    if (container['p:grpSpPr']) {
      headerXml += this.#builder.build({ 'p:grpSpPr': container['p:grpSpPr'] })
    }

    // Gather drawing children
    const drawingElements = new Map()
    for (const tag of ['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp']) {
      let items = container[tag] || []
      if (!Array.isArray(items)) items = [items]
      for (const item of items) {
        let id = null
        if (tag === 'p:sp') id = item?.['p:nvSpPr']?.['p:cNvPr']?.['@_id']
        else if (tag === 'p:pic') id = item?.['p:nvPicPr']?.['p:cNvPr']?.['@_id']
        else if (tag === 'p:graphicFrame') id = item?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_id']
        else if (tag === 'p:grpSp') id = item?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_id']
        else if (tag === 'p:cxnSp') id = item?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_id']

        if (id !== undefined && id !== null) {
          drawingElements.set(String(id), { tag, obj: item })
        }
      }
    }

    // Append items not in the explicit Z-order list
    const fullZOrder = [...zOrder]
    for (const id of drawingElements.keys()) {
      if (!fullZOrder.includes(id)) {
        fullZOrder.push(id)
      }
    }

    // Build children in Z-order
    let childrenXml = ''
    for (const id of fullZOrder) {
      const el = drawingElements.get(id)
      if (!el) continue

      if (el.tag === 'p:grpSp') {
        childrenXml += this.serializeContainer(el.obj, 'p:grpSp')
      } else {
        childrenXml += this.#builder.build({ [el.tag]: el.obj })
      }
    }

    // Container attributes
    let attrsStr = ''
    const attrs = {}
    for (const k in container) {
      if (k.startsWith('@_')) {
        attrs[k] = container[k]
      }
    }
    if (Object.keys(attrs).length > 0) {
      const attrXml = this.#builder.build({ [containerTagName]: { ...attrs } })
      const match = attrXml.match(/<[^>]+>/)
      if (match) {
        attrsStr = match[0].slice(containerTagName.length + 1, -1)
      }
    }

    return `<${containerTagName}${attrsStr}>${headerXml}${childrenXml}</${containerTagName}>`
  }

  /**
   * Extracts the XML declaration line from an XML string.
   *
   * @param {string} xmlString - Raw XML string.
   * @returns {string} Declaration line or empty string.
   */
  extractDeclaration(xmlString) {
    const match = xmlString.match(/^<\?xml[^>]+\?>/)
    return match ? match[0] : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
  }

  /**
   * Performs a deep clone of a parsed XML object.
   * Used when copying slides to avoid shared object references.
   *
   * @param {Object} obj - Object to clone.
   * @returns {Object} Deep clone.
   */
  deepClone(obj) {
    return JSON.parse(JSON.stringify(obj))
  }

  /**
   * Finds all nodes matching a key path in a parsed XML object.
   * Uses a simple dot-notation path (e.g., 'p:sp.p:txBody.a:p.a:r').
   *
   * @param {Object} obj - Root object to search.
   * @param {string} path - Dot-separated key path.
   * @returns {Array} Array of matching nodes.
   *
   * @example
   * const runs = parser.findAll(slideObj, 'p:cSld.p:spTree.p:sp.p:txBody.a:p.a:r');
   */
  findAll(obj, path) {
    const keys = path.split('.')
    let current = [obj]

    for (const key of keys) {
      const next = []
      for (const node of current) {
        if (node && typeof node === 'object') {
          const val = node[key]
          if (Array.isArray(val)) {
            next.push(...val)
          } else if (val !== undefined) {
            next.push(val)
          }
        }
      }
      current = next
    }

    return current
  }

  /**
   * Gets a single node by path (returns first match).
   *
   * @param {Object} obj - Root object to search.
   * @param {string} path - Dot-separated key path.
   * @returns {*} First matching node or undefined.
   */
  getNode(obj, path) {
    return this.findAll(obj, path)[0]
  }

  /**
   * Sets a value at a dot-notation path in an object, creating intermediate
   * objects as needed.
   *
   * @param {Object} obj - Root object.
   * @param {string} path - Dot-separated key path.
   * @param {*} value - Value to set.
   */
  setNode(obj, path, value) {
    const keys = path.split('.')
    let current = obj

    for (let i = 0; i < keys.length - 1; i++) {
      const key = keys[i]
      if (!current[key] || typeof current[key] !== 'object') {
        current[key] = {}
      }
      current = Array.isArray(current[key]) ? current[key][0] : current[key]
    }

    current[keys[keys.length - 1]] = value
  }

  /**
   * Performs a string replacement directly on the raw XML string.
   * Faster than parse → modify → build for simple text replacements.
   *
   * @param {string} xmlString - Raw XML.
   * @param {string} search - Substring to find.
   * @param {string} replace - Replacement string.
   * @param {boolean} [all=true] - Replace all occurrences or just first.
   * @returns {string} Modified XML string.
   */
  replaceInXml(xmlString, search, replace, all = true) {
    const escaped = search.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
    const flags = all ? 'g' : ''
    return xmlString.replace(new RegExp(escaped, flags), replace)
  }

  /**
   * Extracts all text content from a slide XML string.
   * Useful for debugging or searching slide content.
   *
   * @param {string} xmlString - Slide XML content.
   * @returns {string[]} Array of text strings found in the slide.
   */
  extractTextContent(xmlString) {
    const texts = []
    const textPattern = /<a:t>([^<]*)<\/a:t>/g
    let match
    while ((match = textPattern.exec(xmlString)) !== null) {
      if (match[1].trim()) texts.push(match[1])
    }
    return texts
  }

  /**
   * Validates that an XML string is well-formed.
   *
   * @param {string} xmlString - XML to validate.
   * @returns {{ valid: boolean, error: string|null }}
   */
  validate(xmlString) {
    try {
      this.parse(xmlString)
      return { valid: true, error: null }
    } catch (err) {
      return { valid: false, error: err.message }
    }
  }
}

module.exports = {
  XMLParser,
  Z_ORDER_SYMBOL,
}
