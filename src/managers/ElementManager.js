/**
 * @fileoverview ElementManager - Unified inspection, query, removal, and lifecycle management for slide elements.
 */

'use strict'

const { PPTXError } = require('../utils/errors.js')
const { createLogger } = require('../utils/logger.js')
const { Z_ORDER_SYMBOL } = require('../parsers/XMLParser.js')
const { getBounds, checkCollision, isContained, toEmu } = require('../spatial/units.js')

const logger = createLogger('ElementManager')

const DRAWING_TAGS = ['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp']

function normalizeBounds(b) {
  if (!b || typeof b !== 'object') return b
  return {
    x: b.x !== undefined ? toEmu(b.x) : 0,
    y: b.y !== undefined ? toEmu(b.y) : 0,
    width: b.width !== undefined ? toEmu(b.width) : 0,
    height: b.height !== undefined ? toEmu(b.height) : 0,
  }
}

/**
 * Detects the normalized element type from its OpenXML node and tag.
 *
 * @param {string} tag - OpenXML tag name.
 * @param {Object} item - Element AST node.
 * @returns {string} Normalized element type: 'textbox' | 'shape' | 'image' | 'video' | 'audio' | 'table' | 'chart' | 'smartart' | 'group' | 'connector' | 'graphicFrame'
 */
function detectElementType(tag, item) {
  if (tag === 'p:sp') {
    const isTxBox =
      item?.['p:spPr']?.['@_txBox'] === '1' ||
      item?.['p:spPr']?.['@_txBox'] === true ||
      item?.['p:nvSpPr']?.['p:cNvSpPr']?.['@_txBox'] === '1' ||
      item?.['p:nvSpPr']?.['p:cNvSpPr']?.['@_txBox'] === true
    if (isTxBox) return 'textbox'

    const phType = item?.['p:nvSpPr']?.['p:nvPr']?.['p:ph']?.['@_type']
    if (phType === 'title' || phType === 'body' || phType === 'subTitle') {
      return 'textbox'
    }
    return 'shape'
  }

  if (tag === 'p:pic') {
    const nvPr = item?.['p:nvPicPr']?.['p:nvPr']
    const isVideo =
      Boolean(nvPr?.['a:videoFile']) ||
      Boolean(nvPr?.['p:extLst']?.['p:ext']?.['p14:media']) ||
      (Array.isArray(nvPr?.['p:extLst']?.['p:ext']) &&
        nvPr?.['p:extLst']?.['p:ext']?.some(e => e['p14:media']))
    if (isVideo) return 'video'

    const isAudio = Boolean(nvPr?.['a:audioFile'])
    if (isAudio) return 'audio'

    return 'image'
  }

  if (tag === 'p:graphicFrame') {
    if (item?.['a:graphic']?.['a:graphicData']?.['a:tbl']) return 'table'
    if (item?.['a:graphic']?.['a:graphicData']?.['c:chart']) return 'chart'
    const uri = item?.['a:graphic']?.['a:graphicData']?.['@_uri'] || ''
    if (uri.includes('table')) return 'table'
    if (uri.includes('chart')) return 'chart'
    if (uri.includes('diagram')) return 'smartart'
    return 'graphicFrame'
  }

  if (tag === 'p:grpSp') {
    return 'group'
  }

  if (tag === 'p:cxnSp') {
    return 'connector'
  }

  return 'unknown'
}

/**
 * Extracts element ID from OpenXML non-visual properties.
 */
function getElementId(tag, item) {
  if (tag === 'p:sp') return item?.['p:nvSpPr']?.['p:cNvPr']?.['@_id']
  if (tag === 'p:pic') return item?.['p:nvPicPr']?.['p:cNvPr']?.['@_id']
  if (tag === 'p:graphicFrame') return item?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_id']
  if (tag === 'p:grpSp') return item?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_id']
  if (tag === 'p:cxnSp') return item?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_id']
  return null
}

/**
 * Extracts element Name from OpenXML non-visual properties.
 */
function getElementName(tag, item) {
  if (tag === 'p:sp') return item?.['p:nvSpPr']?.['p:cNvPr']?.['@_name']
  if (tag === 'p:pic') return item?.['p:nvPicPr']?.['p:cNvPr']?.['@_name']
  if (tag === 'p:graphicFrame') return item?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_name']
  if (tag === 'p:grpSp') return item?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_name']
  if (tag === 'p:cxnSp') return item?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_name']
  return null
}

/**
 * Recursively extracts relationship IDs (e.g. 'rId2') referenced within an element AST.
 */
function extractRelationshipIds(node, acc = new Set()) {
  if (!node || typeof node !== 'object') return acc
  if (Array.isArray(node)) {
    for (const item of node) {
      extractRelationshipIds(item, acc)
    }
    return acc
  }

  for (const [key, val] of Object.entries(node)) {
    if (typeof val === 'string' && val.startsWith('rId')) {
      if (
        key.startsWith('@_r:') ||
        key === '@_rId' ||
        key === '@_id' ||
        key.toLowerCase().includes('rid')
      ) {
        acc.add(val)
      }
    } else if (typeof val === 'object') {
      extractRelationshipIds(val, acc)
    }
  }

  return acc
}

/**
 * Checks whether a given relationship ID is referenced anywhere in container outside targetElement.
 */
function isRelationshipReferencedElsewhere(container, targetElement, rId) {
  let referenced = false

  function walk(node) {
    if (!node || typeof node !== 'object' || referenced) return
    if (node === targetElement) return

    if (Array.isArray(node)) {
      for (const item of node) walk(item)
      return
    }

    for (const val of Object.values(node)) {
      if (typeof val === 'string' && val === rId) {
        referenced = true
        return
      }
      if (typeof val === 'object') walk(val)
    }
  }

  walk(container)
  return referenced
}

/**
 * @class ElementManager
 * @description Manages slide visual elements inspection, query, removal, and cleanup.
 */
class ElementManager {
  /** @private @type {SlideManager} */
  #slideManager
  /** @private @type {RelationshipManager} */
  #relationshipManager
  /** @private @type {ZipManager} */
  #zipManager
  /** @private @type {XMLParser} */
  #xmlParser

  constructor(slideManager, relationshipManager, zipManager, xmlParser) {
    this.#slideManager = slideManager
    this.#relationshipManager = relationshipManager
    this.#zipManager = zipManager
    this.#xmlParser = xmlParser
  }

  /**
   * Retrieves all visual elements on a slide with normalized metadata.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {Object} [filter={}] - Optional filter criteria (type, name, id, bounds, within, intersects).
   * @returns {Array<Object>} Normalized element descriptors.
   */
  getElements(slideIndex, filter = {}) {
    const slideObj = this.#slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    if (!spTree) return []

    // Stacking order
    const zOrder = spTree[Z_ORDER_SYMBOL] || []
    const elementsMap = new Map()

    for (const tag of DRAWING_TAGS) {
      let items = spTree[tag] || []
      if (!Array.isArray(items)) items = [items]

      for (const item of items) {
        const rawId = getElementId(tag, item)
        const id = rawId !== undefined && rawId !== null ? String(rawId) : ''
        const name = getElementName(tag, item) || ''
        const type = detectElementType(tag, item)
        const bounds = getBounds(item)

        const descriptor = {
          id,
          name,
          type,
          bounds,
          slideIndex,
          element: item,
          frame: tag === 'p:graphicFrame' ? item : null,
        }

        if (id) {
          elementsMap.set(id, descriptor)
        } else {
          // If no ID, generate a fallback key
          elementsMap.set(`elem_${elementsMap.size + 1}`, descriptor)
        }
      }
    }

    // Order elements according to zOrder sequence if present
    let orderedElements = []
    const visited = new Set()

    if (zOrder.length > 0) {
      for (const zid of zOrder) {
        const el = elementsMap.get(zid)
        if (el) {
          orderedElements.push(el)
          visited.add(zid)
        }
      }
    }

    for (const [key, el] of elementsMap.entries()) {
      if (!visited.has(key)) {
        orderedElements.push(el)
      }
    }

    // Apply filtering
    if (filter.type && filter.type !== '*' && filter.type !== 'all') {
      const targetType = String(filter.type).toLowerCase()
      orderedElements = orderedElements.filter(el => {
        if (targetType === 'image' && (el.type === 'image' || el.type === 'picture')) return true
        if (
          targetType === 'shape' &&
          (el.type === 'shape' || el.type === 'textbox' || el.type === 'connector')
        )
          return true
        if (targetType === 'text' && el.type === 'textbox') return true
        return el.type.toLowerCase() === targetType
      })
    }

    if (filter.name) {
      if (filter.name instanceof RegExp) {
        orderedElements = orderedElements.filter(el => filter.name.test(el.name))
      } else if (typeof filter.name === 'string') {
        orderedElements = orderedElements.filter(
          el => el.name === filter.name || el.id === filter.name
        )
      }
    }

    if (filter.id !== undefined) {
      const idStr = String(filter.id)
      orderedElements = orderedElements.filter(el => el.id === idStr)
    }

    if (filter.bounds || filter.within) {
      const outer = normalizeBounds(filter.bounds || filter.within)
      orderedElements = orderedElements.filter(el => isContained(el.bounds, outer))
    }

    if (filter.intersects || filter.overlaps) {
      const box = normalizeBounds(filter.intersects || filter.overlaps)
      orderedElements = orderedElements.filter(el => checkCollision(el.bounds, box))
    }

    return orderedElements
  }

  /**
   * Retrieves single normalized element descriptor by ID or name.
   *
   * @param {string|number} elementId - Element ID or name.
   * @param {number} slideIndex - 1-based slide index.
   * @returns {Object|null} Normalized element descriptor or null.
   */
  getElement(elementId, slideIndex) {
    const idStr = String(elementId)
    const elements = this.getElements(slideIndex)
    return elements.find(el => el.id === idStr || el.name === idStr) || null
  }

  /**
   * Safely purges an element from the slide and unlinks orphaned relationships.
   *
   * @param {string|number} elementId - Element ID or name to remove.
   * @param {number} slideIndex - 1-based slide index.
   * @returns {boolean} True if element was removed.
   */
  removeElement(elementId, slideIndex) {
    const slideInfo = this.#slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) {
      throw new PPTXError(`Slide ${slideIndex} does not exist`)
    }

    const slideObj = this.#slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    if (!spTree) return false

    const idStr = String(elementId)

    // Locate the element and parent container
    let found = null
    for (const tag of DRAWING_TAGS) {
      let items = spTree[tag] || []
      if (!Array.isArray(items)) items = [items]

      for (let i = 0; i < items.length; i++) {
        const item = items[i]
        const rawId = getElementId(tag, item)
        const name = getElementName(tag, item)
        if (String(rawId) === idStr || name === idStr) {
          found = { tag, index: i, item, parent: spTree, id: rawId ? String(rawId) : null }
          break
        }
      }
      if (found) break
    }

    if (!found) {
      logger.warn(`Element "${elementId}" not found on slide ${slideIndex}`)
      return false
    }

    const { tag, item, id } = found

    // Extract all relationship IDs used by this element
    const relIds = extractRelationshipIds(item)

    // Remove element node from container array
    if (Array.isArray(spTree[tag])) {
      spTree[tag] = spTree[tag].filter(x => x !== item)
    } else if (spTree[tag] === item) {
      delete spTree[tag]
    }

    // Remove from Z-order array if present
    if (id && Array.isArray(spTree[Z_ORDER_SYMBOL])) {
      spTree[Z_ORDER_SYMBOL] = spTree[Z_ORDER_SYMBOL].filter(zid => zid !== id)
    }

    // Clean orphaned relationships in slideN.xml.rels
    if (relIds.size > 0 && this.#relationshipManager) {
      for (const rId of relIds) {
        if (!isRelationshipReferencedElsewhere(spTree, null, rId)) {
          try {
            this.#relationshipManager.removeRelationship(slideInfo.zipPath, rId)
            logger.debug(`Unlinked orphaned relationship "${rId}" from ${slideInfo.zipPath}`)
          } catch (err) {
            logger.debug(`Failed to unlink relationship ${rId}: ${err.message}`)
          }
        }
      }
    }

    // Clean timing condition if video or audio
    if (id && slideObj?.['p:sld']?.['p:timing']) {
      this.#cleanTimingForShape(slideObj, id)
    }

    this.#slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Removed element "${elementId}" from slide ${slideIndex}`)
    return true
  }

  /**
   * Clears all visual elements from the slide canvas while preserving layout bindings and group headers.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {Object} [options={}] - Options (preservePlaceholders).
   */
  removeAllElements(slideIndex, options = {}) {
    const slideInfo = this.#slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) {
      throw new PPTXError(`Slide ${slideIndex} does not exist`)
    }

    const slideObj = this.#slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    if (!spTree) return

    const preservePlaceholders = options.preservePlaceholders === true
    const allRemovedRelIds = new Set()

    for (const tag of DRAWING_TAGS) {
      let items = spTree[tag] || []
      if (!Array.isArray(items)) items = [items]

      if (tag === 'p:sp' && preservePlaceholders) {
        const kept = []
        for (const sp of items) {
          const ph = sp?.['p:nvSpPr']?.['p:nvPr']?.['p:ph']
          if (ph !== undefined) {
            kept.push(sp)
          } else {
            extractRelationshipIds(sp, allRemovedRelIds)
          }
        }
        spTree['p:sp'] = kept
      } else {
        for (const item of items) {
          extractRelationshipIds(item, allRemovedRelIds)
        }
        spTree[tag] = []
      }
    }

    // Reset Z-Order tracking
    if (preservePlaceholders && Array.isArray(spTree['p:sp'])) {
      const keptIds = []
      for (const sp of spTree['p:sp']) {
        const id = sp?.['p:nvSpPr']?.['p:cNvPr']?.['@_id']
        if (id) keptIds.push(String(id))
      }
      spTree[Z_ORDER_SYMBOL] = keptIds
    } else {
      spTree[Z_ORDER_SYMBOL] = []
    }

    // Clean orphaned relationships that were referenced by removed elements
    if (allRemovedRelIds.size > 0 && this.#relationshipManager) {
      for (const rId of allRemovedRelIds) {
        if (!isRelationshipReferencedElsewhere(spTree, null, rId)) {
          try {
            this.#relationshipManager.removeRelationship(slideInfo.zipPath, rId)
            logger.debug(`Unlinked relationship "${rId}" on removeAllElements`)
          } catch (err) {
            logger.debug(`Failed to unlink relationship ${rId}: ${err.message}`)
          }
        }
      }
    }

    // Remove timing node if no elements remain
    if (!preservePlaceholders && slideObj?.['p:sld']?.['p:timing']) {
      delete slideObj['p:sld']['p:timing']
    }

    this.#slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Cleared all elements from slide ${slideIndex}`)
  }

  /**
   * Cleans any timing sequence referencing a specific shape ID.
   * @private
   */
  #cleanTimingForShape(slideObj, shapeId) {
    const idStr = String(shapeId)
    const timing = slideObj?.['p:sld']?.['p:timing']
    if (!timing) return

    function purgeInNode(node) {
      if (!node || typeof node !== 'object') return
      if (Array.isArray(node)) {
        for (let i = node.length - 1; i >= 0; i--) {
          const item = node[i]
          const spId =
            item?.['p:cMediaNode']?.['p:cTn']?.['p:stCondLst']?.['p:cond']?.['p:tgtEl']?.[
              'p:spTgt'
            ]?.['@_spid'] || item?.['@_spid']
          if (String(spId) === idStr) {
            node.splice(i, 1)
          } else {
            purgeInNode(item)
          }
        }
        return
      }

      for (const key of Object.keys(node)) {
        purgeInNode(node[key])
      }
    }

    purgeInNode(timing)
  }
}

module.exports = { ElementManager }
