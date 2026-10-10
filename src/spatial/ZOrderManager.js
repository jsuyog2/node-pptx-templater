/**
 * @fileoverview ZOrderManager - Handles slide element Z-order (layer stacking) operations.
 */

const { PPTXError } = require('../utils/errors.js')
const { Z_ORDER_SYMBOL } = require('../parsers/XMLParser.js')

function detectElementType(tag, item) {
  if (tag === 'p:sp') {
    const isTxBox =
      item?.['p:nvSpPr']?.['p:cNvSpPr']?.['@_txBox'] === '1' ||
      item?.['p:nvSpPr']?.['p:cNvSpPr']?.['@_txBox'] === true
    const phType = item?.['p:nvSpPr']?.['p:nvPr']?.['p:ph']?.['@_type']
    if (isTxBox || phType === 'title' || phType === 'body' || item?.['p:txBody']) {
      return 'text'
    }
    return 'shape'
  }
  if (tag === 'p:pic') {
    return 'image'
  }
  if (tag === 'p:graphicFrame') {
    const uri = item?.['a:graphic']?.['a:graphicData']?.['@_uri'] || ''
    if (uri.includes('chart')) return 'chart'
    if (uri.includes('table')) return 'table'
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
 * @class ZOrderManager
 * @description Manages stacking layers and z-index of shapes on slides.
 */
class ZOrderManager {
  /** @private @type {XMLParser} */
  #xmlParser

  constructor(xmlParser) {
    this.#xmlParser = xmlParser
  }

  /**
   * Helper to parse Z-order array for a container or initialize it if missing.
   */
  getOrInitZOrder(container) {
    if (!container[Z_ORDER_SYMBOL]) {
      const list = []
      for (const tag of ['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp']) {
        let items = container[tag] || []
        if (!Array.isArray(items)) items = [items]
        for (const item of items) {
          let id = null
          if (tag === 'p:sp') id = item?.['p:nvSpPr']?.['p:cNvPr']?.['@_id']
          else if (tag === 'p:pic') id = item?.['p:nvPicPr']?.['p:cNvPr']?.['@_id']
          else if (tag === 'p:graphicFrame')
            id = item?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_id']
          else if (tag === 'p:grpSp') id = item?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_id']
          else if (tag === 'p:cxnSp') id = item?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_id']

          if (id !== undefined && id !== null) {
            list.push(String(id))
          }
        }
      }
      container[Z_ORDER_SYMBOL] = list
    }
    return container[Z_ORDER_SYMBOL]
  }

  /**
   * Helper to find drawing element and its parent container.
   */
  findObjectByIdOrName(container, targetId) {
    if (!container) return null

    for (const tag of ['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp']) {
      let items = container[tag] || []
      if (!Array.isArray(items)) items = [items]
      for (const item of items) {
        let id = null
        let name = null
        if (tag === 'p:sp') {
          id = item?.['p:nvSpPr']?.['p:cNvPr']?.['@_id']
          name = item?.['p:nvSpPr']?.['p:cNvPr']?.['@_name']
        } else if (tag === 'p:pic') {
          id = item?.['p:nvPicPr']?.['p:cNvPr']?.['@_id']
          name = item?.['p:nvPicPr']?.['p:cNvPr']?.['@_name']
        } else if (tag === 'p:graphicFrame') {
          id = item?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_id']
          name = item?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_name']
        } else if (tag === 'p:grpSp') {
          id = item?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_id']
          name = item?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_name']
        } else if (tag === 'p:cxnSp') {
          id = item?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_id']
          name = item?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_name']
        }

        if (String(id) === String(targetId) || name === targetId) {
          return { tag, obj: item, id: String(id), name, parent: container }
        }

        if (tag === 'p:grpSp') {
          const res = this.findObjectByIdOrName(item, targetId)
          if (res) return res
        }
      }
    }
    return null
  }

  /**
   * Retrieves the Z-order sequence of objects on a slide.
   *
   * @param {number} slideIndex
   * @param {SlideManager} slideManager
   * @returns {Array<Object>} List of object metadata in stacking order.
   */
  getObjectOrder(slideIndex, slideManager) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    if (!spTree) return []

    const zOrder = this.getOrInitZOrder(spTree)

    const drawingElements = new Map()
    for (const tag of ['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp']) {
      let items = spTree[tag] || []
      if (!Array.isArray(items)) items = [items]
      for (const item of items) {
        let id = null
        let name = null
        if (tag === 'p:sp') {
          id = item?.['p:nvSpPr']?.['p:cNvPr']?.['@_id']
          name = item?.['p:nvSpPr']?.['p:cNvPr']?.['@_name']
        } else if (tag === 'p:pic') {
          id = item?.['p:nvPicPr']?.['p:cNvPr']?.['@_id']
          name = item?.['p:nvPicPr']?.['p:cNvPr']?.['@_name']
        } else if (tag === 'p:graphicFrame') {
          id = item?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_id']
          name = item?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_name']
        } else if (tag === 'p:grpSp') {
          id = item?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_id']
          name = item?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_name']
        } else if (tag === 'p:cxnSp') {
          id = item?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_id']
          name = item?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_name']
        }

        if (id !== undefined && id !== null) {
          drawingElements.set(String(id), { tag, obj: item, name })
        }
      }
    }

    const fullZOrder = [...zOrder]
    for (const id of drawingElements.keys()) {
      if (!fullZOrder.includes(id)) {
        fullZOrder.push(id)
      }
    }

    const result = []
    let zIndex = 1
    for (const id of fullZOrder) {
      const el = drawingElements.get(id)
      if (!el) continue

      result.push({
        id: el.name || id,
        type: detectElementType(el.tag, el.obj),
        zIndex: zIndex++,
      })
    }

    return result
  }

  /**
   * Core reordering orchestrator.
   * Runs the reorder callback on the correct container and saves the slide XML.
   */
  #modifyZOrder(slideIndex, objectId, slideManager, callback) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    if (!spTree) {
      throw new PPTXError(`Invalid slide structure for slide ${slideIndex}`)
    }

    const res = this.findObjectByIdOrName(spTree, objectId)
    if (!res) {
      throw new PPTXError(`Object "${objectId}" not found on slide ${slideIndex}`)
    }

    const container = res.parent
    const zOrder = this.getOrInitZOrder(container)

    callback(zOrder, res.id, spTree)

    const decl = this.#xmlParser.extractDeclaration(slideXml)
    slideManager.setSlideXml(slideIndex, this.#xmlParser.build(slideObj, decl))
  }

  bringForward(slideIndex, objectId, slideManager) {
    this.#modifyZOrder(slideIndex, objectId, slideManager, (zOrder, elementId) => {
      const idx = zOrder.indexOf(elementId)
      if (idx !== -1 && idx < zOrder.length - 1) {
        zOrder[idx] = zOrder[idx + 1]
        zOrder[idx + 1] = elementId
      }
    })
  }

  sendBackward(slideIndex, objectId, slideManager) {
    this.#modifyZOrder(slideIndex, objectId, slideManager, (zOrder, elementId) => {
      const idx = zOrder.indexOf(elementId)
      if (idx > 0) {
        zOrder[idx] = zOrder[idx - 1]
        zOrder[idx - 1] = elementId
      }
    })
  }

  bringToFront(slideIndex, objectId, slideManager) {
    this.#modifyZOrder(slideIndex, objectId, slideManager, (zOrder, elementId) => {
      const idx = zOrder.indexOf(elementId)
      if (idx !== -1) {
        zOrder.splice(idx, 1)
        zOrder.push(elementId)
      }
    })
  }

  sendToBack(slideIndex, objectId, slideManager) {
    this.#modifyZOrder(slideIndex, objectId, slideManager, (zOrder, elementId) => {
      const idx = zOrder.indexOf(elementId)
      if (idx !== -1) {
        zOrder.splice(idx, 1)
        zOrder.unshift(elementId)
      }
    })
  }

  setZIndex(slideIndex, objectId, zIndex, slideManager) {
    this.#modifyZOrder(slideIndex, objectId, slideManager, (zOrder, elementId) => {
      const idx = zOrder.indexOf(elementId)
      if (idx !== -1) {
        zOrder.splice(idx, 1)
        const targetIdx = Math.max(0, Math.min(zIndex - 1, zOrder.length))
        zOrder.splice(targetIdx, 0, elementId)
      }
    })
  }

  moveObjectBefore(slideIndex, objectId, targetId, slideManager) {
    this.#modifyZOrder(slideIndex, objectId, slideManager, (zOrder, elementId, spTree) => {
      const targetRes = this.findObjectByIdOrName(spTree, targetId)
      if (!targetRes) {
        throw new PPTXError(`Target object "${targetId}" not found on slide ${slideIndex}`)
      }
      if (targetRes.parent !== targetRes.parent) {
        throw new PPTXError('Cannot move elements across different group containers')
      }

      const idx = zOrder.indexOf(elementId)
      if (idx !== -1) {
        zOrder.splice(idx, 1)
      }
      const targetIdx = zOrder.indexOf(targetRes.id)
      if (targetIdx !== -1) {
        zOrder.splice(targetIdx, 0, elementId)
      } else {
        zOrder.push(elementId)
      }
    })
  }

  moveObjectAfter(slideIndex, objectId, targetId, slideManager) {
    this.#modifyZOrder(slideIndex, objectId, slideManager, (zOrder, elementId, spTree) => {
      const targetRes = this.findObjectByIdOrName(spTree, targetId)
      if (!targetRes) {
        throw new PPTXError(`Target object "${targetId}" not found on slide ${slideIndex}`)
      }

      const idx = zOrder.indexOf(elementId)
      if (idx !== -1) {
        zOrder.splice(idx, 1)
      }
      const targetIdx = zOrder.indexOf(targetRes.id)
      if (targetIdx !== -1) {
        zOrder.splice(targetIdx + 1, 0, elementId)
      } else {
        zOrder.push(elementId)
      }
    })
  }

  reorderObjects(slideIndex, order, slideManager) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    if (!spTree) return

    const zOrder = this.getOrInitZOrder(spTree)

    // Resolve ordered names/IDs to existing drawing IDs
    const resolvedIds = []
    for (const item of order) {
      const res = this.findObjectByIdOrName(spTree, item)
      if (res && res.parent === spTree) {
        resolvedIds.push(res.id)
      }
    }

    // Keep track of unspecified IDs
    const unspecifiedIds = zOrder.filter(id => !resolvedIds.includes(id))

    // Reconstruct the z-order: unspecified bottom, specified top
    spTree[Z_ORDER_SYMBOL] = [...unspecifiedIds, ...resolvedIds]

    const decl = this.#xmlParser.extractDeclaration(slideXml)
    slideManager.setSlideXml(slideIndex, this.#xmlParser.build(slideObj, decl))
  }

  applyZOrder(slideIndex, configs, slideManager) {
    if (!Array.isArray(configs)) return

    for (const config of configs) {
      if (!config.id) continue

      if (config.bringToFront) {
        this.bringToFront(slideIndex, config.id, slideManager)
      } else if (config.sendToBack) {
        this.sendToBack(slideIndex, config.id, slideManager)
      } else if (config.bringForward) {
        this.bringForward(slideIndex, config.id, slideManager)
      } else if (config.sendBackward) {
        this.sendBackward(slideIndex, config.id, slideManager)
      } else if (config.zIndex !== undefined) {
        this.setZIndex(slideIndex, config.id, config.zIndex, slideManager)
      }
    }
  }

  // Layer Utilities
  getTopMostObject(slideIndex, slideManager) {
    const order = this.getObjectOrder(slideIndex, slideManager)
    return order.length > 0 ? order[order.length - 1] : null
  }

  getBottomMostObject(slideIndex, slideManager) {
    const order = this.getObjectOrder(slideIndex, slideManager)
    return order.length > 0 ? order[0] : null
  }

  swapObjects(slideIndex, objectId1, objectId2, slideManager) {
    this.#modifyZOrder(slideIndex, objectId1, slideManager, (zOrder, elementId1, spTree) => {
      const res2 = this.findObjectByIdOrName(spTree, objectId2)
      if (!res2) {
        throw new PPTXError(`Object "${objectId2}" not found on slide ${slideIndex}`)
      }
      const elementId2 = res2.id
      const idx1 = zOrder.indexOf(elementId1)
      const idx2 = zOrder.indexOf(elementId2)
      if (idx1 !== -1 && idx2 !== -1) {
        zOrder[idx1] = elementId2
        zOrder[idx2] = elementId1
      }
    })
  }

  sortObjects(slideIndex, compareFn, slideManager) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    if (!spTree) return

    const zOrder = this.getOrInitZOrder(spTree)

    // Build complete info list
    const order = this.getObjectOrder(slideIndex, slideManager)

    // Sort order using compareFn
    order.sort(compareFn)

    // Map sorted IDs back
    const sortedIds = []
    for (const item of order) {
      // Find ID by name or matching ID in container
      const res = this.findObjectByIdOrName(spTree, item.id)
      if (res && res.parent === spTree) {
        sortedIds.push(res.id)
      }
    }

    // Preserve unspecified ones at bottom
    const unspecifiedIds = zOrder.filter(id => !sortedIds.includes(id))
    spTree[Z_ORDER_SYMBOL] = [...unspecifiedIds, ...sortedIds]

    const decl = this.#xmlParser.extractDeclaration(slideXml)
    slideManager.setSlideXml(slideIndex, this.#xmlParser.build(slideObj, decl))
  }

  normalizeZOrder(slideIndex, slideManager) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    if (!spTree) return

    // Re-initialize and clean up
    delete spTree[Z_ORDER_SYMBOL]
    this.getOrInitZOrder(spTree)

    const decl = this.#xmlParser.extractDeclaration(slideXml)
    slideManager.setSlideXml(slideIndex, this.#xmlParser.build(slideObj, decl))
  }
}

module.exports = { ZOrderManager }
