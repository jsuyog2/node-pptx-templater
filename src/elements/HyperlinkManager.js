/**
 * @fileoverview HyperlinkManager - Manages OpenXML PresentationML hyperlinks in PPTX slides.
 *
 * Fully compliant with ISO/IEC 29500 (ECMA-376) PresentationML & DrawingML specifications.
 * Supports:
 *  - External URLs (http, https, mailto, ftp)
 *  - Internal slide jumps (ppaction://hlinksldjump with slide relationship target)
 *  - Interactive slideshow actions (ppaction://hlinkshowjump?s=nextslide, etc.)
 *  - Shape-level hyperlinks (<p:cNvPr><a:hlinkClick>)
 *  - Substring run-level hyperlinks (<a:rPr><a:hlinkClick>) with run splitting & typography styling
 *  - Read-only hyperlink introspection (getHyperlinks, getHyperlink)
 *  - Automatic orphaned relationship purging
 */

const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')
const { sanitizeUrl } = require('../utils/xmlUtils.js')
const { REL_TYPES } = require('../managers/RelationshipManager.js')
const { StyleBuilder } = require('../parsers/StyleBuilder.js')

const logger = createLogger('HyperlinkManager')

/**
 * Normalizes navigation action shortcuts to OpenXML ppaction URI strings.
 * Conforms strictly to PowerPoint / ISO/IEC 29500 standard query strings (?jump=...).
 *
 * @param {string} action
 * @returns {string}
 */
function normalizeAction(action) {
  if (!action || typeof action !== 'string') return ''
  const trimmed = action.trim()
  const lower = trimmed.toLowerCase()
  switch (lower) {
    case 'next':
    case 'nextslide':
      return 'ppaction://hlinkshowjump?jump=nextslide'
    case 'previous':
    case 'prev':
    case 'prevslide':
    case 'previousslide':
      return 'ppaction://hlinkshowjump?jump=previousslide'
    case 'first':
    case 'firstslide':
      return 'ppaction://hlinkshowjump?jump=firstslide'
    case 'last':
    case 'lastslide':
      return 'ppaction://hlinkshowjump?jump=lastslide'
    case 'end':
    case 'endshow':
      return 'ppaction://hlinkshowjump?jump=endshow'
    default:
      if (lower.startsWith('ppaction://hlinkshowjump?s=')) {
        return trimmed.replace(/\?s=/i, '?jump=')
      }
      if (lower.startsWith('ppaction://')) return trimmed
      return `ppaction://${trimmed}`
  }
}

/**
 * Checks if a URL or action string is a slideshow action (e.g. next, prev, ppaction://hlinkshowjump).
 *
 * @param {string} val
 * @returns {boolean}
 */
function isShowJumpAction(val) {
  if (!val || typeof val !== 'string') return false
  const lower = val.trim().toLowerCase()
  return (
    lower.startsWith('ppaction://hlinkshowjump') ||
    [
      'next',
      'nextslide',
      'prev',
      'previous',
      'prevslide',
      'previousslide',
      'first',
      'firstslide',
      'last',
      'lastslide',
      'end',
      'endshow',
    ].includes(lower)
  )
}

/**
 * Ensures strict PresentationML sequence of child elements inside <p:cNvPr>:
 * 1. Attributes (@_id, @_name, etc.)
 * 2. <a:hlinkClick>
 * 3. <a:hlinkHover>
 * 4. Other children
 * 5. <a:extLst>
 *
 * @param {Object} cNvPr
 * @returns {Object}
 */
function orderCnvPrChildren(cNvPr) {
  if (!cNvPr || typeof cNvPr !== 'object') return cNvPr

  const ordered = {}
  // 1. Attributes
  for (const key of Object.keys(cNvPr)) {
    if (key.startsWith('@_')) {
      ordered[key] = cNvPr[key]
    }
  }

  // 2. a:hlinkClick
  if (cNvPr['a:hlinkClick']) {
    ordered['a:hlinkClick'] = cNvPr['a:hlinkClick']
  }

  // 3. a:hlinkHover
  if (cNvPr['a:hlinkHover']) {
    ordered['a:hlinkHover'] = cNvPr['a:hlinkHover']
  }

  // 4. Other child elements except extLst
  for (const key of Object.keys(cNvPr)) {
    if (
      !key.startsWith('@_') &&
      key !== 'a:hlinkClick' &&
      key !== 'a:hlinkHover' &&
      key !== 'a:extLst'
    ) {
      ordered[key] = cNvPr[key]
    }
  }

  // 5. a:extLst
  if (cNvPr['a:extLst']) {
    ordered['a:extLst'] = cNvPr['a:extLst']
  }

  // Clean in-place keys and re-assign
  for (const key of Object.keys(cNvPr)) {
    delete cNvPr[key]
  }
  Object.assign(cNvPr, ordered)
  return cNvPr
}

/**
 * Sanitizes a color value to a valid DrawingML 6-character hex string without leading '#'.
 * Supports strings ('#F43F5E', 'F43F5E', '#F00'), numbers, and standard color names.
 *
 * @param {string|number|null|undefined} hex - Input color representation.
 * @returns {string|null} 6-character uppercase hex string (e.g. 'F43F5E'), or null if empty/invalid.
 */
function sanitizeHexColor(hex) {
  if (hex === undefined || hex === null || hex === '') return null
  if (typeof hex === 'number') {
    return Math.max(0, Math.floor(hex)).toString(16).padStart(6, '0').slice(-6).toUpperCase()
  }
  if (typeof hex !== 'string') return null
  let clean = hex.replace(/^#+/, '').trim()
  if (clean.length === 3 && /^[0-9A-Fa-f]{3}$/.test(clean)) {
    clean = clean
      .split('')
      .map(c => c + c)
      .join('')
  }
  if (/^[0-9A-Fa-f]{6}$/.test(clean)) {
    return clean.toUpperCase()
  }
  const fallback = StyleBuilder.normalizeHexColor(hex)
  return fallback !== '000000' || hex.toLowerCase() === 'black'
    ? fallback
    : clean.slice(0, 6).toUpperCase()
}

/**
 * Applies text formatting styles (color, underline, bold, italic, fontSize, fontFace)
 * to all text runs within a shape's txBody if style properties are passed.
 *
 * @param {Object} shape - Target shape element.
 * @param {Object} style - Styling options.
 */
function applyTextStyleToShape(shape, style) {
  if (!shape || !shape['p:txBody'] || !style || typeof style !== 'object') return
  const flatStyle =
    style.style && typeof style.style === 'object' ? { ...style.style, ...style } : style
  const hasStyle =
    flatStyle.color !== undefined ||
    flatStyle.fontColor !== undefined ||
    flatStyle.textColor !== undefined ||
    flatStyle.underline !== undefined ||
    flatStyle.bold !== undefined ||
    flatStyle.italic !== undefined ||
    flatStyle.fontSize !== undefined ||
    flatStyle.fontFace !== undefined ||
    flatStyle.fontFamily !== undefined ||
    flatStyle.highlight !== undefined ||
    flatStyle.strikethrough !== undefined ||
    flatStyle.strike !== undefined

  if (!hasStyle) return

  let paras = shape['p:txBody']['a:p'] || []
  if (!Array.isArray(paras)) {
    paras = [paras]
    shape['p:txBody']['a:p'] = paras
  }

  for (const p of paras) {
    let runs = p['a:r'] || []
    if (!Array.isArray(runs)) {
      runs = [runs]
      p['a:r'] = runs
    }
    for (const r of runs) {
      r['a:rPr'] = StyleBuilder.buildRunProperties(r['a:rPr'] || {}, flatStyle)
    }
  }
}

/**
 * Extracts non-visual drawing properties (<p:cNvPr>) from any slide element.
 *
 * @param {Object} element - Shape, picture, group, or graphicFrame AST node.
 * @returns {Object|null}
 */
function getCnvPr(element) {
  if (!element || typeof element !== 'object') return null
  return (
    element['p:nvSpPr']?.['p:cNvPr'] ||
    element['p:nvPicPr']?.['p:cNvPr'] ||
    element['p:nvGrpSpPr']?.['p:cNvPr'] ||
    element['p:nvGraphicFramePr']?.['p:cNvPr'] ||
    element['p:nvCxnSpPr']?.['p:cNvPr'] ||
    element['p:cNvPr'] ||
    null
  )
}

/**
 * Recursively locates a drawable element inside an spTree container by ID or name.
 *
 * @param {Object} container - spTree or grpSp AST node.
 * @param {string|number} targetId - Shape ID or name.
 * @returns {{ element: Object, cNvPr: Object, type: string, parent: Object }|null}
 */
function findElementInContainer(container, targetId) {
  if (!container || typeof container !== 'object') return null
  const targetStr = String(targetId).trim().toLowerCase()

  const tags = [
    { tag: 'p:sp', type: 'shape' },
    { tag: 'p:pic', type: 'image' },
    { tag: 'p:graphicFrame', type: 'graphicFrame' },
    { tag: 'p:grpSp', type: 'group' },
    { tag: 'p:cxnSp', type: 'connector' },
  ]

  for (const { tag, type } of tags) {
    let items = container[tag] || []
    if (!Array.isArray(items)) items = [items]
    for (const item of items) {
      const cNvPr = getCnvPr(item)
      if (cNvPr) {
        const id = String(cNvPr['@_id'] || '')
          .trim()
          .toLowerCase()
        const name = String(cNvPr['@_name'] || '')
          .trim()
          .toLowerCase()
        if (targetStr === 'first' || id === targetStr || name === targetStr) {
          return { element: item, cNvPr, type, parent: container }
        }
      }
      if (tag === 'p:grpSp') {
        const found = findElementInContainer(item, targetId)
        if (found) return found
      }
    }
  }

  return null
}

/**
 * Recursively checks if an rId is referenced anywhere in an AST node.
 *
 * @param {Object} node
 * @param {string} rId
 * @returns {boolean}
 */
function isRIdReferencedInAst(node, rId) {
  if (!node || typeof node !== 'object') return false
  if (node['@_r:id'] === rId || node['@_id'] === rId) return true

  for (const key of Object.keys(node)) {
    const val = node[key]
    if (Array.isArray(val)) {
      for (const item of val) {
        if (isRIdReferencedInAst(item, rId)) return true
      }
    } else if (typeof val === 'object') {
      if (isRIdReferencedInAst(val, rId)) return true
    }
  }
  return false
}

/**
 * @class HyperlinkManager
 * @description Manages creation, modification, removal, and introspection of hyperlinks in PPTX slides.
 */
class HyperlinkManager {
  /** @private @type {XMLParser} */
  #xmlParser
  /** @private @type {RelationshipManager} */
  #relationshipManager

  /**
   * @param {XMLParser} xmlParser
   * @param {RelationshipManager} relationshipManager
   */
  constructor(xmlParser, relationshipManager) {
    this.#xmlParser = xmlParser
    this.#relationshipManager = relationshipManager
  }

  // ==========================================================================
  // 1. addHyperlink(slideIndex, targetId, url, options, slideManager, relationshipManager)
  // ==========================================================================
  /**
   * Binds a hyperlink to an entire shape, textbox, image, or container.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string|number} targetId - ID or name of the target shape/element.
   * @param {string|number|Object} url - Destination URL, action string, target slide number, or options bag.
   * @param {Object} [options={}] - Additional options (tooltip, action, targetSlide, styling).
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   */
  addHyperlink(slideIndex, targetId, url, options = {}, slideManager, relationshipManager) {
    let effectiveUrl = url
    let effectiveOptions = options || {}
    if (typeof url === 'object' && url !== null) {
      effectiveOptions = { ...url, ...effectiveOptions }
      effectiveUrl =
        effectiveOptions.url !== undefined
          ? effectiveOptions.url
          : effectiveOptions.action !== undefined
            ? effectiveOptions.action
            : effectiveOptions.targetSlide
    }

    const relMgr = relationshipManager || this.#relationshipManager
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) {
      throw new PPTXError(`Slide ${slideIndex} not found`)
    }

    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    if (!spTree) {
      throw new PPTXError(`Slide ${slideIndex} has no valid spTree structure`)
    }

    const found = findElementInContainer(spTree, targetId)
    if (!found || !found.cNvPr) {
      throw new PPTXError(`Target element "${targetId}" not found on slide ${slideIndex}`)
    }

    const cNvPr = found.cNvPr

    // Clean up existing hyperlink relationship if replacing
    if (cNvPr['a:hlinkClick'] && cNvPr['a:hlinkClick']['@_r:id']) {
      const oldRId = cNvPr['a:hlinkClick']['@_r:id']
      delete cNvPr['a:hlinkClick']
      this.#purgeRelationshipIfOrphaned(
        slideIndex,
        oldRId,
        slideObj,
        slideInfo,
        slideManager,
        relMgr
      )
    }

    const { rId, action, tooltip } = this.#resolveLinkTarget(
      effectiveUrl,
      effectiveOptions,
      slideInfo,
      slideManager,
      relMgr
    )

    const hlinkClickObj = {}
    if (rId !== null && rId !== undefined) hlinkClickObj['@_r:id'] = rId
    if (action) hlinkClickObj['@_action'] = action
    if (tooltip) hlinkClickObj['@_tooltip'] = tooltip

    // Ensure OpenXML schema order inside <p:cNvPr>
    cNvPr['a:hlinkClick'] = hlinkClickObj
    orderCnvPrChildren(cNvPr)

    // Apply text styling if specified in options and target shape has text body
    applyTextStyleToShape(found.element, effectiveOptions)

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`addHyperlink: Attached hyperlink to target "${targetId}" on slide ${slideIndex}`)
  }

  // ==========================================================================
  // 2. updateHyperlink(slideIndex, targetId, url, options, slideManager, relationshipManager)
  // ==========================================================================
  /**
   * Updates an existing hyperlink on a target element.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string|number} targetId - ID or name of the target shape/element.
   * @param {string|number|Object} [url] - New destination URL, action, target slide number, or options bag.
   * @param {Object} [options={}] - Options (tooltip, action, targetSlide, styling).
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   */
  updateHyperlink(slideIndex, targetId, url, options = {}, slideManager, relationshipManager) {
    let effectiveUrl = url
    let effectiveOptions = options || {}
    if (typeof url === 'object' && url !== null) {
      effectiveOptions = { ...url, ...effectiveOptions }
      effectiveUrl =
        effectiveOptions.url !== undefined
          ? effectiveOptions.url
          : effectiveOptions.action !== undefined
            ? effectiveOptions.action
            : effectiveOptions.targetSlide
    }

    const relMgr = relationshipManager || this.#relationshipManager
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) {
      throw new PPTXError(`Slide ${slideIndex} not found`)
    }

    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    if (!spTree) {
      throw new PPTXError(`Slide ${slideIndex} has no valid spTree structure`)
    }

    const found = findElementInContainer(spTree, targetId)
    if (!found || !found.cNvPr) {
      throw new PPTXError(`Target element "${targetId}" not found on slide ${slideIndex}`)
    }

    const cNvPr = found.cNvPr
    const existingHlink = cNvPr['a:hlinkClick']

    // If no hyperlink existed previously, delegate to addHyperlink
    if (!existingHlink) {
      return this.addHyperlink(
        slideIndex,
        targetId,
        effectiveUrl,
        effectiveOptions,
        slideManager,
        relMgr
      )
    }

    // If URL or targetSlide is provided, re-resolve target and clean up old rel if needed
    if (
      effectiveUrl !== undefined ||
      effectiveOptions.targetSlide !== undefined ||
      effectiveOptions.action !== undefined
    ) {
      const oldRId = existingHlink['@_r:id']
      delete cNvPr['a:hlinkClick']
      if (oldRId) {
        this.#purgeRelationshipIfOrphaned(
          slideIndex,
          oldRId,
          slideObj,
          slideInfo,
          slideManager,
          relMgr
        )
      }

      const effectiveTarget =
        effectiveUrl !== undefined ? effectiveUrl : existingHlink['@_action'] || ''
      const {
        rId,
        action,
        tooltip: resolvedTip,
      } = this.#resolveLinkTarget(
        effectiveTarget,
        effectiveOptions,
        slideInfo,
        slideManager,
        relMgr
      )

      const hlinkClickObj = {}
      if (rId !== null && rId !== undefined) hlinkClickObj['@_r:id'] = rId
      if (action) hlinkClickObj['@_action'] = action
      const effectiveTip =
        effectiveOptions.tooltip !== undefined
          ? effectiveOptions.tooltip
          : resolvedTip || existingHlink['@_tooltip']
      if (effectiveTip) hlinkClickObj['@_tooltip'] = effectiveTip

      cNvPr['a:hlinkClick'] = hlinkClickObj
      orderCnvPrChildren(cNvPr)
    } else {
      // Just updating tooltip or properties
      if (effectiveOptions.tooltip !== undefined) {
        if (effectiveOptions.tooltip === null || effectiveOptions.tooltip === '') {
          delete existingHlink['@_tooltip']
        } else {
          existingHlink['@_tooltip'] = String(effectiveOptions.tooltip)
        }
      }
    }

    // Apply text styling if specified
    applyTextStyleToShape(found.element, effectiveOptions)

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(
      `updateHyperlink: Updated hyperlink on target "${targetId}" on slide ${slideIndex}`
    )
  }

  // ==========================================================================
  // 3. removeHyperlink(slideIndex, targetId, slideManager, relationshipManager)
  // ==========================================================================
  /**
   * Detaches a hyperlink from a shape, textbox, image, or text run, purging orphaned relationships.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string|number|Object} targetId - ID, shape name, or options object { targetId, shapeName, text }.
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   */
  removeHyperlink(slideIndex, targetId, slideManager, relationshipManager) {
    const relMgr = relationshipManager || this.#relationshipManager
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) return

    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    if (!spTree) return

    let targetKey = targetId
    let targetText = null

    if (typeof targetId === 'object' && targetId !== null) {
      targetKey = targetId.targetId || targetId.shapeId || targetId.shapeName || null
      targetText = targetId.text || null
    }

    const removedRIds = new Set()

    // 1. Try removing shape-level hyperlink
    if (targetKey) {
      const found = findElementInContainer(spTree, targetKey)
      if (found && found.cNvPr && found.cNvPr['a:hlinkClick']) {
        const rId = found.cNvPr['a:hlinkClick']['@_r:id']
        if (rId) removedRIds.add(rId)
        delete found.cNvPr['a:hlinkClick']
      }
    }

    // 2. Try removing text-level hyperlink from runs matching targetText or within targetKey shape
    const shapes = spTree['p:sp']
      ? Array.isArray(spTree['p:sp'])
        ? spTree['p:sp']
        : [spTree['p:sp']]
      : []
    for (const shape of shapes) {
      const cNvPr = getCnvPr(shape)
      const shapeId = cNvPr ? String(cNvPr['@_id'] || '').toLowerCase() : ''
      const shapeName = cNvPr ? String(cNvPr['@_name'] || '').toLowerCase() : ''
      const keyLower = targetKey ? String(targetKey).trim().toLowerCase() : null

      const isMatchingShape = !keyLower || shapeId === keyLower || shapeName === keyLower

      if (isMatchingShape && shape['p:txBody']) {
        let paras = shape['p:txBody']['a:p'] || []
        if (!Array.isArray(paras)) paras = [paras]
        for (const p of paras) {
          let runs = p['a:r'] || []
          if (!Array.isArray(runs)) runs = [runs]
          for (const r of runs) {
            const rPr = r['a:rPr']
            if (rPr && rPr['a:hlinkClick']) {
              const runText = String(r['a:t'] || '')
              if (
                !targetText ||
                runText.includes(targetText) ||
                (keyLower && !targetKey && runText.includes(keyLower))
              ) {
                const rId = rPr['a:hlinkClick']['@_r:id']
                if (rId) removedRIds.add(rId)
                delete rPr['a:hlinkClick']
              }
            }
          }
        }
      }
    }

    // Purge orphaned relationships
    for (const rId of removedRIds) {
      this.#purgeRelationshipIfOrphaned(slideIndex, rId, slideObj, slideInfo, slideManager, relMgr)
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(
      `removeHyperlink: Removed hyperlink for "${targetKey || targetText}" on slide ${slideIndex}`
    )
  }

  // ==========================================================================
  // 4. getHyperlinks(slideIndex, slideManager, relationshipManager)
  // ==========================================================================
  /**
   * Retrieves all hyperlinks (shape-level and textRun-level) defined on a slide.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @returns {Array<Object>} List of normalized hyperlink descriptor objects.
   */
  getHyperlinks(slideIndex, slideManager, relationshipManager) {
    const relMgr = relationshipManager || this.#relationshipManager
    let slideInfo = null
    try {
      slideInfo = slideManager.getSlideInfo(slideIndex)
    } catch {
      return []
    }
    if (!slideInfo) return []

    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    if (!spTree) return []

    const rels = relMgr.getRelationships(slideInfo.zipPath)
    const relMap = new Map()
    for (const rel of rels) {
      relMap.set(rel.id, rel)
    }

    const allSlides = slideManager.getAllSlideInfo()
    const resolveSlideIndex = target => {
      if (!target) return undefined
      const targetFilename = String(target).split('/').pop()
      const match = allSlides.find(s => s.zipPath && s.zipPath.endsWith(targetFilename))
      return match ? match.index : undefined
    }

    const results = []

    // Helper to extract hyperlink descriptor from an a:hlinkClick node
    const buildDescriptor = (hlinkClick, meta = {}) => {
      if (!hlinkClick) return null
      const rId = hlinkClick['@_r:id'] || hlinkClick['@_id'] || null
      const action = hlinkClick['@_action'] || null
      const tooltip = hlinkClick['@_tooltip'] || null

      const type = meta.type || 'shape'
      let isInternal = false
      let targetSlide = undefined
      let url = null

      const rel = rId ? relMap.get(rId) : null
      if (rel) {
        if (rel.type === REL_TYPES.SLIDE || (action && action.includes('hlinksldjump'))) {
          isInternal = true
          targetSlide = resolveSlideIndex(rel.target)
          url = null
        } else if (rel.type === REL_TYPES.HYPERLINK || rel.targetMode === 'External') {
          isInternal = false
          url = rel.target
        } else {
          url = rel.target
          isInternal = false
        }
      } else if (action) {
        isInternal = true
        url = null
      }

      return {
        targetId: meta.targetId || meta.shapeId || meta.shapeName || 'unknown',
        url: url || undefined,
        isInternal,
        targetSlide,
        action: action || undefined,
        tooltip: tooltip || undefined,
        type,
        shapeName: meta.shapeName || undefined,
        shapeId: meta.shapeId || undefined,
        text: meta.text || undefined,
        rId: rId || undefined,
      }
    }

    // 1. Traverse all drawable shapes/pictures/containers for shape-level links
    const inspectElement = (item, elementType) => {
      const cNvPr = getCnvPr(item)
      if (cNvPr) {
        const shapeId = cNvPr['@_id'] ? String(cNvPr['@_id']) : undefined
        const shapeName = cNvPr['@_name'] ? String(cNvPr['@_name']) : undefined
        const targetId = shapeName || shapeId || 'unknown'

        if (cNvPr['a:hlinkClick']) {
          const desc = buildDescriptor(cNvPr['a:hlinkClick'], {
            targetId,
            shapeId,
            shapeName,
            type: elementType === 'p:pic' ? 'image' : 'shape',
          })
          if (desc) results.push(desc)
        }

        // 2. Traverse text runs for inline textRun-level links
        if (item['p:txBody']) {
          let paras = item['p:txBody']['a:p'] || []
          if (!Array.isArray(paras)) paras = [paras]
          for (const p of paras) {
            let runs = p['a:r'] || []
            if (!Array.isArray(runs)) runs = [runs]
            for (const r of runs) {
              const rPr = r['a:rPr']
              if (rPr && rPr['a:hlinkClick']) {
                const text = r['a:t'] !== undefined ? String(r['a:t']) : undefined
                const desc = buildDescriptor(rPr['a:hlinkClick'], {
                  targetId,
                  shapeId,
                  shapeName,
                  text,
                  type: 'textRun',
                })
                if (desc) results.push(desc)
              }
            }
          }
        }
      }

      if (elementType === 'p:grpSp') {
        const grpChildren = ['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp']
        for (const childTag of grpChildren) {
          let childItems = item[childTag] || []
          if (!Array.isArray(childItems)) childItems = [childItems]
          for (const child of childItems) {
            inspectElement(child, childTag)
          }
        }
      }
    }

    const topTags = ['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp']
    for (const tag of topTags) {
      let items = spTree[tag] || []
      if (!Array.isArray(items)) items = [items]
      for (const item of items) {
        inspectElement(item, tag)
      }
    }

    return results
  }

  // ==========================================================================
  // 5. getHyperlink(slideIndex, targetId, slideManager, relationshipManager)
  // ==========================================================================
  /**
   * Retrieves the hyperlink configuration for a specific target element or text.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string|number} targetId - Shape ID, shape name, or text string.
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @returns {Object|null} Hyperlink descriptor object or null.
   */
  getHyperlink(slideIndex, targetId, slideManager, relationshipManager) {
    const links = this.getHyperlinks(slideIndex, slideManager, relationshipManager)
    const targetStr = String(targetId).trim().toLowerCase()

    const match = links.find(l => {
      const idMatch = l.shapeId && String(l.shapeId).toLowerCase() === targetStr
      const nameMatch = l.shapeName && String(l.shapeName).toLowerCase() === targetStr
      const targetMatch = l.targetId && String(l.targetId).toLowerCase() === targetStr
      const textMatch = l.text && String(l.text).toLowerCase() === targetStr
      return idMatch || nameMatch || targetMatch || textMatch
    })

    return match || null
  }

  // ==========================================================================
  // 6. setHyperlink(slideIndex, targetId, text, url, options, slideManager, relationshipManager)
  // ==========================================================================
  /**
   * Locates a substring inside a shape/textbox, isolates it into a split run,
   * attaches an <a:hlinkClick> hyperlink, and applies typography/styling overrides.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string|number|null} targetId - ID or name of the target shape (or null to search across slide).
   * @param {string} text - Substring text to make clickable.
   * @param {string|number|Object} url - Target URL, action, target slide number, or options bag.
   * @param {Object} [options={}] - Styling and hyperlink options (color, underline, bold, italic, tooltip).
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   */
  setHyperlink(slideIndex, targetId, text, url, options = {}, slideManager, relationshipManager) {
    if (!text || typeof text !== 'string') {
      throw new PPTXError('setHyperlink requires a non-empty text string to link')
    }

    let effectiveUrl = url
    let effectiveOptions = options || {}
    if (typeof url === 'object' && url !== null) {
      effectiveOptions = { ...url, ...effectiveOptions }
      effectiveUrl =
        effectiveOptions.url !== undefined
          ? effectiveOptions.url
          : effectiveOptions.action !== undefined
            ? effectiveOptions.action
            : effectiveOptions.targetSlide
    }

    const relMgr = relationshipManager || this.#relationshipManager
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) {
      throw new PPTXError(`Slide ${slideIndex} not found`)
    }

    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
    if (!spTree) {
      throw new PPTXError(`Slide ${slideIndex} has no valid spTree structure`)
    }

    // 1. Locate target shape
    let targetShape = null
    if (targetId) {
      const found = findElementInContainer(spTree, targetId)
      if (found && found.element && found.element['p:txBody']) {
        targetShape = found.element
      }
    }

    // Fallback: search all shapes on the slide for one containing the text
    if (!targetShape) {
      let shapes = spTree['p:sp'] || []
      if (!Array.isArray(shapes)) shapes = [shapes]
      for (const s of shapes) {
        if (s['p:txBody']) {
          let paras = s['p:txBody']['a:p'] || []
          if (!Array.isArray(paras)) paras = [paras]
          for (const p of paras) {
            let runs = p['a:r'] || []
            if (!Array.isArray(runs)) runs = [runs]
            const fullText = runs.map(r => String(r['a:t'] || '')).join('')
            if (fullText.includes(text)) {
              targetShape = s
              break
            }
          }
        }
        if (targetShape) break
      }
    }

    if (!targetShape || !targetShape['p:txBody']) {
      throw new PPTXError(`Target shape for text "${text}" not found on slide ${slideIndex}`)
    }

    // 2. Resolve relationship / action for the hyperlink
    const { rId, action, tooltip } = this.#resolveLinkTarget(
      effectiveUrl,
      effectiveOptions,
      slideInfo,
      slideManager,
      relMgr
    )

    const hlinkClickObj = {}
    if (rId !== null && rId !== undefined) hlinkClickObj['@_r:id'] = rId
    if (action) hlinkClickObj['@_action'] = action
    if (tooltip) hlinkClickObj['@_tooltip'] = tooltip

    // Helper to build styled and schema-ordered run properties
    const buildStyledRunProperties = existingRPr => {
      let targetRPr = this.#xmlParser.deepClone(existingRPr || {})
      delete targetRPr['@_err']

      const styleObj =
        effectiveOptions.style && typeof effectiveOptions.style === 'object'
          ? effectiveOptions.style
          : {}
      const styleOverrides = { ...styleObj, ...effectiveOptions }

      const rawColor =
        effectiveOptions.color ??
        styleObj.color ??
        effectiveOptions.fontColor ??
        styleObj.fontColor ??
        effectiveOptions.textColor ??
        styleObj.textColor

      const sanitizedColor = sanitizeHexColor(rawColor)

      const isAction =
        effectiveOptions.action ||
        (typeof effectiveUrl === 'string' && isShowJumpAction(effectiveUrl)) ||
        effectiveOptions.targetSlide !== undefined ||
        typeof effectiveUrl === 'number'

      if (sanitizedColor) {
        styleOverrides.color = sanitizedColor
      } else if (!isAction) {
        styleOverrides.color = '0563C1'
      }

      if (styleOverrides.underline === undefined && !isAction) {
        styleOverrides.underline = true
      }

      // If underline is enabled and no explicit underlineColor provided, match text color
      if (
        styleOverrides.underline &&
        styleOverrides.underline !== 'none' &&
        styleOverrides.underlineColor === undefined &&
        styleOverrides.color
      ) {
        styleOverrides.underlineColor = styleOverrides.color
      }

      targetRPr = StyleBuilder.buildRunProperties(targetRPr, styleOverrides)
      targetRPr['a:hlinkClick'] = hlinkClickObj
      return StyleBuilder.orderRunProperties(targetRPr)
    }

    // 3. Locate the paragraph and perform run splitting
    let paras = targetShape['p:txBody']['a:p'] || []
    if (!Array.isArray(paras)) {
      paras = [paras]
      targetShape['p:txBody']['a:p'] = paras
    }

    let modified = false

    for (const p of paras) {
      let runs = p['a:r'] || []
      if (!Array.isArray(runs)) {
        runs = [runs]
        p['a:r'] = runs
      }

      // Check single run containment first (most common case)
      const singleRunIdx = runs.findIndex(r => String(r['a:t'] || '').includes(text))
      if (singleRunIdx !== -1) {
        const r = runs[singleRunIdx]
        const fullStr = String(r['a:t'])
        const matchIdx = fullStr.indexOf(text)
        const beforeText = fullStr.substring(0, matchIdx)
        const matchedText = text
        const afterText = fullStr.substring(matchIdx + text.length)

        const targetRPr = buildStyledRunProperties(r['a:rPr'])

        const replacementRuns = []
        if (beforeText.length > 0) {
          replacementRuns.push({
            'a:rPr': this.#xmlParser.deepClone(r['a:rPr'] || {}),
            'a:t': beforeText,
            '@_xml:space': 'preserve',
          })
        }
        replacementRuns.push({
          'a:rPr': targetRPr,
          'a:t': matchedText,
        })
        if (afterText.length > 0) {
          replacementRuns.push({
            'a:rPr': this.#xmlParser.deepClone(r['a:rPr'] || {}),
            'a:t': afterText,
            '@_xml:space': 'preserve',
          })
        }

        runs.splice(singleRunIdx, 1, ...replacementRuns)
        modified = true
        break
      }

      // Cross-run match: concatenate runs and split across boundaries
      const runTexts = runs.map(r => String(r['a:t'] || ''))
      const joinedText = runTexts.join('')
      const crossMatchIdx = joinedText.indexOf(text)
      if (crossMatchIdx !== -1) {
        const startPos = crossMatchIdx
        const endPos = crossMatchIdx + text.length

        let currentPos = 0
        const newRuns = []
        let matchedRunCreated = false

        for (let i = 0; i < runs.length; i++) {
          const r = runs[i]
          const rText = runTexts[i]
          const rStart = currentPos
          const rEnd = currentPos + rText.length
          currentPos = rEnd

          if (rEnd <= startPos || rStart >= endPos) {
            // Entirely before or after the match
            newRuns.push(r)
          } else {
            // Overlaps with the match
            if (rStart < startPos) {
              const beforePart = rText.substring(0, startPos - rStart)
              newRuns.push({
                'a:rPr': this.#xmlParser.deepClone(r['a:rPr'] || {}),
                'a:t': beforePart,
                '@_xml:space': 'preserve',
              })
            }

            if (!matchedRunCreated) {
              const targetRPr = buildStyledRunProperties(r['a:rPr'])
              newRuns.push({
                'a:rPr': targetRPr,
                'a:t': text,
              })
              matchedRunCreated = true
            }

            if (rEnd > endPos) {
              const afterPart = rText.substring(endPos - rStart)
              newRuns.push({
                'a:rPr': this.#xmlParser.deepClone(r['a:rPr'] || {}),
                'a:t': afterPart,
                '@_xml:space': 'preserve',
              })
            }
          }
        }

        p['a:r'] = newRuns
        modified = true
        break
      }
    }

    if (!modified) {
      throw new PPTXError(
        `Text "${text}" not found in shape "${targetId || 'active'}" on slide ${slideIndex}`
      )
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(
      `setHyperlink: Successfully set styled hyperlink for "${text}" on slide ${slideIndex}`
    )
  }

  // ==========================================================================
  // Private Helper: Resolve target URL / action / slide jump
  // ==========================================================================
  #resolveLinkTarget(url, options, slideInfo, slideManager, relationshipManager) {
    let rId = null
    let action = null
    const tooltip = options.tooltip ? String(options.tooltip) : undefined

    // 1. Navigation Action (next, prev, first, last, endshow, ppaction://hlinkshowjump)
    if (options.action && isShowJumpAction(options.action)) {
      action = normalizeAction(options.action)
      return { rId: '', action, tooltip }
    }
    if (typeof url === 'string' && isShowJumpAction(url)) {
      action = normalizeAction(url)
      return { rId: '', action, tooltip }
    }

    // 2. Internal Slide Target
    const isTargetSlide =
      options.targetSlide !== undefined ||
      options.isInternal === true ||
      typeof url === 'number' ||
      (typeof url === 'string' && /^#?slide\s*:?\s*(\d+)$/i.test(url.trim())) ||
      (typeof url === 'string' && /^\d+$/.test(url.trim()))

    if (isTargetSlide) {
      let targetSlideNumber = options.targetSlide
      if (targetSlideNumber === undefined) {
        if (typeof url === 'number') {
          targetSlideNumber = url
        } else {
          targetSlideNumber = parseInt(String(url).replace(/[^\d]/g, ''), 10)
        }
      }

      const targetInfo = slideManager.getSlideInfo(targetSlideNumber)
      if (!targetInfo) {
        throw new PPTXError(`Target slide ${targetSlideNumber} not found`)
      }

      const relativeTarget = targetInfo.zipPath.split('/').pop()

      // Deduplicate relationship if one already exists for this target slide
      const existingRels = relationshipManager.getRelationships(slideInfo.zipPath)
      const existingRel = existingRels.find(
        r => r.type === REL_TYPES.SLIDE && r.target === relativeTarget
      )
      if (existingRel) {
        rId = existingRel.id
      } else {
        rId = relationshipManager.addRelationship(
          slideInfo.zipPath,
          REL_TYPES.SLIDE,
          relativeTarget
        )
      }

      action = 'ppaction://hlinksldjump'
      return { rId, action, tooltip }
    }

    // 3. Custom ppaction URI
    if (typeof url === 'string' && url.startsWith('ppaction://')) {
      action = url
      return { rId: '', action, tooltip }
    }

    // 4. External URL
    const sanitizedUrl = sanitizeUrl(url)

    // Deduplicate relationship if one already exists for this external URL
    const existingRels = relationshipManager.getRelationships(slideInfo.zipPath)
    const existingRel = existingRels.find(
      r =>
        r.type === REL_TYPES.HYPERLINK &&
        r.target === sanitizedUrl &&
        (r.targetMode === 'External' || !r.targetMode)
    )
    if (existingRel) {
      rId = existingRel.id
    } else {
      rId = relationshipManager.addRelationship(
        slideInfo.zipPath,
        REL_TYPES.HYPERLINK,
        sanitizedUrl,
        'External'
      )
    }

    return {
      rId,
      action: options.action ? normalizeAction(options.action) : null,
      tooltip,
    }
  }

  // ==========================================================================
  // Private Helper: Purge orphaned relationships
  // ==========================================================================
  #purgeRelationshipIfOrphaned(
    slideIndex,
    rId,
    slideObj,
    slideInfo,
    slideManager,
    relationshipManager
  ) {
    if (!rId) return
    const isStillUsed = isRIdReferencedInAst(slideObj, rId)
    if (!isStillUsed) {
      relationshipManager.removeRelationship(slideInfo.zipPath, rId)
      logger.debug(`Purged orphaned relationship "${rId}" from slide ${slideIndex}`)
    }
  }
}

module.exports = { HyperlinkManager }
