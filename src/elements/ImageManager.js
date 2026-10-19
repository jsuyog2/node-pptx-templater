/**
 * @fileoverview ImageManager - Manages slide-level image operations (AddImage, UpdateImage, RemoveImage).
 */

const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')
const { REL_TYPES } = require('../managers/RelationshipManager.js')
const { resolveElementBounds } = require('../utils/coordinateUtils.js')

const logger = createLogger('ImageManager')

function ensureSpTreeOrder(spTree) {
  if (!spTree) return
  if (spTree['p:extLst']) {
    const extLst = spTree['p:extLst']
    delete spTree['p:extLst']
    spTree['p:extLst'] = extLst
  }
}

/**
 * @class ImageManager
 * @description Manages image elements and relationships on individual slides.
 */
class ImageManager {
  /** @private @type {XMLParser} */
  #xmlParser

  /**
   * @param {XMLParser} xmlParser
   */
  constructor(xmlParser) {
    this.#xmlParser = xmlParser
  }

  /**
   * Adds an image to a slide.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string|Buffer|Uint8Array} source - Image source (file path, URL, data URI, or Buffer).
   * @param {Object} [options={}] - Options (position, dimensions, rotation, opacity, crop, border, shadow).
   * @param {SlideManager} slideManager
   * @param {MediaManager} mediaManager
   * @param {RelationshipManager} relationshipManager
   * @returns {Promise<{ id: number, name: string, rId: string }>} Added image descriptor.
   */
  async addImage(
    slideIndex,
    source,
    options = {},
    slideManager,
    mediaManager,
    relationshipManager
  ) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) {
      throw new PPTXError(`Slide ${slideIndex} not found`)
    }
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    if (!spTree) {
      throw new PPTXError(`Invalid slide structure for slide ${slideIndex}`)
    }

    // Embed image asset into ppt/media/
    const destMediaZipPath = await mediaManager.embedMedia(source, 'image')
    const relativeTarget = `../media/${destMediaZipPath.split('/').pop()}`

    // Add relationship
    const rId = relationshipManager.addRelationship(
      slideInfo.zipPath,
      REL_TYPES.IMAGE,
      relativeTarget
    )

    // Generate unique shape ID
    const existingIds = this.#getAllShapeIds(spTree)
    const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 1000
    const newId = maxId + 1

    // Resolve coordinates & bounds
    const bounds = resolveElementBounds(
      {
        x: options.x !== undefined ? options.x : 0,
        y: options.y !== undefined ? options.y : 0,
        width:
          options.width !== undefined
            ? options.width
            : options.cx !== undefined
              ? options.cx
              : '30%',
        height:
          options.height !== undefined
            ? options.height
            : options.cy !== undefined
              ? options.cy
              : '30%',
        unit: options.unit,
        isPercent: options.isPercent,
      },
      slideManager
    )

    const name = options.name || `Picture ${newId}`

    // Rotation (degrees to EMU: deg * 60000)
    let rotAttr = ''
    if (options.rotation !== undefined && options.rotation !== 0) {
      const rotVal = Math.round(options.rotation * 60000)
      rotAttr = ` rot="${rotVal}"`
    }

    // Opacity / transparency in thousandths of a percent (100% = 100000)
    let alphaModFix = ''
    if (options.opacity !== undefined) {
      const opacityVal = options.opacity <= 1 ? options.opacity * 100000 : options.opacity * 1000
      alphaModFix = `<a:alphaModFix amt="${Math.round(opacityVal)}"/>`
    } else if (options.transparency !== undefined) {
      const transVal =
        options.transparency <= 1
          ? (1 - options.transparency) * 100000
          : (100 - options.transparency) * 1000
      alphaModFix = `<a:alphaModFix amt="${Math.round(transVal)}"/>`
    }

    // Crop: <a:srcRect l="..." t="..." r="..." b="..."/>
    let srcRectXml = ''
    const crop = options.crop || options.cropTo
    if (crop) {
      const toThousandths = v =>
        v === undefined ? 0 : v <= 1 ? Math.round(v * 100000) : Math.round(v * 1000)
      const l = toThousandths(crop.l !== undefined ? crop.l : crop.left)
      const t = toThousandths(crop.t !== undefined ? crop.t : crop.top)
      const r = toThousandths(crop.r !== undefined ? crop.r : crop.right)
      const b = toThousandths(crop.b !== undefined ? crop.b : crop.bottom)
      srcRectXml = `<a:srcRect l="${l}" t="${t}" r="${r}" b="${b}"/>`
    }

    // Border line: <a:ln w="..."><a:solidFill><a:srgbClr val="..."/></a:solidFill></a:ln>
    let lnXml = ''
    if (options.border) {
      const borderWidth =
        typeof options.border.width === 'number'
          ? options.border.width < 50
            ? Math.round(options.border.width * 12700)
            : options.border.width
          : 12700
      const borderColor = String(options.border.color || '000000')
        .replace('#', '')
        .toUpperCase()
      lnXml = `<a:ln w="${borderWidth}"><a:solidFill><a:srgbClr val="${borderColor}"/></a:solidFill></a:ln>`
    }

    // Shadow: <a:effectLst><a:outerShdw ...></a:effectLst>
    let effectLstXml = ''
    if (options.shadow) {
      const blur = options.shadow.blur ? Math.round(options.shadow.blur * 12700) : 50800
      const dist = options.shadow.distance ? Math.round(options.shadow.distance * 12700) : 38100
      const dir = options.shadow.angle ? Math.round(options.shadow.angle * 60000) : 5400000
      const color = String(options.shadow.color || '000000')
        .replace('#', '')
        .toUpperCase()
      effectLstXml = `<a:effectLst><a:outerShdw blurRad="${blur}" dist="${dist}" dir="${dir}" algn="tl"><a:srgbClr val="${color}"/></a:outerShdw></a:effectLst>`
    }

    const picXml = `<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="${newId}" name="${name}"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rId}" cstate="print">
      ${alphaModFix}
      <a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst>
    </a:blip>
    ${srcRectXml}
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm${rotAttr}>
      <a:off x="${bounds.x}" y="${bounds.y}"/>
      <a:ext cx="${bounds.width}" cy="${bounds.height}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    ${lnXml}
    ${effectLstXml}
  </p:spPr>
</p:pic>`

    const parsedPicRaw = this.#xmlParser.parse(picXml, 'pic.xml')['p:pic']
    const picObj = Array.isArray(parsedPicRaw) ? parsedPicRaw[0] : parsedPicRaw

    if (!spTree['p:pic']) {
      spTree['p:pic'] = []
    }
    if (!Array.isArray(spTree['p:pic'])) {
      spTree['p:pic'] = [spTree['p:pic']]
    }
    spTree['p:pic'].push(picObj)
    ensureSpTreeOrder(spTree)

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Added image "${name}" (ID: ${newId}, rId: ${rId}) to slide ${slideIndex}`)

    return { id: newId, name, rId }
  }

  /**
   * Alias for addImage.
   */
  async AddImage(...args) {
    return this.addImage(...args)
  }

  /**
   * Updates an existing image on a slide.
   *
   * @param {number} slideIndex
   * @param {string|number} imageIdOrName
   * @param {Object} options
   * @param {SlideManager} slideManager
   * @param {MediaManager} mediaManager
   * @param {RelationshipManager} relationshipManager
   * @returns {Promise<boolean>} True if updated, false if not found.
   */
  async updateImage(
    slideIndex,
    imageIdOrName,
    options = {},
    slideManager,
    mediaManager,
    relationshipManager
  ) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) return false
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    if (!spTree) return false

    const targetIdStr = String(imageIdOrName)
    const match = this.#findPicRecursive(spTree, targetIdStr)
    if (!match) return false

    const { pic } = match

    // 1. Swap source if provided
    if (options.source) {
      const destMediaZipPath = await mediaManager.embedMedia(options.source, 'image')
      const relativeTarget = `../media/${destMediaZipPath.split('/').pop()}`
      const existingRId = pic?.['p:blipFill']?.['a:blip']?.['@_r:embed']
      if (existingRId) {
        relationshipManager.updateRelationshipTarget(slideInfo.zipPath, existingRId, relativeTarget)
      } else {
        const newRId = relationshipManager.addRelationship(
          slideInfo.zipPath,
          REL_TYPES.IMAGE,
          relativeTarget
        )
        if (!pic['p:blipFill']) pic['p:blipFill'] = {}
        if (!pic['p:blipFill']['a:blip']) pic['p:blipFill']['a:blip'] = {}
        pic['p:blipFill']['a:blip']['@_r:embed'] = newRId
      }
    }

    // 2. Spatial coordinates / bounds
    if (
      options.x !== undefined ||
      options.y !== undefined ||
      options.width !== undefined ||
      options.height !== undefined ||
      options.cx !== undefined ||
      options.cy !== undefined ||
      options.unit ||
      options.isPercent
    ) {
      if (!pic['p:spPr']) pic['p:spPr'] = {}
      if (!pic['p:spPr']['a:xfrm']) {
        pic['p:spPr']['a:xfrm'] = {
          'a:off': { '@_x': '0', '@_y': '0' },
          'a:ext': { '@_cx': '2743200', '@_cy': '1828800' },
        }
      }
      const xfrm = pic['p:spPr']['a:xfrm']
      const bounds = resolveElementBounds(
        {
          x: options.x !== undefined ? options.x : xfrm['a:off']?.['@_x'],
          y: options.y !== undefined ? options.y : xfrm['a:off']?.['@_y'],
          width:
            options.width !== undefined
              ? options.width
              : options.cx !== undefined
                ? options.cx
                : xfrm['a:ext']?.['@_cx'],
          height:
            options.height !== undefined
              ? options.height
              : options.cy !== undefined
                ? options.cy
                : xfrm['a:ext']?.['@_cy'],
          unit: options.unit,
          isPercent: options.isPercent,
        },
        slideManager
      )
      if (!xfrm['a:off']) xfrm['a:off'] = {}
      if (!xfrm['a:ext']) xfrm['a:ext'] = {}
      xfrm['a:off']['@_x'] = String(bounds.x)
      xfrm['a:off']['@_y'] = String(bounds.y)
      xfrm['a:ext']['@_cx'] = String(bounds.width)
      xfrm['a:ext']['@_cy'] = String(bounds.height)
    }

    // 3. Rotation
    if (options.rotation !== undefined) {
      if (!pic['p:spPr']) pic['p:spPr'] = {}
      if (!pic['p:spPr']['a:xfrm']) pic['p:spPr']['a:xfrm'] = {}
      pic['p:spPr']['a:xfrm']['@_rot'] = String(Math.round(options.rotation * 60000))
    }

    // 4. Crop
    const crop = options.crop || options.cropTo
    if (crop) {
      if (!pic['p:blipFill']) pic['p:blipFill'] = {}
      const toThousandths = v =>
        v === undefined ? 0 : v <= 1 ? Math.round(v * 100000) : Math.round(v * 1000)
      const l = String(toThousandths(crop.l !== undefined ? crop.l : crop.left))
      const t = String(toThousandths(crop.t !== undefined ? crop.t : crop.top))
      const r = String(toThousandths(crop.r !== undefined ? crop.r : crop.right))
      const b = String(toThousandths(crop.b !== undefined ? crop.b : crop.bottom))
      pic['p:blipFill']['a:srcRect'] = {
        '@_l': l,
        '@_t': t,
        '@_r': r,
        '@_b': b,
      }
    }

    // 5. Opacity / transparency
    if (options.opacity !== undefined || options.transparency !== undefined) {
      if (!pic['p:blipFill']) pic['p:blipFill'] = {}
      if (!pic['p:blipFill']['a:blip']) pic['p:blipFill']['a:blip'] = {}
      const opacityVal =
        options.opacity !== undefined
          ? options.opacity <= 1
            ? options.opacity * 100000
            : options.opacity * 1000
          : options.transparency <= 1
            ? (1 - options.transparency) * 100000
            : (100 - options.transparency) * 1000
      pic['p:blipFill']['a:blip']['a:alphaModFix'] = {
        '@_amt': String(Math.round(opacityVal)),
      }
    }

    // 6. Border
    if (options.border) {
      if (!pic['p:spPr']) pic['p:spPr'] = {}
      const borderWidth =
        typeof options.border.width === 'number'
          ? options.border.width < 50
            ? Math.round(options.border.width * 12700)
            : options.border.width
          : 12700
      const borderColor = String(options.border.color || '000000')
        .replace('#', '')
        .toUpperCase()
      pic['p:spPr']['a:ln'] = {
        '@_w': String(borderWidth),
        'a:solidFill': {
          'a:srgbClr': { '@_val': borderColor },
        },
      }
    }

    // 7. Shadow
    if (options.shadow) {
      if (!pic['p:spPr']) pic['p:spPr'] = {}
      const blur = options.shadow.blur ? Math.round(options.shadow.blur * 12700) : 50800
      const dist = options.shadow.distance ? Math.round(options.shadow.distance * 12700) : 38100
      const dir = options.shadow.angle ? Math.round(options.shadow.angle * 60000) : 5400000
      const color = String(options.shadow.color || '000000')
        .replace('#', '')
        .toUpperCase()
      pic['p:spPr']['a:effectLst'] = {
        'a:outerShdw': {
          '@_blurRad': String(blur),
          '@_dist': String(dist),
          '@_dir': String(dir),
          '@_algn': 'tl',
          'a:srgbClr': { '@_val': color },
        },
      }
    }

    ensureSpTreeOrder(spTree)
    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Updated image "${imageIdOrName}" on slide ${slideIndex}`)
    return true
  }

  /**
   * Alias for updateImage.
   */
  async UpdateImage(...args) {
    return this.updateImage(...args)
  }

  /**
   * Removes an image from a slide and purges its relationship if orphaned.
   *
   * @param {number} slideIndex
   * @param {string|number} imageIdOrName
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @returns {boolean} True if removed, false if not found.
   */
  removeImage(slideIndex, imageIdOrName, slideManager, relationshipManager) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) return false
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    if (!spTree) return false

    const targetIdStr = String(imageIdOrName)
    const match = this.#findPicRecursive(spTree, targetIdStr)
    if (!match) return false

    const { pic, parent } = match

    // Remove from shape array
    if (parent['p:pic']) {
      if (Array.isArray(parent['p:pic'])) {
        parent['p:pic'] = parent['p:pic'].filter(p => p !== pic)
      } else if (parent['p:pic'] === pic) {
        delete parent['p:pic']
      }
    }

    // Check relationship and remove if not used elsewhere on this slide
    const rId = pic?.['p:blipFill']?.['a:blip']?.['@_r:embed']
    if (rId) {
      const isUsedElsewhere = this.#isRelationshipUsedInTree(spTree, rId)
      if (!isUsedElsewhere) {
        relationshipManager.removeRelationship(slideInfo.zipPath, rId)
      }
    }

    ensureSpTreeOrder(spTree)
    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Removed image "${imageIdOrName}" from slide ${slideIndex}`)
    return true
  }

  /**
   * Alias for removeImage.
   */
  RemoveImage(...args) {
    return this.removeImage(...args)
  }

  /**
   * Replaces a placeholder shape with an image, locking its bounding box.
   * Internal helper used by PPTXTemplater.prototype.replaceShape.
   */
  async replaceShapeWithImage(
    slideIndex,
    shapeIdOrName,
    sourcePathOrBuffer,
    slideManager,
    mediaManager,
    relationshipManager
  ) {
    const res = slideManager.getSlideShape(slideIndex, shapeIdOrName)
    if (!res) {
      throw new PPTXError(`Shape "${shapeIdOrName}" not found on slide ${slideIndex}`)
    }

    const xfrm = res.shape['p:spPr']?.['a:xfrm']
    if (!xfrm || !xfrm['a:off'] || !xfrm['a:ext']) {
      throw new PPTXError(`Shape "${shapeIdOrName}" lacks valid bounding box (xfrm)`)
    }

    const origCx = parseInt(xfrm['a:ext']['@_cx'], 10) || 1000000
    const origCy = parseInt(xfrm['a:ext']['@_cy'], 10) || 1000000
    const origX = parseInt(xfrm['a:off']['@_x'], 10) || 0
    const origY = parseInt(xfrm['a:off']['@_y'], 10) || 0

    const destMediaZipPath = await mediaManager.embedMedia(sourcePathOrBuffer, 'image')
    const relativeTarget = `../media/${destMediaZipPath.split('/').pop()}`
    const slideInfo = slideManager.getSlideInfo(slideIndex)

    const rId = relationshipManager.addRelationship(
      slideInfo.zipPath,
      REL_TYPES.IMAGE,
      relativeTarget
    )

    const picXml = `<p:pic><p:nvPicPr><p:cNvPr id="${res.id}" name="${res.name}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${origX}" y="${origY}"/><a:ext cx="${origCx}" cy="${origCy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`

    const parsedRaw = this.#xmlParser.parse(picXml, 'pic.xml')['p:pic']
    const parsedPic = Array.isArray(parsedRaw) ? parsedRaw[0] : parsedRaw

    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']

    if (spTree) {
      const shapes = Array.isArray(spTree['p:sp']) ? spTree['p:sp'] : [spTree['p:sp']]
      const newShapes = shapes.filter(s => s !== res.shape)
      if (newShapes.length === 1) spTree['p:sp'] = newShapes[0]
      else if (newShapes.length > 1) spTree['p:sp'] = newShapes
      else delete spTree['p:sp']

      if (!spTree['p:pic']) {
        spTree['p:pic'] = parsedPic
      } else if (Array.isArray(spTree['p:pic'])) {
        spTree['p:pic'].push(parsedPic)
      } else {
        spTree['p:pic'] = [spTree['p:pic'], parsedPic]
      }
      slideManager.markSlideObjDirty(slideIndex)
    }

    logger.debug(`Replaced shape "${shapeIdOrName}" with image at "${relativeTarget}" (ID: ${rId})`)
  }

  #findPicRecursive(container, targetId) {
    if (!container) return null

    let pics = container['p:pic'] || []
    if (!Array.isArray(pics)) pics = [pics]

    for (const pic of pics) {
      const cNvPr = pic?.['p:nvPicPr']?.['p:cNvPr']
      if (cNvPr) {
        const name = cNvPr['@_name']
        const id = String(cNvPr['@_id'])
        const embedId = pic?.['p:blipFill']?.['a:blip']?.['@_r:embed']
        if (name === targetId || id === targetId || embedId === targetId) {
          return { pic, parent: container }
        }
      }
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const group of groups) {
      const res = this.#findPicRecursive(group, targetId)
      if (res) return res
    }

    return null
  }

  #isRelationshipUsedInTree(container, rId) {
    if (!container) return false

    let pics = container['p:pic'] || []
    if (!Array.isArray(pics)) pics = [pics]
    for (const pic of pics) {
      if (pic?.['p:blipFill']?.['a:blip']?.['@_r:embed'] === rId) return true
      if (pic?.['p:nvPicPr']?.['p:nvPr']?.['a:videoFile']?.['@_r:link'] === rId) return true
      if (pic?.['p:nvPicPr']?.['p:nvPr']?.['a:audioFile']?.['@_r:link'] === rId) return true
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const group of groups) {
      if (this.#isRelationshipUsedInTree(group, rId)) return true
    }

    return false
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
}

module.exports = { ImageManager }
