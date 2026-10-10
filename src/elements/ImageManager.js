/**
 * @fileoverview ImageManager - Manages slide-level image operations like replace, add, remove, and list.
 */

const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')
const { REL_TYPES } = require('../managers/RelationshipManager.js')

const logger = createLogger('ImageManager')

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
   * Replaces an existing image on a slide, locking its bounding box and maintaining aspect ratio.
   *
   * @param {number} slideIndex
   * @param {string} imageIdOrName - Shape name or shape ID.
   * @param {string|Buffer} sourcePathOrBuffer - New image source path or Buffer.
   * @param {SlideManager} slideManager
   * @param {MediaManager} mediaManager
   * @param {RelationshipManager} relationshipManager
   */
  async replaceImage(
    slideIndex,
    imageIdOrName,
    sourcePathOrBuffer,
    slideManager,
    mediaManager,
    relationshipManager
  ) {
    const picRes = slideManager.getSlidePic(slideIndex, imageIdOrName)

    if (!picRes) {
      throw new PPTXError(`Image shape "${imageIdOrName}" not found on slide ${slideIndex}`)
    }

    const rId = picRes.pic?.['p:blipFill']?.['a:blip']?.['@_r:embed']
    if (!rId) {
      throw new PPTXError(`No relationship ID found on image shape "${imageIdOrName}"`)
    }

    // Lock bounding box & aspect ratio
    const { getImageMetadata } = require('../utils/imageMetadata.js')
    const meta = await getImageMetadata(sourcePathOrBuffer)

    if (meta && meta.width > 0 && meta.height > 0) {
      const xfrm = picRes.pic['p:spPr']?.['a:xfrm']
      if (xfrm && xfrm['a:ext'] && xfrm['a:off']) {
        const origCx = parseInt(xfrm['a:ext']['@_cx'], 10)
        const origCy = parseInt(xfrm['a:ext']['@_cy'], 10)
        const origX = parseInt(xfrm['a:off']['@_x'], 10)
        const origY = parseInt(xfrm['a:off']['@_y'], 10)

        if (!isNaN(origCx) && !isNaN(origCy) && origCx > 0 && origCy > 0) {
          const origAspect = origCx / origCy
          const newAspect = meta.width / meta.height

          let newCx = origCx
          let newCy = origCy
          let newX = origX
          let newY = origY

          if (newAspect > origAspect) {
            // New image is wider than bounding box
            newCy = Math.round(origCx / newAspect)
            newY = origY + Math.round((origCy - newCy) / 2)
          } else if (newAspect < origAspect) {
            // New image is taller than bounding box
            newCx = Math.round(origCy * newAspect)
            newX = origX + Math.round((origCx - newCx) / 2)
          }

          xfrm['a:ext']['@_cx'] = String(newCx)
          xfrm['a:ext']['@_cy'] = String(newCy)
          xfrm['a:off']['@_x'] = String(newX)
          xfrm['a:off']['@_y'] = String(newY)
          slideManager.markSlideObjDirty(slideIndex)
        }
      }
    }

    // Embed the new image bytes
    const destMediaZipPath = await mediaManager.embedImage(sourcePathOrBuffer)
    const relativeTarget = `../media/${destMediaZipPath.split('/').pop()}`

    // Update relationship target path
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    relationshipManager.updateRelationshipTarget(slideInfo.zipPath, rId, relativeTarget)

    logger.debug(`Replaced image "${imageIdOrName}" target with "${relativeTarget}"`)
  }

  /**
   * Replaces a placeholder shape with an image, locking its bounding box.
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

    // Read shape dimensions
    const xfrm = res.shape['p:spPr']?.['a:xfrm']
    if (!xfrm || !xfrm['a:off'] || !xfrm['a:ext']) {
      throw new PPTXError(`Shape "${shapeIdOrName}" lacks valid bounding box (xfrm)`)
    }

    const { getImageMetadata } = require('../utils/imageMetadata.js')
    const meta = await getImageMetadata(sourcePathOrBuffer)

    const origCx = parseInt(xfrm['a:ext']['@_cx'], 10) || 1000000
    const origCy = parseInt(xfrm['a:ext']['@_cy'], 10) || 1000000
    const origX = parseInt(xfrm['a:off']['@_x'], 10) || 0
    const origY = parseInt(xfrm['a:off']['@_y'], 10) || 0

    let newCx = origCx
    let newCy = origCy
    let newX = origX
    let newY = origY

    if (meta && meta.width > 0 && meta.height > 0) {
      const origAspect = origCx / origCy
      const newAspect = meta.width / meta.height

      if (newAspect > origAspect) {
        newCy = Math.round(origCx / newAspect)
        newY = origY + Math.round((origCy - newCy) / 2)
      } else if (newAspect < origAspect) {
        newCx = Math.round(origCy * newAspect)
        newX = origX + Math.round((origCx - newCx) / 2)
      }
    }

    // Delete the old shape and replace with a p:pic
    const destMediaZipPath = await mediaManager.embedImage(sourcePathOrBuffer)
    const relativeTarget = `../media/${destMediaZipPath.split('/').pop()}`
    const slideInfo = slideManager.getSlideInfo(slideIndex)

    const rId = relationshipManager.addRelationship(
      slideInfo.zipPath,
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
      relativeTarget
    )

    const picXml = `<p:pic><p:nvPicPr><p:cNvPr id="${res.id}" name="${res.name}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${newX}" y="${newY}"/><a:ext cx="${newCx}" cy="${newCy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`

    const parsedRaw = this.#xmlParser.parse(picXml, 'pic.xml')['p:pic']
    const parsedPic = Array.isArray(parsedRaw) ? parsedRaw[0] : parsedRaw

    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']

    if (spTree) {
      // Find and remove the old shape
      const shapes = Array.isArray(spTree['p:sp']) ? spTree['p:sp'] : [spTree['p:sp']]
      const newShapes = shapes.filter(s => s !== res.shape)
      if (newShapes.length === 1) spTree['p:sp'] = newShapes[0]
      else if (newShapes.length > 1) spTree['p:sp'] = newShapes
      else delete spTree['p:sp']

      // Add the new pic
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

  /**
   * Adds a new image to a slide.
   *
   * @param {number} slideIndex
   * @param {string|Buffer} sourcePathOrBuffer
   * @param {Object} options - Position options (x, y, width, height in EMUs or inches).
   * @param {SlideManager} slideManager
   * @param {MediaManager} mediaManager
   * @param {RelationshipManager} relationshipManager
   */
  async addImage(
    slideIndex,
    sourcePathOrBuffer,
    options = {},
    slideManager,
    mediaManager,
    relationshipManager
  ) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    if (!spTree) {
      throw new PPTXError(`Invalid slide structure for slide ${slideIndex}`)
    }

    // Embed the image
    const destMediaZipPath = await mediaManager.embedImage(sourcePathOrBuffer)
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

    // Convert positions
    const x =
      options.x === undefined
        ? 0
        : options.x < 100
          ? Math.round(options.x * 914400)
          : Math.round(options.x)
    const y =
      options.y === undefined
        ? 0
        : options.y < 100
          ? Math.round(options.y * 914400)
          : Math.round(options.y)
    const cx =
      options.width === undefined
        ? 2743200
        : options.width < 100
          ? Math.round(options.width * 914400)
          : Math.round(options.width)
    const cy =
      options.height === undefined
        ? 1828800
        : options.height < 100
          ? Math.round(options.height * 914400)
          : Math.round(options.height)
    const name = options.name || `Picture ${newId}`

    // Build the pic XML snippet using mediaManager's builder and parse it
    const picXml = mediaManager.buildImageXml(rId, {
      x,
      y,
      width: cx,
      height: cy,
      name,
      shapeId: newId,
    })
    const parsedPicRaw = this.#xmlParser.parse(picXml, 'pic.xml')['p:pic']
    const picObj = Array.isArray(parsedPicRaw) ? parsedPicRaw[0] : parsedPicRaw

    if (!spTree['p:pic']) {
      spTree['p:pic'] = []
    }
    if (!Array.isArray(spTree['p:pic'])) {
      spTree['p:pic'] = [spTree['p:pic']]
    }
    spTree['p:pic'].push(picObj)

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Added image "${name}" with ID ${newId} and rId ${rId} to slide ${slideIndex}`)
  }

  /**
   * Removes an image from a slide.
   *
   * @param {number} slideIndex
   * @param {string} imageIdOrName
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   */
  removeImage(slideIndex, imageIdOrName, slideManager, relationshipManager) {
    const picRes = slideManager.getSlidePic(slideIndex, imageIdOrName)

    if (!picRes) {
      throw new PPTXError(`Image shape "${imageIdOrName}" not found on slide ${slideIndex}`)
    }

    // Remove from shape list
    const parent = picRes.parent
    if (parent['p:pic']) {
      if (Array.isArray(parent['p:pic'])) {
        parent['p:pic'] = parent['p:pic'].filter(p => p !== picRes.pic)
      } else if (parent['p:pic'] === picRes.pic) {
        delete parent['p:pic']
      }
    }

    // Remove relationship
    const rId = picRes.pic?.['p:blipFill']?.['a:blip']?.['@_r:embed']
    if (rId) {
      const slideInfo = slideManager.getSlideInfo(slideIndex)
      relationshipManager.removeRelationship(slideInfo.zipPath, rId)
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Removed image "${imageIdOrName}" from slide ${slideIndex}`)
  }

  /**
   * Enumerates all images on a slide.
   *
   * @param {number} slideIndex
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @returns {Array<Object>} List of image elements found.
   */
  getImages(slideIndex, slideManager, relationshipManager) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    const imagesInfo = []
    this.#collectImagesInfo(spTree, slideInfo.zipPath, relationshipManager, imagesInfo)
    return imagesInfo
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

  #collectImagesInfo(container, slideZipPath, relationshipManager, results) {
    if (!container) return

    let pics = container['p:pic'] || []
    if (!Array.isArray(pics)) pics = [pics]

    for (const pic of pics) {
      const cNvPr = pic?.['p:nvPicPr']?.['p:cNvPr']
      if (!cNvPr) continue

      const name = cNvPr['@_name']
      const id = String(cNvPr['@_id'])

      const rId = pic?.['p:blipFill']?.['a:blip']?.['@_r:embed']
      let targetPath = ''
      if (rId) {
        const rel = relationshipManager.getRelationshipById(slideZipPath, rId)
        if (rel) {
          targetPath = relationshipManager.resolveTarget(slideZipPath, rel.target)
        }
      }

      const xfrm = pic['p:spPr']?.['a:xfrm']
      const position = xfrm
        ? {
            x: parseInt(xfrm['a:off']?.['@_x'] || 0, 10),
            y: parseInt(xfrm['a:off']?.['@_y'] || 0, 10),
            cx: parseInt(xfrm['a:ext']?.['@_cx'] || 0, 10),
            cy: parseInt(xfrm['a:ext']?.['@_cy'] || 0, 10),
          }
        : null

      results.push({
        type: 'image',
        id,
        name,
        relationshipId: rId,
        targetPath,
        position,
      })
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const g of groups) {
      this.#collectImagesInfo(g, slideZipPath, relationshipManager, results)
    }
  }
}

module.exports = { ImageManager }
