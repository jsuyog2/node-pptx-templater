/**
 * @fileoverview MediaManager - Manages media files (images, videos, audio) in PPTX.
 *
 * All media files are stored in ppt/media/ within the PPTX ZIP.
 * They are referenced via relationships from slides.
 */

const { createHash } = require('crypto')
const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')
const { REL_TYPES } = require('./RelationshipManager.js')
const { resolveElementBounds } = require('../utils/coordinateUtils.js')
const {
  EXT_TO_MIME,
  DEFAULT_POSTER_PNG,
  getDefaultAudioIcon,
  resolveMediaPayload,
} = require('../utils/mediaUtils.js')

const logger = createLogger('MediaManager')

function ensureSpTreeOrder(spTree) {
  if (!spTree) return
  if (spTree['p:extLst']) {
    const extLst = spTree['p:extLst']
    delete spTree['p:extLst']
    spTree['p:extLst'] = extLst
  }
}

/**
 * @class MediaManager
 * @description Manages media embedding, deduplication, and retrieval in PPTX files.
 */
class MediaManager {
  /** @private @type {ContentTypesManager} */
  #contentTypesManager
  /** @private @type {ZipManager} */
  #zipManager

  /**
   * Content hash → existing media ZIP path for deduplication.
   * @private @type {Map<string, string>}
   */
  #mediaHashIndex = new Map()

  /**
   * All known media files.
   * @private @type {Map<string, Object>}
   */
  #mediaRegistry = new Map()

  /**
   * Counter for generating unique media file names.
   * @private @type {number}
   */
  #nextMediaId = 1

  /**
   * @param {ContentTypesManager} contentTypesManager
   */
  constructor(contentTypesManager) {
    this.#contentTypesManager = contentTypesManager
  }

  /**
   * Initializes by scanning existing media files in the PPTX.
   *
   * @param {ZipManager} zipManager
   * @returns {Promise<void>}
   */
  async initialize(zipManager) {
    this.#zipManager = zipManager
    const mediaFiles = zipManager.listFiles('ppt/media/')

    for (const mediaPath of mediaFiles) {
      const ext = mediaPath.split('.').pop().toLowerCase()
      const mimeType = EXT_TO_MIME[ext] || 'application/octet-stream'

      const mediaInfo = { zipPath: mediaPath, hash: null, mimeType, size: null }
      this.#mediaRegistry.set(mediaPath, mediaInfo)

      const numMatch = /\d+/.exec(mediaPath.split('/').pop())
      if (numMatch) {
        const num = parseInt(numMatch[0], 10)
        if (num >= this.#nextMediaId) this.#nextMediaId = num + 1
      }
    }

    logger.debug(`Registered ${this.#mediaRegistry.size} media file(s) (lazy loading enabled)`)
  }

  async #ensureAllMediaHashed() {
    for (const [mediaPath, mediaInfo] of this.#mediaRegistry.entries()) {
      if (mediaInfo.hash === null) {
        const data = await this.#zipManager.readBinaryFile(mediaPath)
        if (data) {
          const hash = this.#hashBytes(data)
          mediaInfo.hash = hash
          mediaInfo.size = data.length
          this.#mediaHashIndex.set(hash, mediaPath)
        }
      }
    }
  }

  /**
   * Returns the total number of media files.
   * @returns {number}
   */
  get mediaCount() {
    return this.#mediaRegistry.size
  }

  /**
   * Embeds media (image, video, or audio) from file path, URL, data URI, Buffer, or Stream.
   * Deduplicates automatically using SHA-1 hashing.
   *
   * @param {string|Buffer|Uint8Array} source
   * @param {string} [defaultCategory='image']
   * @param {string} [mimeType]
   * @returns {Promise<string>} ZIP path of embedded media (e.g. 'ppt/media/image1.png')
   */
  async embedMedia(source, defaultCategory = 'image', mimeType) {
    const payload = await resolveMediaPayload(source, defaultCategory, mimeType)
    const { data, ext } = payload
    mimeType = payload.mimeType

    const hash = this.#hashBytes(data)
    const size = data.length

    if (this.#mediaHashIndex.has(hash)) {
      const existingPath = this.#mediaHashIndex.get(hash)
      logger.debug(`Reusing existing media: ${existingPath} (hash: ${hash.substring(0, 8)}...)`)
      return existingPath
    }

    await this.#ensureAllMediaHashed()

    if (this.#mediaHashIndex.has(hash)) {
      const existingPath = this.#mediaHashIndex.get(hash)
      logger.debug(`Reusing existing media: ${existingPath} (hash: ${hash.substring(0, 8)}...)`)
      return existingPath
    }

    const mediaId = this.#nextMediaId++
    const prefix =
      defaultCategory === 'video' ? 'media' : defaultCategory === 'audio' ? 'media' : 'image'
    const zipPath = `ppt/media/${prefix}${mediaId}.${ext}`

    this.#zipManager.writeBinaryFile(zipPath, data)
    this.#mediaHashIndex.set(hash, zipPath)
    this.#mediaRegistry.set(zipPath, { zipPath, hash, mimeType, size })

    this.#registerContentType(ext, mimeType)
    logger.debug(`Embedded new media: ${zipPath} (${size} bytes)`)
    return zipPath
  }

  /**
   * Embeds an image file. Delegates to embedMedia.
   *
   * @param {string|Buffer|Uint8Array} source
   * @param {string} [mimeType]
   * @returns {Promise<string>}
   */
  async embedImage(source, mimeType) {
    return this.embedMedia(source, 'image', mimeType)
  }

  /**
   * Copies an existing media part to a new unique file without deduplication.
   *
   * @param {string} sourceZipPath
   * @returns {Promise<string>}
   */
  async copyMediaAsNewPart(sourceZipPath) {
    const data = await this.#zipManager.readBinaryFile(sourceZipPath)
    if (!data) {
      throw new PPTXError(`Cannot clone slide: media not found at ${sourceZipPath}`)
    }

    const fileName = sourceZipPath.split('/').pop()
    const ext = fileName.includes('.') ? fileName.split('.').pop().toLowerCase() : 'png'
    const mimeType = EXT_TO_MIME[ext] || 'application/octet-stream'

    let mediaId = this.#nextMediaId
    let zipPath = `ppt/media/image${mediaId}.${ext}`
    while (this.#zipManager.hasFile(zipPath)) {
      mediaId++
      zipPath = `ppt/media/image${mediaId}.${ext}`
    }
    this.#nextMediaId = mediaId + 1

    this.#zipManager.writeBinaryFile(zipPath, data)
    this.#mediaRegistry.set(zipPath, {
      zipPath,
      hash: null,
      mimeType,
      size: data.length,
    })
    this.#registerContentType(ext, mimeType)

    logger.debug(`Cloned media to new part: ${zipPath}`)
    return zipPath
  }

  /**
   * Builds XML snippet for an image.
   *
   * @param {string} rId
   * @param {Object} opts
   * @returns {string}
   */
  buildImageXml(rId, opts = {}) {
    const { x = 0, y = 0, width = 2743200, height = 1828800, name = 'image', shapeId = 1 } = opts

    return `<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="${shapeId}" name="${name}"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rId}" cstate="print">
      <a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst>
    </a:blip>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="${x}" y="${y}"/>
      <a:ext cx="${width}" cy="${height}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`
  }

  /**
   * Builds XML snippet for video element.
   *
   * @param {string} rIdVideo
   * @param {string} rIdMedia
   * @param {string} rIdPoster
   * @param {Object} opts
   * @returns {string}
   */
  buildVideoXml(rIdVideo, rIdMedia, rIdPoster, opts = {}) {
    const {
      x = 0,
      y = 0,
      width = 5486400,
      height = 3086100,
      name = 'Video',
      shapeId = 1,
      loop = false,
      mute = false,
      startTime,
      endTime,
    } = opts

    let trimXml = ''
    if (startTime !== undefined || endTime !== undefined) {
      const st = startTime !== undefined ? Math.round(startTime * 1000) : 0
      const end = endTime !== undefined ? Math.round(endTime * 1000) : 0
      trimXml = `<p14:trim st="${st}" end="${end}"/>`
    }

    const loopAttr = loop ? ' loop="1"' : ''
    const muteAttr = mute ? ' mute="1"' : ''

    return `<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="${shapeId}" name="${name}">
      <a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="" action="ppaction://media"/>
    </p:cNvPr>
    <p:cNvPicPr>
      <a:picLocks noChangeAspect="1"/>
    </p:cNvPicPr>
    <p:nvPr>
      <a:videoFile xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:link="${rIdVideo}"/>
      <p:extLst>
        <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
          <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rIdMedia}"${loopAttr}${muteAttr}>
            ${trimXml}
          </p14:media>
        </p:ext>
      </p:extLst>
    </p:nvPr>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rIdPoster}" cstate="print"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="${x}" y="${y}"/>
      <a:ext cx="${width}" cy="${height}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`
  }

  /**
   * Builds XML snippet for audio element.
   *
   * @param {string} rIdAudio
   * @param {string} rIdMedia
   * @param {string} rIdIcon
   * @param {Object} opts
   * @returns {string}
   */
  buildAudioXml(rIdAudio, rIdMedia, rIdIcon, opts = {}) {
    const {
      x = 457200,
      y = 457200,
      width = 731520,
      height = 731520,
      name = 'Audio',
      shapeId = 1,
      loop = false,
      volume = 100,
      startTime,
      endTime,
    } = opts

    let trimXml = ''
    if (startTime !== undefined || endTime !== undefined) {
      const st = startTime !== undefined ? Math.round(startTime * 1000) : 0
      const end = endTime !== undefined ? Math.round(endTime * 1000) : 0
      trimXml = `<p14:trim st="${st}" end="${end}"/>`
    }

    const loopAttr = loop ? ' loop="1"' : ''
    const volVal = Math.round((volume / 100) * 100000)

    return `<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="${shapeId}" name="${name}">
      <a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="" action="ppaction://media"/>
    </p:cNvPr>
    <p:cNvPicPr>
      <a:picLocks noChangeAspect="1"/>
    </p:cNvPicPr>
    <p:nvPr>
      <a:audioFile xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:link="${rIdAudio}"/>
      <p:extLst>
        <p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}">
          <p14:media xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rIdMedia}" vol="${volVal}"${loopAttr}>
            ${trimXml}
          </p14:media>
        </p:ext>
      </p:extLst>
    </p:nvPr>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rIdIcon}" cstate="print"/>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="${x}" y="${y}"/>
      <a:ext cx="${width}" cy="${height}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`
  }

  /**
   * Embeds and attaches a video to a slide.
   *
   * @param {number} slideIndex
   * @param {string|Buffer|Uint8Array} source
   * @param {Object} options
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @param {XMLParser} xmlParser
   */
  async addVideo(slideIndex, source, options = {}, slideManager, relationshipManager, xmlParser) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) throw new PPTXError(`Slide ${slideIndex} not found`)

    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']
    if (!spTree) throw new PPTXError(`Invalid slide structure for slide ${slideIndex}`)

    // 1. Embed video stream
    const destVideoZipPath = await this.embedMedia(source, 'video')
    const relativeVideoTarget = `../media/${destVideoZipPath.split('/').pop()}`

    // 2. Add video & media relationships
    const rIdVideo = relationshipManager.addRelationship(
      slideInfo.zipPath,
      REL_TYPES.VIDEO,
      relativeVideoTarget
    )
    const rIdMedia = relationshipManager.addRelationship(
      slideInfo.zipPath,
      REL_TYPES.MEDIA,
      relativeVideoTarget
    )

    // 3. Poster image
    let posterSource = options.poster || options.posterFrame
    if (!posterSource) {
      posterSource = DEFAULT_POSTER_PNG
    }
    const destPosterZipPath = await this.embedMedia(posterSource, 'image')
    const relativePosterTarget = `../media/${destPosterZipPath.split('/').pop()}`
    const rIdPoster = relationshipManager.addRelationship(
      slideInfo.zipPath,
      REL_TYPES.IMAGE,
      relativePosterTarget
    )

    // 4. Generate shape ID & resolve bounds
    const existingIds = this.#getAllShapeIds(spTree)
    const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 1000
    const newId = maxId + 1

    const bounds = resolveElementBounds(
      {
        x: options.x !== undefined ? options.x : 0,
        y: options.y !== undefined ? options.y : 0,
        width: options.width !== undefined ? options.width : '50%',
        height: options.height !== undefined ? options.height : '50%',
        unit: options.unit,
        isPercent: options.isPercent,
      },
      slideManager
    )

    const name = options.name || `Video ${newId}`

    const videoXml = this.buildVideoXml(rIdVideo, rIdMedia, rIdPoster, {
      x: bounds.x,
      y: bounds.y,
      width: bounds.width,
      height: bounds.height,
      name,
      shapeId: newId,
      loop: Boolean(options.loop),
      mute: Boolean(options.mute),
      startTime: options.startTime,
      endTime: options.endTime,
    })

    const parsedRaw = xmlParser.parse(videoXml, 'video.xml')['p:pic']
    const picObj = Array.isArray(parsedRaw) ? parsedRaw[0] : parsedRaw

    if (!spTree['p:pic']) spTree['p:pic'] = []
    if (!Array.isArray(spTree['p:pic'])) spTree['p:pic'] = [spTree['p:pic']]
    spTree['p:pic'].push(picObj)
    ensureSpTreeOrder(spTree)

    this.#ensureVideoTiming(slideObj, newId, options)

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Added video "${name}" (ID: ${newId}) to slide ${slideIndex}`)
    return { id: newId, name, rIdVideo, rIdMedia }
  }

  /**
   * Alias for addVideo.
   */
  async AddVideo(...args) {
    return this.addVideo(...args)
  }

  /**
   * Updates an existing video shape and playback properties.
   *
   * @param {number} slideIndex
   * @param {string|number} videoId
   * @param {Object} options
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @returns {Promise<boolean>} True if found and updated.
   */
  async updateVideo(slideIndex, videoId, options = {}, slideManager, relationshipManager) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) return false
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']
    if (!spTree) return false

    const targetIdStr = String(videoId)
    const match = this.#findVideoRecursive(spTree, targetIdStr)
    if (!match) return false

    const { pic } = match

    // 0. Source swap
    if (options.source) {
      const destVideoZipPath = await this.embedMedia(options.source, 'video')
      const relativeVideoTarget = `../media/${destVideoZipPath.split('/').pop()}`
      const rIdVideo = pic?.['p:nvPicPr']?.['p:nvPr']?.['a:videoFile']?.['@_r:link']
      let rIdMedia =
        pic?.['p:nvPicPr']?.['p:nvPr']?.['p:extLst']?.['p:ext']?.['p14:media']?.['@_r:embed']
      if (!rIdMedia && pic?.['p:nvPicPr']?.['p:nvPr']?.['p:extLst']?.['p:ext']) {
        const extList = pic['p:nvPicPr']['p:nvPr']['p:extLst']['p:ext']
        if (Array.isArray(extList)) {
          for (const e of extList) {
            if (e['p14:media']) {
              rIdMedia = e['p14:media']['@_r:embed']
              break
            }
          }
        }
      }
      if (rIdVideo) {
        relationshipManager.updateRelationshipTarget(
          slideInfo.zipPath,
          rIdVideo,
          relativeVideoTarget
        )
      }
      if (rIdMedia && rIdMedia !== rIdVideo) {
        relationshipManager.updateRelationshipTarget(
          slideInfo.zipPath,
          rIdMedia,
          relativeVideoTarget
        )
      }
    }

    // 1. Spatial bounds
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
          'a:ext': { '@_cx': '5486400', '@_cy': '3086100' },
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

    // 2. Poster frame update
    if (options.poster || options.posterFrame) {
      const posterSource = options.poster || options.posterFrame
      const destPosterZipPath = await this.embedMedia(posterSource, 'image')
      const relativePosterTarget = `../media/${destPosterZipPath.split('/').pop()}`
      const existingPosterRId = pic?.['p:blipFill']?.['a:blip']?.['@_r:embed']
      const newRId = relationshipManager.addRelationship(
        slideInfo.zipPath,
        REL_TYPES.IMAGE,
        relativePosterTarget
      )
      if (!pic['p:blipFill']) pic['p:blipFill'] = {}
      if (!pic['p:blipFill']['a:blip']) pic['p:blipFill']['a:blip'] = {}
      pic['p:blipFill']['a:blip']['@_r:embed'] = newRId
      if (
        existingPosterRId &&
        existingPosterRId !== newRId &&
        !this.#isRelationshipUsedInTree(spTree, existingPosterRId)
      ) {
        relationshipManager.removeRelationship(slideInfo.zipPath, existingPosterRId)
      }
    }

    // 3. Playback options in p14:media
    let p14Media = pic['p:nvPicPr']?.['p:nvPr']?.['p:extLst']?.['p:ext']?.['p14:media']
    if (!p14Media && pic['p:nvPicPr']?.['p:nvPr']?.['p:extLst']?.['p:ext']) {
      const ext = pic['p:nvPicPr']['p:nvPr']['p:extLst']['p:ext']
      if (Array.isArray(ext)) {
        for (const e of ext) {
          if (e['p14:media']) {
            p14Media = e['p14:media']
            break
          }
        }
      }
    }

    if (p14Media) {
      if (options.loop !== undefined) {
        if (options.loop) p14Media['@_loop'] = '1'
        else delete p14Media['@_loop']
      }
      if (options.mute !== undefined) {
        if (options.mute) p14Media['@_mute'] = '1'
        else delete p14Media['@_mute']
      }
      if (options.startTime !== undefined || options.endTime !== undefined) {
        const st = options.startTime !== undefined ? Math.round(options.startTime * 1000) : 0
        const end = options.endTime !== undefined ? Math.round(options.endTime * 1000) : 0
        p14Media['p14:trim'] = { '@_st': String(st), '@_end': String(end) }
      }
    }

    // 4. Timing update
    const shapeId = pic?.['p:nvPicPr']?.['p:cNvPr']?.['@_id']
    if (shapeId && (options.autoplay !== undefined || options.volume !== undefined)) {
      this.#ensureVideoTiming(slideObj, shapeId, options)
    }

    ensureSpTreeOrder(spTree)
    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Updated video "${videoId}" on slide ${slideIndex}`)
    return true
  }

  /**
   * Alias for updateVideo.
   */
  async UpdateVideo(...args) {
    return this.updateVideo(...args)
  }

  /**
   * Removes a video shape and associated relationships/timing.
   *
   * @param {number} slideIndex
   * @param {string|number} videoId
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @returns {boolean} True if removed.
   */
  removeVideo(slideIndex, videoId, slideManager, relationshipManager) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) return false
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']
    if (!spTree) return false

    const targetIdStr = String(videoId)
    const match = this.#findVideoRecursive(spTree, targetIdStr)
    if (!match) return false

    const { pic, parent } = match
    const shapeId = pic?.['p:nvPicPr']?.['p:cNvPr']?.['@_id']

    // Remove from shape array
    if (parent['p:pic']) {
      if (Array.isArray(parent['p:pic'])) {
        parent['p:pic'] = parent['p:pic'].filter(p => p !== pic)
      } else if (parent['p:pic'] === pic) {
        delete parent['p:pic']
      }
    }

    // Clean up relationships
    const rIdVideo = pic?.['p:nvPicPr']?.['p:nvPr']?.['a:videoFile']?.['@_r:link']
    let rIdMedia =
      pic?.['p:nvPicPr']?.['p:nvPr']?.['p:extLst']?.['p:ext']?.['p14:media']?.['@_r:embed']
    if (!rIdMedia && pic?.['p:nvPicPr']?.['p:nvPr']?.['p:extLst']?.['p:ext']) {
      const extList = pic['p:nvPicPr']['p:nvPr']['p:extLst']['p:ext']
      if (Array.isArray(extList)) {
        for (const e of extList) {
          if (e['p14:media']) {
            rIdMedia = e['p14:media']['@_r:embed']
            break
          }
        }
      }
    }
    const rIdPoster = pic?.['p:blipFill']?.['a:blip']?.['@_r:embed']

    if (rIdVideo) relationshipManager.removeRelationship(slideInfo.zipPath, rIdVideo)
    if (rIdMedia && rIdMedia !== rIdVideo)
      relationshipManager.removeRelationship(slideInfo.zipPath, rIdMedia)
    if (rIdPoster && !this.#isRelationshipUsedInTree(spTree, rIdPoster)) {
      relationshipManager.removeRelationship(slideInfo.zipPath, rIdPoster)
    }

    // Remove timing node
    if (shapeId) {
      this.#removeSlideTiming(slideObj, shapeId)
    }

    ensureSpTreeOrder(spTree)
    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Removed video "${videoId}" from slide ${slideIndex}`)
    return true
  }

  /**
   * Alias for removeVideo.
   */
  RemoveVideo(...args) {
    return this.removeVideo(...args)
  }

  /**
   * Embeds and attaches an audio file to a slide.
   *
   * @param {number} slideIndex
   * @param {string|Buffer|Uint8Array} source
   * @param {Object} options
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @param {XMLParser} xmlParser
   */
  async addAudio(slideIndex, source, options = {}, slideManager, relationshipManager, xmlParser) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) throw new PPTXError(`Slide ${slideIndex} not found`)

    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']
    if (!spTree) throw new PPTXError(`Invalid slide structure for slide ${slideIndex}`)

    // 1. Embed audio stream
    const destAudioZipPath = await this.embedMedia(source, 'audio')
    const relativeAudioTarget = `../media/${destAudioZipPath.split('/').pop()}`

    // 2. Add audio & media relationships
    const rIdAudio = relationshipManager.addRelationship(
      slideInfo.zipPath,
      REL_TYPES.AUDIO,
      relativeAudioTarget
    )
    const rIdMedia = relationshipManager.addRelationship(
      slideInfo.zipPath,
      REL_TYPES.MEDIA,
      relativeAudioTarget
    )

    // 3. Audio icon poster
    let iconSource = options.icon || options.poster
    if (!iconSource) {
      iconSource = getDefaultAudioIcon()
    }
    const destIconZipPath = await this.embedMedia(iconSource, 'image')
    const relativeIconTarget = `../media/${destIconZipPath.split('/').pop()}`
    const rIdIcon = relationshipManager.addRelationship(
      slideInfo.zipPath,
      REL_TYPES.IMAGE,
      relativeIconTarget
    )

    // 4. Shape ID & bounds
    const existingIds = this.#getAllShapeIds(spTree)
    const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 1000
    const newId = maxId + 1

    const bounds = resolveElementBounds(
      {
        x: options.x !== undefined ? options.x : 457200,
        y: options.y !== undefined ? options.y : 457200,
        width: options.width !== undefined ? options.width : 731520,
        height: options.height !== undefined ? options.height : 731520,
        unit: options.unit,
        isPercent: options.isPercent,
      },
      slideManager
    )

    const name = options.name || `Audio ${newId}`

    const audioXml = this.buildAudioXml(rIdAudio, rIdMedia, rIdIcon, {
      x: bounds.x,
      y: bounds.y,
      width: bounds.width,
      height: bounds.height,
      name,
      shapeId: newId,
      loop: Boolean(options.loop),
      volume: options.volume !== undefined ? options.volume : 100,
      startTime: options.startTime,
      endTime: options.endTime,
    })

    const parsedRaw = xmlParser.parse(audioXml, 'audio.xml')['p:pic']
    const picObj = Array.isArray(parsedRaw) ? parsedRaw[0] : parsedRaw

    if (!spTree['p:pic']) spTree['p:pic'] = []
    if (!Array.isArray(spTree['p:pic'])) spTree['p:pic'] = [spTree['p:pic']]
    spTree['p:pic'].push(picObj)
    ensureSpTreeOrder(spTree)

    this.#ensureAudioTiming(slideObj, newId, options)

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Added audio "${name}" (ID: ${newId}) to slide ${slideIndex}`)
    return { id: newId, name, rIdAudio, rIdMedia }
  }

  /**
   * Alias for addAudio.
   */
  async AddAudio(...args) {
    return this.addAudio(...args)
  }

  /**
   * Updates an existing audio element.
   *
   * @param {number} slideIndex
   * @param {string|number} audioId
   * @param {Object} options
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @returns {Promise<boolean>} True if updated.
   */
  async updateAudio(slideIndex, audioId, options = {}, slideManager, relationshipManager) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) return false
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']
    if (!spTree) return false

    const targetIdStr = String(audioId)
    const match = this.#findAudioRecursive(spTree, targetIdStr)
    if (!match) return false

    const { pic } = match

    // 0. Source swap
    if (options.source && relationshipManager) {
      const destAudioZipPath = await this.embedMedia(options.source, 'audio')
      const relativeAudioTarget = `../media/${destAudioZipPath.split('/').pop()}`
      const rIdAudio = pic?.['p:nvPicPr']?.['p:nvPr']?.['a:audioFile']?.['@_r:link']
      let rIdMedia =
        pic?.['p:nvPicPr']?.['p:nvPr']?.['p:extLst']?.['p:ext']?.['p14:media']?.['@_r:embed']
      if (!rIdMedia && pic?.['p:nvPicPr']?.['p:nvPr']?.['p:extLst']?.['p:ext']) {
        const extList = pic['p:nvPicPr']['p:nvPr']['p:extLst']['p:ext']
        if (Array.isArray(extList)) {
          for (const e of extList) {
            if (e['p14:media']) {
              rIdMedia = e['p14:media']['@_r:embed']
              break
            }
          }
        }
      }
      if (rIdAudio) {
        relationshipManager.updateRelationshipTarget(
          slideInfo.zipPath,
          rIdAudio,
          relativeAudioTarget
        )
      }
      if (rIdMedia && rIdMedia !== rIdAudio) {
        relationshipManager.updateRelationshipTarget(
          slideInfo.zipPath,
          rIdMedia,
          relativeAudioTarget
        )
      }
    }

    // 1. Spatial bounds
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
          'a:off': { '@_x': '457200', '@_y': '457200' },
          'a:ext': { '@_cx': '731520', '@_cy': '731520' },
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

    // 2. Playback settings in p14:media
    let p14Media = pic['p:nvPicPr']?.['p:nvPr']?.['p:extLst']?.['p:ext']?.['p14:media']
    if (!p14Media && pic['p:nvPicPr']?.['p:nvPr']?.['p:extLst']?.['p:ext']) {
      const ext = pic['p:nvPicPr']['p:nvPr']['p:extLst']['p:ext']
      if (Array.isArray(ext)) {
        for (const e of ext) {
          if (e['p14:media']) {
            p14Media = e['p14:media']
            break
          }
        }
      }
    }

    if (p14Media) {
      if (options.loop !== undefined) {
        if (options.loop) p14Media['@_loop'] = '1'
        else delete p14Media['@_loop']
      }
      if (options.volume !== undefined) {
        const volVal = Math.round((options.volume / 100) * 100000)
        p14Media['@_vol'] = String(volVal)
      }
      if (options.startTime !== undefined || options.endTime !== undefined) {
        const st = options.startTime !== undefined ? Math.round(options.startTime * 1000) : 0
        const end = options.endTime !== undefined ? Math.round(options.endTime * 1000) : 0
        p14Media['p14:trim'] = { '@_st': String(st), '@_end': String(end) }
      }
    }

    // 3. Timing and visibility update
    if (pic?.['p:nvPicPr']?.['p:cNvPicPr']?.['p14:showWhenStopped']) {
      delete pic['p:nvPicPr']['p:cNvPicPr']['p14:showWhenStopped']
    }

    const shapeId = pic?.['p:nvPicPr']?.['p:cNvPr']?.['@_id']
    if (
      shapeId &&
      (options.autoplay !== undefined ||
        options.volume !== undefined ||
        options.hideDuringShow !== undefined ||
        options.hideIcon !== undefined)
    ) {
      this.#ensureAudioTiming(slideObj, shapeId, options)
    }

    ensureSpTreeOrder(spTree)
    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Updated audio "${audioId}" on slide ${slideIndex}`)
    return true
  }

  /**
   * Alias for updateAudio.
   */
  async UpdateAudio(...args) {
    return this.updateAudio(...args)
  }

  /**
   * Removes an audio shape and associated relationships.
   *
   * @param {number} slideIndex
   * @param {string|number} audioId
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @returns {boolean} True if removed.
   */
  removeAudio(slideIndex, audioId, slideManager, relationshipManager) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) return false
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']
    if (!spTree) return false

    const targetIdStr = String(audioId)
    const match = this.#findAudioRecursive(spTree, targetIdStr)
    if (!match) return false

    const { pic, parent } = match
    const shapeId = pic?.['p:nvPicPr']?.['p:cNvPr']?.['@_id']

    if (parent['p:pic']) {
      if (Array.isArray(parent['p:pic'])) {
        parent['p:pic'] = parent['p:pic'].filter(p => p !== pic)
      } else if (parent['p:pic'] === pic) {
        delete parent['p:pic']
      }
    }

    const rIdAudio = pic?.['p:nvPicPr']?.['p:nvPr']?.['a:audioFile']?.['@_r:link']
    let rIdMedia =
      pic?.['p:nvPicPr']?.['p:nvPr']?.['p:extLst']?.['p:ext']?.['p14:media']?.['@_r:embed']
    if (!rIdMedia && pic?.['p:nvPicPr']?.['p:nvPr']?.['p:extLst']?.['p:ext']) {
      const extList = pic['p:nvPicPr']['p:nvPr']['p:extLst']['p:ext']
      if (Array.isArray(extList)) {
        for (const e of extList) {
          if (e['p14:media']) {
            rIdMedia = e['p14:media']['@_r:embed']
            break
          }
        }
      }
    }
    const rIdIcon = pic?.['p:blipFill']?.['a:blip']?.['@_r:embed']

    if (rIdAudio) relationshipManager.removeRelationship(slideInfo.zipPath, rIdAudio)
    if (rIdMedia && rIdMedia !== rIdAudio)
      relationshipManager.removeRelationship(slideInfo.zipPath, rIdMedia)
    if (rIdIcon && !this.#isRelationshipUsedInTree(spTree, rIdIcon)) {
      relationshipManager.removeRelationship(slideInfo.zipPath, rIdIcon)
    }

    if (shapeId) {
      this.#removeSlideTiming(slideObj, shapeId)
    }

    ensureSpTreeOrder(spTree)
    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Removed audio "${audioId}" from slide ${slideIndex}`)
    return true
  }

  /**
   * Alias for removeAudio.
   */
  RemoveAudio(...args) {
    return this.removeAudio(...args)
  }

  /**
   * Queries all media items (images, videos, audio) on a slide.
   *
   * @param {number} slideIndex
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @returns {Array<Object>} List of media metadata items.
   */
  getMedias(slideIndex, slideManager, relationshipManager) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) return []
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']

    const medias = []
    this.#collectMedias(spTree, slideInfo.zipPath, relationshipManager, medias, slideIndex)
    return medias
  }

  /**
   * Alias for getMedias.
   */
  GetMedias(...args) {
    return this.getMedias(...args)
  }

  /**
   * Retrieves metadata for a specific media item by ID.
   *
   * @param {string|number} mediaId
   * @param {number} slideIndex
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @returns {Object|null}
   */
  getMedia(mediaId, slideIndex, slideManager, relationshipManager) {
    const medias = this.getMedias(slideIndex, slideManager, relationshipManager)
    const targetStr = String(mediaId)
    return (
      medias.find(
        m =>
          m.id === targetStr ||
          m.name === targetStr ||
          m.relationshipId === targetStr ||
          m.rId === targetStr
      ) || null
    )
  }

  /**
   * Alias for getMedia.
   */
  GetMedia(...args) {
    return this.getMedia(...args)
  }

  /**
   * Duplicates an existing media shape on the slide while deduplicating underlying binary assets.
   *
   * @param {string|number} mediaId
   * @param {string|number} [clonedMediaId]
   * @param {Object} [clonedOptions={}]
   * @param {number} slideIndex
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @param {XMLParser} xmlParser
   * @returns {boolean} True if cloned.
   */
  cloneMedia(
    mediaId,
    clonedMediaId,
    clonedOptions = {},
    slideIndex,
    slideManager,
    _relationshipManager,
    _xmlParser
  ) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) return false
    const slideObj = slideManager.getSlideObj(slideIndex)
    const spTree =
      slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
      slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']
    if (!spTree) return false

    const targetIdStr = String(mediaId)
    const match = this.#findMediaRecursive(spTree, targetIdStr)
    if (!match) return false

    const { pic } = match

    // Deep clone AST
    const clonedPic = JSON.parse(JSON.stringify(pic))

    // Assign new unique shape ID
    const existingIds = this.#getAllShapeIds(spTree)
    const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 1000
    const newId = maxId + 1

    const newName =
      clonedMediaId !== undefined && clonedMediaId !== null
        ? String(clonedMediaId)
        : clonedOptions.name || `Clone of ${pic?.['p:nvPicPr']?.['p:cNvPr']?.['@_name'] || 'Media'}`

    if (!clonedPic['p:nvPicPr']) clonedPic['p:nvPicPr'] = {}
    if (!clonedPic['p:nvPicPr']['p:cNvPr']) clonedPic['p:nvPicPr']['p:cNvPr'] = {}
    clonedPic['p:nvPicPr']['p:cNvPr']['@_id'] = String(newId)
    clonedPic['p:nvPicPr']['p:cNvPr']['@_name'] = newName
    if (clonedPic['p:nvPicPr']['p:cNvPr']['a:extLst']) {
      delete clonedPic['p:nvPicPr']['p:cNvPr']['a:extLst']
    }

    // Apply spatial overrides
    if (
      clonedOptions.x !== undefined ||
      clonedOptions.y !== undefined ||
      clonedOptions.width !== undefined ||
      clonedOptions.height !== undefined ||
      clonedOptions.cx !== undefined ||
      clonedOptions.cy !== undefined ||
      clonedOptions.unit ||
      clonedOptions.isPercent
    ) {
      if (!clonedPic['p:spPr']) clonedPic['p:spPr'] = {}
      if (!clonedPic['p:spPr']['a:xfrm']) {
        clonedPic['p:spPr']['a:xfrm'] = {
          'a:off': { '@_x': '0', '@_y': '0' },
          'a:ext': { '@_cx': '2743200', '@_cy': '1828800' },
        }
      }
      const xfrm = clonedPic['p:spPr']['a:xfrm']
      const bounds = resolveElementBounds(
        {
          x: clonedOptions.x !== undefined ? clonedOptions.x : xfrm['a:off']?.['@_x'],
          y: clonedOptions.y !== undefined ? clonedOptions.y : xfrm['a:off']?.['@_y'],
          width:
            clonedOptions.width !== undefined
              ? clonedOptions.width
              : clonedOptions.cx !== undefined
                ? clonedOptions.cx
                : xfrm['a:ext']?.['@_cx'],
          height:
            clonedOptions.height !== undefined
              ? clonedOptions.height
              : clonedOptions.cy !== undefined
                ? clonedOptions.cy
                : xfrm['a:ext']?.['@_cy'],
          unit: clonedOptions.unit,
          isPercent: clonedOptions.isPercent,
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

    if (clonedOptions.rotation !== undefined) {
      if (!clonedPic['p:spPr']) clonedPic['p:spPr'] = {}
      if (!clonedPic['p:spPr']['a:xfrm']) clonedPic['p:spPr']['a:xfrm'] = {}
      clonedPic['p:spPr']['a:xfrm']['@_rot'] = String(Math.round(clonedOptions.rotation * 60000))
    }

    if (!spTree['p:pic']) spTree['p:pic'] = []
    if (!Array.isArray(spTree['p:pic'])) spTree['p:pic'] = [spTree['p:pic']]
    spTree['p:pic'].push(clonedPic)
    ensureSpTreeOrder(spTree)

    const isVideo = Boolean(clonedPic?.['p:nvPicPr']?.['p:nvPr']?.['a:videoFile'])
    const isAudio = Boolean(clonedPic?.['p:nvPicPr']?.['p:nvPr']?.['a:audioFile'])
    if (isVideo) {
      this.#ensureVideoTiming(slideObj, newId, clonedOptions)
    } else if (isAudio) {
      this.#ensureAudioTiming(slideObj, newId, clonedOptions)
    }

    slideManager.markSlideObjDirty(slideIndex)
    logger.debug(`Cloned media "${mediaId}" as "${newName}" (ID: ${newId}) on slide ${slideIndex}`)
    return true
  }

  /**
   * Alias for cloneMedia.
   */
  CloneMedia(...args) {
    return this.cloneMedia(...args)
  }

  #collectMedias(container, slideZipPath, relationshipManager, results, slideIndex) {
    if (!container) return

    let pics = container['p:pic'] || []
    if (!Array.isArray(pics)) pics = [pics]

    for (const pic of pics) {
      const cNvPr = pic?.['p:nvPicPr']?.['p:cNvPr']
      if (!cNvPr) continue

      const name = cNvPr['@_name'] || ''
      const id = String(cNvPr['@_id'] || '')

      const isVideo = Boolean(pic?.['p:nvPicPr']?.['p:nvPr']?.['a:videoFile'])
      const isAudio = Boolean(pic?.['p:nvPicPr']?.['p:nvPr']?.['a:audioFile'])
      const type = isVideo ? 'video' : isAudio ? 'audio' : 'image'

      let rId = ''
      if (isVideo) {
        rId = pic?.['p:nvPicPr']?.['p:nvPr']?.['a:videoFile']?.['@_r:link'] || ''
      } else if (isAudio) {
        rId = pic?.['p:nvPicPr']?.['p:nvPr']?.['a:audioFile']?.['@_r:link'] || ''
      } else {
        rId = pic?.['p:blipFill']?.['a:blip']?.['@_r:embed'] || ''
      }

      let targetPath = ''
      if (rId) {
        const rel = relationshipManager.getRelationshipById(slideZipPath, rId)
        if (rel) {
          targetPath = relationshipManager.resolveTarget(slideZipPath, rel.target)
        }
      }

      const ext = targetPath ? targetPath.split('.').pop().toLowerCase() : ''
      const mediaInfo = targetPath ? this.getMediaInfo(targetPath) : null
      const xfrm = pic['p:spPr']?.['a:xfrm']

      const x = parseInt(xfrm?.['a:off']?.['@_x'] || 0, 10)
      const y = parseInt(xfrm?.['a:off']?.['@_y'] || 0, 10)
      const cx = parseInt(xfrm?.['a:ext']?.['@_cx'] || 0, 10)
      const cy = parseInt(xfrm?.['a:ext']?.['@_cy'] || 0, 10)

      const position = { x, y, width: cx, height: cy, cx, cy }
      const rotVal = xfrm?.['@_rot'] ? parseInt(xfrm['@_rot'], 10) : 0
      const rotation = rotVal ? Math.round(rotVal / 60000) : 0

      const item = {
        id,
        name,
        type,
        source: targetPath,
        targetPath,
        relationshipId: rId,
        rId,
        position,
        x,
        y,
        width: cx,
        height: cy,
        rotation,
        format: ext,
        mimeType: mediaInfo?.mimeType || EXT_TO_MIME[ext] || undefined,
        size: mediaInfo?.size,
        slideIndex,
      }

      // Check p14:media settings
      let p14Media = pic?.['p:nvPicPr']?.['p:nvPr']?.['p:extLst']?.['p:ext']?.['p14:media']
      if (!p14Media && pic?.['p:nvPicPr']?.['p:nvPr']?.['p:extLst']?.['p:ext']) {
        const extList = pic['p:nvPicPr']['p:nvPr']['p:extLst']['p:ext']
        if (Array.isArray(extList)) {
          for (const e of extList) {
            if (e['p14:media']) {
              p14Media = e['p14:media']
              break
            }
          }
        }
      }

      if (p14Media) {
        if (p14Media['@_loop'] === '1') item.loop = true
        if (p14Media['@_mute'] === '1') item.mute = true
        if (p14Media['@_vol'])
          item.volume = Math.round((parseInt(p14Media['@_vol'], 10) / 100000) * 100)
        if (p14Media['p14:trim']) {
          item.startTime = parseInt(p14Media['p14:trim']['@_st'] || 0, 10) / 1000
          item.endTime = parseInt(p14Media['p14:trim']['@_end'] || 0, 10) / 1000
        }
      }

      results.push(item)
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const g of groups) {
      this.#collectMedias(g, slideZipPath, relationshipManager, results, slideIndex)
    }
  }

  #findMediaRecursive(container, targetId) {
    if (!container) return null

    let pics = container['p:pic'] || []
    if (!Array.isArray(pics)) pics = [pics]

    for (const pic of pics) {
      const cNvPr = pic?.['p:nvPicPr']?.['p:cNvPr']
      if (cNvPr) {
        const name = cNvPr['@_name']
        const id = String(cNvPr['@_id'])
        const embedId = pic?.['p:blipFill']?.['a:blip']?.['@_r:embed']
        const videoId = pic?.['p:nvPicPr']?.['p:nvPr']?.['a:videoFile']?.['@_r:link']
        const audioId = pic?.['p:nvPicPr']?.['p:nvPr']?.['a:audioFile']?.['@_r:link']
        if (
          name === targetId ||
          id === targetId ||
          embedId === targetId ||
          videoId === targetId ||
          audioId === targetId
        ) {
          return { pic, parent: container }
        }
      }
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const group of groups) {
      const res = this.#findMediaRecursive(group, targetId)
      if (res) return res
    }

    return null
  }

  #findVideoRecursive(container, targetId) {
    if (!container) return null

    let pics = container['p:pic'] || []
    if (!Array.isArray(pics)) pics = [pics]

    for (const pic of pics) {
      if (!pic?.['p:nvPicPr']?.['p:nvPr']?.['a:videoFile']) continue
      const cNvPr = pic?.['p:nvPicPr']?.['p:cNvPr']
      if (cNvPr) {
        const name = cNvPr['@_name']
        const id = String(cNvPr['@_id'])
        const linkId = pic['p:nvPicPr']['p:nvPr']['a:videoFile']['@_r:link']
        if (name === targetId || id === targetId || linkId === targetId) {
          return { pic, parent: container }
        }
      }
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const group of groups) {
      const res = this.#findVideoRecursive(group, targetId)
      if (res) return res
    }

    return null
  }

  #findAudioRecursive(container, targetId) {
    if (!container) return null

    let pics = container['p:pic'] || []
    if (!Array.isArray(pics)) pics = [pics]

    for (const pic of pics) {
      if (!pic?.['p:nvPicPr']?.['p:nvPr']?.['a:audioFile']) continue
      const cNvPr = pic?.['p:nvPicPr']?.['p:cNvPr']
      if (cNvPr) {
        const name = cNvPr['@_name']
        const id = String(cNvPr['@_id'])
        const linkId = pic['p:nvPicPr']['p:nvPr']['a:audioFile']['@_r:link']
        if (name === targetId || id === targetId || linkId === targetId) {
          return { pic, parent: container }
        }
      }
    }

    let groups = container['p:grpSp'] || []
    if (!Array.isArray(groups)) groups = [groups]
    for (const group of groups) {
      const res = this.#findAudioRecursive(group, targetId)
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

  #getNextTimingId(slideObj) {
    let maxId = 0
    const timing = slideObj?.['p:sld']?.['p:timing']
    if (!timing) return 1
    const scan = node => {
      if (!node || typeof node !== 'object') return
      if (node['@_id'] && !isNaN(parseInt(node['@_id'], 10))) {
        maxId = Math.max(maxId, parseInt(node['@_id'], 10))
      }
      for (const key of Object.keys(node)) {
        const val = node[key]
        if (Array.isArray(val)) {
          for (const item of val) scan(item)
        } else if (typeof val === 'object') {
          scan(val)
        }
      }
    }
    scan(timing)
    return maxId + 1
  }

  #ensureBaseTiming(slideObj) {
    if (!slideObj) return null
    if (!slideObj['p:sld']) slideObj['p:sld'] = {}
    if (!slideObj['p:sld']['p:timing']) {
      slideObj['p:sld']['p:timing'] = {
        'p:tnLst': {
          'p:par': {
            'p:cTn': {
              '@_id': '1',
              '@_dur': 'indefinite',
              '@_restart': 'never',
              '@_nodeType': 'tmRoot',
              'p:childTnLst': {
                'p:seq': [
                  {
                    '@_concurrent': '1',
                    '@_nextAc': 'seek',
                    'p:cTn': {
                      '@_id': '2',
                      '@_dur': 'indefinite',
                      '@_nodeType': 'mainSeq',
                      'p:childTnLst': {
                        'p:par': [],
                      },
                    },
                    'p:prevCondLst': {
                      'p:cond': {
                        '@_evt': 'onPrev',
                        '@_delay': '0',
                        'p:tgtEl': { 'p:sldTgt': {} },
                      },
                    },
                    'p:nextCondLst': {
                      'p:cond': {
                        '@_evt': 'onNext',
                        '@_delay': '0',
                        'p:tgtEl': { 'p:sldTgt': {} },
                      },
                    },
                  },
                ],
              },
            },
          },
        },
      }
    }

    const timing = slideObj['p:sld']['p:timing']
    if (!timing['p:tnLst']) timing['p:tnLst'] = {}
    if (!timing['p:tnLst']['p:par']) timing['p:tnLst']['p:par'] = {}
    const rootPar = timing['p:tnLst']['p:par']
    if (!rootPar['p:cTn']) {
      rootPar['p:cTn'] = {
        '@_id': '1',
        '@_dur': 'indefinite',
        '@_restart': 'never',
        '@_nodeType': 'tmRoot',
        'p:childTnLst': {},
      }
    }
    if (!rootPar['p:cTn']['p:childTnLst']) {
      rootPar['p:cTn']['p:childTnLst'] = {}
    }
    const rootChildTnLst = rootPar['p:cTn']['p:childTnLst']

    if (!rootChildTnLst['p:seq']) {
      rootChildTnLst['p:seq'] = []
    } else if (!Array.isArray(rootChildTnLst['p:seq'])) {
      rootChildTnLst['p:seq'] = [rootChildTnLst['p:seq']]
    }

    let mainSeq = rootChildTnLst['p:seq'].find(s => s?.['p:cTn']?.['@_nodeType'] === 'mainSeq')
    if (!mainSeq) {
      const maxId = this.#getNextTimingId(slideObj)
      mainSeq = {
        '@_concurrent': '1',
        '@_nextAc': 'seek',
        'p:cTn': {
          '@_id': String(maxId),
          '@_dur': 'indefinite',
          '@_nodeType': 'mainSeq',
          'p:childTnLst': {
            'p:par': [],
          },
        },
        'p:prevCondLst': {
          'p:cond': {
            '@_evt': 'onPrev',
            '@_delay': '0',
            'p:tgtEl': { 'p:sldTgt': {} },
          },
        },
        'p:nextCondLst': {
          'p:cond': {
            '@_evt': 'onNext',
            '@_delay': '0',
            'p:tgtEl': { 'p:sldTgt': {} },
          },
        },
      }
      rootChildTnLst['p:seq'].unshift(mainSeq)
    }

    if (!mainSeq['p:cTn']) mainSeq['p:cTn'] = {}
    if (!mainSeq['p:cTn']['p:childTnLst']) {
      mainSeq['p:cTn']['p:childTnLst'] = { 'p:par': [] }
    } else if (!mainSeq['p:cTn']['p:childTnLst']['p:par']) {
      mainSeq['p:cTn']['p:childTnLst']['p:par'] = []
    } else if (!Array.isArray(mainSeq['p:cTn']['p:childTnLst']['p:par'])) {
      mainSeq['p:cTn']['p:childTnLst']['p:par'] = [mainSeq['p:cTn']['p:childTnLst']['p:par']]
    }

    return { rootChildTnLst, mainSeq }
  }

  #ensureVideoTiming(slideObj, shapeId, options = {}) {
    if (!slideObj) return
    const spid = String(shapeId)
    this.#removeSlideTiming(slideObj, shapeId)

    const base = this.#ensureBaseTiming(slideObj)
    if (!base) return
    const { rootChildTnLst, mainSeq } = base

    let nextTnId = this.#getNextTimingId(slideObj)
    const isAutoplay = Boolean(options.autoplay)

    // 1. Add mainSeq playFrom par
    const mainSeqPar = {
      'p:cTn': {
        '@_id': String(nextTnId++),
        '@_fill': 'hold',
        'p:stCondLst': {
          'p:cond': { '@_delay': isAutoplay ? '0' : 'indefinite' },
        },
        'p:childTnLst': {
          'p:par': {
            'p:cTn': {
              '@_id': String(nextTnId++),
              '@_fill': 'hold',
              'p:stCondLst': {
                'p:cond': { '@_delay': '0' },
              },
              'p:childTnLst': {
                'p:par': {
                  'p:cTn': {
                    '@_id': String(nextTnId++),
                    '@_presetID': '1',
                    '@_presetClass': 'mediacall',
                    '@_presetSubtype': '0',
                    '@_fill': 'hold',
                    '@_nodeType': isAutoplay ? 'afterEffect' : 'clickEffect',
                    'p:stCondLst': {
                      'p:cond': { '@_delay': '0' },
                    },
                    'p:childTnLst': {
                      'p:cmd': {
                        '@_type': 'call',
                        '@_cmd': 'playFrom(0.0)',
                        'p:cBhvr': {
                          'p:cTn': {
                            '@_id': String(nextTnId++),
                            '@_dur': '22165',
                            '@_fill': 'hold',
                          },
                          'p:tgtEl': {
                            'p:spTgt': { '@_spid': spid },
                          },
                        },
                      },
                    },
                  },
                },
              },
            },
          },
        },
      },
    }
    mainSeq['p:cTn']['p:childTnLst']['p:par'].push(mainSeqPar)

    // 2. Add interactiveSeq (click to togglePause)
    const interactiveSeq = {
      '@_concurrent': '1',
      '@_nextAc': 'seek',
      'p:cTn': {
        '@_id': String(nextTnId++),
        '@_restart': 'whenNotActive',
        '@_fill': 'hold',
        '@_evtFilter': 'cancelBubble',
        '@_nodeType': 'interactiveSeq',
        'p:stCondLst': {
          'p:cond': {
            '@_evt': 'onClick',
            '@_delay': '0',
            'p:tgtEl': { 'p:spTgt': { '@_spid': spid } },
          },
        },
        'p:endSync': {
          '@_evt': 'end',
          '@_delay': '0',
          'p:rtn': { '@_val': 'all' },
        },
        'p:childTnLst': {
          'p:par': {
            'p:cTn': {
              '@_id': String(nextTnId++),
              '@_fill': 'hold',
              'p:stCondLst': {
                'p:cond': { '@_delay': '0' },
              },
              'p:childTnLst': {
                'p:par': {
                  'p:cTn': {
                    '@_id': String(nextTnId++),
                    '@_fill': 'hold',
                    'p:stCondLst': {
                      'p:cond': { '@_delay': '0' },
                    },
                    'p:childTnLst': {
                      'p:par': {
                        'p:cTn': {
                          '@_id': String(nextTnId++),
                          '@_presetID': '2',
                          '@_presetClass': 'mediacall',
                          '@_presetSubtype': '0',
                          '@_fill': 'hold',
                          '@_nodeType': 'clickEffect',
                          'p:stCondLst': {
                            'p:cond': { '@_delay': '0' },
                          },
                          'p:childTnLst': {
                            'p:cmd': {
                              '@_type': 'call',
                              '@_cmd': 'togglePause',
                              'p:cBhvr': {
                                'p:cTn': {
                                  '@_id': String(nextTnId++),
                                  '@_dur': '1',
                                  '@_fill': 'hold',
                                },
                                'p:tgtEl': {
                                  'p:spTgt': { '@_spid': spid },
                                },
                              },
                            },
                          },
                        },
                      },
                    },
                  },
                },
              },
            },
          },
        },
        'p:nextCondLst': {
          'p:cond': {
            '@_evt': 'onClick',
            '@_delay': '0',
            'p:tgtEl': { 'p:spTgt': { '@_spid': spid } },
          },
        },
      },
    }
    rootChildTnLst['p:seq'].push(interactiveSeq)

    // 3. Add p:video node
    const volVal = Math.round(((options.volume !== undefined ? options.volume : 80) / 100) * 100000)
    const videoNode = {
      'p:cMediaNode': {
        '@_vol': String(volVal),
        'p:cTn': {
          '@_id': String(nextTnId++),
          '@_fill': 'hold',
          '@_display': '0',
          'p:stCondLst': {
            'p:cond': { '@_delay': 'indefinite' },
          },
        },
        'p:tgtEl': {
          'p:spTgt': { '@_spid': spid },
        },
      },
    }
    if (!rootChildTnLst['p:video']) {
      rootChildTnLst['p:video'] = []
    } else if (!Array.isArray(rootChildTnLst['p:video'])) {
      rootChildTnLst['p:video'] = [rootChildTnLst['p:video']]
    }
    rootChildTnLst['p:video'].push(videoNode)
  }

  #ensureAudioTiming(slideObj, shapeId, options = {}) {
    if (!slideObj) return
    const spid = String(shapeId)
    this.#removeSlideTiming(slideObj, shapeId)

    const base = this.#ensureBaseTiming(slideObj)
    if (!base) return
    const { rootChildTnLst, mainSeq } = base

    let nextTnId = this.#getNextTimingId(slideObj)
    const isAutoplay = Boolean(options.autoplay)

    // 1. Add mainSeq playFrom par
    const mainSeqPar = {
      'p:cTn': {
        '@_id': String(nextTnId++),
        '@_fill': 'hold',
        'p:stCondLst': {
          'p:cond': { '@_delay': isAutoplay ? '0' : 'indefinite' },
        },
        'p:childTnLst': {
          'p:par': {
            'p:cTn': {
              '@_id': String(nextTnId++),
              '@_fill': 'hold',
              'p:stCondLst': {
                'p:cond': { '@_delay': '0' },
              },
              'p:childTnLst': {
                'p:par': {
                  'p:cTn': {
                    '@_id': String(nextTnId++),
                    '@_presetID': '1',
                    '@_presetClass': 'mediacall',
                    '@_presetSubtype': '0',
                    '@_fill': 'hold',
                    '@_nodeType': isAutoplay ? 'afterEffect' : 'clickEffect',
                    'p:stCondLst': {
                      'p:cond': { '@_delay': '0' },
                    },
                    'p:childTnLst': {
                      'p:cmd': {
                        '@_type': 'call',
                        '@_cmd': 'playFrom(0.0)',
                        'p:cBhvr': {
                          'p:cTn': {
                            '@_id': String(nextTnId++),
                            '@_dur': '11337',
                            '@_fill': 'hold',
                          },
                          'p:tgtEl': {
                            'p:spTgt': { '@_spid': spid },
                          },
                        },
                      },
                    },
                  },
                },
              },
            },
          },
        },
      },
    }
    mainSeq['p:cTn']['p:childTnLst']['p:par'].push(mainSeqPar)

    // 2. Add p:audio node
    const volVal = Math.round(((options.volume !== undefined ? options.volume : 80) / 100) * 100000)
    const cMediaNode = {
      '@_vol': String(volVal),
      'p:cTn': {
        '@_id': String(nextTnId++),
        '@_fill': 'hold',
        '@_display': '0',
        'p:stCondLst': {
          'p:cond': { '@_delay': 'indefinite' },
        },
        'p:endCondLst': {
          'p:cond': {
            '@_evt': 'onStopAudio',
            '@_delay': '0',
            'p:tgtEl': { 'p:sldTgt': {} },
          },
        },
      },
      'p:tgtEl': {
        'p:spTgt': { '@_spid': spid },
      },
    }
    if (options.hideDuringShow || options.hideIcon) {
      cMediaNode['@_showWhenStopped'] = '0'
    }

    const audioNode = {
      'p:cMediaNode': cMediaNode,
    }
    if (!rootChildTnLst['p:audio']) {
      rootChildTnLst['p:audio'] = []
    } else if (!Array.isArray(rootChildTnLst['p:audio'])) {
      rootChildTnLst['p:audio'] = [rootChildTnLst['p:audio']]
    }
    rootChildTnLst['p:audio'].push(audioNode)
  }

  #ensureSlideTiming(slideObj, shapeId) {
    this.#ensureVideoTiming(slideObj, shapeId, { autoplay: true })
  }

  #containsSpid(node, targetSpid) {
    if (!node || typeof node !== 'object') return false
    if (node['p:spTgt'] && String(node['p:spTgt']['@_spid']) === String(targetSpid)) return true
    if (node['@_spid'] && String(node['@_spid']) === String(targetSpid)) return true
    for (const key of Object.keys(node)) {
      const val = node[key]
      if (Array.isArray(val)) {
        for (const item of val) {
          if (this.#containsSpid(item, targetSpid)) return true
        }
      } else if (typeof val === 'object') {
        if (this.#containsSpid(val, targetSpid)) return true
      }
    }
    return false
  }

  #removeSlideTiming(slideObj, shapeId) {
    if (!slideObj?.['p:sld']?.['p:timing']) return
    const targetSpid = String(shapeId)
    const timing = slideObj['p:sld']['p:timing']
    const rootChildTnLst = timing?.['p:tnLst']?.['p:par']?.['p:cTn']?.['p:childTnLst']
    if (!rootChildTnLst) return

    // 1. Remove video nodes
    if (rootChildTnLst['p:video']) {
      let videos = Array.isArray(rootChildTnLst['p:video'])
        ? rootChildTnLst['p:video']
        : [rootChildTnLst['p:video']]
      videos = videos.filter(v => !this.#containsSpid(v, targetSpid))
      if (videos.length === 0) {
        delete rootChildTnLst['p:video']
      } else {
        rootChildTnLst['p:video'] = videos.length === 1 ? videos[0] : videos
      }
    }

    // 2. Remove audio nodes
    if (rootChildTnLst['p:audio']) {
      let audios = Array.isArray(rootChildTnLst['p:audio'])
        ? rootChildTnLst['p:audio']
        : [rootChildTnLst['p:audio']]
      audios = audios.filter(a => !this.#containsSpid(a, targetSpid))
      if (audios.length === 0) {
        delete rootChildTnLst['p:audio']
      } else {
        rootChildTnLst['p:audio'] = audios.length === 1 ? audios[0] : audios
      }
    }

    // 3. Remove sequence nodes
    if (rootChildTnLst['p:seq']) {
      let seqs = Array.isArray(rootChildTnLst['p:seq'])
        ? rootChildTnLst['p:seq']
        : [rootChildTnLst['p:seq']]
      seqs = seqs.filter(seq => {
        const isInteractive = seq['p:cTn']?.['@_nodeType'] === 'interactiveSeq'
        if (isInteractive && this.#containsSpid(seq, targetSpid)) {
          return false
        }
        const isMain = seq['p:cTn']?.['@_nodeType'] === 'mainSeq'
        if (isMain) {
          const childTnLst = seq['p:cTn']?.['p:childTnLst']
          if (childTnLst?.['p:par']) {
            let pars = Array.isArray(childTnLst['p:par'])
              ? childTnLst['p:par']
              : [childTnLst['p:par']]
            pars = pars.filter(p => !this.#containsSpid(p, targetSpid))
            if (pars.length === 0) {
              delete childTnLst['p:par']
            } else {
              childTnLst['p:par'] = pars.length === 1 ? pars[0] : pars
            }
          }
        }
        return true
      })

      if (seqs.length === 0) {
        delete rootChildTnLst['p:seq']
      } else {
        rootChildTnLst['p:seq'] = seqs.length === 1 ? seqs[0] : seqs
      }
    }

    // Check if any active animation or media node remains
    const hasVideo = Boolean(rootChildTnLst['p:video'])
    const hasAudio = Boolean(rootChildTnLst['p:audio'])
    let hasMainSeqPars = false
    let hasInteractiveSeq = false
    if (rootChildTnLst['p:seq']) {
      const seqs = Array.isArray(rootChildTnLst['p:seq'])
        ? rootChildTnLst['p:seq']
        : [rootChildTnLst['p:seq']]
      for (const seq of seqs) {
        if (seq['p:cTn']?.['@_nodeType'] === 'interactiveSeq') hasInteractiveSeq = true
        if (
          seq['p:cTn']?.['@_nodeType'] === 'mainSeq' &&
          seq['p:cTn']?.['p:childTnLst']?.['p:par']
        ) {
          hasMainSeqPars = true
        }
      }
    }

    if (!hasVideo && !hasAudio && !hasInteractiveSeq && !hasMainSeqPars) {
      delete slideObj['p:sld']['p:timing']
    }
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

  /**
   * Gets info about a media file by its ZIP path.
   *
   * @param {string} zipPath
   * @returns {Object|undefined}
   */
  getMediaInfo(zipPath) {
    return this.#mediaRegistry.get(zipPath)
  }

  /**
   * Returns all registered media files.
   * @returns {Object[]}
   */
  getAllMedia() {
    return Array.from(this.#mediaRegistry.values())
  }

  #hashBytes(data) {
    return createHash('sha1').update(data).digest('hex')
  }

  #detectExtension(data) {
    if (!data || data.length < 4) return 'png'
    const sig = data.slice(0, 16)

    // PNG: 89 50 4E 47 0D 0A 1A 0A
    if (sig[0] === 0x89 && sig[1] === 0x50 && sig[2] === 0x4e && sig[3] === 0x47) return 'png'
    // JPEG: FF D8 FF
    if (sig[0] === 0xff && sig[1] === 0xd8 && sig[2] === 0xff) return 'jpg'
    // GIF: 47 49 46
    if (sig[0] === 0x47 && sig[1] === 0x49 && sig[2] === 0x46) return 'gif'
    // WEBP / WAV / AVI
    if (sig[0] === 0x52 && sig[1] === 0x49 && sig[2] === 0x46 && sig[3] === 0x46) {
      if (
        sig.length >= 12 &&
        sig[8] === 0x57 &&
        sig[9] === 0x45 &&
        sig[10] === 0x42 &&
        sig[11] === 0x50
      ) {
        return 'webp'
      }
      if (
        sig.length >= 12 &&
        sig[8] === 0x57 &&
        sig[9] === 0x41 &&
        sig[10] === 0x56 &&
        sig[11] === 0x45
      ) {
        return 'wav'
      }
      if (sig.length >= 12 && sig[8] === 0x41 && sig[9] === 0x56 && sig[10] === 0x49) {
        return 'avi'
      }
    }
    // TIFF: II*. or MM.*
    if (
      (sig[0] === 0x49 && sig[1] === 0x49 && sig[2] === 0x2a) ||
      (sig[0] === 0x4d && sig[1] === 0x4d && sig[3] === 0x2a)
    ) {
      return 'tiff'
    }
    // BMP: 42 4D
    if (sig[0] === 0x42 && sig[1] === 0x4d) return 'bmp'
    // MP3: ID3 or frame sync
    if (sig[0] === 0x49 && sig[1] === 0x44 && sig[2] === 0x33) return 'mp3'
    if (sig[0] === 0xff && (sig[1] & 0xe0) === 0xe0) return 'mp3'
    // MP4 / MOV: ....ftyp or ....moov
    if (
      sig.length >= 8 &&
      (sig.slice(4, 8).toString('latin1') === 'ftyp' ||
        sig.slice(4, 8).toString('latin1') === 'moov')
    ) {
      return 'mp4'
    }

    const textStart = data.slice(0, 100).toString('utf8').trim()
    if (
      textStart.startsWith('<svg') ||
      textStart.includes('<svg') ||
      textStart.startsWith('<?xml')
    ) {
      return 'svg'
    }

    return 'png'
  }

  #registerContentType(ext, mimeType) {
    this.#contentTypesManager.addDefault(ext, mimeType)
  }
}

module.exports = { MediaManager }
