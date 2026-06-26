/**
 * @fileoverview MediaManager - Manages media files (images, videos, audio) in PPTX.
 *
 * Media in OpenXML PPTX:
 * ─────────────────────────────────────────────────────────────────
 * All media files are stored in ppt/media/ within the PPTX ZIP.
 * They are referenced via relationships from slides, layouts, etc.
 *
 * Media Reference Chain:
 *   slide.xml → slide.xml.rels → ppt/media/imageN.{ext}
 *
 * Image in slide XML:
 *   <p:pic>
 *     <p:nvPicPr>
 *       <p:cNvPr id="3" name="logo"/>
 *     </p:nvPicPr>
 *     <p:blipFill>
 *       <a:blip r:embed="rId4"/>   ← references image via rId
 *     </p:blipFill>
 *     <p:spPr>
 *       <a:xfrm>
 *         <a:off x="457200" y="274638"/>      ← position (EMU)
 *         <a:ext cx="2743200" cy="1828800"/>  ← size (EMU)
 *       </a:xfrm>
 *     </p:spPr>
 *   </p:pic>
 *
 * Media Deduplication:
 *   Multiple slides may embed the same logo/background.
 *   We hash file content (SHA-1) and reuse existing media files
 *   instead of adding duplicates — reducing output file size.
 *
 * Supported formats:
 *   Images: PNG, JPEG, GIF, SVG, TIFF, BMP, WMF, EMF
 *   Video: MP4, AVI, MOV, WMV
 *   Audio: MP3, WAV, M4A
 */

const { createHash } = require('crypto')
const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')
const fsExtra = require('fs-extra')

const logger = createLogger('MediaManager')

function streamToBuffer(stream) {
  return new Promise((resolve, reject) => {
    const chunks = []
    stream.on('data', chunk => chunks.push(chunk))
    stream.on('end', () => resolve(Buffer.concat(chunks)))
    stream.on('error', reject)
  })
}

/**
 * Extension to MIME type mapping.
 */
const EXT_TO_MIME = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  svg: 'image/svg+xml',
  tiff: 'image/tiff',
  bmp: 'image/bmp',
  wmf: 'image/x-wmf',
  emf: 'image/x-emf',
  webp: 'image/webp',
  mp4: 'video/mp4',
  mp3: 'audio/mpeg',
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
   * @param {ContentTypesManager} contentTypesManager
   */
  constructor(contentTypesManager) {
    this.#contentTypesManager = contentTypesManager
  }

  /**
   * Content hash → existing media ZIP path for deduplication.
   * @private @type {Map<string, string>}
   */
  #mediaHashIndex = new Map()

  /**
   * All known media files.
   * @private @type {Map<string, MediaInfo>}
   */
  #mediaRegistry = new Map()

  /**
   * Counter for generating unique media file names.
   * @private @type {number}
   */
  #nextMediaId = 1

  /**
   * Initializes by scanning existing media files in the PPTX.
   *
   * @param {ZipManager} zipManager
   * @returns {Promise<void>}
   */
  async initialize(zipManager) {
    this.#zipManager = zipManager
    const mediaFiles = zipManager.listFiles('ppt/media/')

    // Register all existing media files without loading/hashing them yet
    for (const mediaPath of mediaFiles) {
      const ext = mediaPath.split('.').pop().toLowerCase()
      const mimeType = EXT_TO_MIME[ext] || 'application/octet-stream'

      const mediaInfo = { zipPath: mediaPath, hash: null, mimeType, size: null }
      this.#mediaRegistry.set(mediaPath, mediaInfo)

      // Track the highest media ID to avoid collisions
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
   * Embeds a new image from a file path or Buffer.
   * Automatically deduplicates — if the same image already exists,
   * returns the existing ZIP path instead of creating a duplicate.
   *
   * @param {string|Buffer} source - File path or image Buffer.
   * @param {string} [mimeType] - MIME type (auto-detected from extension if omitted).
   * @returns {Promise<string>} ZIP path of the embedded image (e.g., 'ppt/media/image5.png').
   */
  async embedImage(source, mimeType) {
    let data
    let ext
    let hash
    let size
    const isStream = source && typeof source.on === 'function' && typeof source.pipe === 'function'

    if (isStream) {
      // Buffer the stream to avoid JSZip streaming pipeline crashes and file locks
      data = await streamToBuffer(source)
      if (source.path && typeof source.path === 'string') {
        const filePath = source.path
        ext = filePath.split('.').pop().toLowerCase()
      } else {
        ext = this.#detectExtension(data)
      }
      mimeType = mimeType || EXT_TO_MIME[ext] || 'image/png'
      hash = this.#hashBytes(data)
      size = data.length
    } else if (typeof source === 'string') {
      // Load from file path directly to buffer
      data = await fsExtra.readFile(source)
      ext = source.split('.').pop().toLowerCase()
      mimeType = mimeType || EXT_TO_MIME[ext] || 'image/png'
      hash = this.#hashBytes(data)
      size = data.length
    } else if (Buffer.isBuffer(source) || source instanceof Uint8Array) {
      data = source
      ext = this.#detectExtension(data)
      mimeType = mimeType || EXT_TO_MIME[ext] || 'image/png'
      hash = this.#hashBytes(data)
      size = data.length
    } else {
      throw new PPTXError(
        'embedImage: source must be a file path string, Buffer, or Readable Stream'
      )
    }

    if (this.#mediaHashIndex.has(hash)) {
      const existingPath = this.#mediaHashIndex.get(hash)
      logger.debug(`Reusing existing media: ${existingPath} (hash: ${hash.substring(0, 8)}...)`)
      return existingPath
    }

    // Ensure all media from template is hashed to check for duplicates
    await this.#ensureAllMediaHashed()

    if (this.#mediaHashIndex.has(hash)) {
      const existingPath = this.#mediaHashIndex.get(hash)
      logger.debug(`Reusing existing media: ${existingPath} (hash: ${hash.substring(0, 8)}...)`)
      return existingPath
    }

    // Create a new media file
    const mediaId = this.#nextMediaId++
    const zipPath = `ppt/media/image${mediaId}.${ext}`

    this.#zipManager.writeBinaryFile(zipPath, data)

    this.#mediaHashIndex.set(hash, zipPath)
    this.#mediaRegistry.set(zipPath, { zipPath, hash, mimeType, size })

    // Register content type
    this.#registerContentType(ext, mimeType)

    logger.debug(`Embedded new media: ${zipPath} (${size} bytes)`)
    return zipPath
  }

  /**
   * Copies an existing media part to a new unique file without deduplication.
   * Used when cloning slides so the copy does not share media relationships with the source.
   *
   * @param {string} sourceZipPath - Absolute ZIP path to the source media file.
   * @returns {Promise<string>} ZIP path of the new media file.
   */
  async copyMediaAsNewPart(sourceZipPath) {
    const data = await this.#zipManager.readBinaryFile(sourceZipPath)
    if (!data) {
      throw new PPTXError(`Cannot clone slide: media not found at ${sourceZipPath}`)
    }

    const fileName = sourceZipPath.split('/').pop()
    const ext = fileName.includes('.') ? fileName.split('.').pop().toLowerCase() : 'png'
    const mimeType = EXT_TO_MIME[ext] || 'application/octet-stream'

    const mediaId = this.#nextMediaId++
    const zipPath = `ppt/media/image${mediaId}.${ext}`

    this.#zipManager.writeBinaryFile(zipPath, data)
    const hash = this.#hashBytes(data)
    this.#mediaHashIndex.set(hash, zipPath)
    this.#mediaRegistry.set(zipPath, { zipPath, hash, mimeType, size: data.length })
    this.#registerContentType(ext, mimeType)

    logger.debug(`Cloned media to new part: ${zipPath}`)
    return zipPath
  }

  /**
   * Generates the slide XML snippet for an image element.
   *
   * @param {string} rId - Relationship ID pointing to the media file.
   * @param {ImageElement} opts - Image element options.
   * @param {number} opts.x - X position in EMU.
   * @param {number} opts.y - Y position in EMU.
   * @param {number} opts.width - Width in EMU.
   * @param {number} opts.height - Height in EMU.
   * @param {string} [opts.name] - Shape name.
   * @param {number} [opts.shapeId] - Shape ID.
   * @returns {string} XML snippet for the image.
   */
  buildImageXml(rId, opts) {
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
   * Gets info about a media file by its ZIP path.
   *
   * @param {string} zipPath
   * @returns {MediaInfo|undefined}
   */
  getMediaInfo(zipPath) {
    return this.#mediaRegistry.get(zipPath)
  }

  /**
   * Returns all registered media files.
   * @returns {MediaInfo[]}
   */
  getAllMedia() {
    return Array.from(this.#mediaRegistry.values())
  }

  /**
   * Computes a SHA-1 hash of binary data.
   * Used for content-addressable deduplication.
   *
   * @private
   * @param {Buffer|Uint8Array} data
   * @returns {string} Hex digest.
   */
  #hashBytes(data) {
    return createHash('sha1').update(data).digest('hex')
  }

  /**
   * Detects image format from magic bytes.
   *
   * @private
   * @param {Buffer|Uint8Array} data
   * @returns {string} File extension.
   */
  #detectExtension(data) {
    const sig = data.slice(0, 8)

    // PNG: 89 50 4E 47 0D 0A 1A 0A
    if (sig[0] === 0x89 && sig[1] === 0x50) return 'png'
    // JPEG: FF D8 FF
    if (sig[0] === 0xff && sig[1] === 0xd8) return 'jpg'
    // GIF: 47 49 46
    if (sig[0] === 0x47 && sig[1] === 0x49) return 'gif'
    // WEBP: 52 49 46 46 ... 57 45 42 50
    if (sig[0] === 0x52 && sig[1] === 0x49 && sig[8] === 0x57) return 'webp'
    // BMP: 42 4D
    if (sig[0] === 0x42 && sig[1] === 0x4d) return 'bmp'

    return 'png' // Default fallback
  }

  #registerContentType(ext, mimeType) {
    this.#contentTypesManager.addDefault(ext, mimeType)
  }
}

module.exports = { MediaManager }
