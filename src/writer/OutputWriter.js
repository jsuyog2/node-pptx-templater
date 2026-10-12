/**
 * @fileoverview OutputWriter — Handles multi-target PPTX serialization and export.
 *
 * Responsibilities:
 *  1. Flush all pending changes from all managers into the ZipManager
 *  2. Ensure [Content_Types].xml is up to date
 *  3. Generate the final ZIP archive
 *  4. Write to disk, Buffer, Stream, ArrayBuffer, Base64, Express, Fastify, or Web Response
 *  5. Extract presentation parts directly to ZIP or uncompressed XML folders
 *
 * @module OutputWriter
 */

'use strict'

const fsExtra = require('fs-extra')
const { writeFile, ensureDir } = fsExtra
const path = require('path')
const { XMLParser } = require('../parsers/XMLParser.js')
const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')
const {
  sendExpress,
  sendFastify,
  createWebResponse,
} = require('./FrameworkAdapters.js')

const logger = createLogger('OutputWriter')

/**
 * Formats an XML string with indentation for inspection and git version control.
 *
 * @param {string} xml - Raw XML string.
 * @returns {string} Formatted XML string.
 */
function formatXml(xml) {
  if (typeof xml !== 'string' || !xml.trim()) return xml
  try {
    let formatted = ''
    let pad = 0
    const PADDING = '  '
    // Split into tags and content
    const tokens = xml.replace(/>\s*</g, '><').split(/(<[^>]+>)/g).filter(Boolean)

    for (const token of tokens) {
      if (token.startsWith('<?') || token.startsWith('<!')) {
        formatted += token + '\n'
      } else if (token.startsWith('</')) {
        pad = Math.max(0, pad - 1)
        formatted += PADDING.repeat(pad) + token + '\n'
      } else if (token.startsWith('<') && token.endsWith('/>')) {
        formatted += PADDING.repeat(pad) + token + '\n'
      } else if (token.startsWith('<')) {
        formatted += PADDING.repeat(pad) + token + '\n'
        pad += 1
      } else {
        const text = token.trim()
        if (text) {
          formatted += PADDING.repeat(pad) + text + '\n'
        }
      }
    }
    return formatted.trim() + '\n'
  } catch {
    return xml
  }
}

/**
 * @class OutputWriter
 * @description Serializes the modified PPTX to various output formats and targets.
 */
class OutputWriter {
  /** @private @type {ZipManager} */
  #zipManager
  /** @private @type {ContentTypesManager} */
  #contentTypesManager
  /** @private @type {RelationshipManager} */
  #relationshipManager

  /** @type {boolean} */
  debugZip = false

  /**
   * @param {ZipManager} zipManager
   * @param {ContentTypesManager} contentTypesManager
   * @param {RelationshipManager} [relationshipManager]
   */
  constructor(zipManager, contentTypesManager, relationshipManager = null) {
    this.#zipManager = zipManager
    this.#contentTypesManager = contentTypesManager
    this.#relationshipManager = relationshipManager
  }

  /**
   * Writes the final PPTX to a file on disk.
   * Creates parent directories if needed.
   *
   * @param {string} filePath - Output file path.
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {Object} [options={}]
   * @returns {Promise<void>}
   */
  async saveToFile(filePath, slideManager, zipManager, options = {}) {
    try {
      const buffer = await this.toBuffer(slideManager, zipManager, options)
      const dir = path.dirname(filePath)
      await ensureDir(dir)
      await writeFile(filePath, buffer)
      logger.info(`Saved to ${filePath} (${(buffer.length / 1024).toFixed(1)} KB)`)
    } catch (err) {
      if (err instanceof PPTXError) throw err
      throw new PPTXError(`Failed to save file to ${filePath}: ${err.message}`, err)
    }
  }

  /**
   * Flushes all pending changes from all managers into the ZipManager.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @returns {Promise<void>}
   */
  async flush(slideManager, zipManager) {
    if (slideManager && typeof slideManager.flush === 'function') {
      slideManager.flush()
    }
    // Complete all pending async mutations before structural normalization
    await zipManager.waitForPendingWrites()
    if (
      slideManager &&
      typeof slideManager.normalizeStructure === 'function' &&
      this.#relationshipManager
    ) {
      await slideManager.normalizeStructure(this.#relationshipManager, this.#contentTypesManager)
    }
    await this.#flushAllSlides(slideManager, zipManager)
    this.#contentTypesManager.flush(zipManager)
  }

  /**
   * Returns the PPTX as an ArrayBuffer.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {Object} [options={}]
   * @returns {Promise<ArrayBuffer>}
   */
  async toArrayBuffer(slideManager, zipManager, options = {}) {
    const buffer = await this.toBuffer(slideManager, zipManager, options)
    return buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength)
  }

  /**
   * Returns the PPTX as a Base64 string or Data URI.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {Object} [options={}] - Pass { dataUri: true } for data URI scheme.
   * @returns {Promise<string>}
   */
  async toBase64(slideManager, zipManager, options = {}) {
    const buffer = await this.toBuffer(slideManager, zipManager, options)
    if (options.dataUri) {
      return (
        'data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,' +
        buffer.toString('base64')
      )
    }
    return buffer.toString('base64')
  }

  /**
   * Pipes the presentation directly to an Express.js HTTP response.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {Object} res - Express response object.
   * @param {string} [filename='presentation.pptx']
   * @param {Object} [options={}]
   * @returns {Promise<void>}
   */
  async toExpress(
    slideManager,
    zipManager,
    res,
    filename = 'presentation.pptx',
    options = {}
  ) {
    const stream = await this.toStream(slideManager, zipManager, options)
    sendExpress(res, stream, filename, options)
  }

  /**
   * Emits a standard Web API Response object (Fetch API / Edge runtime compliant).
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {Object} [options={}]
   * @param {string} [options.filename='presentation.pptx']
   * @returns {Promise<Response>}
   */
  async toResponse(slideManager, zipManager, options = {}) {
    const buffer = await this.toBuffer(slideManager, zipManager, options)
    const filename = options.filename || 'presentation.pptx'
    return createWebResponse(buffer, filename, options)
  }

  /**
   * Sends the presentation via a Fastify reply object.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {Object} reply - Fastify reply object.
   * @param {string} [filename='presentation.pptx']
   * @param {Object} [options={}]
   * @returns {Promise<void>}
   */
  async toFastify(
    slideManager,
    zipManager,
    reply,
    filename = 'presentation.pptx',
    options = {}
  ) {
    const buffer = await this.toBuffer(slideManager, zipManager, options)
    sendFastify(reply, buffer, filename, options)
  }

  /**
   * Returns the PPTX presentation as a Node.js Buffer.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {Object} [options={}]
   * @returns {Promise<Buffer>}
   */
  async toBuffer(slideManager, zipManager, options = {}) {
    await this.flush(slideManager, zipManager)

    const buffer = await zipManager.toBuffer(options)
    logger.debug(`Generated buffer: ${(buffer.length / 1024).toFixed(1)} KB`)

    if (this.debugZip) {
      this.printDebugZip(buffer)
    }

    return buffer
  }

  /**
   * Alias for toBuffer to dump raw compressed ZIP content.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {Object} [options={}]
   * @returns {Promise<Buffer>}
   */
  async toZipBuffer(slideManager, zipManager, options = {}) {
    return this.toBuffer(slideManager, zipManager, options)
  }

  /**
   * Extracts and packages the presentation into a compressed ZIP file at destinationPath.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {string} destinationPath
   * @param {Object} [options={}]
   * @returns {Promise<void>}
   */
  async extractToZip(slideManager, zipManager, destinationPath, options = {}) {
    return this.saveToFile(destinationPath, slideManager, zipManager, options)
  }

  /**
   * Returns the PPTX as a readable Node.js stream.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {Object} [options={}]
   * @returns {Promise<import('stream').Readable>}
   */
  async toStream(slideManager, zipManager, options = {}) {
    await this.flush(slideManager, zipManager)
    const nodeStream = await zipManager.toStream(options)

    if (this.debugZip) {
      const buffer = await zipManager.toBuffer(options)
      this.printDebugZip(buffer)
    }

    return nodeStream
  }

  /**
   * Unpacks all presentation parts into an uncompressed folder on disk.
   * Supports formatting XML files for version control inspection when options.formatXml is true.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {string} destinationDir
   * @param {Object} [options={}]
   * @param {boolean} [options.formatXml=false] - Whether to pretty-format XML and .rels files.
   * @returns {Promise<void>}
   */
  async extractToXml(slideManager, zipManager, destinationDir, options = {}) {
    await this.flush(slideManager, zipManager)
    const resolvedDest = path.resolve(destinationDir)
    await ensureDir(resolvedDest)

    const fileList = zipManager.listFiles()

    for (const relPath of fileList) {
      const targetPath = path.join(resolvedDest, relPath)
      await ensureDir(path.dirname(targetPath))

      const isXml = relPath.endsWith('.xml') || relPath.endsWith('.rels')
      if (isXml) {
        let content = await zipManager.readFile(relPath)
        if (content !== null) {
          if (options.formatXml) {
            content = formatXml(content)
          }
          await writeFile(targetPath, content, 'utf8')
        }
      } else {
        const bin = await zipManager.readBinaryFile(relPath)
        if (bin !== null) {
          await writeFile(targetPath, bin)
        }
      }
    }

    logger.info(`Extracted presentation XML parts to ${destinationDir}`)
  }

  /**
   * Saves the presentation directly to a folder structure on disk.
   * Alias for extractToXml.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {string} destinationDir
   * @param {Object} [options={}]
   * @returns {Promise<void>}
   */
  async saveToFolder(slideManager, zipManager, destinationDir, options = {}) {
    return this.extractToXml(slideManager, zipManager, destinationDir, options)
  }

  /**
   * Serializes the presentation into a Microsoft Flat OPC XML string (<pkg:package>).
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {Object} [options={}]
   * @returns {Promise<string>}
   */
  async toFlatXml(slideManager, zipManager, options = {}) {
    await this.flush(slideManager, zipManager)
    const { FlatOpcConverter } = require('../converters/FlatOpcConverter.js')
    const { VirtualFileSystem } = require('../core/VirtualFileSystem.js')

    const vfs = new VirtualFileSystem()
    const fileList = zipManager.listFiles()
    for (const relPath of fileList) {
      const data = await zipManager.readFile(relPath)
      if (typeof data === 'string') {
        vfs.writeText(relPath, data)
      } else if (data) {
        vfs.writeBytes(relPath, data)
      }
    }

    return FlatOpcConverter.serialize(vfs, options)
  }

  /**
   * Parses the Central Directory of a ZIP buffer and logs debug info for every entry.
   *
   * @param {Buffer} buffer
   */
  printDebugZip(buffer) {
    let offset = 0
    const entries = []

    while (offset < buffer.length - 46) {
      const sig = buffer.readUInt32LE(offset)
      if (sig === 0x02014b50) {
        const compressionMethod = buffer.readUInt16LE(offset + 10)
        const crc32 = buffer.readUInt32LE(offset + 16)
        const compressedSize = buffer.readUInt32LE(offset + 20)
        const uncompressedSize = buffer.readUInt32LE(offset + 24)
        const fileNameLength = buffer.readUInt16LE(offset + 28)
        const extraFieldLength = buffer.readUInt16LE(offset + 30)
        const fileCommentLength = buffer.readUInt16LE(offset + 32)

        const fileName = buffer.toString('utf8', offset + 46, offset + 46 + fileNameLength)

        entries.push({
          name: fileName,
          compressionMethod,
          crc32: crc32.toString(16).toLowerCase(),
          compressedSize,
          uncompressedSize,
        })

        offset += 46 + fileNameLength + extraFieldLength + fileCommentLength
      } else {
        offset++
      }
    }

    logger.info(`--- ZIP debug output (${entries.length} entries) ---`)
    entries.forEach(e => {
      const methodStr =
        e.compressionMethod === 8
          ? 'DEFLATE'
          : e.compressionMethod === 0
            ? 'STORE'
            : `UNKNOWN(${e.compressionMethod})`
      logger.info(e.name)
      logger.info(`compressed: ${e.compressedSize}`)
      logger.info(`uncompressed: ${e.uncompressedSize}`)
      logger.info(`crc: ${e.crc32}`)
      logger.info(`method: ${methodStr}`)
    })
    logger.info('--- End of ZIP debug output ---')
  }

  /**
   * Ensures all dirty slide XML is committed to the ZipManager.
   *
   * @private
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @returns {Promise<void>}
   */
  async #flushAllSlides(slideManager, zipManager) {
    const info = slideManager.getAllSlideInfo()

    for (const slide of info) {
      if (!zipManager.hasFile(slide.zipPath)) {
        logger.warn(`Slide file missing in ZIP: ${slide.zipPath}`)
      }
    }

    if (zipManager.hasFile('docProps/app.xml')) {
      const content = await zipManager.readFile('docProps/app.xml')
      if (content) {
        const parser = new XMLParser()
        const appObj = parser.parse(content, 'app.xml')
        const properties = appObj.Properties

        if (properties) {
          properties.Slides = info.length

          let oldSlideTitlesCount = 0
          const variants = properties.HeadingPairs?.['vt:vector']?.['vt:variant']
          if (Array.isArray(variants)) {
            for (let i = 0; i < variants.length; i++) {
              if (variants[i]['vt:lpstr'] === 'Slide Titles') {
                const countVar = variants[i + 1]
                if (countVar) {
                  oldSlideTitlesCount = parseInt(countVar['vt:i4'], 10) || 0
                  countVar['vt:i4'] = info.length
                }
                break
              }
            }
          }

          const titlesVector = properties.TitlesOfParts?.['vt:vector']
          if (titlesVector) {
            let lpstrs = titlesVector['vt:lpstr']
            if (lpstrs) {
              if (!Array.isArray(lpstrs)) lpstrs = [lpstrs]
              if (oldSlideTitlesCount > 0 && lpstrs.length >= oldSlideTitlesCount) {
                lpstrs = lpstrs.slice(0, lpstrs.length - oldSlideTitlesCount)
              }
              const newSlideTitles = info.map(slide => slide.title || `Slide ${slide.index}`)
              lpstrs.push(...newSlideTitles)

              titlesVector['vt:lpstr'] = lpstrs
              titlesVector['@_size'] = String(lpstrs.length)
            }
          }

          const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          const updatedXml = parser.build(appObj, declaration)
          zipManager.writeFile('docProps/app.xml', updatedXml)
        }
      }
    }

    logger.debug(`Flushed ${info.length} slide(s) to ZIP`)
  }
}

module.exports = { OutputWriter, formatXml }
