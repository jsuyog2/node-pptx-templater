/**
 * @fileoverview OutputWriter - Handles PPTX serialization and output.
 *
 * Responsibilities:
 *  1. Flush all pending changes from all managers into the ZipManager
 *  2. Ensure [Content_Types].xml is up to date
 *  3. Generate the final ZIP archive
 *  4. Write to file, buffer, or stream
 */

const fsExtra = require('fs-extra')
const { writeFile, ensureDir } = fsExtra
const path = require('path')
const { XMLParser } = require('../parsers/XMLParser.js')
const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')
const logger = createLogger('OutputWriter')

/**
 * @class OutputWriter
 * @description Serializes the modified PPTX to various output formats.
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
    // Complete all pending async mutations (slide duplication, chart writes, etc.)
    // BEFORE structural normalization so sldIdLst/rels match the final slide set.
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
   * Returns the PPTX as a Node.js Buffer.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {Object} [options]
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
   * Returns the PPTX as a readable Node.js stream.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {Object} [options]
   * @returns {Promise<Readable>}
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
   * This is called before any output operation.
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

          // Writing it safely here now that the function block is strictly sequential
          zipManager.writeFile('docProps/app.xml', updatedXml)
        }
      }
    }

    logger.debug(`Flushed ${info.length} slide(s) to ZIP`)
  }
}

module.exports = { OutputWriter }
