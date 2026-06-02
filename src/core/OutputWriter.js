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

  /**
   * @param {ZipManager} zipManager
   * @param {ContentTypesManager} contentTypesManager
   */
  constructor(zipManager, contentTypesManager) {
    this.#zipManager = zipManager
    this.#contentTypesManager = contentTypesManager
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
  async saveToFile(filePath, slideManager, zipManager) {
    try {
      const buffer = await this.toBuffer(slideManager, zipManager)
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
   * Returns the PPTX as a Node.js Buffer.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @returns {Promise<Buffer>}
   */
  async toBuffer(slideManager, zipManager) {
    // Ensure all slides are flushed to the ZIP
    await this.#flushAllSlides(slideManager, zipManager)

    // Flush Content Types safely
    this.#contentTypesManager.flush(zipManager)

    // Wait for any queued asynchronous writes (like content types, media hashing)
    await zipManager.waitForPendingWrites()

    const buffer = await zipManager.toBuffer()
    logger.debug(`Generated buffer: ${(buffer.length / 1024).toFixed(1)} KB`)
    return buffer
  }

  /**
   * Returns the PPTX as a readable Node.js stream.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @returns {Promise<Readable>}
   */
  async toStream(slideManager, zipManager) {
    await this.#flushAllSlides(slideManager, zipManager)

    // Flush Content Types safely
    this.#contentTypesManager.flush(zipManager)

    await zipManager.waitForPendingWrites()
    const nodeStream = await zipManager.toStream()
    return nodeStream
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
    // SlideManager already writes to zipManager via setSlideXml,
    // so this is mostly a no-op with a validation step.
    const info = slideManager.getAllSlideInfo()

    for (const slide of info) {
      if (!zipManager.hasFile(slide.zipPath)) {
        logger.warn(`Slide file missing in ZIP: ${slide.zipPath}`)
      }
    }

    // Update the slide count and titles in docProps/app.xml to prevent repair mode issues
    if (zipManager.hasFile('docProps/app.xml')) {
      zipManager.addPendingPromise(
        zipManager.rawZip
          .file('docProps/app.xml')
          .async('text')
          .then(content => {
            const parser = new XMLParser()
            const appObj = parser.parse(content, 'app.xml')
            const properties = appObj.Properties

            if (properties) {
              // 1. Update Slides count
              properties.Slides = info.length

              // 2. Find old slide titles count and update HeadingPairs
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

              // 3. Update TitlesOfParts
              const titlesVector = properties.TitlesOfParts?.['vt:vector']
              if (titlesVector) {
                let lpstrs = titlesVector['vt:lpstr']
                if (lpstrs) {
                  if (!Array.isArray(lpstrs)) lpstrs = [lpstrs]

                  // Remove the old slide titles (which are at the end)
                  if (oldSlideTitlesCount > 0 && lpstrs.length >= oldSlideTitlesCount) {
                    lpstrs = lpstrs.slice(0, lpstrs.length - oldSlideTitlesCount)
                  }

                  // Append new slide titles
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
          })
      )
    }

    logger.debug(`Flushed ${info.length} slide(s) to ZIP`)
  }
}

module.exports = { OutputWriter }
