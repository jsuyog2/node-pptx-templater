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
const { sendExpress, sendFastify, createWebResponse } = require('./FrameworkAdapters.js')
const { EXT_TO_MIME } = require('../utils/mediaUtils.js')

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
    const tokens = xml
      .replace(/>\s*</g, '><')
      .split(/(<[^>]+>)/g)
      .filter(Boolean)

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
    // Ensure all media files in ppt/media/ have default content types registered
    const mediaFiles = zipManager.listFiles('ppt/media/')
    for (const mf of mediaFiles) {
      const ext = mf.split('.').pop().toLowerCase()
      const mime = EXT_TO_MIME[ext] || 'application/octet-stream'
      this.#contentTypesManager.addDefault(ext, mime)
    }
    // Sanitize slide tree ordering and timing references
    if (slideManager && typeof slideManager.slideCount === 'number') {
      for (let idx = 1; idx <= slideManager.slideCount; idx++) {
        const slideObj = slideManager.getSlideObj(idx)
        if (slideObj) {
          const spTree =
            slideObj?.['p:sld']?.['p:cSld']?.['p:spTree'] ||
            slideObj?.['p:sldLayout']?.['p:cSld']?.['p:spTree'] ||
            slideObj?.['p:sldMaster']?.['p:cSld']?.['p:spTree']
          if (spTree && spTree['p:extLst']) {
            const extLst = spTree['p:extLst']
            delete spTree['p:extLst']
            spTree['p:extLst'] = extLst
            slideManager.markSlideObjDirty(idx)
          }
          if (slideObj['p:sld']?.['p:timing']) {
            const modified = this.#sanitizeSlideTiming(slideObj, spTree)
            if (modified) {
              slideManager.markSlideObjDirty(idx)
            }
          }
        }
      }
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
  async toExpress(slideManager, zipManager, res, filename = 'presentation.pptx', options = {}) {
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
  async toFastify(slideManager, zipManager, reply, filename = 'presentation.pptx', options = {}) {
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
   * Serializes the presentation into a Microsoft Flat OPC XML string or file (<pkg:package>).
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @param {string|Object} [destinationPathOrOptions={}] - File path or serialization options.
   * @returns {Promise<string>}
   */
  async toFlatXml(slideManager, zipManager, destinationPathOrOptions = {}, options = {}) {
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

    const isPath = typeof destinationPathOrOptions === 'string'
    const opts = isPath ? options || {} : destinationPathOrOptions || {}
    const xml = FlatOpcConverter.serialize(vfs, opts)

    if (isPath) {
      const resolvedPath = path.resolve(destinationPathOrOptions)
      await ensureDir(path.dirname(resolvedPath))
      await writeFile(resolvedPath, xml, 'utf8')
    }

    return xml
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

  /**
   * Sanitizes slide timing nodes ensuring no orphaned shape IDs exist.
   *
   * @private
   * @param {Object} slideObj
   * @param {Object} spTree
   * @returns {boolean} True if modified.
   */
  #sanitizeSlideTiming(slideObj, spTree) {
    if (!slideObj?.['p:sld']?.['p:timing'] || !spTree) return false
    const existingShapeIds = new Set(this.#collectAllShapeIds(spTree).map(String))
    const referencedSpids = new Set()
    this.#collectReferencedSpids(slideObj['p:sld']['p:timing'], referencedSpids)

    let modified = false
    for (const spid of referencedSpids) {
      if (!existingShapeIds.has(spid)) {
        this.#removeSlideTimingForSpid(slideObj, spid)
        modified = true
      }
    }
    return modified
  }

  /**
   * @private
   * @param {Object} container
   * @returns {Array<number>}
   */
  #collectAllShapeIds(container) {
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
      ids.push(...this.#collectAllShapeIds(g))
    }

    return ids
  }

  /**
   * @private
   * @param {Object} node
   * @param {Set<string>} spids
   */
  #collectReferencedSpids(node, spids) {
    if (!node || typeof node !== 'object') return
    if (node['p:spTgt'] && node['p:spTgt']['@_spid']) {
      spids.add(String(node['p:spTgt']['@_spid']))
    }
    if (node['@_spid']) {
      spids.add(String(node['@_spid']))
    }
    for (const key of Object.keys(node)) {
      const val = node[key]
      if (Array.isArray(val)) {
        for (const item of val) this.#collectReferencedSpids(item, spids)
      } else if (typeof val === 'object') {
        this.#collectReferencedSpids(val, spids)
      }
    }
  }

  /**
   * @private
   * @param {Object} node
   * @param {string} targetSpid
   * @returns {boolean}
   */
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

  /**
   * @private
   * @param {Object} slideObj
   * @param {string|number} shapeId
   */
  #removeSlideTimingForSpid(slideObj, shapeId) {
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
}

module.exports = { OutputWriter, formatXml }
