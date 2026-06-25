/**
 * @fileoverview ZipManager - Handles PPTX ZIP archive operations.
 *
 * A PPTX file is a ZIP archive following the Open Packaging Convention (OPC).
 * This manager wraps JSZip to provide:
 *  - Loading and parsing PPTX ZIP archives
 *  - Reading individual XML parts
 *  - Writing/replacing individual XML parts
 *  - Re-packaging the modified archive
 *  - Media file deduplication
 *
 * ZipManager is the lowest layer — all other managers use it to read/write
 * raw file content within the ZIP.
 */

const JSZip = require('jszip')
const fsExtra = require('fs-extra')
const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')
const { BLANK_PPTX_BASE64 } = require('../templates/blankPptx.js')

const logger = createLogger('ZipManager')

/**
 * @class ZipManager
 * @description Manages the PPTX ZIP archive — reading, modifying, and re-packaging it.
 *
 * All file paths within the ZIP use forward slashes (e.g., 'ppt/slides/slide1.xml').
 */
class ZipManager {
  /**
   * @private
   * @type {JSZip}
   */
  #zip = null

  /**
   * @private
   * @type {Map<string, string>} Cache of decoded XML strings for fast repeated access.
   */
  #xmlCache = new Map()

  /**
   * @private
   * @type {Map<string, string>} Dirty (modified) files that need to be re-written.
   */
  #dirtyFiles = new Map()

  /**
   * @private
   * @type {Map<string, string>} Core properties (dc:title, dc:creator, etc.)
   */
  #coreProperties = new Map()

  /**
   * @private
   * @type {boolean}
   */
  #isFolderMode = false

  /**
   * @private
   * @type {string|null}
   */
  #folderRoot = null

  /**
   * @private
   * @type {Set<string>}
   */
  #folderFiles = new Set()

  /**
   * @private
   * @type {Map<string, Buffer|Uint8Array>}
   */
  #dirtyBinaryFiles = new Map()

  /**
   * @private
   * @type {Set<string>}
   */
  #removedFiles = new Set()

  /**
   * @private
   * @type {Map<string, { type: string, content: string|Buffer|Uint8Array }>|null}
   */
  #cachedFiles = null

  async loadFromCache(cachedFilesMap) {
    this.#cachedFiles = cachedFilesMap
    await this.#loadCoreProperties()
    logger.debug(`Loaded from cache. Files: ${cachedFilesMap.size}`)
  }

  async load(source) {
    try {
      const path = require('path')
      const fs = require('fs-extra')

      let isFolder = false
      let rootDir = null
      let presentationPath = null

      if (typeof source === 'object' && source !== null && (source.presentation || source.root)) {
        isFolder = true
        presentationPath = source.presentation
        rootDir = source.root
      } else if (typeof source === 'string') {
        try {
          const stat = fs.statSync(source)
          if (stat.isDirectory()) {
            isFolder = true
            rootDir = source
          } else if (source.endsWith('.xml')) {
            isFolder = true
            presentationPath = source
          }
        } catch (e) {
          if (source.endsWith('.xml')) {
            isFolder = true
            presentationPath = source
          }
        }
      }

      if (isFolder) {
        this.#isFolderMode = true

        // Resolve presentationPath and rootDir
        if (presentationPath && !rootDir) {
          const resolvedPresentation = path.resolve(presentationPath)
          const dir = path.dirname(resolvedPresentation)
          if (path.basename(dir).toLowerCase() === 'ppt') {
            rootDir = path.dirname(dir)
          } else {
            rootDir = dir
          }
        } else if (rootDir && !presentationPath) {
          const candidates = [
            path.join(rootDir, 'ppt/presentation.xml'),
            path.join(rootDir, 'presentation.xml'),
          ]
          for (const cand of candidates) {
            if (fs.existsSync(cand)) {
              presentationPath = cand
              break
            }
          }
          if (!presentationPath) {
            presentationPath = path.join(rootDir, 'ppt/presentation.xml')
          }
        }

        this.#folderRoot = path.resolve(rootDir)
        logger.debug(`Loading from uncompressed OpenXML folder: ${this.#folderRoot}`)

        // Populate #folderFiles recursively
        const getFiles = async (dir, baseDir) => {
          const results = []
          if (!(await fs.pathExists(dir))) return results
          const list = await fs.readdir(dir)
          for (const file of list) {
            const filePath = path.join(dir, file)
            const stat = await fs.stat(filePath)
            if (stat && stat.isDirectory()) {
              results.push(...(await getFiles(filePath, baseDir)))
            } else {
              const rel = path.relative(baseDir, filePath).replace(/\\/g, '/')
              results.push(rel)
            }
          }
          return results
        }

        const files = await getFiles(this.#folderRoot, this.#folderRoot)
        this.#folderFiles = new Set(files)

        await this.#loadCoreProperties()
        logger.debug(`Folder loaded successfully. Files: ${this.#folderFiles.size}`)
      } else {
        let data
        if (typeof source === 'string') {
          logger.debug(`Reading file: ${source}`)
          data = await fsExtra.readFile(source)
        } else if (Buffer.isBuffer(source) || source instanceof Uint8Array) {
          data = source
        } else {
          throw new PPTXError(
            `Invalid source type: ${typeof source}. Expected string path or Buffer.`
          )
        }

        this.#zip = await JSZip.loadAsync(data)
        await this.#loadCoreProperties()
        logger.debug(`ZIP loaded successfully. Files: ${Object.keys(this.#zip.files).length}`)
      }
    } catch (err) {
      if (err instanceof PPTXError) throw err
      throw new PPTXError(`Failed to load template: ${err.message}`, err)
    }
  }

  /**
   * Creates a blank PPTX structure from the embedded minimal template.
   * @returns {Promise<void>}
   */
  async createBlank() {
    const buffer = Buffer.from(BLANK_PPTX_BASE64, 'base64')
    this.#zip = await JSZip.loadAsync(buffer)
    await this.#loadCoreProperties()
    logger.debug('Created blank PPTX structure')
  }

  /**
   * Reads and caches a text file from the ZIP archive.
   *
   * @param {string} zipPath - Path within the ZIP (e.g., 'ppt/slides/slide1.xml').
   * @returns {Promise<string|null>} File content as UTF-8 string, or null if not found.
   */
  async readFile(zipPath) {
    // Normalize path separators
    const normalPath = zipPath.replace(/\\/g, '/')

    // Return cached version if available and not dirty
    if (this.#xmlCache.has(normalPath) && !this.#dirtyFiles.has(normalPath)) {
      return this.#xmlCache.get(normalPath)
    }

    // Check dirty files (pending writes)
    if (this.#dirtyFiles.has(normalPath)) {
      return this.#dirtyFiles.get(normalPath)
    }

    if (this.#cachedFiles && this.#cachedFiles.has(normalPath)) {
      const entry = this.#cachedFiles.get(normalPath)
      let content
      if (entry.type === 'text') {
        content = entry.content
      } else {
        const { TextDecoder } = require('util')
        content = new TextDecoder('utf-8').decode(entry.content)
      }
      this.#xmlCache.set(normalPath, content)
      return content
    }

    if (this.#isFolderMode) {
      const path = require('path')
      const fs = require('fs-extra')
      const diskPath = path.join(this.#folderRoot, normalPath)
      if (!(await fs.pathExists(diskPath))) {
        logger.debug(`File not found in folder: ${normalPath}`)
        return null
      }
      const content = await fs.readFile(diskPath, 'utf8')
      this.#xmlCache.set(normalPath, content)
      return content
    }

    if (!this.#zip) return null
    const file = this.#zip.file(normalPath)
    if (!file) {
      logger.debug(`File not found in ZIP: ${normalPath}`)
      return null
    }

    const content = await file.async('text')
    this.#xmlCache.set(normalPath, content)
    return content
  }

  /**
   * Synchronously reads a cached text file.
   *
   * @param {string} zipPath - Path within the ZIP.
   * @returns {string|null} Cached content or null.
   */
  readCachedFile(zipPath) {
    const normalPath = zipPath.replace(/\\/g, '/')
    if (this.#dirtyFiles.has(normalPath)) {
      return this.#dirtyFiles.get(normalPath)
    }
    if (this.#xmlCache.has(normalPath)) {
      return this.#xmlCache.get(normalPath)
    }
    if (this.#cachedFiles && this.#cachedFiles.has(normalPath)) {
      const entry = this.#cachedFiles.get(normalPath)
      let content
      if (entry.type === 'text') {
        content = entry.content
      } else {
        const { TextDecoder } = require('util')
        content = new TextDecoder('utf-8').decode(entry.content)
      }
      this.#xmlCache.set(normalPath, content)
      return content
    }
    return null
  }

  /**
   * Reads a binary file from the ZIP archive.
   *
   * @param {string} zipPath - Path within the ZIP.
   * @returns {Promise<Uint8Array|null>} Binary content or null if not found.
   */
  async readBinaryFile(zipPath) {
    const normalPath = zipPath.replace(/\\/g, '/')
    if (this.#dirtyBinaryFiles.has(normalPath)) {
      return this.#dirtyBinaryFiles.get(normalPath)
    }
    if (this.#cachedFiles && this.#cachedFiles.has(normalPath)) {
      const entry = this.#cachedFiles.get(normalPath)
      return entry.content
    }
    if (this.#isFolderMode) {
      const path = require('path')
      const fs = require('fs-extra')
      const diskPath = path.join(this.#folderRoot, normalPath)
      if (!(await fs.pathExists(diskPath))) return null
      return fs.readFile(diskPath)
    }
    if (!this.#zip) return null
    const file = this.#zip.file(normalPath)
    if (!file) return null
    return file.async('uint8array')
  }

  writeFile(zipPath, content) {
    const normalPath = zipPath.replace(/\\/g, '/')
    this.#dirtyFiles.set(normalPath, content)
    this.#xmlCache.set(normalPath, content)
    this.#removedFiles.delete(normalPath)
    if (this.#zip) {
      this.#zip.file(normalPath, content)
    }
    logger.debug(`Queued write: ${normalPath}`)
  }

  writeBinaryFile(zipPath, data) {
    const normalPath = zipPath.replace(/\\/g, '/')
    this.#dirtyBinaryFiles.set(normalPath, data)
    this.#removedFiles.delete(normalPath)
    if (this.#zip) {
      this.#zip.file(normalPath, data)
    }
    logger.debug(`Queued binary write: ${normalPath}`)
  }

  /**
   * @private
   * @type {Promise[]}
   */
  #pendingPromises = []

  /**
   * Adds a promise to the pending queue to be awaited before saving.
   * @param {Promise} promise
   */
  addPendingPromise(promise) {
    this.#pendingPromises.push(promise)
  }

  /**
   * Waits for all pending asynchronous operations (like async ZIP reads/writes) to complete.
   * @returns {Promise<void>}
   */
  async waitForPendingWrites() {
    if (this.#pendingPromises.length > 0) {
      await Promise.all(this.#pendingPromises)
      this.#pendingPromises = []
    }
  }

  /**
   * Removes a file from the ZIP archive.
   *
   * @param {string} zipPath - Path to remove.
   */
  removeFile(zipPath) {
    const normalPath = zipPath.replace(/\\/g, '/')
    this.#removedFiles.add(normalPath)
    this.#xmlCache.delete(normalPath)
    this.#dirtyFiles.delete(normalPath)
    this.#dirtyBinaryFiles.delete(normalPath)
    if (this.#zip) {
      this.#zip.remove(normalPath)
    }
  }

  /**
   * Checks if a file exists in the ZIP archive.
   *
   * @param {string} zipPath - Path to check.
   * @returns {boolean}
   */
  hasFile(zipPath) {
    const normalPath = zipPath.replace(/\\/g, '/')
    if (this.#removedFiles.has(normalPath)) return false
    if (this.#dirtyFiles.has(normalPath) || this.#dirtyBinaryFiles.has(normalPath)) return true
    if (this.#cachedFiles && this.#cachedFiles.has(normalPath)) return true
    if (this.#isFolderMode) {
      return this.#folderFiles.has(normalPath)
    }
    return this.#zip && this.#zip.file(normalPath) !== null
  }

  /**
   * Lists all files in the ZIP archive matching an optional prefix.
   *
   * @param {string} [prefix] - Optional path prefix filter.
   * @returns {string[]} Array of matching file paths.
   */
  listFiles(prefix = '') {
    if (this.#cachedFiles) {
      const allFiles = new Set([
        ...this.#cachedFiles.keys(),
        ...this.#dirtyFiles.keys(),
        ...this.#dirtyBinaryFiles.keys(),
      ])
      return Array.from(allFiles).filter(f => !this.#removedFiles.has(f) && f.startsWith(prefix))
    }
    if (this.#isFolderMode) {
      const allFiles = new Set([
        ...this.#folderFiles,
        ...this.#dirtyFiles.keys(),
        ...this.#dirtyBinaryFiles.keys(),
      ])
      return Array.from(allFiles).filter(f => !this.#removedFiles.has(f) && f.startsWith(prefix))
    }
    if (!this.#zip) return []
    return Object.keys(this.#zip.files).filter(f => !this.#zip.files[f].dir && f.startsWith(prefix))
  }

  /**
   * Generates the final ZIP archive as a Buffer.
   * All pending changes are applied before compressing.
   *
   * @returns {Promise<Buffer>} Compressed PPTX as a Buffer.
   */
  async toBuffer(options = {}) {
    await this.#ensureZipForExport()
    const zipOptions = this.#getZipOptions(options)
    return this.#zip.generateAsync(zipOptions)
  }

  async toStream(options = {}) {
    await this.#ensureZipForExport()
    const zipOptions = this.#getZipOptions(options)
    zipOptions.streamFiles = true
    return this.#zip.generateNodeStream(zipOptions)
  }

  #getZipOptions(options = {}) {
    const compression = options.compression || 'balanced'
    let method = 'DEFLATE'
    let level = 6

    if (compression === 'none' || compression === 'store') {
      method = 'STORE'
      level = 0
    } else if (compression === 'fast') {
      method = 'DEFLATE'
      level = 1
    } else if (compression === 'balanced') {
      method = 'DEFLATE'
      level = 6
    } else if (compression === 'maximum') {
      method = 'DEFLATE'
      level = 9
    }

    return {
      type: 'nodebuffer',
      compression: method,
      compressionOptions: method === 'DEFLATE' ? { level } : undefined,
    }
  }

  /**
   * Returns a core document property (from docProps/core.xml).
   *
   * @param {string} key - Property key (e.g., 'dc:title', 'dc:creator').
   * @returns {string|undefined}
   */
  getCoreProperty(key) {
    return this.#coreProperties.get(key)
  }

  /**
   * Sets a core document property.
   * Updates docProps/core.xml in the ZIP.
   *
   * @param {string} key - Property key.
   * @param {string} value - Property value.
   */
  setCoreProperty(key, value) {
    this.#coreProperties.set(key, value)
  }

  /**
   * Parses and loads core properties from docProps/core.xml.
   * @private
   */
  async #loadCoreProperties() {
    const coreXml = await this.readFile('docProps/core.xml')
    if (!coreXml) return

    // Simple regex extraction for core properties (lightweight vs full parse)
    const propPattern = /<(dc:[a-zA-Z]+|dcterms:[a-zA-Z]+)[^>]*>([^<]*)<\/\1>/g
    let match
    while ((match = propPattern.exec(coreXml)) !== null) {
      this.#coreProperties.set(match[1], match[2])
    }
  }

  /**
   * Validates the integrity of the ZIP archive.
   * Checks for CRC integrity, entry sizes, duplicate entries, missing critical entries, and invalid binary data.
   *
   * @returns {Promise<void>}
   * @throws {PPTXError} If any validation issue is found.
   */
  async validateArchive() {
    await this.#ensureZipForExport()
    const files = this.#zip.files
    const errors = []
    const seenPaths = new Set()
    const { XMLParser } = require('../parsers/XMLParser.js')
    const parser = new XMLParser()

    for (const [name, file] of Object.entries(files)) {
      if (file.dir) continue

      const lowerPath = name.toLowerCase()
      if (seenPaths.has(lowerPath)) {
        errors.push(`Duplicate entry found (case-insensitive): ${name}`)
      }
      seenPaths.add(lowerPath)

      try {
        // file.async('uint8array') forces decompression and checks CRC32 & uncompressed size
        const content = await file.async('uint8array')

        // If XML/rels file, verify it's not empty and is valid XML
        if (name.endsWith('.xml') || name.endsWith('.rels')) {
          const { TextDecoder } = require('util')
          const text = new TextDecoder('utf-8').decode(content)
          if (!text.trim()) {
            errors.push(`XML/rels entry is empty: ${name}`)
          } else {
            try {
              parser.parse(text, name)
            } catch (xmlErr) {
              errors.push(`Invalid XML/rels structure in ${name}: ${xmlErr.message}`)
            }
          }
        }

        // Verify media files (like images) are not empty and have correct magic numbers
        if (name.startsWith('ppt/media/')) {
          if (content.length === 0) {
            errors.push(`Media entry is empty: ${name}`)
          }
          const ext = name.split('.').pop().toLowerCase()
          if (ext === 'png') {
            if (
              content[0] !== 0x89 ||
              content[1] !== 0x50 ||
              content[2] !== 0x4e ||
              content[3] !== 0x47
            ) {
              errors.push(`Invalid PNG signature in media file: ${name}`)
            }
          } else if (ext === 'jpg' || ext === 'jpeg') {
            if (content[0] !== 0xff || content[1] !== 0xd8) {
              errors.push(`Invalid JPEG signature in media file: ${name}`)
            }
          }
        }
      } catch (err) {
        errors.push(`CRC32 or uncompressed size integrity failure on entry ${name}: ${err.message}`)
      }
    }

    // Check for critical missing OpenXML files
    const criticalFiles = [
      '[Content_Types].xml',
      '_rels/.rels',
      'ppt/presentation.xml',
      'ppt/_rels/presentation.xml.rels',
    ]

    for (const critical of criticalFiles) {
      if (!this.hasFile(critical)) {
        errors.push(`Critical OpenXML package file is missing: ${critical}`)
      }
    }

    if (errors.length > 0) {
      throw new PPTXError(
        `ZIP archive validation failed:\n${errors.map(e => `  • ${e}`).join('\n')}`
      )
    }

    logger.info('ZIP archive integrity validation passed.')
  }

  /**
   * Returns the raw JSZip instance (for advanced use cases).
   * @returns {JSZip}
   */
  /**
   * Saves the presentation to a folder structure on the filesystem.
   *
   * @param {string} destPath - Target directory path.
   * @returns {Promise<void>}
   */
  async toFolder(destPath) {
    const path = require('path')
    const fs = require('fs-extra')
    await fs.ensureDir(destPath)

    if (this.#isFolderMode) {
      const allFiles = new Set([
        ...this.#folderFiles,
        ...this.#dirtyFiles.keys(),
        ...this.#dirtyBinaryFiles.keys(),
      ])

      for (const relPath of allFiles) {
        if (this.#removedFiles.has(relPath)) {
          const targetPath = path.join(destPath, relPath)
          await fs.remove(targetPath)
          continue
        }

        const targetPath = path.join(destPath, relPath)
        await fs.ensureDir(path.dirname(targetPath))

        if (this.#dirtyFiles.has(relPath)) {
          await fs.writeFile(targetPath, this.#dirtyFiles.get(relPath), 'utf8')
        } else if (this.#dirtyBinaryFiles.has(relPath)) {
          await fs.writeFile(targetPath, this.#dirtyBinaryFiles.get(relPath))
        } else {
          const srcPath = path.join(this.#folderRoot, relPath)
          const resolvedSrc = path.resolve(srcPath)
          const resolvedDest = path.resolve(targetPath)
          if (resolvedSrc !== resolvedDest) {
            await fs.copy(srcPath, targetPath)
          }
        }
      }
    } else {
      const files = this.#zip.files
      for (const [name, file] of Object.entries(files)) {
        if (file.dir) continue
        const targetPath = path.join(destPath, name)
        await fs.ensureDir(path.dirname(targetPath))
        const buffer = await file.async('nodebuffer')
        await fs.writeFile(targetPath, buffer)
      }
    }
  }

  async #ensureZipForExport() {
    if (this.#zip) return this.#zip

    const JSZip = require('jszip')
    const path = require('path')
    const fs = require('fs-extra')

    const zip = new JSZip()

    if (this.#cachedFiles) {
      // 1. Read all files from cache (that are not removed)
      for (const [relPath, entry] of this.#cachedFiles.entries()) {
        if (this.#removedFiles.has(relPath)) continue

        if (this.#dirtyFiles.has(relPath)) {
          zip.file(relPath, this.#dirtyFiles.get(relPath))
        } else if (this.#dirtyBinaryFiles.has(relPath)) {
          zip.file(relPath, this.#dirtyBinaryFiles.get(relPath))
        } else {
          zip.file(relPath, entry.content)
        }
      }

      // 2. Write any new files that were added (and not already in cache)
      for (const [relPath, content] of this.#dirtyFiles.entries()) {
        if (!this.#cachedFiles.has(relPath) && !this.#removedFiles.has(relPath)) {
          zip.file(relPath, content)
        }
      }
      for (const [relPath, data] of this.#dirtyBinaryFiles.entries()) {
        if (!this.#cachedFiles.has(relPath) && !this.#removedFiles.has(relPath)) {
          zip.file(relPath, data)
        }
      }

      this.#zip = zip
      return zip
    }

    // 1. Read all files from the original folder structure (that are not removed)
    for (const relPath of this.#folderFiles) {
      if (this.#removedFiles.has(relPath)) continue

      if (this.#dirtyFiles.has(relPath)) {
        zip.file(relPath, this.#dirtyFiles.get(relPath))
      } else if (this.#dirtyBinaryFiles.has(relPath)) {
        zip.file(relPath, this.#dirtyBinaryFiles.get(relPath))
      } else {
        const diskPath = path.join(this.#folderRoot, relPath)
        const data = await fs.readFile(diskPath)
        zip.file(relPath, data)
      }
    }

    // 2. Write any new files that were added (and not already in folderFiles)
    for (const [relPath, content] of this.#dirtyFiles.entries()) {
      if (!this.#folderFiles.has(relPath) && !this.#removedFiles.has(relPath)) {
        zip.file(relPath, content)
      }
    }
    for (const [relPath, data] of this.#dirtyBinaryFiles.entries()) {
      if (!this.#folderFiles.has(relPath) && !this.#removedFiles.has(relPath)) {
        zip.file(relPath, data)
      }
    }

    this.#zip = zip
    return zip
  }

  get rawZip() {
    return this.#zip
  }
}

module.exports = { ZipManager }
