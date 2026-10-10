/**
 * @fileoverview PackageLoader — multi-format OpenXML package ingestion engine.
 *
 * Supports three distinct input formats:
 *   1. Direct PPTX — file path, Node.js Buffer, Uint8Array, or ReadableStream.
 *   2. ZIP archives — validated by PKZip magic bytes 0x50 0x4B 0x03 0x04.
 *   3. Unpacked OpenXML directory trees — folder containing [Content_Types].xml and ppt/.
 *
 * @module PackageLoader
 */

'use strict'

const path = require('path')
const fs = require('fs/promises')
const { VirtualFileSystem } = require('./VirtualFileSystem.js')

/** PKZip local file header magic: PK\x03\x04 */
const PKZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04]

/**
 * Returns true if the first four bytes of buf match the PKZip magic number.
 *
 * @param {Uint8Array|Buffer} buf
 * @returns {boolean}
 */
function isPkZip(buf) {
  if (!buf || buf.length < 4) return false
  return (
    buf[0] === PKZIP_MAGIC[0] &&
    buf[1] === PKZIP_MAGIC[1] &&
    buf[2] === PKZIP_MAGIC[2] &&
    buf[3] === PKZIP_MAGIC[3]
  )
}

/**
 * Reads all bytes from a Node.js or Web Readable stream into a Buffer.
 *
 * @param {import('stream').Readable|ReadableStream} stream
 * @returns {Promise<Buffer>}
 */
async function streamToBuffer(stream) {
  // Web ReadableStream support (has getReader)
  if (stream && typeof stream.getReader === 'function') {
    const reader = stream.getReader()
    const chunks = []
    while (true) {
      const { done, value } = await reader.read()
      if (done) break
      if (value) {
        chunks.push(Buffer.isBuffer(value) ? value : Buffer.from(value))
      }
    }
    return Buffer.concat(chunks)
  }

  // Node.js Readable stream support (event emitter)
  return new Promise((resolve, reject) => {
    const chunks = []
    stream.on('data', chunk => chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk)))
    stream.on('end', () => resolve(Buffer.concat(chunks)))
    stream.on('error', reject)
  })
}

const REQUIRED_PARTS = ['[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml']

/**
 * Loads an OpenXML PPTX package from multiple source formats into a VirtualFileSystem.
 */
class PackageLoader {
  /**
   * Loads a PPTX package from any supported source into a fresh VirtualFileSystem.
   *
   * @param {string|Buffer|Uint8Array|import('stream').Readable|ReadableStream} source
   * @returns {Promise<VirtualFileSystem>}
   */
  static async load(source) {
    if (!source) {
      throw new Error('PackageLoader.load: source must not be null or undefined')
    }

    // Web ReadableStream or Node.js ReadableStream
    if (
      (source && typeof source === 'object' && typeof source.pipe === 'function') ||
      (source && typeof source === 'object' && typeof source.getReader === 'function')
    ) {
      const buf = await streamToBuffer(source)
      return PackageLoader.#loadFromBytes(buf)
    }

    // Buffer / Uint8Array
    if (Buffer.isBuffer(source) || source instanceof Uint8Array) {
      return PackageLoader.#loadFromBytes(source)
    }

    // String path
    if (typeof source === 'string') {
      const resolved = path.resolve(source)

      let stat
      try {
        stat = await fs.stat(resolved)
      } catch {
        if (resolved.endsWith('.xml')) {
          return PackageLoader.#loadFromXmlRoot(resolved)
        }
        throw new Error(`PackageLoader: path not found: ${resolved}`)
      }

      if (stat.isDirectory()) {
        return PackageLoader.#loadFromDirectory(resolved)
      }

      if (stat.isFile()) {
        if (resolved.endsWith('.xml')) {
          return PackageLoader.#loadFromXmlRoot(resolved)
        }
        const bytes = await fs.readFile(resolved)
        return PackageLoader.#loadFromBytes(bytes)
      }

      throw new Error(`PackageLoader: unsupported path type: ${resolved}`)
    }

    throw new Error(
      `PackageLoader.load: unsupported source type "${typeof source}". ` +
        'Expected string path, Buffer, Uint8Array, or ReadableStream.'
    )
  }

  static async #loadFromBytes(bytes) {
    if (!isPkZip(bytes)) {
      throw new Error(
        'PackageLoader: invalid PKZip magic bytes — not a valid ZIP/PPTX archive ' +
          `(got 0x${bytes[0]?.toString(16)} 0x${bytes[1]?.toString(16)} ` +
          `0x${bytes[2]?.toString(16)} 0x${bytes[3]?.toString(16)})`
      )
    }

    const JSZip = require('jszip')
    const zip = await JSZip.loadAsync(bytes)

    const vfs = new VirtualFileSystem()
    const promises = []

    zip.forEach((relativePath, zipEntry) => {
      if (zipEntry.dir) return
      promises.push(
        zipEntry.async('uint8array').then(data => {
          vfs.writeBytes(relativePath, data)
        })
      )
    })

    await Promise.all(promises)
    PackageLoader.#validateCriticalParts(vfs, 'ZIP archive')
    return vfs
  }

  static async #loadFromDirectory(dirRoot) {
    const vfs = new VirtualFileSystem()

    async function recurse(absDir) {
      const entries = await fs.readdir(absDir, { withFileTypes: true })
      await Promise.all(
        entries.map(async entry => {
          const absPath = path.join(absDir, entry.name)
          if (entry.isDirectory()) {
            await recurse(absPath)
          } else if (entry.isFile()) {
            const relPath = path.relative(dirRoot, absPath).replace(/\\/g, '/')
            const data = await fs.readFile(absPath)
            vfs.writeBytes(relPath, data)
          }
        })
      )
    }

    await recurse(dirRoot)
    PackageLoader.#validateCriticalParts(vfs, `directory "${dirRoot}"`)
    return vfs
  }

  static async #loadFromXmlRoot(xmlPath) {
    const dir = path.dirname(xmlPath)
    const root = path.basename(dir).toLowerCase() === 'ppt' ? path.dirname(dir) : dir
    return PackageLoader.#loadFromDirectory(root)
  }

  static #validateCriticalParts(vfs, sourceDescription) {
    const missing = REQUIRED_PARTS.filter(p => !vfs.has(p))
    if (missing.length > 0) {
      throw new Error(
        `PackageLoader: ${sourceDescription} is missing required OpenXML parts: ` +
          missing.join(', ')
      )
    }
  }
}

module.exports = { PackageLoader, isPkZip, streamToBuffer }
