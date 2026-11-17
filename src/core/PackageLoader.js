/**
 * @fileoverview PackageLoader — multi-format OpenXML package ingestion engine.
 *
 * Supports polymorphic loading of presentations from:
 *   1. OpenXML files & templates (.pptx, .pptm, .potx, .potm, .ppsx, .ppsm, .ppam)
 *   2. Legacy & alternative formats (.ppt and .odp) with automated conversion
 *   3. ZIP archives (Node.js Buffer, Uint8Array, or .zip files)
 *   4. Uncompressed OpenXML directory trees containing [Content_Types].xml, _rels/, and ppt/
 *   5. Web URLs (HTTP/HTTPS streaming via fetch into in-memory VFS)
 *   6. Cloud object storage (AWS S3, Google Cloud Storage, Azure Blob Storage)
 *
 * @module PackageLoader
 */

'use strict'

const path = require('path')
const fs = require('fs/promises')
const { VirtualFileSystem } = require('./VirtualFileSystem.js')
const {
  LegacyFormatConverter,
  isLegacyPpt,
  isOdp,
  isPkZip,
  PKZIP_MAGIC,
} = require('../converters/LegacyFormatConverter.js')
const { FlatOpcConverter, isFlatOpc } = require('../converters/FlatOpcConverter.js')
const { createLogger } = require('../utils/logger.js')

const { PPTXError } = require('../utils/errors.js')

const logger = createLogger('PackageLoader')

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
   * @param {string|Buffer|Uint8Array|import('stream').Readable|ReadableStream|VirtualFileSystem|Object} source
   * @param {Object} [options={}]
   * @returns {Promise<VirtualFileSystem>}
   */
  static async load(source, options = {}) {
    if (!source) {
      throw new Error('PackageLoader.load: source must not be null or undefined')
    }

    // Direct VirtualFileSystem instance
    if (source instanceof VirtualFileSystem) {
      return source.clone()
    }

    // Folder object configuration { presentation, root }
    if (
      typeof source === 'object' &&
      source !== null &&
      !Buffer.isBuffer(source) &&
      !(source instanceof Uint8Array) &&
      (source.presentation || source.root)
    ) {
      let rootDir = source.root
      if (!rootDir && source.presentation) {
        const resolvedPres = path.resolve(source.presentation)
        const dir = path.dirname(resolvedPres)
        rootDir = path.basename(dir).toLowerCase() === 'ppt' ? path.dirname(dir) : dir
      }
      return PackageLoader.#loadFromDirectory(path.resolve(rootDir))
    }

    // Cloud object configuration { provider: 's3' | 'gcs' | 'azure', ... }
    if (
      typeof source === 'object' &&
      !Buffer.isBuffer(source) &&
      !(source instanceof Uint8Array) &&
      source.provider
    ) {
      return PackageLoader.#loadFromCloud(source, options)
    }

    // Web ReadableStream or Node.js ReadableStream
    if (
      (source && typeof source === 'object' && typeof source.pipe === 'function') ||
      (source && typeof source === 'object' && typeof source.getReader === 'function')
    ) {
      const buf = await streamToBuffer(source)
      return PackageLoader.#loadFromBytes(buf, options)
    }

    // Buffer / Uint8Array
    if (Buffer.isBuffer(source) || source instanceof Uint8Array) {
      return PackageLoader.#loadFromBytes(source, options)
    }

    // String URL, XML content, or file path
    if (typeof source === 'string') {
      const trimmed = source.trim()

      // Direct Flat OPC XML string
      if (isFlatOpc(trimmed)) {
        return FlatOpcConverter.convert(trimmed, options)
      }

      if (trimmed.startsWith('http://') || trimmed.startsWith('https://')) {
        return PackageLoader.#loadFromUrl(trimmed, options)
      }

      const resolved = path.resolve(trimmed)
      let stat
      try {
        stat = await fs.stat(resolved)
      } catch {
        if (resolved.endsWith('.xml')) {
          return PackageLoader.#loadFromXmlRoot(resolved)
        }
        throw new PPTXError(`PackageLoader: path not found: ${resolved}`)
      }

      if (stat.isDirectory()) {
        return PackageLoader.#loadFromDirectory(resolved)
      }

      if (stat.isFile()) {
        const bytes = await fs.readFile(resolved)
        if (isFlatOpc(bytes)) {
          return FlatOpcConverter.convert(bytes, options)
        }
        if (resolved.endsWith('.xml')) {
          return PackageLoader.#loadFromXmlRoot(resolved)
        }
        return PackageLoader.#loadFromBytes(bytes, options)
      }

      throw new Error(`PackageLoader: unsupported path type: ${resolved}`)
    }

    throw new Error(
      `PackageLoader.load: unsupported source type "${typeof source}". ` +
        'Expected string path/URL, Buffer, Uint8Array, ReadableStream, or cloud configuration.'
    )
  }

  /**
   * Fetches a presentation from a remote HTTP or HTTPS URL into the in-memory VFS.
   *
   * @param {string} url - Remote URL pointing to a PPTX presentation.
   * @param {Object} [options={}] - HTTP fetch and parser options.
   * @returns {Promise<VirtualFileSystem>}
   */
  static async #loadFromUrl(url, options = {}) {
    logger.info(`Fetching presentation from URL: ${url}`)
    const fetchFn =
      typeof globalThis.fetch === 'function'
        ? globalThis.fetch
        : async (u, opt) => {
            const https = u.startsWith('https:') ? require('https') : require('http')
            return new Promise((resolve, reject) => {
              https
                .get(u, opt, res => {
                  const chunks = []
                  res.on('data', c => chunks.push(c))
                  res.on('end', () => {
                    resolve({
                      ok: res.statusCode >= 200 && res.statusCode < 300,
                      status: res.statusCode,
                      statusText: res.statusMessage,
                      arrayBuffer: async () => Buffer.concat(chunks),
                    })
                  })
                  res.on('error', reject)
                })
                .on('error', reject)
            })
          }

    const fetchOptions = {
      headers: options.headers || {},
      signal: options.signal,
    }

    const response = await fetchFn(url, fetchOptions)
    if (!response.ok) {
      throw new Error(
        `PackageLoader: HTTP request failed with status ${response.status} ${response.statusText}`
      )
    }

    const data = await response.arrayBuffer()
    const buffer = Buffer.isBuffer(data) ? data : Buffer.from(data)
    return PackageLoader.#loadFromBytes(buffer, options)
  }

  /**
   * Ingests a presentation from cloud object storage (AWS S3, Google Cloud Storage, Azure Blob).
   *
   * @private
   * @param {Object} cloudConfig
   * @param {'s3'|'gcs'|'azure'} cloudConfig.provider - Cloud provider.
   * @param {import('stream').Readable|Buffer|Uint8Array} [cloudConfig.stream] - Direct stream.
   * @param {Buffer|Uint8Array} [cloudConfig.buffer] - Direct buffer.
   * @param {Object} [cloudConfig.client] - SDK client instance (AWS S3, GCS, or Azure).
   * @param {string} [cloudConfig.bucket] - S3 bucket or GCS bucket name.
   * @param {string} [cloudConfig.key] - S3 object key or GCS file path.
   * @param {Object} [cloudConfig.blobClient] - Azure BlobClient instance.
   * @param {Object} [options={}]
   * @returns {Promise<VirtualFileSystem>}
   */
  static async #loadFromCloud(cloudConfig, options = {}) {
    if (!cloudConfig || typeof cloudConfig !== 'object') {
      throw new Error('PackageLoader.fromCloud: cloudConfig must be a valid object')
    }

    const { provider, stream, buffer } = cloudConfig

    // 1. Direct stream or buffer passed inside cloud config
    if (buffer) {
      return PackageLoader.#loadFromBytes(buffer, options)
    }
    if (stream) {
      const buf = await streamToBuffer(stream)
      return PackageLoader.#loadFromBytes(buf, options)
    }

    // 2. Custom downloader function (mock or custom SDK)
    if (typeof cloudConfig.download === 'function') {
      const downloaded = await cloudConfig.download()
      const buf = Buffer.isBuffer(downloaded) ? downloaded : Buffer.from(downloaded)
      return PackageLoader.#loadFromBytes(buf, options)
    }
    if (typeof cloudConfig.getStream === 'function') {
      const s = await cloudConfig.getStream()
      const buf = await streamToBuffer(s)
      return PackageLoader.#loadFromBytes(buf, options)
    }

    // 3. AWS S3 integration
    if (provider === 's3') {
      const { client, s3, bucket, key } = cloudConfig
      const targetBucket = bucket || cloudConfig.Bucket
      const targetKey = key || cloudConfig.Key

      if (client && typeof client.send === 'function') {
        // AWS SDK v3: s3Client.send(new GetObjectCommand({ Bucket, Key }))
        const res = await client.send({
          Bucket: targetBucket,
          Key: targetKey,
        })
        const body = res.Body
        if (body) {
          const buf = Buffer.isBuffer(body) ? body : await streamToBuffer(body)
          return PackageLoader.#loadFromBytes(buf, options)
        }
      } else if (s3 && typeof s3.getObject === 'function') {
        // AWS SDK v2: s3.getObject({ Bucket, Key }).promise()
        const res = await s3.getObject({ Bucket: targetBucket, Key: targetKey }).promise()
        const buf = Buffer.isBuffer(res.Body) ? res.Body : Buffer.from(res.Body)
        return PackageLoader.#loadFromBytes(buf, options)
      }

      throw new Error(`PackageLoader.fromCloud: S3 provider requires valid client and bucket/key`)
    }

    // 4. Google Cloud Storage integration
    if (provider === 'gcs') {
      const { storage, client, bucket, key } = cloudConfig
      const storageClient = storage || client
      if (storageClient && typeof storageClient.bucket === 'function') {
        const file = storageClient.bucket(bucket).file(key)
        const [contents] = await file.download()
        return PackageLoader.#loadFromBytes(contents, options)
      }
      throw new Error(
        'PackageLoader.fromCloud: GCS provider requires storage client and bucket/key'
      )
    }

    // 5. Azure Blob Storage integration
    if (provider === 'azure') {
      const { blobClient, client } = cloudConfig
      const targetClient = blobClient || client
      if (targetClient && typeof targetClient.downloadToBuffer === 'function') {
        const buf = await targetClient.downloadToBuffer()
        return PackageLoader.#loadFromBytes(buf, options)
      }
      throw new Error(
        'PackageLoader.fromCloud: Azure provider requires blobClient with downloadToBuffer'
      )
    }

    throw new Error(
      `PackageLoader.fromCloud: unsupported or incomplete provider configuration for "${provider}"`
    )
  }

  static async #loadFromBytes(bytes, options = {}) {
    // 0. Check for Flat OPC XML
    if (isFlatOpc(bytes)) {
      return FlatOpcConverter.convert(bytes, options)
    }

    // 1. Check for legacy PPT (OLE CFBF header)
    if (isLegacyPpt(bytes)) {
      return LegacyFormatConverter.convert(bytes, options)
    }

    // 2. Check for PKZip signature
    if (isPkZip(bytes)) {
      // Check if it's an ODP file
      if (await isOdp(bytes)) {
        return LegacyFormatConverter.convert(bytes, options)
      }

      // Standard OpenXML PKZip container (.pptx, .pptm, .potx, .potm, .ppsx, .ppsm, .ppam, .zip)
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

    throw new Error(
      'PackageLoader: invalid package signature — not a recognized OpenXML PKZip, Flat OPC XML, PPT, or ODP file ' +
        `(got 0x${bytes[0]?.toString(16)} 0x${bytes[1]?.toString(16)} ` +
        `0x${bytes[2]?.toString(16)} 0x${bytes[3]?.toString(16)})`
    )
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

module.exports = {
  PackageLoader,
  FlatOpcConverter,
  isFlatOpc,
  isPkZip,
  streamToBuffer,
  PKZIP_MAGIC,
}
