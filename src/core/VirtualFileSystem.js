/**
 * @fileoverview VirtualFileSystem — in-memory OpenXML package store.
 *
 * All OpenXML package entries (XML parts, binary media, relationship files)
 * are held as raw `Uint8Array` values inside a single `Map<string, Uint8Array>`.
 * No synchronous blocking I/O is ever used; callers must go through the
 * encode/decode helpers provided here.
 *
 * Path convention: forward-slash relative paths identical to ZIP entry names,
 * e.g. "ppt/slides/slide1.xml", "[Content_Types].xml".
 *
 * @module VirtualFileSystem
 */

'use strict'

const encoder =
  typeof globalThis.TextEncoder !== 'undefined'
    ? new globalThis.TextEncoder()
    : new (require('util').TextEncoder)()

const decoder =
  typeof globalThis.TextDecoder !== 'undefined'
    ? new globalThis.TextDecoder('utf-8')
    : new (require('util').TextDecoder)('utf-8')

/**
 * In-memory store for all OpenXML package assets.
 */
class VirtualFileSystem {
  /** @type {Map<string, Uint8Array>} */
  #store = new Map()

  /**
   * Returns the raw bytes for `path`, or `null` if absent.
   *
   * @param {string} path - Forward-slash relative ZIP entry path.
   * @returns {Uint8Array|null}
   */
  readBytes(path) {
    const key = this.#normalize(path)
    return this.#store.get(key) ?? null
  }

  /**
   * Alias for readBytes.
   *
   * @param {string} path
   * @returns {Uint8Array|undefined}
   */
  get(path) {
    const bytes = this.readBytes(path)
    return bytes ?? undefined
  }

  /**
   * Decodes the stored bytes as a UTF-8 string.
   *
   * @param {string} path - ZIP entry path.
   * @returns {string|null} UTF-8 string or null if not found.
   */
  readText(path) {
    const bytes = this.readBytes(path)
    if (bytes === null) return null
    return decoder.decode(bytes)
  }

  /**
   * Alias for readText.
   *
   * @param {string} path
   * @returns {string|undefined}
   */
  getText(path) {
    const text = this.readText(path)
    return text ?? undefined
  }

  /**
   * Stores raw bytes for `path`.
   * Accepts both `Uint8Array` and Node.js `Buffer`.
   *
   * @param {string} path - ZIP entry path.
   * @param {Uint8Array|Buffer} bytes - Raw binary content.
   */
  writeBytes(path, bytes) {
    if (!(bytes instanceof Uint8Array) && !Buffer.isBuffer(bytes)) {
      throw new TypeError(`VirtualFileSystem.writeBytes: expected Uint8Array/Buffer for "${path}"`)
    }
    const key = this.#normalize(path)
    this.#store.set(key, bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes))
  }

  /**
   * Alias for writeBytes/writeText.
   *
   * @param {string} path
   * @param {Uint8Array|Buffer|string} data
   */
  set(path, data) {
    if (typeof data === 'string') {
      this.writeText(path, data)
    } else {
      this.writeBytes(path, data)
    }
  }

  /**
   * UTF-8–encodes `text` and stores it as bytes for `path`.
   *
   * @param {string} path - ZIP entry path.
   * @param {string} text - UTF-8 string content.
   */
  writeText(path, text) {
    if (typeof text !== 'string') {
      throw new TypeError(`VirtualFileSystem.writeText: expected string for "${path}"`)
    }
    const key = this.#normalize(path)
    this.#store.set(key, encoder.encode(text))
  }

  /**
   * Alias for writeText.
   *
   * @param {string} path
   * @param {string} text
   */
  setText(path, text) {
    this.writeText(path, text)
  }

  /**
   * Returns `true` if the entry exists.
   *
   * @param {string} path - ZIP entry path.
   * @returns {boolean}
   */
  has(path) {
    return this.#store.has(this.#normalize(path))
  }

  /**
   * Removes an entry.
   *
   * @param {string} path - ZIP entry path.
   * @returns {boolean} `true` if an entry was actually deleted.
   */
  delete(path) {
    return this.#store.delete(this.#normalize(path))
  }

  /**
   * Lists all stored paths, optionally filtered by a prefix.
   *
   * @param {string} [prefix=''] - Optional path prefix.
   * @returns {string[]}
   */
  list(prefix = '') {
    const results = []
    for (const key of this.#store.keys()) {
      if (key.startsWith(prefix)) results.push(key)
    }
    return results
  }

  /**
   * Returns the total number of entries currently stored.
   * @returns {number}
   */
  get size() {
    return this.#store.size
  }

  /**
   * Getter for entries map.
   * @returns {Map<string, Uint8Array>}
   */
  get entries() {
    return this.#store
  }

  /**
   * Imports all entries from a Map.
   *
   * @param {Map<string, any>} map - Source map to import from.
   */
  importMap(map) {
    for (const [rawPath, value] of map.entries()) {
      const key = this.#normalize(rawPath)
      if (value instanceof Uint8Array || Buffer.isBuffer(value)) {
        this.#store.set(key, value instanceof Uint8Array ? value : new Uint8Array(value))
      } else if (value && typeof value === 'object' && 'content' in value) {
        const { type, content } = value
        if (type === 'text' && typeof content === 'string') {
          this.#store.set(key, encoder.encode(content))
        } else if (content instanceof Uint8Array || Buffer.isBuffer(content)) {
          this.#store.set(key, content instanceof Uint8Array ? content : new Uint8Array(content))
        }
      }
    }
  }

  /**
   * Exports all entries as a plain Map<string, Uint8Array>.
   * @returns {Map<string, Uint8Array>}
   */
  exportMap() {
    return new Map(this.#store)
  }

  /**
   * Creates a deep clone of this VFS.
   * @returns {VirtualFileSystem}
   */
  clone() {
    const copy = new VirtualFileSystem()
    for (const [key, bytes] of this.#store.entries()) {
      copy.#store.set(key, bytes.slice())
    }
    return copy
  }

  /**
   * Removes all entries from the store.
   */
  clear() {
    this.#store.clear()
  }

  /**
   * Normalises a path to forward-slash, no leading slash.
   * @param {string} p
   * @returns {string}
   */
  #normalize(p) {
    if (!p) return ''
    return p.replace(/\\/g, '/').replace(/^\/+/, '')
  }
}

module.exports = { VirtualFileSystem }
