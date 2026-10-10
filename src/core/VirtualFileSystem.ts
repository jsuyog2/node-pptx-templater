/**
 * @fileoverview VirtualFileSystem — in-memory OpenXML package store.
 *
 * All OpenXML package entries (XML parts, binary media, relationship files)
 * are held as raw `Uint8Array` values inside a single `Map<string, Uint8Array>`.
 * No synchronous blocking I/O is ever used; callers must go through the
 * encode/decode helpers provided here.
 *
 * Path convention: forward-slash relative paths identical to ZIP entry names,
 * e.g. `"ppt/slides/slide1.xml"`, `"[Content_Types].xml"`.
 *
 * @module VirtualFileSystem
 */

'use strict'

// ─── Codec singletons (reused for efficiency) ────────────────────────────────

const encoder: TextEncoder =
  typeof globalThis.TextEncoder !== 'undefined'
    ? new globalThis.TextEncoder()
    : new (require('util').TextEncoder)()

const decoder: TextDecoder =
  typeof globalThis.TextDecoder !== 'undefined'
    ? new globalThis.TextDecoder('utf-8')
    : new (require('util').TextDecoder)('utf-8')

// ─── Legacy cache entry format (from ZipManager.preload) ─────────────────────

/** Entry format used by the legacy `ZipManager` cache. */
export interface LegacyCacheEntry {
  type: 'text' | 'binary'
  content: string | Uint8Array | Buffer
}

/** Any value accepted by `importMap`. */
export type VFSMapValue = Uint8Array | Buffer | LegacyCacheEntry

// ─── VirtualFileSystem ────────────────────────────────────────────────────────

/**
 * In-memory store for all OpenXML package assets.
 *
 * All entries are stored as `Uint8Array` regardless of whether they originate
 * from text (XML/rels) or binary (images, embedded workbooks) sources.
 *
 * @example
 * ```ts
 * const vfs = new VirtualFileSystem()
 * vfs.writeText('ppt/slides/slide1.xml', '<p:sld>…</p:sld>')
 * const xml = vfs.readText('ppt/slides/slide1.xml')
 * ```
 */
export class VirtualFileSystem {
  /** Core store: all package entries as raw bytes. */
  readonly #store: Map<string, Uint8Array> = new Map()

  // ─── Read ────────────────────────────────────────────────────────────────

  /**
   * Returns the raw bytes for `path`, or `null` if absent.
   *
   * @param path - Forward-slash relative ZIP entry path.
   */
  readBytes(path: string): Uint8Array | null {
    const key = this.#normalize(path)
    return this.#store.get(key) ?? null
  }

  /**
   * Decodes the stored bytes as a UTF-8 string.
   *
   * @param path - ZIP entry path.
   * @returns UTF-8 string or `null` if not found.
   */
  readText(path: string): string | null {
    const bytes = this.readBytes(path)
    if (bytes === null) return null
    return decoder.decode(bytes)
  }

  // ─── Write ───────────────────────────────────────────────────────────────

  /**
   * Stores raw bytes for `path`.
   * Accepts both `Uint8Array` and Node.js `Buffer` (which is a `Uint8Array` subclass).
   *
   * @param path - ZIP entry path.
   * @param bytes - Raw binary content.
   */
  writeBytes(path: string, bytes: Uint8Array | Buffer): void {
    if (!(bytes instanceof Uint8Array) && !Buffer.isBuffer(bytes)) {
      throw new TypeError(
        `VirtualFileSystem.writeBytes: expected Uint8Array/Buffer for "${path}"`
      )
    }
    const key = this.#normalize(path)
    this.#store.set(key, bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes))
  }

  /**
   * UTF-8–encodes `text` and stores it as bytes for `path`.
   *
   * @param path - ZIP entry path.
   * @param text - UTF-8 string content.
   */
  writeText(path: string, text: string): void {
    if (typeof text !== 'string') {
      throw new TypeError(`VirtualFileSystem.writeText: expected string for "${path}"`)
    }
    const key = this.#normalize(path)
    this.#store.set(key, encoder.encode(text))
  }

  // ─── Existence & Enumeration ──────────────────────────────────────────────

  /**
   * Returns `true` if the entry exists.
   *
   * @param path - ZIP entry path.
   */
  has(path: string): boolean {
    return this.#store.has(this.#normalize(path))
  }

  /**
   * Removes an entry.
   *
   * @param path - ZIP entry path.
   * @returns `true` if an entry was actually deleted.
   */
  delete(path: string): boolean {
    return this.#store.delete(this.#normalize(path))
  }

  /**
   * Lists all stored paths, optionally filtered by a prefix.
   *
   * @param prefix - Optional path prefix (e.g. `"ppt/slides/"`).
   */
  list(prefix = ''): string[] {
    const results: string[] = []
    for (const key of this.#store.keys()) {
      if (key.startsWith(prefix)) results.push(key)
    }
    return results
  }

  /**
   * Returns the total number of entries currently stored.
   */
  get size(): number {
    return this.#store.size
  }

  // ─── Bulk operations ──────────────────────────────────────────────────────

  /**
   * Imports all entries from a `Map<string, Uint8Array>` or a legacy cache map
   * with `{ type: string, content: string | Uint8Array }` values.
   *
   * @param map - Source map to import from.
   */
  importMap(map: Map<string, VFSMapValue>): void {
    for (const [rawPath, value] of map.entries()) {
      const key = this.#normalize(rawPath)
      if (value instanceof Uint8Array || Buffer.isBuffer(value)) {
        this.#store.set(
          key,
          value instanceof Uint8Array ? value : new Uint8Array(value as Buffer)
        )
      } else if (value && typeof value === 'object' && 'content' in value) {
        // Legacy format: { type: 'text'|'binary', content: string|Uint8Array }
        const { type, content } = value as LegacyCacheEntry
        if (type === 'text' && typeof content === 'string') {
          this.#store.set(key, encoder.encode(content))
        } else if (content instanceof Uint8Array || Buffer.isBuffer(content)) {
          this.#store.set(
            key,
            content instanceof Uint8Array ? content : new Uint8Array(content as Buffer)
          )
        }
      }
    }
  }

  /**
   * Exports all entries as a plain `Map<string, Uint8Array>`.
   * The returned map is a shallow copy — byte arrays are shared references.
   */
  exportMap(): Map<string, Uint8Array> {
    return new Map(this.#store)
  }

  /**
   * Creates a deep clone of this VFS.
   * All byte arrays are copied so mutations on the clone do not affect the original.
   */
  clone(): VirtualFileSystem {
    const copy = new VirtualFileSystem()
    for (const [key, bytes] of this.#store.entries()) {
      copy.#store.set(key, bytes.slice())
    }
    return copy
  }

  /**
   * Removes all entries from the store.
   */
  clear(): void {
    this.#store.clear()
  }

  // ─── Internal helpers ─────────────────────────────────────────────────────

  /**
   * Normalises a path to the canonical forward-slash, no-leading-slash form.
   *
   * @param p - Raw path string (may use backslashes or have a leading `/`).
   * @returns Normalised path.
   */
  #normalize(p: string): string {
    if (!p) return ''
    return p.replace(/\\/g, '/').replace(/^\/+/, '')
  }
}
