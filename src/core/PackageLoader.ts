/**
 * @fileoverview PackageLoader — multi-format OpenXML package ingestion engine.
 *
 * Supports three distinct input formats as mandated by AGENTS.md:
 *
 *   1. **Direct PPTX** — file path (string), Node.js Buffer, or ReadableStream.
 *   2. **ZIP archives** — validated by PKZip magic bytes `0x50 0x4B 0x03 0x04`
 *      regardless of file extension (.pptx, .zip, .tmp, etc.).
 *   3. **Unpacked OpenXML directory trees** — a folder containing
 *      `[Content_Types].xml` and `ppt/` is read recursively into memory
 *      without any intermediate zip step.
 *
 * All output is produced as a populated {@link VirtualFileSystem} instance. No
 * synchronous blocking I/O (`fs.promises.readFile`, `fs.promises.stat`, `readdirSync`, etc.)
 * is ever used.
 *
 * @module PackageLoader
 */

'use strict'

import * as path from 'path'
import * as fs from 'fs/promises'
import type { Readable } from 'stream'
import { VirtualFileSystem } from './VirtualFileSystem.js'

// ─── PKZip signature ──────────────────────────────────────────────────────────

/** PKZip local file header magic: `PK\x03\x04` */
const PKZIP_MAGIC = [0x50, 0x4b, 0x03, 0x04] as const

/**
 * Returns `true` if the first four bytes of `buf` match the PKZip magic number.
 *
 * @param buf - Buffer or Uint8Array to inspect.
 */
export function isPkZip(buf: Uint8Array | Buffer): boolean {
  if (!buf || buf.length < 4) return false
  return (
    buf[0] === PKZIP_MAGIC[0] &&
    buf[1] === PKZIP_MAGIC[1] &&
    buf[2] === PKZIP_MAGIC[2] &&
    buf[3] === PKZIP_MAGIC[3]
  )
}

// ─── Stream helper ────────────────────────────────────────────────────────────

/**
 * Reads all bytes from a Node.js `Readable` stream into a single `Buffer`.
 * Non-blocking: consumes the stream asynchronously via event listeners.
 *
 * @param stream - A Node.js readable stream.
 * @returns Concatenated `Buffer` of all chunks.
 */
export async function streamToBuffer(stream: Readable): Promise<Buffer> {
  return new Promise<Buffer>((resolve, reject) => {
    const chunks: Buffer[] = []
    stream.on('data', (chunk: unknown) =>
      chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk as ArrayBuffer))
    )
    stream.on('end', () => resolve(Buffer.concat(chunks)))
    stream.on('error', reject)
  })
}

// ─── Accepted source types ────────────────────────────────────────────────────

/**
 * Union of all source types accepted by {@link PackageLoader.load}.
 *
 * - `string` — file-system path to a `.pptx` file, `.xml` file, or directory root.
 * - `Buffer` / `Uint8Array` — raw ZIP/PPTX bytes already in memory.
 * - `Readable` — streamed binary data (e.g. `fs.createReadStream()`).
 */
export type PackageSource = string | Buffer | Uint8Array | Readable

// ─── Required OpenXML parts ───────────────────────────────────────────────────

const REQUIRED_PARTS = ['[Content_Types].xml', '_rels/.rels', 'ppt/presentation.xml'] as const

// ─── PackageLoader ────────────────────────────────────────────────────────────

/**
 * Loads an OpenXML PPTX package from multiple source formats into a
 * {@link VirtualFileSystem}.
 *
 * @example
 * ```ts
 * // From file path
 * const vfs = await PackageLoader.load('./template.pptx')
 *
 * // From Buffer
 * const buf = await fs.readFile('./template.pptx')
 * const vfs = await PackageLoader.load(buf)
 *
 * // From ReadableStream
 * import { createReadStream } from 'fs'
 * const vfs = await PackageLoader.load(createReadStream('./template.pptx'))
 *
 * // From unpacked XML folder
 * const vfs = await PackageLoader.load('./extracted-pptx-folder/')
 * ```
 */
export class PackageLoader {
  /**
   * Loads a PPTX package from any supported source into a fresh VirtualFileSystem.
   *
   * @param source - File path, Buffer, Uint8Array, or ReadableStream.
   * @returns Populated VirtualFileSystem ready for manipulation.
   * @throws {Error} If the source is invalid, missing, or lacks required OpenXML parts.
   */
  static async load(source: PackageSource): Promise<VirtualFileSystem> {
    if (!source) {
      throw new Error('PackageLoader.load: source must not be null or undefined')
    }

    // ── ReadableStream ────────────────────────────────────────────────────
    if (source && typeof source === 'object' && typeof (source as Readable).pipe === 'function') {
      const buf = await streamToBuffer(source as Readable)
      return PackageLoader.#loadFromBytes(buf)
    }

    // ── Buffer / Uint8Array ───────────────────────────────────────────────
    if (Buffer.isBuffer(source) || source instanceof Uint8Array) {
      return PackageLoader.#loadFromBytes(source as Buffer | Uint8Array)
    }

    // ── String path ───────────────────────────────────────────────────────
    if (typeof source === 'string') {
      const resolved = path.resolve(source)

      let stat: Awaited<ReturnType<typeof fs.stat>>
      try {
        stat = await fs.stat(resolved)
      } catch {
        // If stat fails and path looks like a presentation.xml path, try as directory root
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

  // ─── Private: load from bytes (ZIP / PPTX) ───────────────────────────────

  /**
   * Inflates a ZIP/PPTX buffer into the VFS after verifying the PKZip header.
   *
   * @param bytes - Raw ZIP bytes.
   */
  static async #loadFromBytes(bytes: Buffer | Uint8Array): Promise<VirtualFileSystem> {
    if (!isPkZip(bytes)) {
      throw new Error(
        'PackageLoader: invalid PKZip magic bytes — not a valid ZIP/PPTX archive ' +
          `(got 0x${bytes[0]?.toString(16)} 0x${bytes[1]?.toString(16)} ` +
          `0x${bytes[2]?.toString(16)} 0x${bytes[3]?.toString(16)})`
      )
    }

    // jszip is a runtime CJS/ESM dependency — lazy-require avoids circular ESM issues.
    const JSZip: typeof import('jszip') = require('jszip')
    const zip = await (JSZip as unknown as { loadAsync: Function }).loadAsync(bytes)

    const vfs = new VirtualFileSystem()
    const promises: Promise<void>[] = []

    zip.forEach((relativePath: string, zipEntry: { dir: boolean; async: Function }) => {
      if (zipEntry.dir) return
      promises.push(
        zipEntry.async('uint8array').then((data: Uint8Array) => {
          vfs.writeBytes(relativePath, data)
        })
      )
    })

    await Promise.all(promises)

    PackageLoader.#validateCriticalParts(vfs, 'ZIP archive')
    return vfs
  }

  // ─── Private: load from unpacked directory ────────────────────────────────

  /**
   * Recursively reads an unpacked OpenXML directory tree into memory.
   * Uses only `fs.readdir` / `fs.readFile` (non-blocking).
   *
   * @param dirRoot - Absolute path to the directory root.
   */
  static async #loadFromDirectory(dirRoot: string): Promise<VirtualFileSystem> {
    const vfs = new VirtualFileSystem()

    async function recurse(absDir: string): Promise<void> {
      const entries = await fs.readdir(absDir, { withFileTypes: true })
      await Promise.all(
        entries.map(async (entry) => {
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

  // ─── Private: resolve from presentation.xml path ─────────────────────────

  /**
   * Resolves the OpenXML root from a `presentation.xml` path and delegates to
   * {@link #loadFromDirectory}.
   *
   * @param xmlPath - Absolute path to `presentation.xml`.
   */
  static async #loadFromXmlRoot(xmlPath: string): Promise<VirtualFileSystem> {
    const dir = path.dirname(xmlPath)
    // If the file lives inside a `ppt/` folder, the root is one level up.
    const root = path.basename(dir).toLowerCase() === 'ppt' ? path.dirname(dir) : dir
    return PackageLoader.#loadFromDirectory(root)
  }

  // ─── Validation ───────────────────────────────────────────────────────────

  /**
   * Asserts that the minimum required OpenXML parts are present in `vfs`.
   *
   * @param vfs - The VirtualFileSystem to validate.
   * @param sourceDescription - Human-readable source label for error messages.
   */
  static #validateCriticalParts(vfs: VirtualFileSystem, sourceDescription: string): void {
    const missing = REQUIRED_PARTS.filter((p) => !vfs.has(p))
    if (missing.length > 0) {
      throw new Error(
        `PackageLoader: ${sourceDescription} is missing required OpenXML parts: ` +
          missing.join(', ')
      )
    }
  }
}
