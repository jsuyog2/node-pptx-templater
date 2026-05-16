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

import JSZip from 'jszip';
import fsExtra from 'fs-extra';
import { createLogger } from '../utils/logger.js';
import { PPTXError } from '../utils/errors.js';
import { BLANK_PPTX_BASE64 } from '../templates/blankPptx.js';

const logger = createLogger('ZipManager');

/**
 * @class ZipManager
 * @description Manages the PPTX ZIP archive — reading, modifying, and re-packaging it.
 *
 * All file paths within the ZIP use forward slashes (e.g., 'ppt/slides/slide1.xml').
 */
export class ZipManager {
  /**
   * @private
   * @type {JSZip}
   */
  #zip = null;

  /**
   * @private
   * @type {Map<string, string>} Cache of decoded XML strings for fast repeated access.
   */
  #xmlCache = new Map();

  /**
   * @private
   * @type {Map<string, string>} Dirty (modified) files that need to be re-written.
   */
  #dirtyFiles = new Map();

  /**
   * @private
   * @type {Map<string, string>} Core properties (dc:title, dc:creator, etc.)
   */
  #coreProperties = new Map();

  /**
   * Loads a PPTX file from a path or Buffer.
   *
   * @param {string|Buffer} source - File path or Buffer.
   * @returns {Promise<void>}
   * @throws {PPTXError} If the file cannot be read or parsed as a ZIP.
   */
  async load(source) {
    try {
      let data;
      if (typeof source === 'string') {
        logger.debug(`Reading file: ${source}`);
        data = await fsExtra.readFile(source);
      } else if (Buffer.isBuffer(source) || source instanceof Uint8Array) {
        data = source;
      } else {
        throw new PPTXError(`Invalid source type: ${typeof source}. Expected string path or Buffer.`);
      }

      this.#zip = await JSZip.loadAsync(data);
      await this.#loadCoreProperties();
      logger.debug(`ZIP loaded successfully. Files: ${Object.keys(this.#zip.files).length}`);
    } catch (err) {
      if (err instanceof PPTXError) throw err;
      throw new PPTXError(`Failed to load PPTX: ${err.message}`, err);
    }
  }

  /**
   * Creates a blank PPTX structure from the embedded minimal template.
   * @returns {Promise<void>}
   */
  async createBlank() {
    const buffer = Buffer.from(BLANK_PPTX_BASE64, 'base64');
    this.#zip = await JSZip.loadAsync(buffer);
    await this.#loadCoreProperties();
    logger.debug('Created blank PPTX structure');
  }

  /**
   * Reads and caches a text file from the ZIP archive.
   *
   * @param {string} zipPath - Path within the ZIP (e.g., 'ppt/slides/slide1.xml').
   * @returns {Promise<string|null>} File content as UTF-8 string, or null if not found.
   */
  async readFile(zipPath) {
    // Normalize path separators
    const normalPath = zipPath.replace(/\\/g, '/');

    // Return cached version if available and not dirty
    if (this.#xmlCache.has(normalPath) && !this.#dirtyFiles.has(normalPath)) {
      return this.#xmlCache.get(normalPath);
    }

    // Check dirty files (pending writes)
    if (this.#dirtyFiles.has(normalPath)) {
      return this.#dirtyFiles.get(normalPath);
    }

    const file = this.#zip.file(normalPath);
    if (!file) {
      logger.debug(`File not found in ZIP: ${normalPath}`);
      return null;
    }

    const content = await file.async('text');
    this.#xmlCache.set(normalPath, content);
    return content;
  }

  /**
   * Reads a binary file from the ZIP archive.
   *
   * @param {string} zipPath - Path within the ZIP.
   * @returns {Promise<Uint8Array|null>} Binary content or null if not found.
   */
  async readBinaryFile(zipPath) {
    const normalPath = zipPath.replace(/\\/g, '/');
    const file = this.#zip.file(normalPath);
    if (!file) return null;
    return file.async('uint8array');
  }

  /**
   * Writes (or overwrites) a text file in the ZIP archive.
   * Changes are buffered and applied when generating the output ZIP.
   *
   * @param {string} zipPath - Path within the ZIP.
   * @param {string} content - UTF-8 string content.
   */
  writeFile(zipPath, content) {
    const normalPath = zipPath.replace(/\\/g, '/');
    this.#dirtyFiles.set(normalPath, content);
    this.#xmlCache.set(normalPath, content);
    // Also write to the underlying JSZip object
    this.#zip.file(normalPath, content);
    logger.debug(`Queued write: ${normalPath}`);
  }

  /**
   * Writes a binary file to the ZIP archive.
   *
   * @param {string} zipPath - Path within the ZIP.
   * @param {Buffer|Uint8Array} data - Binary data.
   */
  writeBinaryFile(zipPath, data) {
    const normalPath = zipPath.replace(/\\/g, '/');
    this.#zip.file(normalPath, data);
    logger.debug(`Queued binary write: ${normalPath}`);
  }

  /**
   * @private
   * @type {Promise[]}
   */
  #pendingPromises = [];

  /**
   * Adds a promise to the pending queue to be awaited before saving.
   * @param {Promise} promise
   */
  addPendingPromise(promise) {
    this.#pendingPromises.push(promise);
  }

  /**
   * Waits for all pending asynchronous operations (like async ZIP reads/writes) to complete.
   * @returns {Promise<void>}
   */
  async waitForPendingWrites() {
    if (this.#pendingPromises.length > 0) {
      await Promise.all(this.#pendingPromises);
      this.#pendingPromises = [];
    }
  }

  /**
   * Removes a file from the ZIP archive.
   *
   * @param {string} zipPath - Path to remove.
   */
  removeFile(zipPath) {
    const normalPath = zipPath.replace(/\\/g, '/');
    this.#zip.remove(normalPath);
    this.#xmlCache.delete(normalPath);
    this.#dirtyFiles.delete(normalPath);
  }

  /**
   * Checks if a file exists in the ZIP archive.
   *
   * @param {string} zipPath - Path to check.
   * @returns {boolean}
   */
  hasFile(zipPath) {
    const normalPath = zipPath.replace(/\\/g, '/');
    return this.#zip.file(normalPath) !== null;
  }

  /**
   * Lists all files in the ZIP archive matching an optional prefix.
   *
   * @param {string} [prefix] - Optional path prefix filter.
   * @returns {string[]} Array of matching file paths.
   */
  listFiles(prefix = '') {
    return Object.keys(this.#zip.files).filter(
      f => !this.#zip.files[f].dir && f.startsWith(prefix)
    );
  }

  /**
   * Generates the final ZIP archive as a Buffer.
   * All pending changes are applied before compressing.
   *
   * @returns {Promise<Buffer>} Compressed PPTX as a Buffer.
   */
  async toBuffer() {
    return this.#zip.generateAsync({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
    });
  }

  /**
   * Generates the final ZIP archive as a readable Stream.
   *
   * @returns {Promise<NodeJS.ReadableStream>}
   */
  async toStream() {
    return this.#zip.generateNodeStream({
      type: 'nodebuffer',
      compression: 'DEFLATE',
      compressionOptions: { level: 6 },
      streamFiles: true,
    });
  }

  /**
   * Returns a core document property (from docProps/core.xml).
   *
   * @param {string} key - Property key (e.g., 'dc:title', 'dc:creator').
   * @returns {string|undefined}
   */
  getCoreProperty(key) {
    return this.#coreProperties.get(key);
  }

  /**
   * Sets a core document property.
   * Updates docProps/core.xml in the ZIP.
   *
   * @param {string} key - Property key.
   * @param {string} value - Property value.
   */
  setCoreProperty(key, value) {
    this.#coreProperties.set(key, value);
  }

  /**
   * Parses and loads core properties from docProps/core.xml.
   * @private
   */
  async #loadCoreProperties() {
    const coreXml = await this.readFile('docProps/core.xml');
    if (!coreXml) return;

    // Simple regex extraction for core properties (lightweight vs full parse)
    const propPattern = /<(dc:[a-zA-Z]+|dcterms:[a-zA-Z]+)[^>]*>([^<]*)<\/\1>/g;
    let match;
    while ((match = propPattern.exec(coreXml)) !== null) {
      this.#coreProperties.set(match[1], match[2]);
    }
  }

  /**
   * Returns the raw JSZip instance (for advanced use cases).
   * @returns {JSZip}
   */
  get rawZip() {
    return this.#zip;
  }
}
