/**
 * @fileoverview OutputWriter - Handles PPTX serialization and output.
 *
 * Responsibilities:
 *  1. Flush all pending changes from all managers into the ZipManager
 *  2. Ensure [Content_Types].xml is up to date
 *  3. Generate the final ZIP archive
 *  4. Write to file, buffer, or stream
 */

import fsExtra from 'fs-extra';
const { writeFile, ensureDir } = fsExtra;
import path from 'path';
import { createLogger } from '../utils/logger.js';
import { PPTXError } from '../utils/errors.js';
import { Readable } from 'stream';

const logger = createLogger('OutputWriter');

/**
 * @class OutputWriter
 * @description Serializes the modified PPTX to various output formats.
 */
export class OutputWriter {
  /** @private @type {ZipManager} */
  #zipManager;

  /**
   * @param {ZipManager} zipManager
   */
  constructor(zipManager) {
    this.#zipManager = zipManager;
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
      const buffer = await this.toBuffer(slideManager, zipManager);
      const dir = path.dirname(filePath);
      await ensureDir(dir);
      await writeFile(filePath, buffer);
      logger.info(`Saved to ${filePath} (${(buffer.length / 1024).toFixed(1)} KB)`);
    } catch (err) {
      if (err instanceof PPTXError) throw err;
      throw new PPTXError(`Failed to save file to ${filePath}: ${err.message}`, err);
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
    await this.#flushAllSlides(slideManager, zipManager);

    // Wait for any queued asynchronous writes (like content types, media hashing)
    await zipManager.waitForPendingWrites();

    const buffer = await zipManager.toBuffer();
    logger.debug(`Generated buffer: ${(buffer.length / 1024).toFixed(1)} KB`);
    return buffer;
  }

  /**
   * Returns the PPTX as a readable Node.js stream.
   *
   * @param {SlideManager} slideManager
   * @param {ZipManager} zipManager
   * @returns {Promise<Readable>}
   */
  async toStream(slideManager, zipManager) {
    await this.#flushAllSlides(slideManager, zipManager);
    await zipManager.waitForPendingWrites();
    const nodeStream = await zipManager.toStream();
    return nodeStream;
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
    const info = slideManager.getAllSlideInfo();

    for (const slide of info) {
      if (!zipManager.hasFile(slide.zipPath)) {
        logger.warn(`Slide file missing in ZIP: ${slide.zipPath}`);
      }
    }

    // Update the slide count in docProps/app.xml to prevent repair mode issues
    if (zipManager.hasFile('docProps/app.xml')) {
      zipManager.addPendingPromise(
        zipManager.rawZip.file('docProps/app.xml').async('text').then(content => {
          const updated = content.replace(/<Slides>[0-9]+<\/Slides>/, `<Slides>${info.length}</Slides>`);
          zipManager.writeFile('docProps/app.xml', updated);
        })
      );
    }

    logger.debug(`Flushed ${info.length} slide(s) to ZIP`);
  }
}
