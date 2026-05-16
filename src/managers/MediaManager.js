/**
 * @fileoverview MediaManager - Manages media files (images, videos, audio) in PPTX.
 *
 * Media in OpenXML PPTX:
 * ─────────────────────────────────────────────────────────────────
 * All media files are stored in ppt/media/ within the PPTX ZIP.
 * They are referenced via relationships from slides, layouts, etc.
 *
 * Media Reference Chain:
 *   slide.xml → slide.xml.rels → ppt/media/imageN.{ext}
 *
 * Image in slide XML:
 *   <p:pic>
 *     <p:nvPicPr>
 *       <p:cNvPr id="3" name="logo"/>
 *     </p:nvPicPr>
 *     <p:blipFill>
 *       <a:blip r:embed="rId4"/>   ← references image via rId
 *     </p:blipFill>
 *     <p:spPr>
 *       <a:xfrm>
 *         <a:off x="457200" y="274638"/>      ← position (EMU)
 *         <a:ext cx="2743200" cy="1828800"/>  ← size (EMU)
 *       </a:xfrm>
 *     </p:spPr>
 *   </p:pic>
 *
 * Media Deduplication:
 *   Multiple slides may embed the same logo/background.
 *   We hash file content (SHA-1) and reuse existing media files
 *   instead of adding duplicates — reducing output file size.
 *
 * Supported formats:
 *   Images: PNG, JPEG, GIF, SVG, TIFF, BMP, WMF, EMF
 *   Video: MP4, AVI, MOV, WMV
 *   Audio: MP3, WAV, M4A
 */

import { createHash } from 'crypto';
import { createLogger } from '../utils/logger.js';
import { PPTXError } from '../utils/errors.js';
import { REL_TYPES } from './RelationshipManager.js';
import fsExtra from 'fs-extra';

const logger = createLogger('MediaManager');

/**
 * MIME type to extension mapping for media files.
 */
const MEDIA_TYPES = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/jpg': 'jpg',
  'image/gif': 'gif',
  'image/svg+xml': 'svg',
  'image/tiff': 'tiff',
  'image/bmp': 'bmp',
  'image/x-wmf': 'wmf',
  'image/x-emf': 'emf',
  'image/webp': 'webp',
  'video/mp4': 'mp4',
  'audio/mpeg': 'mp3',
};

/**
 * Extension to MIME type mapping.
 */
const EXT_TO_MIME = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  gif: 'image/gif',
  svg: 'image/svg+xml',
  tiff: 'image/tiff',
  bmp: 'image/bmp',
  wmf: 'image/x-wmf',
  emf: 'image/x-emf',
  webp: 'image/webp',
  mp4: 'video/mp4',
  mp3: 'audio/mpeg',
};

/**
 * @class MediaManager
 * @description Manages media embedding, deduplication, and retrieval in PPTX files.
 */
export class MediaManager {
  /** @private @type {ZipManager} */
  #zipManager;

  /**
   * Content hash → existing media ZIP path for deduplication.
   * @private @type {Map<string, string>}
   */
  #mediaHashIndex = new Map();

  /**
   * All known media files.
   * @private @type {Map<string, MediaInfo>}
   */
  #mediaRegistry = new Map();

  /**
   * Counter for generating unique media file names.
   * @private @type {number}
   */
  #nextMediaId = 1;

  /**
   * Initializes by scanning existing media files in the PPTX.
   *
   * @param {ZipManager} zipManager
   * @returns {Promise<void>}
   */
  async initialize(zipManager) {
    this.#zipManager = zipManager;
    const mediaFiles = zipManager.listFiles('ppt/media/');

    // Index all existing media files by content hash for deduplication
    await Promise.all(
      mediaFiles.map(async mediaPath => {
        const data = await zipManager.readBinaryFile(mediaPath);
        if (data) {
          const hash = this.#hashBytes(data);
          const ext = mediaPath.split('.').pop().toLowerCase();
          const mimeType = EXT_TO_MIME[ext] || 'application/octet-stream';

          const mediaInfo = { zipPath: mediaPath, hash, mimeType, size: data.length };
          this.#mediaHashIndex.set(hash, mediaPath);
          this.#mediaRegistry.set(mediaPath, mediaInfo);

          // Track the highest media ID to avoid collisions
          const numMatch = /\d+/.exec(mediaPath.split('/').pop());
          if (numMatch) {
            const num = parseInt(numMatch[0], 10);
            if (num >= this.#nextMediaId) this.#nextMediaId = num + 1;
          }
        }
      })
    );

    logger.debug(`Indexed ${this.#mediaRegistry.size} media file(s)`);
  }

  /**
   * Returns the total number of media files.
   * @returns {number}
   */
  get mediaCount() {
    return this.#mediaRegistry.size;
  }

  /**
   * Embeds a new image from a file path or Buffer.
   * Automatically deduplicates — if the same image already exists,
   * returns the existing ZIP path instead of creating a duplicate.
   *
   * @param {string|Buffer} source - File path or image Buffer.
   * @param {string} [mimeType] - MIME type (auto-detected from extension if omitted).
   * @returns {Promise<string>} ZIP path of the embedded image (e.g., 'ppt/media/image5.png').
   */
  async embedImage(source, mimeType) {
    let data;
    let ext;

    if (typeof source === 'string') {
      // Load from file path
      data = await fsExtra.readFile(source);
      ext = source.split('.').pop().toLowerCase();
      mimeType = mimeType || EXT_TO_MIME[ext] || 'image/png';
    } else if (Buffer.isBuffer(source) || source instanceof Uint8Array) {
      data = source;
      // Detect format from magic bytes
      ext = this.#detectExtension(data);
      mimeType = mimeType || EXT_TO_MIME[ext] || 'image/png';
    } else {
      throw new PPTXError('embedImage: source must be a file path string or Buffer');
    }

    // Check for duplicate (content-addressable dedup)
    const hash = this.#hashBytes(data);
    if (this.#mediaHashIndex.has(hash)) {
      const existingPath = this.#mediaHashIndex.get(hash);
      logger.debug(`Reusing existing media: ${existingPath} (hash: ${hash.substring(0, 8)}...)`);
      return existingPath;
    }

    // Create a new media file
    const mediaId = this.#nextMediaId++;
    const zipPath = `ppt/media/image${mediaId}.${ext}`;

    this.#zipManager.writeBinaryFile(zipPath, data);
    this.#mediaHashIndex.set(hash, zipPath);
    this.#mediaRegistry.set(zipPath, { zipPath, hash, mimeType, size: data.length });

    // Register content type
    this.#registerContentType(ext, mimeType);

    logger.debug(`Embedded new media: ${zipPath} (${data.length} bytes)`);
    return zipPath;
  }

  /**
   * Generates the slide XML snippet for an image element.
   *
   * @param {string} rId - Relationship ID pointing to the media file.
   * @param {ImageElement} opts - Image element options.
   * @param {number} opts.x - X position in EMU.
   * @param {number} opts.y - Y position in EMU.
   * @param {number} opts.width - Width in EMU.
   * @param {number} opts.height - Height in EMU.
   * @param {string} [opts.name] - Shape name.
   * @param {number} [opts.shapeId] - Shape ID.
   * @returns {string} XML snippet for the image.
   */
  buildImageXml(rId, opts) {
    const { x = 0, y = 0, width = 2743200, height = 1828800, name = 'image', shapeId = 1 } = opts;

    return `<p:pic>
  <p:nvPicPr>
    <p:cNvPr id="${shapeId}" name="${name}"/>
    <p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>
    <p:nvPr/>
  </p:nvPicPr>
  <p:blipFill>
    <a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="${rId}" cstate="print">
      <a:extLst><a:ext uri="{28A0092B-C50C-407E-A947-70E740481C1C}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst>
    </a:blip>
    <a:stretch><a:fillRect/></a:stretch>
  </p:blipFill>
  <p:spPr>
    <a:xfrm>
      <a:off x="${x}" y="${y}"/>
      <a:ext cx="${width}" cy="${height}"/>
    </a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
</p:pic>`;
  }

  /**
   * Gets info about a media file by its ZIP path.
   *
   * @param {string} zipPath
   * @returns {MediaInfo|undefined}
   */
  getMediaInfo(zipPath) {
    return this.#mediaRegistry.get(zipPath);
  }

  /**
   * Returns all registered media files.
   * @returns {MediaInfo[]}
   */
  getAllMedia() {
    return Array.from(this.#mediaRegistry.values());
  }

  /**
   * Computes a SHA-1 hash of binary data.
   * Used for content-addressable deduplication.
   *
   * @private
   * @param {Buffer|Uint8Array} data
   * @returns {string} Hex digest.
   */
  #hashBytes(data) {
    return createHash('sha1').update(data).digest('hex');
  }

  /**
   * Detects image format from magic bytes.
   *
   * @private
   * @param {Buffer|Uint8Array} data
   * @returns {string} File extension.
   */
  #detectExtension(data) {
    const sig = data.slice(0, 8);

    // PNG: 89 50 4E 47 0D 0A 1A 0A
    if (sig[0] === 0x89 && sig[1] === 0x50) return 'png';
    // JPEG: FF D8 FF
    if (sig[0] === 0xFF && sig[1] === 0xD8) return 'jpg';
    // GIF: 47 49 46
    if (sig[0] === 0x47 && sig[1] === 0x49) return 'gif';
    // WEBP: 52 49 46 46 ... 57 45 42 50
    if (sig[0] === 0x52 && sig[1] === 0x49 && sig[8] === 0x57) return 'webp';
    // BMP: 42 4D
    if (sig[0] === 0x42 && sig[1] === 0x4D) return 'bmp';

    return 'png'; // Default fallback
  }

  /**
   * Registers a new content type for the media format in [Content_Types].xml.
   * @private
   */
  #registerContentType(ext, mimeType) {
    if (!this.#zipManager) return;

    const contentTypesXml = this.#zipManager.rawZip.file('[Content_Types].xml');
    if (!contentTypesXml) return;

    // Read synchronously (in-memory) since we're in constructor context
    this.#zipManager.addPendingPromise(
      contentTypesXml.async('text').then(xml => {
        const defaultEntry = `Extension="${ext}" ContentType="${mimeType}"`;
        if (!xml.includes(defaultEntry)) {
          const newDefault = `<Default ${defaultEntry}/>`;
          const updated = xml.replace(
            '</Types>',
            `  ${newDefault}\n</Types>`
          );
          this.#zipManager.writeFile('[Content_Types].xml', updated);
        }
      })
    );
  }
}
