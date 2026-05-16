/**
 * @fileoverview ContentTypesManager - Manages registrations in [Content_Types].xml.
 *
 * Implements structured, XML-safe manipulation of the OPC manifest.
 */

import { createLogger } from '../utils/logger.js';
import { PPTXError } from '../utils/errors.js';

const logger = createLogger('ContentTypesManager');

const TYPES_XML_PATH = '[Content_Types].xml';

export class ContentTypesManager {
  /** @private @type {XMLParser} */
  #xmlParser;

  /** @private @type {Object} */
  #contentTypesObj = null;

  /**
   * @param {XMLParser} xmlParser
   */
  constructor(xmlParser) {
    this.#xmlParser = xmlParser;
  }

  /**
   * Initializes the manager by reading and parsing [Content_Types].xml from the ZIP.
   *
   * @param {ZipManager} zipManager
   * @returns {Promise<void>}
   */
  async initialize(zipManager) {
    const content = await zipManager.readFile(TYPES_XML_PATH);
    if (!content) {
      throw new PPTXError(`${TYPES_XML_PATH} is missing from the archive.`);
    }

    this.#contentTypesObj = this.#xmlParser.parse(content, TYPES_XML_PATH);
    
    // Ensure structure is correct
    if (!this.#contentTypesObj.Types) {
      this.#contentTypesObj.Types = {
        '@_xmlns': 'http://schemas.openxmlformats.org/package/2006/content-types',
        Default: [],
        Override: []
      };
    }

    // Ensure array properties
    if (!this.#contentTypesObj.Types.Default) {
      this.#contentTypesObj.Types.Default = [];
    } else if (!Array.isArray(this.#contentTypesObj.Types.Default)) {
      this.#contentTypesObj.Types.Default = [this.#contentTypesObj.Types.Default];
    }

    if (!this.#contentTypesObj.Types.Override) {
      this.#contentTypesObj.Types.Override = [];
    } else if (!Array.isArray(this.#contentTypesObj.Types.Override)) {
      this.#contentTypesObj.Types.Override = [this.#contentTypesObj.Types.Override];
    }

    logger.debug(`Loaded [Content_Types].xml with ${this.#contentTypesObj.Types.Default.length} Defaults and ${this.#contentTypesObj.Types.Override.length} Overrides`);
  }

  /**
   * Registers a default content type for a file extension.
   *
   * @param {string} extension - The file extension (e.g., 'png').
   * @param {string} contentType - The MIME type.
   */
  addDefault(extension, contentType) {
    const extLower = extension.toLowerCase();
    const defaults = this.#contentTypesObj.Types.Default;

    const existing = defaults.find(d => d['@_Extension']?.toLowerCase() === extLower);
    if (existing) {
      existing['@_ContentType'] = contentType;
    } else {
      defaults.push({
        '@_Extension': extLower,
        '@_ContentType': contentType
      });
      logger.debug(`Registered default content type for extension .${extLower} -> ${contentType}`);
    }
  }

  /**
   * Registers an override content type for a specific package part.
   *
   * @param {string} partName - Absolute part path starting with '/' (e.g., '/ppt/slides/slide1.xml').
   * @param {string} contentType - The MIME type.
   */
  addOverride(partName, contentType) {
    const normalizedPart = partName.startsWith('/') ? partName : `/${partName}`;
    const overrides = this.#contentTypesObj.Types.Override;

    const existing = overrides.find(o => o['@_PartName'] === normalizedPart);
    if (existing) {
      existing['@_ContentType'] = contentType;
    } else {
      overrides.push({
        '@_PartName': normalizedPart,
        '@_ContentType': contentType
      });
      logger.debug(`Registered override content type for ${normalizedPart} -> ${contentType}`);
    }
  }

  /**
   * Removes an override content type registration.
   *
   * @param {string} partName - Absolute part path (e.g., '/ppt/slides/slide1.xml').
   */
  removeOverride(partName) {
    const normalizedPart = partName.startsWith('/') ? partName : `/${partName}`;
    const overrides = this.#contentTypesObj.Types.Override;

    const filtered = overrides.filter(o => o['@_PartName'] !== normalizedPart);
    if (filtered.length !== overrides.length) {
      this.#contentTypesObj.Types.Override = filtered;
      logger.debug(`Removed content type override for ${normalizedPart}`);
    }
  }

  /**
   * Checks if a default registration exists for an extension.
   *
   * @param {string} extension
   * @returns {boolean}
   */
  hasDefault(extension) {
    const extLower = extension.toLowerCase();
    return this.#contentTypesObj.Types.Default.some(d => d['@_Extension']?.toLowerCase() === extLower);
  }

  /**
   * Checks if an override registration exists for a part.
   *
   * @param {string} partName
   * @returns {boolean}
   */
  hasOverride(partName) {
    const normalizedPart = partName.startsWith('/') ? partName : `/${partName}`;
    return this.#contentTypesObj.Types.Override.some(o => o['@_PartName'] === normalizedPart);
  }

  /**
   * Serializes back to [Content_Types].xml and writes to ZIP.
   *
   * @param {ZipManager} zipManager
   */
  flush(zipManager) {
    const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    const xml = this.#xmlParser.build(this.#contentTypesObj, declaration);
    zipManager.writeFile(TYPES_XML_PATH, xml);
    logger.debug(`Flushed ${TYPES_XML_PATH}`);
  }
}
