/**
 * @fileoverview RelationshipManager - Manages OpenXML relationship files (.rels).
 *
 * In OpenXML, every part (slide, layout, master, chart, image) that references
 * another part does so through a "relationship" file stored in a _rels/ folder.
 *
 * Relationship File Structure:
 * ┌─────────────────────────────────────────────────────────────────┐
 * │ <?xml version="1.0" encoding="UTF-8" standalone="yes"?>        │
 * │ <Relationships xmlns="...">                                     │
 * │   <Relationship                                                 │
 * │     Id="rId1"                                                   │
 * │     Type="...slideLayout"                                       │
 * │     Target="../slideLayouts/slideLayout1.xml"/>                 │
 * │   <Relationship                                                 │
 * │     Id="rId2"                                                   │
 * │     Type="...hyperlink"                                         │
 * │     Target="https://example.com"                               │
 * │     TargetMode="External"/>                                     │
 * │ </Relationships>                                                │
 * └─────────────────────────────────────────────────────────────────┘
 *
 * Relationship ID rules:
 *  - Must be unique within each .rels file
 *  - Format: rId1, rId2, rId3, ... (sequential)
 *  - Referenced by r:id="rId1" attributes in the parent part
 *
 * Common relationship types (shortened):
 *  - .../slide              → presentation → slide
 *  - .../slideLayout        → slide → layout
 *  - .../slideMaster        → layout → master
 *  - .../chart              → slide → chart
 *  - .../image              → slide → image
 *  - .../hyperlink          → text run → external URL
 *  - .../slideToSlide       → slide → another slide (inter-slide link)
 */

import { createLogger } from '../utils/logger.js';
import { PPTXError } from '../utils/errors.js';
import { generateRelationshipId } from '../utils/relationshipUtils.js';

const logger = createLogger('RelationshipManager');

/**
 * OpenXML relationship type constants.
 * Using shortened forms; full URIs are in the OpenXML spec.
 */
export const REL_TYPES = {
  SLIDE: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
  SLIDE_LAYOUT: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
  SLIDE_MASTER: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
  CHART: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
  IMAGE: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  HYPERLINK: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
  NOTES_SLIDE: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
  THEME: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
  TABLE_STYLES: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles',
  PRESENTATION: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
  CORE_PROPERTIES: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties',
  EXTENDED_PROPERTIES: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties',
  PACKAGE: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package',
};

/**
 * @class RelationshipManager
 * @description Parses, manages, and writes OpenXML relationship (.rels) files.
 *
 * Each PPTX part that has relationships gets a corresponding .rels file in
 * a _rels/ subdirectory. For example:
 *  - ppt/presentation.xml → ppt/_rels/presentation.xml.rels
 *  - ppt/slides/slide1.xml → ppt/slides/_rels/slide1.xml.rels
 */
export class RelationshipManager {
  /**
   * @private
   * @type {XMLParser}
   */
  #xmlParser;

  /**
   * @private
   * @type {ZipManager}
   */
  #zipManager;

  /**
   * @private
   * @type {Map<string, Relationship[]>}
   * Maps zip path → parsed relationships array.
   * Key: relationship file path (e.g., 'ppt/_rels/presentation.xml.rels')
   */
  #relationships = new Map();

  /**
   * @param {XMLParser} xmlParser
   */
  constructor(xmlParser) {
    this.#xmlParser = xmlParser;
  }

  /**
   * Initializes by discovering all .rels files in the ZIP.
   *
   * @param {ZipManager} zipManager
   * @returns {Promise<void>}
   */
  async initialize(zipManager) {
    this.#zipManager = zipManager;
    const relFiles = zipManager.listFiles('').filter(f => f.endsWith('.rels'));

    await Promise.all(
      relFiles.map(async relsPath => {
        const content = await zipManager.readFile(relsPath);
        if (content) {
          this.#relationships.set(relsPath, this.#parseRels(content, relsPath));
        }
      })
    );

    logger.debug(`Loaded ${this.#relationships.size} relationship files`);
  }

  /**
   * Returns the relationship file path for a given part path.
   *
   * @example
   * getRelsPath('ppt/slides/slide1.xml')
   * // → 'ppt/slides/_rels/slide1.xml.rels'
   *
   * @param {string} partPath - ZIP path of the part.
   * @returns {string} Path to the corresponding .rels file.
   */
  getRelsPath(partPath) {
    const lastSlash = partPath.lastIndexOf('/');
    const dir = lastSlash >= 0 ? partPath.substring(0, lastSlash) : '';
    const file = lastSlash >= 0 ? partPath.substring(lastSlash + 1) : partPath;
    return dir ? `${dir}/_rels/${file}.rels` : `_rels/${file}.rels`;
  }

  /**
   * Gets all relationships for a given part.
   *
   * @param {string} partPath - ZIP path of the part (not the .rels file).
   * @returns {Relationship[]} Array of relationships.
   */
  getRelationships(partPath) {
    const relsPath = this.getRelsPath(partPath);
    return this.#relationships.get(relsPath) || [];
  }

  /**
   * Gets a specific relationship by ID for a given part.
   *
   * @param {string} partPath - ZIP path of the part.
   * @param {string} rId - Relationship ID (e.g., 'rId1').
   * @returns {Relationship|null}
   */
  getRelationshipById(partPath, rId) {
    const rels = this.getRelationships(partPath);
    return rels.find(r => r.id === rId) || null;
  }

  /**
   * Gets relationships filtered by type.
   *
   * @param {string} partPath - ZIP path of the part.
   * @param {string} type - Relationship type (use REL_TYPES constants).
   * @returns {Relationship[]}
   */
  getRelationshipsByType(partPath, type) {
    return this.getRelationships(partPath).filter(r => r.type === type);
  }

  /**
   * Adds a new relationship to a part.
   * Automatically assigns the next available rId.
   *
   * @param {string} partPath - ZIP path of the owning part.
   * @param {string} type - Relationship type (REL_TYPES constant).
   * @param {string} target - Target path or URL.
   * @param {string} [targetMode] - 'External' for URLs, omit for internal parts.
   * @returns {string} The assigned relationship ID (e.g., 'rId3').
   */
  addRelationship(partPath, type, target, targetMode) {
    const relsPath = this.getRelsPath(partPath);

    if (!this.#relationships.has(relsPath)) {
      this.#relationships.set(relsPath, []);
    }

    const existing = this.#relationships.get(relsPath);
    const newId = generateRelationshipId(existing.map(r => r.id));

    const rel = { id: newId, type, target };
    if (targetMode) rel.targetMode = targetMode;

    existing.push(rel);
    this.#flushRels(relsPath, partPath);

    logger.debug(`Added relationship ${newId} (${type.split('/').pop()}) to ${partPath}`);
    return newId;
  }

  /**
   * Removes a relationship from a part.
   *
   * @param {string} partPath - ZIP path of the owning part.
   * @param {string} rId - Relationship ID to remove.
   */
  removeRelationship(partPath, rId) {
    const relsPath = this.getRelsPath(partPath);
    const existing = this.#relationships.get(relsPath) || [];
    const filtered = existing.filter(r => r.id !== rId);
    this.#relationships.set(relsPath, filtered);
    this.#flushRels(relsPath, partPath);
  }

  /**
   * Updates the target of an existing relationship.
   *
   * @param {string} partPath - ZIP path of the owning part.
   * @param {string} rId - Relationship ID to update.
   * @param {string} newTarget - New target value.
   */
  updateRelationshipTarget(partPath, rId, newTarget) {
    const relsPath = this.getRelsPath(partPath);
    const existing = this.#relationships.get(relsPath) || [];
    const rel = existing.find(r => r.id === rId);
    if (rel) {
      rel.target = newTarget;
      this.#flushRels(relsPath, partPath);
    }
  }

  /**
   * Copies all relationships from one part to another.
   * Used when cloning slides — the clone gets the same layout/master references.
   *
   * @param {string} sourcePath - Source part path.
   * @param {string} destPath - Destination part path.
   * @param {string[]} [excludeTypes] - Relationship types to exclude.
   * @returns {Map<string, string>} Map of old rId → new rId for the cloned part.
   */
  copyRelationships(sourcePath, destPath, excludeTypes = []) {
    const sourceRels = this.getRelationships(sourcePath);
    const destRelsPath = this.getRelsPath(destPath);
    const idMap = new Map();

    if (!this.#relationships.has(destRelsPath)) {
      this.#relationships.set(destRelsPath, []);
    }

    const destRels = this.#relationships.get(destRelsPath);

    for (const rel of sourceRels) {
      if (excludeTypes.includes(rel.type)) continue;
      const newId = generateRelationshipId(destRels.map(r => r.id));
      const newRel = { ...rel, id: newId };
      destRels.push(newRel);
      idMap.set(rel.id, newId);
    }

    this.#flushRels(destRelsPath, destPath);
    return idMap;
  }

  /**
   * Resolves a relative target path to an absolute ZIP path.
   *
   * @example
   * resolveTarget('ppt/slides/slide1.xml', '../slideLayouts/slideLayout1.xml')
   * // → 'ppt/slideLayouts/slideLayout1.xml'
   *
   * @param {string} partPath - The part that owns the relationship.
   * @param {string} target - Relative target from the relationship.
   * @returns {string} Absolute ZIP path.
   */
  resolveTarget(partPath, target) {
    if (target.startsWith('http://') || target.startsWith('https://')) {
      return target; // External URL — return as-is
    }

    const baseParts = partPath.split('/');
    baseParts.pop(); // Remove file name, keep directory
    const targetParts = target.split('/');

    for (const part of targetParts) {
      if (part === '..') {
        baseParts.pop();
      } else if (part !== '.') {
        baseParts.push(part);
      }
    }

    return baseParts.join('/');
  }

  /**
   * Parses a .rels XML file into an array of relationship objects.
   * @private
   * @param {string} xmlContent - Raw XML content.
   * @param {string} relsPath - For error reporting.
   * @returns {Relationship[]}
   */
  #parseRels(xmlContent, relsPath) {
    try {
      const obj = this.#xmlParser.parse(xmlContent, relsPath);
      const relationships = obj?.Relationships?.Relationship || [];
      const relsArray = Array.isArray(relationships) ? relationships : [relationships];

      return relsArray.map(rel => ({
        id: rel['@_Id'],
        type: rel['@_Type'],
        target: rel['@_Target'],
        targetMode: rel['@_TargetMode'] || null,
      }));
    } catch (err) {
      logger.warn(`Failed to parse ${relsPath}: ${err.message}`);
      return [];
    }
  }

  /**
   * Serializes the in-memory relationships back to XML and updates the ZIP.
   * @private
   * @param {string} relsPath - Path of the .rels file.
   * @param {string} partPath - For logging.
   */
  #flushRels(relsPath, partPath) {
    const rels = this.#relationships.get(relsPath) || [];
    const xml = this.#buildRelsXml(rels);
    if (this.#zipManager) {
      this.#zipManager.writeFile(relsPath, xml);
    }
  }

  /**
   * Builds the XML string for a relationships file.
   * @private
   * @param {Relationship[]} rels
   * @returns {string}
   */
  #buildRelsXml(rels) {
    const lines = rels.map(rel => {
      const targetMode = rel.targetMode ? ` TargetMode="${rel.targetMode}"` : '';
      return `  <Relationship Id="${rel.id}" Type="${rel.type}" Target="${rel.target}"${targetMode}/>`;
    });

    return [
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">',
      ...lines,
      '</Relationships>',
    ].join('\n');
  }

  /**
   * Flushes all dirty relationship files to the ZIP manager.
   * Called before generating final output.
   *
   * @param {ZipManager} zipManager
   */
  flushAll(zipManager) {
    for (const [relsPath, rels] of this.#relationships) {
      const xml = this.#buildRelsXml(rels);
      zipManager.writeFile(relsPath, xml);
    }
  }

  /**
   * Scans all relationships and removes those pointing to missing internal targets.
   * This is part of the repair functionality.
   *
   * @param {ZipManager} zipManager
   */
  removeOrphanRelationships(zipManager) {
    let removedCount = 0;
    for (const [relsPath, rels] of this.#relationships.entries()) {
      // Determine the base part path from the rels path
      // e.g. ppt/slides/_rels/slide1.xml.rels -> ppt/slides/slide1.xml
      const partPath = relsPath.replace('_rels/', '').replace('.rels', '');

      const filtered = rels.filter(rel => {
        if (rel.targetMode === 'External') return true;
        const targetPath = this.resolveTarget(partPath, rel.target);
        if (!zipManager.hasFile(targetPath)) {
          logger.warn(`Removing orphan relationship ${rel.id} pointing to missing target: ${targetPath}`);
          removedCount++;
          return false;
        }
        return true;
      });

      if (filtered.length !== rels.length) {
        this.#relationships.set(relsPath, filtered);
        this.#flushRels(relsPath, partPath);
      }
    }
    logger.debug(`Removed ${removedCount} orphan relationship(s).`);
  }
}
