/**
 * @fileoverview SlideManager - Manages individual slide operations.
 *
 * Slides in OpenXML PPTX:
 * ─────────────────────────────────────────────────────────────────
 * Each slide is an XML file stored at ppt/slides/slideN.xml.
 * The slide order is defined in ppt/presentation.xml under:
 *   <p:sldIdLst>
 *     <p:sldId id="256" r:id="rId2"/>   ← slide 1
 *     <p:sldId id="257" r:id="rId3"/>   ← slide 2
 *   </p:sldIdLst>
 *
 * Each slide has:
 *  - A content tree: p:sld > p:cSld > p:spTree > [shapes/text/images]
 *  - A layout reference via its .rels file (points to slideLayouts/slideLayoutN.xml)
 *  - Optional notes slide: ppt/notesSlides/notesSlideN.xml
 *  - Optional animation data: embedded in the slide XML itself
 *
 * Shape types in spTree:
 *  - p:sp       → Text/shape placeholders
 *  - p:pic      → Images
 *  - p:graphicFrame → Charts, tables, SmartArt
 *  - p:grpSp    → Grouped shapes
 *  - p:cxnSp    → Connectors
 */

import { createLogger } from '../utils/logger.js';
import { PPTXError, SlideNotFoundError } from '../utils/errors.js';
import { REL_TYPES } from './RelationshipManager.js';
import { buildNewSlideXml } from '../templates/slideTemplate.js';
import { generateUniqueId } from '../utils/idUtils.js';
import { remapRelationshipIds } from '../utils/relationshipUtils.js';

const logger = createLogger('SlideManager');

/** MIME type for PPTX slide parts. */
const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';

/**
 * @typedef {Object} SlideInfo
 * @property {number} index - 1-based slide number.
 * @property {string} zipPath - Path within the ZIP.
 * @property {string} relationshipId - rId in presentation.xml.rels.
 * @property {string} slideId - Unique slide ID from presentation.xml.
 * @property {string[]} tags - Custom tags assigned via tagSlide().
 * @property {string} title - Slide title (extracted from title placeholder).
 */

/**
 * @class SlideManager
 * @description Manages slide loading, ordering, modification, and creation.
 */
export class SlideManager {
  /** @private @type {XMLParser} */
  #xmlParser;
  /** @private @type {RelationshipManager} */
  #relationshipManager;
  /** @private @type {ContentTypesManager} */
  #contentTypesManager;
  /** @private @type {ZipManager} */
  #zipManager;

  /**
   * Slide registry: maps 1-based index → SlideInfo.
   * @private @type {Map<number, SlideInfo>}
   */
  #slides = new Map();

  /**
   * Raw XML cache: maps zipPath → XML string.
   * @private @type {Map<string, string>}
   */
  #slideXmlCache = new Map();

  /**
   * Custom tags: maps tag → array of 1-based indices.
   * @private @type {Map<string, number[]>}
   */
  #tags = new Map();

  /**
   * Parsed presentation.xml object (cached for modification).
   * @private @type {Object}
   */
  #presentationObj = null;

  /**
   * @param {XMLParser} xmlParser
   * @param {RelationshipManager} relationshipManager
   * @param {ContentTypesManager} contentTypesManager
   */
  constructor(xmlParser, relationshipManager, contentTypesManager) {
    this.#xmlParser = xmlParser;
    this.#relationshipManager = relationshipManager;
    this.#contentTypesManager = contentTypesManager;
  }

  /**
   * Initializes by reading presentation.xml and discovering all slides.
   *
   * @param {ZipManager} zipManager
   * @returns {Promise<void>}
   */
  async initialize(zipManager) {
    this.#zipManager = zipManager;
    const presentationXml = await zipManager.readFile('ppt/presentation.xml');

    if (!presentationXml) {
      // If no presentation.xml, create a blank one
      await this.#initializeBlankPresentation();
      return;
    }

    this.#presentationObj = this.#xmlParser.parse(presentationXml, 'presentation.xml');
    await this.#discoverSlides(zipManager);
  }

  /**
   * Discovers all slides from presentation.xml and caches their info.
   * @private
   */
  async #discoverSlides(zipManager) {
    const rels = this.#relationshipManager.getRelationships('ppt/presentation.xml');
    const slideRels = rels.filter(r => r.type === REL_TYPES.SLIDE);

    // Get slide order from presentation.xml sldIdLst
    const sldIdList = this.#xmlParser.findAll(this.#presentationObj, 'p:presentation.p:sldIdLst.p:sldId');

    // Map rId → slide info from sldIdLst
    const rIdToSlideId = new Map();
    for (const sldId of sldIdList) {
      const rId = sldId['@_r:id'];
      const slideId = sldId['@_id'];
      if (rId) rIdToSlideId.set(rId, slideId);
    }

    // Attempt to read slide titles from docProps/app.xml to preserve them
    let slideTitles = [];
    try {
      const appXml = await zipManager.readFile('docProps/app.xml');
      if (appXml) {
        const appObj = this.#xmlParser.parse(appXml, 'app.xml');
        const lpstrs = appObj?.Properties?.TitlesOfParts?.['vt:vector']?.['vt:lpstr'];
        if (lpstrs) {
          const allLpstrs = Array.isArray(lpstrs) ? lpstrs : [lpstrs];
          // Slide titles are usually the last N items where N = slide count
          // We take the last slideRels.length items
          slideTitles = allLpstrs.slice(-slideRels.length);
        }
      }
    } catch (e) {
      logger.warn('Failed to parse app.xml for slide titles', e);
    }

    // Build ordered slide list
    let slideIndex = 1;
    for (const sldId of sldIdList) {
      const rId = sldId['@_r:id'];
      const slideRel = slideRels.find(r => r.id === rId);
      if (!slideRel) continue;

      // Resolve absolute path from relative target
      const zipPath = this.#relationshipManager.resolveTarget('ppt/presentation.xml', slideRel.target);

      const slideInfo = {
        index: slideIndex,
        zipPath,
        relationshipId: rId,
        slideId: rIdToSlideId.get(rId) || String(256 + slideIndex),
        tags: [],
        title: slideTitles[slideIndex - 1] || '',
      };

      this.#slides.set(slideIndex, slideInfo);
      slideIndex++;
    }

    logger.debug(`Discovered ${this.#slides.size} slides`);
  }

  /**
   * Returns the total number of slides.
   * @returns {number}
   */
  get slideCount() {
    return this.#slides.size;
  }

  /**
   * Returns all 1-based slide indices.
   * @returns {number[]}
   */
  getAllSlideIndices() {
    return Array.from(this.#slides.keys()).sort((a, b) => a - b);
  }

  /**
   * Returns info objects for all slides.
   * @returns {SlideInfo[]}
   */
  getAllSlideInfo() {
    return this.getAllSlideIndices().map(i => this.#slides.get(i));
  }

  /**
   * Resolves a string/number ref to an array of 1-based slide indices.
   * - If number: returns [number]
   * - If string: looks up tag registry
   *
   * @param {number|string} ref
   * @returns {number[]}
   */
  resolveSlideRef(ref) {
    if (typeof ref === 'number') {
      this.#assertSlideExists(ref);
      return [ref];
    }
    const taggedIndices = this.#tags.get(ref);
    if (!taggedIndices || taggedIndices.length === 0) {
      throw new SlideNotFoundError(`No slides found with tag: "${ref}"`);
    }
    return taggedIndices;
  }

  /**
   * Gets the raw XML string for a slide.
   * Loads from ZIP on first access, then returns from cache.
   *
   * @param {number} slideIndex - 1-based index.
   * @returns {string} Slide XML content.
   */
  getSlideXml(slideIndex) {
    this.#assertSlideExists(slideIndex);
    const info = this.#slides.get(slideIndex);

    if (this.#slideXmlCache.has(info.zipPath)) {
      return this.#slideXmlCache.get(info.zipPath);
    }

    // This is sync because we pre-load; async callers should use getSlideXmlAsync
    throw new PPTXError(`Slide ${slideIndex} XML not pre-loaded. Use getSlideXmlAsync().`);
  }

  /**
   * Async version of getSlideXml — loads from ZIP if not cached.
   *
   * @param {number} slideIndex - 1-based index.
   * @returns {Promise<string>}
   */
  async getSlideXmlAsync(slideIndex) {
    this.#assertSlideExists(slideIndex);
    const info = this.#slides.get(slideIndex);

    if (!this.#slideXmlCache.has(info.zipPath)) {
      const xml = await this.#zipManager.readFile(info.zipPath);
      if (!xml) throw new SlideNotFoundError(`Slide ${slideIndex} XML not found at ${info.zipPath}`);
      this.#slideXmlCache.set(info.zipPath, xml);
    }

    return this.#slideXmlCache.get(info.zipPath);
  }

  /**
   * Sets (replaces) the XML for a slide and marks it as dirty.
   *
   * @param {number} slideIndex - 1-based index.
   * @param {string} xml - New XML content.
   */
  setSlideXml(slideIndex, xml) {
    this.#assertSlideExists(slideIndex);
    const info = this.#slides.get(slideIndex);
    this.#slideXmlCache.set(info.zipPath, xml);
    this.#zipManager.writeFile(info.zipPath, xml);
  }

  /**
   * Tags a slide with a custom string identifier.
   *
   * @param {number} slideIndex - 1-based index.
   * @param {string} tag - Tag string.
   */
  tagSlide(slideIndex, tag) {
    this.#assertSlideExists(slideIndex);
    const info = this.#slides.get(slideIndex);
    if (!info.tags.includes(tag)) info.tags.push(tag);

    if (!this.#tags.has(tag)) this.#tags.set(tag, []);
    const tagList = this.#tags.get(tag);
    if (!tagList.includes(slideIndex)) tagList.push(slideIndex);
  }

  /**
   * Gets the SlideInfo for a slide.
   *
   * @param {number} slideIndex - 1-based index.
   * @returns {SlideInfo}
   */
  getSlideInfo(slideIndex) {
    this.#assertSlideExists(slideIndex);
    return this.#slides.get(slideIndex);
  }

  /**
   * Adds a completely new slide to the presentation.
   * Creates the XML file, adds it to presentation.xml, and registers relationships.
   *
   * @param {NewSlideOptions} options
   * @param {RelationshipManager} relationshipManager
   * @param {MediaManager} mediaManager
   */
  addNewSlide(options, relationshipManager, mediaManager) {
    const newIndex = this.#slides.size + 1;
    let nextFileIndex = 1;
    while (this.#zipManager.hasFile(`ppt/slides/slide${nextFileIndex}.xml`)) {
      nextFileIndex++;
    }
    const slideFileName = `slide${nextFileIndex}.xml`;
    const slideZipPath = `ppt/slides/${slideFileName}`;

    // Find the first available layout to reference
    const layoutRels = relationshipManager.getRelationshipsByType('ppt/presentation.xml', REL_TYPES.SLIDE_MASTER);
    const masterTarget = layoutRels[0]?.target || '../slideMasters/slideMaster1.xml';

    // Generate the slide XML
    const slideXml = buildNewSlideXml(options, newIndex);

    // Write the slide XML to the ZIP
    this.#zipManager.writeFile(slideZipPath, slideXml);
    this.#slideXmlCache.set(slideZipPath, slideXml);

    // Add slide relationship to presentation.xml.rels
    const rId = relationshipManager.addRelationship(
      'ppt/presentation.xml',
      REL_TYPES.SLIDE,
      `slides/${slideFileName}`
    );

    // Add slide layout relationship to the new slide's .rels
    // Reference the first available layout
    const layoutRelsAll = this.#zipManager.listFiles('ppt/slideLayouts/').filter(f => f.endsWith('.xml'));
    const firstLayout = layoutRelsAll[0] || 'ppt/slideLayouts/slideLayout1.xml';
    const relativeLayoutPath = `../slideLayouts/${firstLayout.split('/').pop()}`;
    relationshipManager.addRelationship(slideZipPath, REL_TYPES.SLIDE_LAYOUT, relativeLayoutPath);

    // Generate a unique slide ID
    const existingIds = Array.from(this.#slides.values()).map(s => parseInt(s.slideId, 10));
    const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 255;
    const newSlideId = String(maxId + 1);

    const slideInfo = {
      index: newIndex,
      zipPath: slideZipPath,
      relationshipId: rId,
      slideId: newSlideId,
      tags: [],
      title: options.title || '',
    };

    this.#slides.set(newIndex, slideInfo);

    // Update presentation.xml sldIdLst
    this.#addSlideToPresentation(rId, newSlideId);

    // Update [Content_Types].xml
    this.#registerSlideContentType(slideFileName);

    logger.debug(`Added new slide ${newIndex} at ${slideZipPath}`);
  }

  /**
   * Clones an existing slide, duplicating its XML and relationships.
   *
   * @param {number} sourceIndex - 1-based source slide number.
   * @param {number} [atPosition] - Insert position (1-based). Default: append.
   * @param {RelationshipManager} relationshipManager
   */
  cloneSlide(sourceIndex, atPosition, relationshipManager) {
    this.#assertSlideExists(sourceIndex);
    const sourceInfo = this.#slides.get(sourceIndex);

    const newIndex = this.#slides.size + 1;
    let nextFileIndex = 1;
    while (this.#zipManager.hasFile(`ppt/slides/slide${nextFileIndex}.xml`)) {
      nextFileIndex++;
    }
    const slideFileName = `slide${nextFileIndex}.xml`;
    const slideZipPath = `ppt/slides/${slideFileName}`;

    // Copy the source XML
    let sourceXml = this.getSlideXml(sourceIndex);

    // Copy relationships from source slide (excluding notes, which are slide-specific)
    const idMap = relationshipManager.copyRelationships(
      sourceInfo.zipPath,
      slideZipPath,
      [REL_TYPES.NOTES_SLIDE]
    );

    // Remap relationship IDs in the cloned XML to match the new targets
    sourceXml = remapRelationshipIds(sourceXml, idMap);

    this.#zipManager.writeFile(slideZipPath, sourceXml);
    this.#slideXmlCache.set(slideZipPath, sourceXml);

    // Add to presentation.xml
    const rId = relationshipManager.addRelationship(
      'ppt/presentation.xml',
      REL_TYPES.SLIDE,
      `slides/${slideFileName}`
    );

    const existingIds = Array.from(this.#slides.values()).map(s => parseInt(s.slideId, 10));
    const maxId = Math.max(...existingIds);
    const newSlideId = String(maxId + 1);

    const slideInfo = {
      index: newIndex,
      zipPath: slideZipPath,
      relationshipId: rId,
      slideId: newSlideId,
      tags: [...sourceInfo.tags],
      title: sourceInfo.title,
    };

    this.#slides.set(newIndex, slideInfo);
    this.#addSlideToPresentation(rId, newSlideId);
    this.#registerSlideContentType(slideFileName);

    logger.debug(`Cloned slide ${sourceIndex} to new slide ${newIndex}`);
  }

  /**
   * Removes a slide from the presentation.
   *
   * @param {number} slideIndex - 1-based index.
   */
  removeSlide(slideIndex) {
    this.#assertSlideExists(slideIndex);
    const info = this.#slides.get(slideIndex);

    // Remove from ZIP
    this.#zipManager.removeFile(info.zipPath);

    // Remove its relationships file
    const relsFileName = info.zipPath.split('/').pop() + '.rels';
    this.#zipManager.removeFile(`ppt/slides/_rels/${relsFileName}`);

    // Remove from cache
    this.#slideXmlCache.delete(info.zipPath);

    // Remove relationship from presentation.xml
    this.#relationshipManager.removeRelationship('ppt/presentation.xml', info.relationshipId);

    // Remove content type from [Content_Types].xml
    this.#contentTypesManager.removeOverride(info.zipPath);

    // Remove from slides map and reindex
    this.#slides.delete(slideIndex);
    this.#reindexSlides();

    // Update presentation.xml
    this.#removeSlideFromPresentation(info.slideId);

    logger.debug(`Removed slide ${slideIndex}`);
  }

  /**
   * Reorders slides to match the given order array.
   *
   * @param {number[]} order - Array of 1-based slide numbers in desired order.
   */
  reorderSlides(order) {
    const current = this.getAllSlideIndices();
    if (order.length !== current.length) {
      throw new PPTXError(`reorderSlides: order array length (${order.length}) must match slide count (${current.length})`);
    }

    const slidesCopy = new Map(this.#slides);
    this.#slides.clear();

    order.forEach((oldIndex, newPos) => {
      const info = slidesCopy.get(oldIndex);
      if (!info) throw new SlideNotFoundError(`Slide ${oldIndex} not found`);
      info.index = newPos + 1;
      this.#slides.set(newPos + 1, info);
    });

    // Rebuild presentation sldIdLst
    this.rebuildPresentationSlideOrder();
    logger.debug(`Reordered slides: [${order.join(', ')}]`);
  }

  /**
   * Resolves a slide reference (index, slideId string, or tag string) to SlideInfo.
   *
   * @param {number|string} slideRef
   * @returns {SlideInfo|null}
   */
  resolveSlideInfo(slideRef) {
    let index;
    if (typeof slideRef === 'number') {
      index = slideRef;
    } else {
      // 1. Try finding by slideId string
      for (const info of this.#slides.values()) {
        if (info.slideId === String(slideRef)) {
          return info;
        }
      }
      // 2. Try finding by tag
      try {
        const indices = this.resolveSlideRef(slideRef);
        if (indices && indices.length > 0) {
          index = indices[0];
        }
      } catch (e) {
        // Fallback: parse as slide index
        const parsedNum = parseInt(slideRef, 10);
        if (!isNaN(parsedNum)) {
          index = parsedNum;
        }
      }
    }

    if (index !== undefined) {
      return this.#slides.get(index) || null;
    }
    return null;
  }

  /**
   * Imports a slide from another PPTX template (PPTXTemplater instance).
   * Preserves all relationships: layouts, media, charts, workbooks, etc.
   *
   * @param {PPTXTemplater} sourceEngine - Source presentation.
   * @param {number|string} slideRef - Slide index (1-based), slide ID, or custom tag.
   * @param {MediaManager} mediaManager - Destination media manager.
   * @returns {Promise<number>} Index of the imported slide.
   */
  async importSlide(sourceEngine, slideRef, mediaManager) {
    const sourceSlideManager = sourceEngine.slideManager;
    const sourceRelManager = sourceEngine.relationshipManager;
    const sourceZip = sourceEngine.zipManager;

    const sourceSlideInfo = sourceSlideManager.resolveSlideInfo(slideRef);
    if (!sourceSlideInfo) {
      throw new SlideNotFoundError(`Source slide "${slideRef}" not found`);
    }

    const newIndex = this.#slides.size + 1;
    let nextFileIndex = 1;
    while (this.#zipManager.hasFile(`ppt/slides/slide${nextFileIndex}.xml`)) {
      nextFileIndex++;
    }
    const slideFileName = `slide${nextFileIndex}.xml`;
    const slideZipPath = `ppt/slides/${slideFileName}`;

    // Read the source slide's XML
    let slideXml = await sourceSlideManager.getSlideXmlAsync(sourceSlideInfo.index);

    // Get relationships from the source slide
    const sourceRels = sourceRelManager.getRelationships(sourceSlideInfo.zipPath);

    // Map to track old rId -> new rId in the destination slide's .rels file
    const idMap = new Map();

    const EXT_TO_MIME_LOCAL = {
      png: 'image/png',
      jpg: 'image/jpeg',
      jpeg: 'image/jpeg',
      gif: 'image/gif',
      webp: 'image/webp',
      bmp: 'image/bmp',
      xml: 'application/xml',
      rels: 'application/vnd.openxmlformats-package.relationships+xml',
    };

    for (const rel of sourceRels) {
      const resolvedTarget = sourceRelManager.resolveTarget(sourceSlideInfo.zipPath, rel.target);

      if (rel.type === REL_TYPES.SLIDE_LAYOUT) {
        // Map to destination's slide layout.
        const layoutFileName = rel.target.split('/').pop();
        const destLayoutPath = `ppt/slideLayouts/${layoutFileName}`;

        let targetLayout = `../slideLayouts/${layoutFileName}`;
        if (!this.#zipManager.hasFile(destLayoutPath)) {
          // Find first available layout
          const layoutFiles = this.#zipManager.listFiles('ppt/slideLayouts/').filter(f => f.endsWith('.xml'));
          if (layoutFiles.length > 0) {
            targetLayout = `../slideLayouts/${layoutFiles[0].split('/').pop()}`;
          } else {
            targetLayout = '../slideLayouts/slideLayout1.xml';
          }
        }

        const newRId = this.#relationshipManager.addRelationship(slideZipPath, rel.type, targetLayout);
        idMap.set(rel.id, newRId);

      } else if (rel.type === REL_TYPES.IMAGE) {
        // Copy media file
        const mediaBytes = await sourceZip.readBinaryFile(resolvedTarget);
        if (mediaBytes) {
          const destMediaZipPath = await mediaManager.embedImage(mediaBytes);
          const relativeMediaTarget = `../media/${destMediaZipPath.split('/').pop()}`;
          const newRId = this.#relationshipManager.addRelationship(slideZipPath, rel.type, relativeMediaTarget);
          idMap.set(rel.id, newRId);
        }

      } else if (rel.type === REL_TYPES.CHART) {
        // Copy chart XML and its relationships
        const chartXml = await sourceZip.readFile(resolvedTarget);
        if (chartXml) {
          const chartRels = sourceRelManager.getRelationships(resolvedTarget);

          let nextChartId = 1;
          while (this.#zipManager.hasFile(`ppt/charts/chart${nextChartId}.xml`)) {
            nextChartId++;
          }
          const destChartZipPath = `ppt/charts/chart${nextChartId}.xml`;
          const chartFileName = `chart${nextChartId}.xml`;

          // Handle workbook packages within charts
          for (const chartRel of chartRels) {
            const resolvedChartTarget = sourceRelManager.resolveTarget(resolvedTarget, chartRel.target);
            const workbookBytes = await sourceZip.readBinaryFile(resolvedChartTarget);

            if (workbookBytes) {
              const workbookFileName = resolvedChartTarget.split('/').pop();
              let nextEmbedId = 1;
              let destWorkbookZipPath = `ppt/embeddings/Microsoft_Excel_Worksheet${nextEmbedId}.xlsx`;
              if (workbookFileName.endsWith('.bin')) {
                destWorkbookZipPath = `ppt/embeddings/oleObject${nextEmbedId}.bin`;
              }
              while (this.#zipManager.hasFile(destWorkbookZipPath)) {
                nextEmbedId++;
                destWorkbookZipPath = workbookFileName.endsWith('.bin')
                  ? `ppt/embeddings/oleObject${nextEmbedId}.bin`
                  : `ppt/embeddings/Microsoft_Excel_Worksheet${nextEmbedId}.xlsx`;
              }

              this.#zipManager.writeBinaryFile(destWorkbookZipPath, workbookBytes);

              const workbookContentType = workbookFileName.endsWith('.bin')
                ? 'application/vnd.ms-office.activeX'
                : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
              this.#contentTypesManager.addOverride(destWorkbookZipPath, workbookContentType);

              const relativeWorkbookPath = `../embeddings/${destWorkbookZipPath.split('/').pop()}`;
              this.#relationshipManager.addRelationship(destChartZipPath, chartRel.type, relativeWorkbookPath);
            }
          }

          this.#zipManager.writeFile(destChartZipPath, chartXml);
          this.#contentTypesManager.addOverride(destChartZipPath, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');

          const relativeChartPath = `../charts/${chartFileName}`;
          const newRId = this.#relationshipManager.addRelationship(slideZipPath, rel.type, relativeChartPath);
          idMap.set(rel.id, newRId);
        }

      } else if (rel.type === REL_TYPES.HYPERLINK) {
        const newRId = this.#relationshipManager.addRelationship(
          slideZipPath,
          rel.type,
          rel.target,
          rel.targetMode
        );
        idMap.set(rel.id, newRId);

      } else {
        // Fallback for notes, themes, styles or custom XML
        if (rel.target && !rel.target.startsWith('http')) {
          const targetBytes = await sourceZip.readBinaryFile(resolvedTarget);
          if (targetBytes && !this.#zipManager.hasFile(resolvedTarget)) {
            this.#zipManager.writeBinaryFile(resolvedTarget, targetBytes);
            const ext = resolvedTarget.split('.').pop().toLowerCase();
            const mime = EXT_TO_MIME_LOCAL[ext] || 'application/octet-stream';
            this.#contentTypesManager.addDefault(ext, mime);
          }
        }
        const newRId = this.#relationshipManager.addRelationship(slideZipPath, rel.type, rel.target, rel.targetMode);
        idMap.set(rel.id, newRId);
      }
    }

    // Remap all relationship IDs inside the imported slide XML
    slideXml = remapRelationshipIds(slideXml, idMap);

    // Save the remapped slide XML to ZIP
    this.#zipManager.writeFile(slideZipPath, slideXml);
    this.#slideXmlCache.set(slideZipPath, slideXml);

    // Generate unique Slide ID
    const existingIds = Array.from(this.#slides.values()).map(s => parseInt(s.slideId, 10));
    const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 255;
    const newSlideId = String(maxId + 1);

    // Add relationship from presentation.xml
    const rId = this.#relationshipManager.addRelationship(
      'ppt/presentation.xml',
      REL_TYPES.SLIDE,
      `slides/${slideFileName}`
    );

    const slideInfo = {
      index: newIndex,
      zipPath: slideZipPath,
      relationshipId: rId,
      slideId: newSlideId,
      tags: [...sourceSlideInfo.tags],
      title: sourceSlideInfo.title || '',
    };

    this.#slides.set(newIndex, slideInfo);

    // Add entry in presentation.xml sldIdLst
    this.#addSlideToPresentation(rId, newSlideId);

    // Register slide in content types
    this.#registerSlideContentType(slideFileName);

    logger.debug(`Successfully imported slide "${slideRef}" to index ${newIndex}`);
    return newIndex;
  }

  /**
   * Exports a subset of slides to a new PPTXTemplater.
   *
   * @param {number[]} slideIndices - 1-based slide indices to export.
   * @param {PPTXTemplater} sourceEngine - Current PPTXTemplater instance.
   * @returns {Promise<PPTXTemplater>}
   */
  async exportSlides(slideIndices, sourceEngine) {
    // Lazy import to avoid circular dep
    const { PPTXTemplater } = await import('../core/PPTXTemplater.js');

    // Create a blank new PPTX
    const newEngine = await PPTXTemplater.create();

    // Remove the default slides from the blank template to avoid orphans
    const defaultSlides = newEngine.slideManager.getAllSlideIndices();
    for (const dIdx of defaultSlides.reverse()) {
      newEngine.slideManager.removeSlide(dIdx);
    }

    // Copy selected slides into the new engine
    for (const idx of slideIndices) {
      this.#assertSlideExists(idx);
      await newEngine.slideManager.importSlide(sourceEngine, idx, newEngine.mediaManager);
    }

    return newEngine;
  }

  /**
   * Validates the slide structure.
   *
   * @param {RelationshipManager} relationshipManager
   * @param {ZipManager} zipManager
   * @returns {ValidationResult}
   */
  validateStructure(relationshipManager, zipManager) {
    const errors = [];
    const warnings = [];

    for (const [index, info] of this.#slides) {
      if (!zipManager.hasFile(info.zipPath)) {
        errors.push(`Slide ${index}: XML file missing at ${info.zipPath}`);
      }

      const rels = relationshipManager.getRelationships(info.zipPath);
      const layoutRel = rels.find(r => r.type === REL_TYPES.SLIDE_LAYOUT);
      if (!layoutRel) {
        warnings.push(`Slide ${index}: No slide layout relationship found`);
      }
    }

    return {
      valid: errors.length === 0,
      errors,
      warnings,
    };
  }

  /**
   * Pre-loads all slide XML into cache (for bulk operations).
   * @returns {Promise<void>}
   */
  async preloadAll() {
    await Promise.all(
      this.getAllSlideIndices().map(i => this.getSlideXmlAsync(i))
    );
  }

  /**
   * Updates the presentation.xml sldIdLst with a new slide entry.
   * @private
   */
  #addSlideToPresentation(rId, slideId) {
    if (!this.#presentationObj) return;

    let sldIdLst = this.#xmlParser.getNode(this.#presentationObj, 'p:presentation.p:sldIdLst');
    if (!sldIdLst) {
      this.#xmlParser.setNode(this.#presentationObj, 'p:presentation.p:sldIdLst', { 'p:sldId': [] });
      sldIdLst = this.#xmlParser.getNode(this.#presentationObj, 'p:presentation.p:sldIdLst');
    }

    if (!sldIdLst['p:sldId']) sldIdLst['p:sldId'] = [];
    if (!Array.isArray(sldIdLst['p:sldId'])) {
      sldIdLst['p:sldId'] = [sldIdLst['p:sldId']];
    }

    sldIdLst['p:sldId'].push({ '@_id': slideId, '@_r:id': rId });
    this.#flushPresentation();
  }

  /**
   * Removes a slide from presentation.xml sldIdLst.
   * @private
   */
  #removeSlideFromPresentation(slideId) {
    if (!this.#presentationObj) return;
    const sldIdLst = this.#xmlParser.getNode(this.#presentationObj, 'p:presentation.p:sldIdLst');
    if (!sldIdLst?.['p:sldId']) return;

    sldIdLst['p:sldId'] = (Array.isArray(sldIdLst['p:sldId'])
      ? sldIdLst['p:sldId']
      : [sldIdLst['p:sldId']]
    ).filter(s => s['@_id'] !== slideId);

    // Also remove from any PowerPoint sections
    this.#removeSlideFromSections(slideId);

    this.#flushPresentation();
  }

  /**
   * Removes a slide ID from all sections in presentation.xml.
   * @private
   * @param {string} slideId - Unique slide ID.
   */
  #removeSlideFromSections(slideId) {
    if (!this.#presentationObj) return;
    const extLst = this.#xmlParser.getNode(this.#presentationObj, 'p:presentation.p:extLst');
    if (!extLst?.['p:ext']) return;

    const exts = Array.isArray(extLst['p:ext']) ? extLst['p:ext'] : [extLst['p:ext']];
    for (const ext of exts) {
      const sectionLst = ext['p14:sectionLst'];
      if (!sectionLst?.['p14:section']) continue;

      const sections = sectionLst['p14:section']; // Guaranteed to be array by XMLParser config

      for (const section of sections) {
        const sldIdLst = section['p14:sldIdLst'];
        if (!sldIdLst?.['p14:sldId']) continue;

        const sldIds = sldIdLst['p14:sldId']; // Guaranteed to be array by XMLParser config
        const targetIdStr = String(slideId);
        const filtered = sldIds.filter(s => String(s['@_id']) !== targetIdStr);

        if (filtered.length !== sldIds.length) {
          logger.debug(`Removing slide ${targetIdStr} from section "${section['@_name']}"`);
          section['p14:sldIdLst']['p14:sldId'] = filtered;
        }
      }
    }
  }

  /**
   * Rebuilds presentation.xml sldIdLst in the current slide order.
   */
  rebuildPresentationSlideOrder() {
    if (!this.#presentationObj) return;
    const sldIdLst = this.#xmlParser.getNode(this.#presentationObj, 'p:presentation.p:sldIdLst');
    if (!sldIdLst) return;

    const ordered = this.getAllSlideIndices().map(i => {
      const info = this.#slides.get(i);
      return { '@_id': info.slideId, '@_r:id': info.relationshipId };
    });

    sldIdLst['p:sldId'] = ordered;
    this.#flushPresentation();
  }

  /**
   * Re-indexes slide map after a removal.
   * @private
   */
  #reindexSlides() {
    const sorted = Array.from(this.#slides.entries()).sort(([a], [b]) => a - b);
    this.#slides.clear();
    sorted.forEach(([, info], i) => {
      info.index = i + 1;
      this.#slides.set(i + 1, info);
    });
  }

  /**
   * Registers a new slide in [Content_Types].xml.
   * @private
   */
  #registerSlideContentType(slideFileName) {
    this.#contentTypesManager.addOverride(`ppt/slides/${slideFileName}`, SLIDE_CONTENT_TYPE);
  }

  /**
   * Writes the updated presentation.xml back to the ZIP.
   * @private
   */
  #flushPresentation() {
    if (!this.#presentationObj || !this.#zipManager) return;
    const declaration = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
    const xml = this.#xmlParser.build(this.#presentationObj, declaration);
    this.#zipManager.writeFile('ppt/presentation.xml', xml);
  }

  /**
   * Initializes a blank presentation.xml structure.
   * @private
   */
  async #initializeBlankPresentation() {
    // Used when creating from scratch
    this.#presentationObj = {
      'p:presentation': {
        '@_xmlns:a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
        '@_xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
        '@_xmlns:p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
        'p:sldMasterIdLst': {},
        'p:sldIdLst': {},
        'p:sldSz': { '@_cx': '9144000', '@_cy': '5143500' },
        'p:notesSz': { '@_cx': '6858000', '@_cy': '9144000' },
      },
    };
    this.#flushPresentation();
  }

  /**
   * Asserts a slide index is valid.
   * @private
   * @param {number} index - 1-based slide index.
   */
  #assertSlideExists(index) {
    if (!this.#slides.has(index)) {
      throw new SlideNotFoundError(`Slide ${index} does not exist. Total slides: ${this.#slides.size}`);
    }
  }
}
