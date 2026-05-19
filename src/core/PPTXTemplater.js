/**
 * @fileoverview PPTXTemplater - The main orchestrator class.
 *
 * This is the primary public API. It coordinates all sub-managers
 * (ZipManager, SlideManager, ChartManager, etc.) and exposes a
 * fluent, chainable interface for template manipulation.
 *
 * OpenXML PPTX Structure:
 * ├── [Content_Types].xml      — lists all parts and their MIME types
 * ├── _rels/.rels              — root relationships (points to presentation)
 * ├── ppt/
 * │   ├── presentation.xml     — slide order, slide masters references
 * │   ├── _rels/presentation.xml.rels
 * │   ├── slides/
 * │   │   ├── slide1.xml       — individual slide content
 * │   │   └── _rels/slide1.xml.rels
 * │   ├── slideLayouts/        — layout templates (title, content, etc.)
 * │   ├── slideMasters/        — master slide designs
 * │   ├── theme/               — color/font themes
 * │   ├── charts/              — embedded chart XML
 * │   └── media/               — embedded images/videos
 * └── docProps/
 *     ├── core.xml             — author, title, etc.
 *     └── app.xml              — application metadata
 */

const { ZipManager } = require('../managers/ZipManager.js');
const { XMLParser } = require('../parsers/XMLParser.js');
const { ContentTypesManager } = require('../managers/ContentTypesManager.js');
const { SlideManager } = require('../managers/SlideManager.js');
const { ChartManager } = require('../managers/ChartManager.js');
const { TableManager } = require('../managers/TableManager.js');
const { HyperlinkManager } = require('../managers/HyperlinkManager.js');
const { MediaManager } = require('../managers/MediaManager.js');
const { RelationshipManager } = require('../managers/RelationshipManager.js');
const { OutputWriter } = require('./OutputWriter.js');
const { TemplateEngine } = require('./TemplateEngine.js');
const { createLogger } = require('../utils/logger.js');
const { PPTXError } = require('../utils/errors.js');

const logger = createLogger('PPTXTemplater');

/**
 * @class PPTXTemplater
 * @description Main engine class for PPTX template manipulation.
 *
 * @example
 * const ppt = await PPTXTemplater.load('template.pptx');
 * ppt.useSlide(1);
 * ppt.replaceText({ '{{title}}': 'My Report' });
 * await ppt.saveToFile('./output/report.pptx');
 */
class PPTXTemplater {
  /**
   * @private
   * @type {ZipManager}
   */
  #zipManager;

  /**
   * @private
   * @type {XMLParser}
   */
  #xmlParser;

  /**
   * @private
   * @type {ContentTypesManager}
   */
  #contentTypesManager;

  /**
   * @private
   * @type {SlideManager}
   */
  #slideManager;

  /**
   * @private
   * @type {ChartManager}
   */
  #chartManager;

  /**
   * @private
   * @type {TableManager}
   */
  #tableManager;

  /**
   * @private
   * @type {HyperlinkManager}
   */
  #hyperlinkManager;

  /**
   * @private
   * @type {MediaManager}
   */
  #mediaManager;

  /**
   * @private
   * @type {RelationshipManager}
   */
  #relationshipManager;

  /**
   * @private
   * @type {OutputWriter}
   */
  #outputWriter;

  /**
   * @private
   * @type {TemplateEngine}
   */
  #templateEngine;

  /**
   * @private
   * @type {number[]} - Currently selected slide indices (1-based)
   */
  #selectedSlides = [];

  /**
   * @private
   * @type {boolean}
   */
  #loaded = false;

  constructor() {
    this.#xmlParser = new XMLParser();
    this.#zipManager = new ZipManager();
    this.#contentTypesManager = new ContentTypesManager(this.#xmlParser);
    this.#relationshipManager = new RelationshipManager(this.#xmlParser);
    this.#slideManager = new SlideManager(this.#xmlParser, this.#relationshipManager, this.#contentTypesManager);
    this.#chartManager = new ChartManager(this.#xmlParser, this.#contentTypesManager);
    this.#tableManager = new TableManager(this.#xmlParser);
    this.#hyperlinkManager = new HyperlinkManager(this.#xmlParser, this.#relationshipManager);
    this.#mediaManager = new MediaManager(this.#contentTypesManager);
    this.#templateEngine = new TemplateEngine(this.#xmlParser);
    this.#outputWriter = new OutputWriter(this.#zipManager, this.#contentTypesManager);
  }

  /**
   * Loads a PPTX template from a file path or buffer.
   *
   * @static
   * @param {string|Buffer} source - Path to PPTX file or Buffer containing PPTX data.
   * @returns {Promise<PPTXTemplater>} Initialized engine instance.
   * @throws {PPTXError} If the file cannot be read or is not a valid PPTX.
   *
   * @example
   * // From file path
   * const ppt = await PPTXTemplater.load('./template.pptx');
   *
   * // From buffer
   * const buffer = fs.readFileSync('./template.pptx');
   * const ppt = await PPTXTemplater.load(buffer);
   */
  static async load(source) {
    const engine = new PPTXTemplater();
    await engine.#initialize(source);
    return engine;
  }

  /**
   * Creates a new blank PPTX from scratch.
   *
   * @static
   * @returns {Promise<PPTXTemplater>} Engine instance with a blank PPTX.
   *
   * @example
   * const ppt = await PPTXTemplater.create();
   * ppt.addSlide({ title: 'First Slide' });
   * await ppt.saveToFile('./new.pptx');
   */
  static async create() {
    const engine = new PPTXTemplater();
    await engine.#initializeBlank();
    return engine;
  }

  /**
   * Initializes the engine by loading a PPTX file/buffer.
   * @private
   * @param {string|Buffer} source
   */
  async #initialize(source) {
    logger.debug(`Loading PPTX from ${typeof source === 'string' ? source : 'buffer'}`);

    // Load and extract the ZIP archive (PPTX is just a ZIP)
    await this.#zipManager.load(source);

    // Initialize content types manager first!
    await this.#contentTypesManager.initialize(this.#zipManager);

    // Parse the core presentation relationships and structure
    await this.#relationshipManager.initialize(this.#zipManager);

    // Load all slide references from presentation.xml
    await this.#slideManager.initialize(this.#zipManager);

    // Pre-load all slide XML into cache to allow synchronous operations like replaceText()
    await this.#slideManager.preloadAll();

    // Initialize chart manager with zip context
    await this.#chartManager.initialize(this.#zipManager);

    // Deduplicate and index media files
    await this.#mediaManager.initialize(this.#zipManager);

    this.#loaded = true;
    logger.debug(`Loaded ${this.#slideManager.slideCount} slides successfully`);
  }

  /**
   * Initializes a blank PPTX structure from embedded template XML.
   * @private
   */
  async #initializeBlank() {
    await this.#zipManager.createBlank();
    await this.#contentTypesManager.initialize(this.#zipManager);
    await this.#relationshipManager.initialize(this.#zipManager);
    await this.#slideManager.initialize(this.#zipManager);
    await this.#chartManager.initialize(this.#zipManager);
    await this.#mediaManager.initialize(this.#zipManager);
    this.#loaded = true;
  }

  /**
   * Asserts the engine is loaded before performing operations.
   * @private
   */
  #assertLoaded() {
    if (!this.#loaded) {
      throw new PPTXError('Engine not initialized. Call PPTXTemplater.load() first.');
    }
  }

  /**
   * Selects one or more slides to work on.
   * All subsequent operations (replaceText, updateChart, etc.) apply to these slides.
   * If not called, operations apply to ALL slides.
   *
   * @param {...number|string} slideRefs - Slide numbers (1-based), IDs, or tags.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.useSlide(1);           // Select slide 1
   * ppt.useSlide(1, 3, 5);    // Select slides 1, 3, and 5
   * ppt.useSlide('intro');     // Select by custom tag
   */
  useSlide(...slideRefs) {
    this.#assertLoaded();
    this.#selectedSlides = slideRefs;
    logger.debug(`Selected slides: ${slideRefs.join(', ')}`);
    return this;
  }

  /**
   * Selects all slides.
   * @returns {PPTXTemplater} this (chainable)
   */
  useAllSlides() {
    this.#assertLoaded();
    this.#selectedSlides = [];
    return this;
  }

  /**
   * Returns the resolved slide indices based on #selectedSlides.
   * If nothing is selected, returns all slide indices.
   * @private
   * @returns {number[]} Array of 1-based slide indices.
   */
  #getTargetSlideIndices() {
    if (this.#selectedSlides.length === 0) {
      return this.#slideManager.getAllSlideIndices();
    }
    return this.#selectedSlides.flatMap(ref => {
      if (typeof ref === 'number') return [ref];
      // Resolve by tag or ID
      return this.#slideManager.resolveSlideRef(ref);
    });
  }

  /**
   * Replaces template placeholders (e.g., {{key}}) with values in the selected slides.
   * Works inside text boxes, titles, grouped shapes, tables, and shapes.
   *
   * @param {Object.<string, string>} replacements - Map of placeholder → replacement value.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.replaceText({
   *   '{{title}}': 'Quarterly Report',
   *   '{{year}}': '2026',
   *   '{{company}}': 'Acme Corp'
   * });
   */
  replaceText(replacements) {
    this.#assertLoaded();
    const targetIndices = this.#getTargetSlideIndices();

    for (const slideIndex of targetIndices) {
      const slideXml = this.#slideManager.getSlideXml(slideIndex);
      const updated = this.#templateEngine.replaceTextInXml(slideXml, replacements);
      this.#slideManager.setSlideXml(slideIndex, updated);
    }

    logger.debug(`Replaced ${Object.keys(replacements).length} placeholder(s) in ${targetIndices.length} slide(s)`);
    return this;
  }

  /**
   * Updates chart data in the selected slide(s).
   * Finds charts by their name/ID and updates categories, series, and values.
   * Preserves original chart styles, themes, and formatting.
   *
   * @param {string} chartId - Chart name or relationship ID.
   * @param {ChartData} data - New chart data.
   * @param {string[]} data.categories - Category labels (X-axis).
   * @param {SeriesData[]} data.series - Data series array.
   * @param {string} data.series[].name - Series name.
   * @param {number[]} data.series[].values - Data values.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.updateChart('sales-chart', {
   *   categories: ['Jan', 'Feb', 'Mar'],
   *   series: [{ name: 'Revenue', values: [120, 150, 180] }]
   * });
   */
  updateChart(chartId, data) {
    this.#assertLoaded();
    const targetIndices = this.#getTargetSlideIndices();

    for (const slideIndex of targetIndices) {
      this.#chartManager.updateChart(
        slideIndex,
        chartId,
        data,
        this.#slideManager,
        this.#relationshipManager
      );
    }

    logger.debug(`Updated chart "${chartId}" in ${targetIndices.length} slide(s)`);
    return this;
  }

  /**
   * Replaces table rows with new data in the selected slide(s).
   * Preserves borders, merged cells, fonts, colors, and alignment from the template.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {string[][]} rows - 2D array of cell values (row × col).
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.updateTable('employees-table', [
   *   ['Name', 'Role', 'Department'],
   *   ['John', 'Engineer', 'Platform'],
   *   ['Jane', 'Designer', 'Product']
   * ]);
   */
  updateTable(tableId, rows) {
    this.#assertLoaded();
    const targetIndices = this.#getTargetSlideIndices();

    for (const slideIndex of targetIndices) {
      this.#tableManager.updateTable(slideIndex, tableId, rows, this.#slideManager);
    }

    logger.debug(`Updated table "${tableId}" in ${targetIndices.length} slide(s)`);
    return this;
  }

  /**
   * Adds or replaces a hyperlink on a text run or shape.
   *
   * @param {HyperlinkOptions} options - Hyperlink configuration.
   * @param {string} options.text - Text to find and make clickable.
   * @param {string} options.url - Target URL.
   * @param {string} [options.tooltip] - Optional tooltip.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.addHyperlink({ text: 'Open Website', url: 'https://example.com' });
   */
  addHyperlink(options) {
    this.#assertLoaded();
    const targetIndices = this.#getTargetSlideIndices();

    for (const slideIndex of targetIndices) {
      this.#hyperlinkManager.addExternalHyperlink(
        slideIndex,
        options,
        this.#slideManager,
        this.#relationshipManager
      );
    }

    return this;
  }

  /**
   * Adds an inter-slide hyperlink to a specific text element.
   *
   * @param {Object} options - Link configuration.
   * @param {number} options.sourceSlide - Source slide number (1-based).
   * @param {number} options.targetSlide - Destination slide number (1-based).
   * @param {string} options.element - Text element to make clickable.
   * @returns {PPTXTemplater} this (chainable)
   */
  addSlideLink(options) {
    this.#assertLoaded();
    const { sourceSlide, targetSlide, element } = options;

    // Fallback: If no element text is provided, link the slide number (legacy behavior)
    if (!element) {
      this.#hyperlinkManager.addSlideHyperlink(
        sourceSlide,
        targetSlide,
        this.#slideManager,
        this.#relationshipManager
      );
    } else {
      // Add a slide hyperlink on specific text
      this.#hyperlinkManager.addTextSlideLink(
        sourceSlide,
        element,
        targetSlide,
        this.#slideManager,
        this.#relationshipManager
      );
    }
    return this;
  }

  /**
   * Adds an inter-slide hyperlink to an image.
   *
   * @param {Object} options
   * @param {number} options.slide - Source slide number.
   * @param {string} options.imageId - Image name/id to make clickable.
   * @param {number} options.targetSlide - Destination slide number.
   * @returns {PPTXTemplater} this
   */
  addImageLink(options) {
    this.#assertLoaded();
    this.#hyperlinkManager.addShapeSlideLink(
      options.slide,
      options.imageId,
      options.targetSlide,
      this.#slideManager,
      this.#relationshipManager
    );
    return this;
  }

  /**
   * Adds an inter-slide hyperlink to a shape.
   *
   * @param {Object} options
   * @param {number} options.slide - Source slide number.
   * @param {string} options.shapeId - Shape name/id to make clickable.
   * @param {number} options.targetSlide - Destination slide number.
   * @returns {PPTXTemplater} this
   */
  addShapeLink(options) {
    this.#assertLoaded();
    this.#hyperlinkManager.addShapeSlideLink(
      options.slide,
      options.shapeId,
      options.targetSlide,
      this.#slideManager,
      this.#relationshipManager
    );
    return this;
  }

  /**
   * Adds a special navigation link (next, previous, first, last slide) to a text element.
   *
   * @param {Object} options
   * @param {number} options.slide - Source slide number (1-based).
   * @param {string} options.element - Text element to make clickable.
   * @param {'next'|'previous'|'first'|'last'} options.action - Navigation action type.
   * @returns {PPTXTemplater} this (chainable)
   */
  addTextNavigationLink(options) {
    this.#assertLoaded();
    const { slide, element, action } = options;
    this.#hyperlinkManager.addTextNavigationLink(slide, element, action, this.#slideManager);
    return this;
  }

  /**
   * Adds a special navigation link (next, previous, first, last slide) to a shape or image.
   *
   * @param {Object} options
   * @param {number} options.slide - Source slide number (1-based).
   * @param {string} options.shapeId - Shape name/id to make clickable.
   * @param {'next'|'previous'|'first'|'last'} options.action - Navigation action type.
   * @returns {PPTXTemplater} this (chainable)
   */
  addShapeNavigationLink(options) {
    this.#assertLoaded();
    const { slide, shapeId, action } = options;
    this.#hyperlinkManager.addShapeNavigationLink(slide, shapeId, action, this.#slideManager);
    return this;
  }

  /**
   * Adds a new slide to the presentation.
   * Automatically generates required XML and relationship entries.
   *
   * @param {NewSlideOptions} options - Slide definition.
   * @param {string} [options.title] - Slide title text.
   * @param {string} [options.layout] - Layout name to use (default: 'blank').
   * @param {SlideElement[]} [options.elements] - Elements to add to the slide.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.addSlide({
   *   title: 'New Slide',
   *   elements: [
   *     { type: 'text', value: 'Hello World', x: 100, y: 200 },
   *     { type: 'image', src: './logo.png', x: 500, y: 100, width: 200, height: 150 }
   *   ]
   * });
   */
  addSlide(options = {}) {
    this.#assertLoaded();
    this.#slideManager.addNewSlide(options, this.#relationshipManager, this.#mediaManager);
    logger.debug(`Added new slide: "${options.title || 'Untitled'}"`);
    return this;
  }

  /**
   * Clones an existing slide and appends it to the end (or at a position).
   *
   * @param {number} sourceSlideNumber - 1-based source slide number.
   * @param {number} [atPosition] - Optional position to insert (1-based). Default: append.
   * @returns {PPTXTemplater} this (chainable)
   */
  cloneSlide(sourceSlideNumber, atPosition) {
    this.#assertLoaded();
    this.#slideManager.cloneSlide(sourceSlideNumber, atPosition, this.#relationshipManager);
    return this;
  }

  /**
   * Removes a slide from the presentation.
   *
   * @param {number} slideNumber - 1-based slide number to remove.
   * @returns {PPTXTemplater} this (chainable)
   */
  removeSlide(slideNumber) {
    this.#assertLoaded();
    this.#slideManager.removeSlide(slideNumber);
    return this;
  }

  /**
   * Reorders slides in the presentation.
   *
   * @param {number[]} order - Array of 1-based slide numbers in desired order.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.reorderSlides([3, 1, 2]); // Move slide 3 to position 1
   */
  reorderSlides(order) {
    this.#assertLoaded();
    this.#slideManager.reorderSlides(order);
    return this;
  }

  /**
   * Tags a slide with a custom string identifier for later selection.
   *
   * @param {number} slideNumber - 1-based slide number.
   * @param {string} tag - Custom tag string.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.tagSlide(1, 'intro');
   * ppt.useSlide('intro').replaceText({ '{{title}}': 'Hello' });
   */
  tagSlide(slideNumber, tag) {
    this.#assertLoaded();
    this.#slideManager.tagSlide(slideNumber, tag);
    return this;
  }

  /**
   * Exports selected slides to a new standalone PPTX engine.
   * Useful for creating "slide decks" from a master template.
   *
   * @param {...number} slideNumbers - 1-based slide numbers to export.
   * @returns {Promise<PPTXTemplater>} New engine with only the selected slides.
   *
   * @example
   * const subset = await ppt.exportSlides(1, 3, 5);
   * await subset.saveToFile('./subset.pptx');
   */
  async exportSlides(...slideNumbers) {
    this.#assertLoaded();
    return this.#slideManager.exportSlides(slideNumbers, this);
  }

  /**
   * Imports a single slide from another PPTXTemplater instance into this presentation.
   * Preserves all slide layouts, charts, relationships, and embedded media.
   *
   * @param {PPTXTemplater} sourceEngine - Source PPTXTemplater instance.
   * @param {number|string} slideRef - Slide index (1-based), ID, or custom tag.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async importSlideFrom(sourceEngine, slideRef) {
    this.#assertLoaded();
    await this.#slideManager.importSlide(sourceEngine, slideRef, this.#mediaManager);
    return this;
  }

  /**
   * Imports selected slides from the current template, discarding the rest.
   * The remaining slides are reordered to match the provided array.
   * Preserves all layouts, themes, relationships, and embedded media.
   *
   * @param {number[]} slideIndices - Array of 1-based slide indices to keep.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.importSlides([1, 3, 5]);
   */
  importSlides(slideIndices) {
    this.#assertLoaded();
    const slidesToKeep = slideIndices.map(i => this.#slideManager.getSlideInfo(i).slideId);

    // Remove unneeded slides from highest to lowest index to avoid shifting issues
    const allIndices = this.#slideManager.getAllSlideIndices();
    for (let i = allIndices.length; i >= 1; i--) {
      const info = this.#slideManager.getSlideInfo(i);
      if (!slidesToKeep.includes(info.slideId)) {
        this.#slideManager.removeSlide(i);
      }
    }

    // Calculate new target order based on the requested slideIndices
    const currentOrder = this.#slideManager.getAllSlideIndices().map(i => this.#slideManager.getSlideInfo(i).slideId);

    const newOrder = slidesToKeep.map(id => {
      return currentOrder.indexOf(id) + 1;
    });

    // Only reorder if needed
    if (newOrder.join(',') !== currentOrder.map((_, i) => i + 1).join(',')) {
      this.#slideManager.reorderSlides(newOrder);
    }

    logger.debug(`Imported ${slideIndices.length} slide(s).`);
    return this;
  }

  /**
   * Returns presentation metadata (title, author, slide count, etc.)
   *
   * @returns {PresentationInfo} Metadata object.
   */
  getInfo() {
    this.#assertLoaded();
    return {
      slideCount: this.#slideManager.slideCount,
      title: this.#zipManager.getCoreProperty('dc:title') || '',
      author: this.#zipManager.getCoreProperty('dc:creator') || '',
      created: this.#zipManager.getCoreProperty('dcterms:created') || '',
      modified: this.#zipManager.getCoreProperty('dcterms:modified') || '',
      slides: this.#slideManager.getAllSlideInfo(),
      mediaCount: this.#mediaManager.mediaCount,
    };
  }

  /**
   * Validates the XML structure of the current PPTX.
   * Reports issues with relationship IDs, missing parts, etc.
   *
   * @returns {ValidationResult} Object with `valid`, `errors`, and `warnings` arrays.
   */
  validate() {
    this.#assertLoaded();
    return this.#slideManager.validateStructure(this.#relationshipManager, this.#zipManager);
  }

  /**
   * Repairs corrupted OpenXML structure, relationships, and content types.
   * Removes orphan relationships, rebuilds slide references, and fixes missing entries.
   *
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async repair() {
    this.#assertLoaded();

    // 1. Rebuild presentation.xml slide mappings
    this.#slideManager.rebuildPresentationSlideOrder();

    // 2. Remove orphan relationships
    this.#relationshipManager.removeOrphanRelationships(this.#zipManager);

    logger.info('PPTX repair complete.');
    return this;
  }

  /**
   * Logs all relationships across the presentation to the console for debugging.
   * @returns {PPTXTemplater} this (chainable)
   */
  debugRelationships() {
    this.#assertLoaded();
    const files = this.#zipManager.listFiles('').filter(f => f.endsWith('.rels'));
    console.log('=== Relationship Graph ===');
    for (const file of files) {
      console.log(`\n${file}:`);
      const rels = this.#relationshipManager.getRelationships(file.replace('_rels/', '').replace('.rels', ''));
      rels.forEach(r => console.log(`  - ${r.id} [${r.type.split('/').pop()}] -> ${r.target}`));
    }
    return this;
  }

  /**
   * Inspects a specific slide's structure and relationships.
   * @param {number} slideIndex - 1-based slide index.
   * @returns {PPTXTemplater} this (chainable)
   */
  inspectSlide(slideIndex) {
    this.#assertLoaded();
    const info = this.#slideManager.getSlideInfo(slideIndex);
    const xml = this.#slideManager.getSlideXml(slideIndex);
    const rels = this.#relationshipManager.getRelationships(info.zipPath);

    console.log(`=== Slide ${slideIndex} Inspection ===`);
    console.log(`Path: ${info.zipPath}`);
    console.log(`ID: ${info.slideId}`);
    console.log(`rId: ${info.relationshipId}`);
    console.log(`Title: ${info.title}`);
    console.log(`XML Size: ${xml.length} characters`);
    console.log(`Relationships (${rels.length}):`);
    rels.forEach(r => console.log(`  - ${r.id} [${r.type.split('/').pop()}] -> ${r.target}`));

    return this;
  }

  /**
   * Inspects and logs the raw XML of any file in the ZIP.
   * @param {string} xmlPath - Path inside the ZIP (e.g., 'ppt/slides/slide1.xml')
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async inspectXML(xmlPath) {
    this.#assertLoaded();
    const xml = await this.#zipManager.readFile(xmlPath);
    console.log(`=== XML Inspection: ${xmlPath} ===`);
    if (!xml) {
      console.log('(File not found or empty)');
    } else {
      console.log(xml.substring(0, 1500) + (xml.length > 1500 ? '...\n[Truncated]' : ''));
    }
    return this;
  }

  /**
   * Validates all charts in the presentation to ensure they are not corrupted.
   * Checks XML, caches, and embedded workbook references.
   *
   * @returns {Promise<Object>} Validation results for charts.
   */
  async validateCharts() {
    this.#assertLoaded();
    const issues = { valid: true, errors: [], warnings: [] };

    // We lazy require ChartRelationshipManager so we don't circularly depend if not needed
    const { ChartRelationshipManager } = require('../managers/charts/ChartRelationshipManager.js');

    const chartFiles = this.#zipManager.listFiles('ppt/charts/')
      .filter(f => {
        const name = f.split('/').pop();
        return name.startsWith('chart') && name.endsWith('.xml') && !f.includes('_rels');
      });

    for (const chartPath of chartFiles) {
      const relIssues = ChartRelationshipManager.validateChartRelationships(this.#relationshipManager, this.#zipManager, chartPath);
      issues.errors.push(...relIssues.errors);
      issues.warnings.push(...relIssues.warnings);
    }

    if (issues.errors.length > 0) issues.valid = false;
    return issues;
  }

  /**
   * Repairs common chart corruption issues such as broken caches,
   * missing embedded workbooks, or orphan nodes.
   *
   * @returns {Promise<PPTXTemplater>} this
   */
  async repairCharts() {
    this.#assertLoaded();
    logger.info('Repairing charts...');

    // Check all charts for missing embedded workbooks
    const chartFiles = this.#zipManager.listFiles('ppt/charts/')
      .filter(f => {
        const name = f.split('/').pop();
        return name.startsWith('chart') && name.endsWith('.xml') && !f.includes('_rels');
      });
    for (const chartPath of chartFiles) {
      const rels = this.#relationshipManager.getRelationshipsByType(chartPath, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package');
      for (const rel of rels) {
        const xlsxPath = this.#relationshipManager.resolveTarget(chartPath, rel.target);
        if (!this.#zipManager.hasFile(xlsxPath)) {
          logger.warn(`Chart ${chartPath} has broken workbook reference ${rel.id}, removing to prevent repair mode.`);
          this.#relationshipManager.removeRelationship(chartPath, rel.id);

          // Also strip c:externalData from chart XML to prevent PowerPoint looking for it
          const xml = await this.#zipManager.readFile(chartPath);
          if (xml) {
            const updated = xml.replace(/<c:externalData[^>]*r:id="[^"]*"[^>]*>/, '').replace(/<\/c:externalData>/, '');
            this.#zipManager.writeFile(chartPath, updated);
          }
        }
      }
    }

    return this;
  }

  /**
   * Inspects a specific chart's metadata and structure.
   *
   * @param {string} chartId
   */
  inspectChart(chartId) {
    this.#assertLoaded();
    console.log(`=== Chart Inspection: ${chartId} ===`);
    // Find chart across all slides to get info
    let found = false;
    for (const i of this.#slideManager.getAllSlideIndices()) {
      try {
        const info = this.#chartManager.getChartsInSlide(i, this.#slideManager, this.#relationshipManager);
        const chart = info.find(c => c.zipPath.toLowerCase().includes(chartId.toLowerCase()) || c.rId === chartId);
        if (chart) {
          console.log(`Found on Slide ${i}`);
          console.log(`ZIP Path: ${chart.zipPath}`);
          console.log(`Relationship ID: ${chart.rId}`);
          found = true;
          break;
        }
      } catch (e) {}
    }
    if (!found) console.log('Chart not found.');
    return this;
  }

  /**
   * Inspects and logs the raw XML of a chart file.
   *
   * @param {string} chartFileName
   */
  async inspectChartXML(chartFileName) {
    const fullPath = chartFileName.includes('/') ? chartFileName : `ppt/charts/${chartFileName}`;
    await this.inspectXML(fullPath);
    return this;
  }

  /**
   * Logs all chart relationships.
   */
  debugChartRelationships() {
    this.#assertLoaded();
    console.log('=== Chart Relationships ===');
    const chartFiles = this.#zipManager.listFiles('ppt/charts/')
      .filter(f => {
        const name = f.split('/').pop();
        return name.startsWith('chart') && name.endsWith('.xml') && !f.includes('_rels');
      });
    for (const chartPath of chartFiles) {
      console.log(`\n${chartPath}:`);
      const rels = this.#relationshipManager.getRelationships(chartPath);
      rels.forEach(r => console.log(`  - ${r.id} [${r.type.split('/').pop()}] -> ${r.target}`));
    }
    return this;
  }

  /**
   * Saves the modified PPTX to a file on disk.
   *
   * @param {string} filePath - Output file path (e.g., './output/report.pptx').
   * @returns {Promise<void>}
   *
   * @example
   * await ppt.saveToFile('./output/report.pptx');
   */
  async saveToFile(filePath) {
    this.#assertLoaded();
    await this.#outputWriter.saveToFile(filePath, this.#slideManager, this.#zipManager);
    logger.info(`Saved PPTX to ${filePath}`);
  }

  /**
   * Returns the PPTX content as a Node.js Buffer.
   * Useful for HTTP responses, email attachments, etc.
   *
   * @returns {Promise<Buffer>} Buffer containing PPTX binary data.
   *
   * @example
   * const buffer = await ppt.toBuffer();
   * res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
   * res.send(buffer);
   */
  async toBuffer() {
    this.#assertLoaded();
    return this.#outputWriter.toBuffer(this.#slideManager, this.#zipManager);
  }

  /**
   * Returns the PPTX content as a readable Node.js Stream.
   * Ideal for streaming large presentations to HTTP responses.
   *
   * @returns {Promise<NodeJS.ReadableStream>} Readable stream of PPTX data.
   *
   * @example
   * const stream = await ppt.toStream();
   * stream.pipe(res);
   */
  async toStream() {
    this.#assertLoaded();
    return this.#outputWriter.toStream(this.#slideManager, this.#zipManager);
  }

  /**
   * Returns the total number of slides in the loaded presentation.
   * @type {number}
   */
  get slideCount() {
    return this.#slideManager.slideCount;
  }

  // --- Public Getters for Internal Managers ---
  get zipManager() { return this.#zipManager; }
  get xmlParser() { return this.#xmlParser; }
  get contentTypesManager() { return this.#contentTypesManager; }
  get relationshipManager() { return this.#relationshipManager; }
  get slideManager() { return this.#slideManager; }
  get chartManager() { return this.#chartManager; }
  get tableManager() { return this.#tableManager; }
  get hyperlinkManager() { return this.#hyperlinkManager; }
  get mediaManager() { return this.#mediaManager; }
}

module.exports = { PPTXTemplater };
