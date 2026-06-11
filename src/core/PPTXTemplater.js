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

const { ZipManager } = require('../managers/ZipManager.js')
const { XMLParser } = require('../parsers/XMLParser.js')
const { ContentTypesManager } = require('../managers/ContentTypesManager.js')
const { SlideManager } = require('../managers/SlideManager.js')
const { ChartManager } = require('../managers/ChartManager.js')
const { TableManager } = require('../managers/TableManager.js')
const { HyperlinkManager } = require('../managers/HyperlinkManager.js')
const { MediaManager } = require('../managers/MediaManager.js')
const { RelationshipManager } = require('../managers/RelationshipManager.js')
const { ShapeManager } = require('../managers/ShapeManager.js')
const { ImageManager } = require('../managers/ImageManager.js')
const { TextManager } = require('../managers/TextManager.js')
const { ZOrderManager } = require('../managers/ZOrderManager.js')
const { ValidationEngine } = require('./ValidationEngine.js')
const { OutputWriter } = require('./OutputWriter.js')
const { TemplateEngine } = require('./TemplateEngine.js')
const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')
const { performance } = require('perf_hooks')

const logger = createLogger('PPTXTemplater')

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
  #zipManager

  /**
   * @private
   * @type {XMLParser}
   */
  #xmlParser

  /**
   * @private
   * @type {ContentTypesManager}
   */
  #contentTypesManager

  /**
   * @private
   * @type {SlideManager}
   */
  #slideManager

  /**
   * @private
   * @type {ChartManager}
   */
  #chartManager

  /**
   * @private
   * @type {TableManager}
   */
  #tableManager

  /**
   * @private
   * @type {HyperlinkManager}
   */
  #hyperlinkManager

  /**
   * @private
   * @type {MediaManager}
   */
  #mediaManager

  /**
   * @private
   * @type {ShapeManager}
   */
  #shapeManager

  /**
   * @private
   * @type {ImageManager}
   */
  #imageManager

  /**
   * @private
   * @type {TextManager}
   */
  #textManager

  /**
   * @private
   * @type {ZOrderManager}
   */
  #zOrderManager

  /**
   * @private
   * @type {RelationshipManager}
   */
  #relationshipManager

  /**
   * @private
   * @type {OutputWriter}
   */
  #outputWriter

  /**
   * @private
   * @type {TemplateEngine}
   */
  #templateEngine

  /**
   * @private
   * @type {number[]} - Currently selected slide indices (1-based)
   */
  #selectedSlides = []

  /**
   * @private
   * @type {boolean}
   */
  #loaded = false

  /**
   * @private
   * @type {Object}
   */
  #profiler

  static #templateCache = new Map()

  constructor() {
    this.#xmlParser = new XMLParser()
    this.#zipManager = new ZipManager()
    this.#contentTypesManager = new ContentTypesManager(this.#xmlParser)
    this.#relationshipManager = new RelationshipManager(this.#xmlParser)
    this.#slideManager = new SlideManager(
      this.#xmlParser,
      this.#relationshipManager,
      this.#contentTypesManager
    )
    this.#chartManager = new ChartManager(this.#xmlParser, this.#contentTypesManager)
    this.#tableManager = new TableManager(this.#xmlParser)
    this.#hyperlinkManager = new HyperlinkManager(this.#xmlParser, this.#relationshipManager)
    this.#mediaManager = new MediaManager(this.#contentTypesManager)
    this.#shapeManager = new ShapeManager(this.#xmlParser)
    this.#imageManager = new ImageManager(this.#xmlParser)
    this.#textManager = new TextManager(this.#xmlParser)
    this.#templateEngine = new TemplateEngine(this.#xmlParser)
    this.#zOrderManager = new ZOrderManager(this.#xmlParser)
    this.#outputWriter = new OutputWriter(this.#zipManager, this.#contentTypesManager)

    this.#profiler = {
      enabled: false,
      templateLoadMs: 0,
      parseMs: 0,
      chartUpdateMs: 0,
      imageUpdateMs: 0,
      zipGenerationMs: 0,
      totalMs: 0,
      memoryUsedMB: 0,
      startTime: performance.now(),
    }
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
    const engine = new PPTXTemplater()
    await engine.#initialize(source)
    return engine
  }

  static async preload(source) {
    let key = source
    if (Buffer.isBuffer(source)) {
      const crypto = require('crypto')
      key = crypto.createHash('sha256').update(source).digest('hex')
    } else if (typeof source === 'object' && source !== null) {
      key = JSON.stringify(source)
    }

    if (PPTXTemplater.#templateCache.has(key)) {
      return PPTXTemplater.#templateCache.get(key)
    }

    const zipManager = new ZipManager()
    await zipManager.load(source)
    const files = zipManager.listFiles()
    const cachedFiles = new Map()
    for (const file of files) {
      const ext = file.split('.').pop().toLowerCase()
      const isText = ext === 'xml' || ext === 'rels' || ext === 'txt'
      if (isText) {
        const content = await zipManager.readFile(file)
        cachedFiles.set(file, { type: 'text', content })
      } else {
        const content = await zipManager.readBinaryFile(file)
        cachedFiles.set(file, { type: 'binary', content })
      }
    }

    PPTXTemplater.#templateCache.set(key, cachedFiles)
    return cachedFiles
  }

  static async cache(source) {
    return PPTXTemplater.preload(source)
  }

  static async fromCache(source) {
    let key = source
    if (Buffer.isBuffer(source)) {
      const crypto = require('crypto')
      key = crypto.createHash('sha256').update(source).digest('hex')
    } else if (typeof source === 'object' && source !== null) {
      key = JSON.stringify(source)
    }

    let cachedFiles = PPTXTemplater.#templateCache.get(key)
    if (!cachedFiles) {
      cachedFiles = await PPTXTemplater.preload(source)
    }

    const engine = new PPTXTemplater()
    await engine.#initializeFromCache(cachedFiles)
    return engine
  }

  static clearCache() {
    PPTXTemplater.#templateCache.clear()
  }

  enablePerformanceProfile() {
    this.#profiler.enabled = true
    this.#profiler.startTime = performance.now()
    return this
  }

  getPerformanceMetrics() {
    if (!this.#profiler.enabled) {
      return {
        enabled: false,
        message: 'Performance profiling not enabled. Call enablePerformanceProfile() first.',
      }
    }
    const endTime = performance.now()
    this.#profiler.totalMs = endTime - this.#profiler.startTime
    this.#profiler.memoryUsedMB = Math.round((process.memoryUsage().rss / 1024 / 1024) * 100) / 100

    return {
      enabled: true,
      templateLoadMs: Math.round(this.#profiler.templateLoadMs * 100) / 100,
      parseMs: Math.round(this.#profiler.parseMs * 100) / 100,
      chartUpdateMs: Math.round(this.#profiler.chartUpdateMs * 100) / 100,
      imageUpdateMs: Math.round(this.#profiler.imageUpdateMs * 100) / 100,
      zipGenerationMs: Math.round(this.#profiler.zipGenerationMs * 100) / 100,
      totalMs: Math.round(this.#profiler.totalMs * 100) / 100,
      memoryUsedMB: this.#profiler.memoryUsedMB,
    }
  }

  /**
   * Loads a template from a PowerPoint XML Presentation format.
   *
   * @static
   * @param {string|Object} options - Path to presentation.xml, folder root, or configuration object.
   * @returns {Promise<PPTXTemplater>} Initialized engine instance.
   */
  static async fromPresentationXml(options) {
    return PPTXTemplater.load(options)
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
    const engine = new PPTXTemplater()
    await engine.#initializeBlank()
    return engine
  }

  /**
   * Extracts a PPTX file into an unzipped OpenXML folder structure.
   *
   * @static
   * @param {string} pptxPath - Path to the source PPTX file.
   * @param {string} outputPath - Path to the destination folder.
   * @param {Object} [options] - Options (e.g. { overwrite: true }).
   * @returns {Promise<void>}
   */
  static async extractPptx(pptxPath, outputPath, options = {}) {
    const fs = require('fs-extra')
    const path = require('path')

    const resolvedPptx = path.resolve(pptxPath)
    const resolvedOut = path.resolve(outputPath)

    if (!fs.existsSync(resolvedPptx)) {
      throw new PPTXError(`Source PPTX file not found: ${pptxPath}`)
    }

    if (fs.existsSync(resolvedOut)) {
      const stats = fs.statSync(resolvedOut)
      if (stats.isFile()) {
        throw new PPTXError(`Destination is a file: ${outputPath}`)
      }
      const files = fs.readdirSync(resolvedOut)
      if (files.length > 0 && !options.overwrite) {
        throw new PPTXError(
          `Destination directory "${outputPath}" is not empty. Set overwrite: true to overwrite.`
        )
      }
    } else {
      await fs.ensureDir(resolvedOut)
    }

    const engine = await PPTXTemplater.load(resolvedPptx)
    await engine.#zipManager.toFolder(resolvedOut)

    // Validation
    const criticalParts = ['ppt/presentation.xml', 'ppt/slides', 'ppt/_rels', '[Content_Types].xml']

    for (const part of criticalParts) {
      const p = path.join(resolvedOut, part)
      if (!fs.existsSync(p)) {
        throw new PPTXError(`Extracted structure is missing critical part: ${part}`)
      }
    }
  }

  /**
   * Rebuilds a PPTX file from an unzipped OpenXML folder structure.
   *
   * @static
   * @param {string} folderPath - Path to the source folder structure.
   * @param {string} pptxPath - Path to the destination PPTX file.
   * @returns {Promise<void>}
   */
  static async buildPptx(folderPath, pptxPath) {
    const fs = require('fs-extra')
    const path = require('path')

    const resolvedFolder = path.resolve(folderPath)
    const resolvedPptx = path.resolve(pptxPath)

    if (!fs.existsSync(resolvedFolder)) {
      throw new PPTXError(`Source folder not found: ${folderPath}`)
    }

    // Validation of the source folder
    const criticalParts = ['ppt/presentation.xml', 'ppt/slides', 'ppt/_rels', '[Content_Types].xml']

    for (const part of criticalParts) {
      const p = path.join(resolvedFolder, part)
      if (!fs.existsSync(p)) {
        throw new PPTXError(`Source folder is missing critical OpenXML part: ${part}`)
      }
    }

    const engine = await PPTXTemplater.load(resolvedFolder)
    await fs.ensureDir(path.dirname(resolvedPptx))
    await engine.saveToFile(resolvedPptx)
  }

  /**
   * Initializes the engine by loading a PPTX file/buffer.
   * @private
   * @param {string|Buffer} source
   */
  async #initialize(source) {
    const t0 = performance.now()
    logger.debug(`Loading PPTX from ${typeof source === 'string' ? source : 'buffer'}`)

    // Load and extract the ZIP archive (PPTX is just a ZIP)
    await this.#zipManager.load(source)
    const t1 = performance.now()
    this.#profiler.templateLoadMs = t1 - t0

    // Initialize content types manager first!
    await this.#contentTypesManager.initialize(this.#zipManager)

    // Parse the core presentation relationships and structure
    await this.#relationshipManager.initialize(this.#zipManager)

    // Load all slide references from presentation.xml
    await this.#slideManager.initialize(this.#zipManager)

    // Pre-load all slide XML into cache to allow synchronous operations like replaceText()
    await this.#slideManager.preloadAll()

    // Initialize chart manager with zip context
    await this.#chartManager.initialize(this.#zipManager)

    // Deduplicate and index media files
    await this.#mediaManager.initialize(this.#zipManager)

    const t2 = performance.now()
    this.#profiler.parseMs = t2 - t1

    this.#loaded = true
    logger.debug(`Loaded ${this.#slideManager.slideCount} slides successfully`)
  }

  async #initializeFromCache(cachedFiles) {
    const t0 = performance.now()
    logger.debug('Initializing PPTX from cached template')

    const clonedCache = new Map(cachedFiles)
    await this.#zipManager.loadFromCache(clonedCache)
    const t1 = performance.now()
    this.#profiler.templateLoadMs = t1 - t0

    // Initialize content types manager first!
    await this.#contentTypesManager.initialize(this.#zipManager)

    // Parse the core presentation relationships and structure
    await this.#relationshipManager.initialize(this.#zipManager)

    // Load all slide references from presentation.xml
    await this.#slideManager.initialize(this.#zipManager)

    // Pre-load all slide XML into cache to allow synchronous operations like replaceText()
    await this.#slideManager.preloadAll()

    // Initialize chart manager with zip context
    await this.#chartManager.initialize(this.#zipManager)

    // Deduplicate and index media files
    await this.#mediaManager.initialize(this.#zipManager)

    const t2 = performance.now()
    this.#profiler.parseMs = t2 - t1

    this.#loaded = true
  }

  /**
   * Initializes a blank PPTX structure from embedded template XML.
   * @private
   */
  async #initializeBlank() {
    await this.#zipManager.createBlank()
    await this.#contentTypesManager.initialize(this.#zipManager)
    await this.#relationshipManager.initialize(this.#zipManager)
    await this.#slideManager.initialize(this.#zipManager)
    // Pre-load all slide XML so synchronous operations work on the blank template's existing slides
    await this.#slideManager.preloadAll()
    await this.#chartManager.initialize(this.#zipManager)
    await this.#mediaManager.initialize(this.#zipManager)
    this.#loaded = true
  }

  /**
   * Asserts the engine is loaded before performing operations.
   * @private
   */
  #assertLoaded() {
    if (!this.#loaded) {
      throw new PPTXError('Engine not initialized. Call PPTXTemplater.load() first.')
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
    this.#assertLoaded()
    this.#selectedSlides = slideRefs
    logger.debug(`Selected slides: ${slideRefs.join(', ')}`)
    return this
  }

  /**
   * Selects all slides.
   * @returns {PPTXTemplater} this (chainable)
   */
  useAllSlides() {
    this.#assertLoaded()
    this.#selectedSlides = []
    return this
  }

  /**
   * Returns the resolved slide indices based on #selectedSlides.
   * If nothing is selected, returns all slide indices.
   * @private
   * @returns {number[]} Array of 1-based slide indices.
   */
  #getTargetSlideIndices() {
    if (this.#selectedSlides.length === 0) {
      return this.#slideManager.getAllSlideIndices()
    }
    return this.#selectedSlides.flatMap(ref => {
      if (typeof ref === 'number') return [ref]
      // Resolve by tag or ID
      return this.#slideManager.resolveSlideRef(ref)
    })
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
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()

    for (const slideIndex of targetIndices) {
      const slideXml = this.#slideManager.getSlideXml(slideIndex)
      const updated = this.#templateEngine.replaceTextInXml(slideXml, replacements)
      this.#slideManager.setSlideXml(slideIndex, updated)
    }

    logger.debug(
      `Replaced ${Object.keys(replacements).length} placeholder(s) in ${targetIndices.length} slide(s)`
    )
    return this
  }

  /**
   * Updates chart data in the selected slide(s).
   * Finds charts by their name/ID and updates categories, series, and values.
   * Preserves original chart styles, themes, and formatting.
   * Supports inline custom data labels by passing objects in the format `{ data: number, label: string }` instead of numbers.
   *
   * @param {string} chartId - Chart name or relationship ID.
   * @param {ChartData} data - New chart data.
   * @param {string[]} data.categories - Category labels (X-axis).
   * @param {SeriesData[]} data.series - Data series array.
   * @param {string} data.series[].name - Series name.
   * @param {number[]|object[]} data.series[].values - Data values (numbers or label objects).
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * // Simple numeric values:
   * ppt.updateChart('sales-chart', {
   *   categories: ['Jan', 'Feb', 'Mar'],
   *   series: [{ name: 'Revenue', values: [120, 150, 180] }]
   * });
   *
   * // Custom inline data labels:
   * ppt.updateChart('sales-chart', {
   *   categories: ['Q1', 'Q2'],
   *   series: [{
   *     name: 'Revenue',
   *     values: [
   *       { data: 120, label: '120 (40%)' },
   *       { data: 180, label: '180 (60%)' }
   *     ]
   *   }]
   * });
   */
  updateChart(chartId, data) {
    this.#assertLoaded()
    const t0 = performance.now()
    const targetIndices = this.#getTargetSlideIndices()

    for (const slideIndex of targetIndices) {
      this.#chartManager.updateChart(
        slideIndex,
        chartId,
        data,
        this.#slideManager,
        this.#relationshipManager
      )
    }

    this.#profiler.chartUpdateMs += performance.now() - t0
    logger.debug(`Updated chart "${chartId}" in ${targetIndices.length} slide(s)`)
    return this
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
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()

    for (const slideIndex of targetIndices) {
      this.#tableManager.updateTable(
        slideIndex,
        tableId,
        rows,
        this.#slideManager,
        this.#shapeManager
      )
    }

    logger.debug(`Updated table "${tableId}" in ${targetIndices.length} slide(s)`)
    return this
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
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()

    for (const slideIndex of targetIndices) {
      this.#hyperlinkManager.addExternalHyperlink(
        slideIndex,
        options,
        this.#slideManager,
        this.#relationshipManager
      )
    }

    return this
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
    this.#assertLoaded()
    const { sourceSlide, targetSlide, element } = options

    // Fallback: If no element text is provided, link the slide number (legacy behavior)
    if (!element) {
      this.#hyperlinkManager.addSlideHyperlink(
        sourceSlide,
        targetSlide,
        this.#slideManager,
        this.#relationshipManager
      )
    } else {
      // Add a slide hyperlink on specific text
      this.#hyperlinkManager.addTextSlideLink(
        sourceSlide,
        element,
        targetSlide,
        this.#slideManager,
        this.#relationshipManager
      )
    }
    return this
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
    this.#assertLoaded()
    this.#hyperlinkManager.addShapeSlideLink(
      options.slide,
      options.imageId,
      options.targetSlide,
      this.#slideManager,
      this.#relationshipManager
    )
    return this
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
    this.#assertLoaded()
    this.#hyperlinkManager.addShapeSlideLink(
      options.slide,
      options.shapeId,
      options.targetSlide,
      this.#slideManager,
      this.#relationshipManager
    )
    return this
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
    this.#assertLoaded()
    const { slide, element, action } = options
    this.#hyperlinkManager.addTextNavigationLink(slide, element, action, this.#slideManager)
    return this
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
    this.#assertLoaded()
    const { slide, shapeId, action } = options
    this.#hyperlinkManager.addShapeNavigationLink(slide, shapeId, action, this.#slideManager)
    return this
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
    this.#assertLoaded()
    this.#slideManager.addNewSlide(options, this.#relationshipManager, this.#mediaManager)
    logger.debug(`Added new slide: "${options.title || 'Untitled'}"`)
    return this
  }

  /**
   * Clones an existing slide and appends it to the end (or at a position).
   *
   * @param {number} sourceSlideNumber - 1-based source slide number.
   * @param {number} [atPosition] - Optional position to insert (1-based). Default: append.
   * @returns {PPTXTemplater} this (chainable)
   */
  cloneSlide(sourceSlideNumber, atPosition) {
    this.#assertLoaded()
    this.#slideManager.cloneSlide(sourceSlideNumber, atPosition, this.#relationshipManager)
    return this
  }

  /**
   * Removes a slide from the presentation.
   *
   * @param {number} slideNumber - 1-based slide number to remove.
   * @returns {PPTXTemplater} this (chainable)
   */
  removeSlide(slideNumber) {
    this.#assertLoaded()
    this.#slideManager.removeSlide(slideNumber)
    return this
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
    this.#assertLoaded()
    this.#slideManager.reorderSlides(order)
    return this
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
    this.#assertLoaded()
    this.#slideManager.tagSlide(slideNumber, tag)
    return this
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
    this.#assertLoaded()
    return this.#slideManager.exportSlides(slideNumbers, this)
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
    this.#assertLoaded()
    await this.#slideManager.importSlide(sourceEngine, slideRef, this.#mediaManager)
    return this
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
    this.#assertLoaded()
    const slidesToKeep = slideIndices.map(i => this.#slideManager.getSlideInfo(i).slideId)

    // Remove unneeded slides from highest to lowest index to avoid shifting issues
    const allIndices = this.#slideManager.getAllSlideIndices()
    for (let i = allIndices.length; i >= 1; i--) {
      const info = this.#slideManager.getSlideInfo(i)
      if (!slidesToKeep.includes(info.slideId)) {
        this.#slideManager.removeSlide(i)
      }
    }

    // Calculate new target order based on the requested slideIndices
    const currentOrder = this.#slideManager
      .getAllSlideIndices()
      .map(i => this.#slideManager.getSlideInfo(i).slideId)

    const newOrder = slidesToKeep.map(id => {
      return currentOrder.indexOf(id) + 1
    })

    // Only reorder if needed
    if (newOrder.join(',') !== currentOrder.map((_, i) => i + 1).join(',')) {
      this.#slideManager.reorderSlides(newOrder)
    }

    logger.debug(`Imported ${slideIndices.length} slide(s).`)
    return this
  }

  /**
   * Returns presentation metadata (title, author, slide count, etc.)
   *
   * @returns {PresentationInfo} Metadata object.
   */
  getInfo() {
    this.#assertLoaded()
    return {
      slideCount: this.#slideManager.slideCount,
      title: this.#zipManager.getCoreProperty('dc:title') || '',
      author: this.#zipManager.getCoreProperty('dc:creator') || '',
      created: this.#zipManager.getCoreProperty('dcterms:created') || '',
      modified: this.#zipManager.getCoreProperty('dcterms:modified') || '',
      slides: this.#slideManager.getAllSlideInfo(),
      mediaCount: this.#mediaManager.mediaCount,
    }
  }

  /**
   * Validates the XML structure of the current PPTX.
   * Reports issues with relationship IDs, missing parts, etc.
   *
   * @returns {ValidationResult} Object with `valid`, `errors`, and `warnings` arrays.
   */
  validate() {
    this.#assertLoaded()
    return this.#slideManager.validateStructure(this.#relationshipManager, this.#zipManager)
  }

  /**
   * Repairs corrupted OpenXML structure, relationships, and content types.
   * Removes orphan relationships, rebuilds slide references, and fixes missing entries.
   *
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async repair() {
    this.#assertLoaded()

    // 1. Rebuild presentation.xml slide mappings
    this.#slideManager.rebuildPresentationSlideOrder()

    // 2. Remove orphan relationships
    this.#relationshipManager.removeOrphanRelationships(this.#zipManager)

    logger.info('PPTX repair complete.')
    return this
  }

  /**
   * Logs all relationships across the presentation to the console for debugging.
   * @returns {PPTXTemplater} this (chainable)
   */
  debugRelationships() {
    this.#assertLoaded()
    const files = this.#zipManager.listFiles('').filter(f => f.endsWith('.rels'))
    console.log('=== Relationship Graph ===')
    for (const file of files) {
      console.log(`\n${file}:`)
      const rels = this.#relationshipManager.getRelationships(
        file.replace('_rels/', '').replace('.rels', '')
      )
      rels.forEach(r => console.log(`  - ${r.id} [${r.type.split('/').pop()}] -> ${r.target}`))
    }
    return this
  }

  /**
   * Inspects a specific slide's structure and relationships.
   * @param {number} slideIndex - 1-based slide index.
   * @returns {PPTXTemplater} this (chainable)
   */
  inspectSlide(slideIndex) {
    this.#assertLoaded()
    const info = this.#slideManager.getSlideInfo(slideIndex)
    const xml = this.#slideManager.getSlideXml(slideIndex)
    const rels = this.#relationshipManager.getRelationships(info.zipPath)

    console.log(`=== Slide ${slideIndex} Inspection ===`)
    console.log(`Path: ${info.zipPath}`)
    console.log(`ID: ${info.slideId}`)
    console.log(`rId: ${info.relationshipId}`)
    console.log(`Title: ${info.title}`)
    console.log(`XML Size: ${xml.length} characters`)
    console.log(`Relationships (${rels.length}):`)
    rels.forEach(r => console.log(`  - ${r.id} [${r.type.split('/').pop()}] -> ${r.target}`))

    return this
  }

  /**
   * Inspects and logs the raw XML of any file in the ZIP.
   * @param {string} xmlPath - Path inside the ZIP (e.g., 'ppt/slides/slide1.xml')
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async inspectXML(xmlPath) {
    this.#assertLoaded()
    const xml = await this.#zipManager.readFile(xmlPath)
    console.log(`=== XML Inspection: ${xmlPath} ===`)
    if (!xml) {
      console.log('(File not found or empty)')
    } else {
      console.log(xml.substring(0, 1500) + (xml.length > 1500 ? '...\n[Truncated]' : ''))
    }
    return this
  }

  /**
   * Validates all charts in the presentation to ensure they are not corrupted.
   * Checks XML, caches, and embedded workbook references.
   *
   * @returns {Promise<Object>} Validation results for charts.
   */
  async validateCharts() {
    this.#assertLoaded()
    const issues = { valid: true, errors: [], warnings: [] }

    // We lazy require ChartRelationshipManager so we don't circularly depend if not needed
    const { ChartRelationshipManager } = require('../managers/charts/ChartRelationshipManager.js')

    const chartFiles = this.#zipManager.listFiles('ppt/charts/').filter(f => {
      const name = f.split('/').pop()
      return name.startsWith('chart') && name.endsWith('.xml') && !f.includes('_rels')
    })

    for (const chartPath of chartFiles) {
      const relIssues = ChartRelationshipManager.validateChartRelationships(
        this.#relationshipManager,
        this.#zipManager,
        chartPath
      )
      issues.errors.push(...relIssues.errors)
      issues.warnings.push(...relIssues.warnings)
    }

    if (issues.errors.length > 0) issues.valid = false
    return issues
  }

  /**
   * Repairs common chart corruption issues such as broken caches,
   * missing embedded workbooks, or orphan nodes.
   *
   * @returns {Promise<PPTXTemplater>} this
   */
  async repairCharts() {
    this.#assertLoaded()
    logger.info('Repairing charts...')

    // Check all charts for missing embedded workbooks
    const chartFiles = this.#zipManager.listFiles('ppt/charts/').filter(f => {
      const name = f.split('/').pop()
      return name.startsWith('chart') && name.endsWith('.xml') && !f.includes('_rels')
    })
    for (const chartPath of chartFiles) {
      const rels = this.#relationshipManager.getRelationshipsByType(
        chartPath,
        'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package'
      )
      for (const rel of rels) {
        const xlsxPath = this.#relationshipManager.resolveTarget(chartPath, rel.target)
        if (!this.#zipManager.hasFile(xlsxPath)) {
          logger.warn(
            `Chart ${chartPath} has broken workbook reference ${rel.id}, removing to prevent repair mode.`
          )
          this.#relationshipManager.removeRelationship(chartPath, rel.id)

          // Also strip c:externalData from chart XML to prevent PowerPoint looking for it
          const xml = await this.#zipManager.readFile(chartPath)
          if (xml) {
            const updated = xml
              .replace(/<c:externalData[^>]*r:id="[^"]*"[^>]*>/, '')
              .replace(/<\/c:externalData>/, '')
            this.#zipManager.writeFile(chartPath, updated)
          }
        }
      }
    }

    return this
  }

  /**
   * Inspects a specific chart's metadata and structure.
   *
   * @param {string} chartId
   */
  inspectChart(chartId) {
    this.#assertLoaded()
    console.log(`=== Chart Inspection: ${chartId} ===`)
    // Find chart across all slides to get info
    let found = false
    for (const i of this.#slideManager.getAllSlideIndices()) {
      try {
        const info = this.#chartManager.getChartsInSlide(
          i,
          this.#slideManager,
          this.#relationshipManager
        )
        const chart = info.find(
          c => c.zipPath.toLowerCase().includes(chartId.toLowerCase()) || c.rId === chartId
        )
        if (chart) {
          console.log(`Found on Slide ${i}`)
          console.log(`ZIP Path: ${chart.zipPath}`)
          console.log(`Relationship ID: ${chart.rId}`)
          found = true
          break
        }
      } catch (e) {}
    }
    if (!found) console.log('Chart not found.')
    return this
  }

  /**
   * Inspects and logs the raw XML of a chart file.
   *
   * @param {string} chartFileName
   */
  async inspectChartXML(chartFileName) {
    const fullPath = chartFileName.includes('/') ? chartFileName : `ppt/charts/${chartFileName}`
    await this.inspectXML(fullPath)
    return this
  }

  /**
   * Logs all chart relationships.
   */
  debugChartRelationships() {
    this.#assertLoaded()
    console.log('=== Chart Relationships ===')
    const chartFiles = this.#zipManager.listFiles('ppt/charts/').filter(f => {
      const name = f.split('/').pop()
      return name.startsWith('chart') && name.endsWith('.xml') && !f.includes('_rels')
    })
    for (const chartPath of chartFiles) {
      console.log(`\n${chartPath}:`)
      const rels = this.#relationshipManager.getRelationships(chartPath)
      rels.forEach(r => console.log(`  - ${r.id} [${r.type.split('/').pop()}] -> ${r.target}`))
    }
    return this
  }

  /**
   * Saves the modified PPTX to a file on disk.
   *
   * @param {string} filePath - Output file path.
   * @param {Object} [options] - Save options.
   * @param {boolean} [options.strict=false] - Throw error on validation failure.
   * @returns {Promise<void>}
   */
  async saveToFile(filePath, options = {}) {
    this.#assertLoaded()
    const result = await this.validatePresentation()
    if (!result.valid) {
      if (options.strict) {
        throw new PPTXError(`Validation failed before save: ${result.errors.join(', ')}`)
      } else {
        logger.warn(
          `Validation issues found before save:\n${result.errors.map(e => `  • ${e}`).join('\n')}`
        )
      }
    }
    await this.validateArchive()

    const t0 = performance.now()
    await this.#outputWriter.saveToFile(filePath, this.#slideManager, this.#zipManager, options)
    this.#profiler.zipGenerationMs += performance.now() - t0

    logger.info(`Saved PPTX to ${filePath}`)
  }

  /**
   * Saves the presentation. Equivalent to saveToFile.
   *
   * @param {string} filePath - Output file path.
   * @param {Object} [options] - Save options.
   * @returns {Promise<void>}
   */
  async save(filePath, options = {}) {
    return this.saveToFile(filePath, options)
  }

  /**
   * Saves the modified presentation XML structures directly to a folder.
   *
   * @param {string} folderPath - Target directory path.
   * @returns {Promise<void>}
   */
  async saveXml(folderPath) {
    this.#assertLoaded()
    await this.#outputWriter.flush(this.#slideManager, this.#zipManager)
    await this.#zipManager.toFolder(folderPath)
    logger.info(`Saved XML presentation to folder ${folderPath}`)
  }

  /**
   * Saves the modified presentation XML structures directly to a folder.
   *
   * @param {string} folderPath - Target directory path.
   * @returns {Promise<void>}
   */
  async saveToFolder(folderPath) {
    return this.saveXml(folderPath)
  }

  /**
   * Returns the PPTX content as a Node.js Buffer.
   *
   * @param {Object} [options] - Save options.
   * @returns {Promise<Buffer>}
   */
  async toBuffer(options = {}) {
    this.#assertLoaded()
    await this.validateArchive()

    const t0 = performance.now()
    const buffer = await this.#outputWriter.toBuffer(this.#slideManager, this.#zipManager, options)
    this.#profiler.zipGenerationMs += performance.now() - t0

    return buffer
  }

  /**
   * Returns the PPTX content as a readable Node.js Stream.
   *
   * @param {Object} [options] - Save options.
   * @returns {Promise<NodeJS.ReadableStream>}
   */
  async toStream(options = {}) {
    this.#assertLoaded()
    await this.validateArchive()

    const t0 = performance.now()
    const stream = await this.#outputWriter.toStream(this.#slideManager, this.#zipManager, options)
    this.#profiler.zipGenerationMs += performance.now() - t0

    return stream
  }

  /**
   * Saves the presentation to a readable stream or pipes it to a writable stream.
   *
   * @param {NodeJS.WritableStream|Object} [writableOrOptions] - Writable stream to pipe to, or options object.
   * @param {Object} [options] - Save options if writable stream was passed first.
   * @returns {Promise<NodeJS.ReadableStream|void>}
   */
  async saveToStream(writableOrOptions, options = {}) {
    this.#assertLoaded()
    let writable = null
    let opts = options
    if (writableOrOptions && typeof writableOrOptions.write === 'function') {
      writable = writableOrOptions
    } else if (writableOrOptions) {
      opts = writableOrOptions
    }

    const stream = await this.toStream(opts)
    if (writable) {
      return new Promise((resolve, reject) => {
        stream.pipe(writable)
        writable.on('finish', resolve)
        writable.on('error', reject)
        stream.on('error', reject)
      })
    }
    return stream
  }

  // === Slide Features ===
  duplicateSlide(slideIndex, atPosition) {
    this.#assertLoaded()
    this.#slideManager.duplicateSlide(slideIndex, atPosition, this.#relationshipManager)
    return this
  }

  deleteSlide(slideIndex) {
    this.#assertLoaded()
    this.#slideManager.removeSlide(slideIndex)
    return this
  }

  moveSlide(fromIndex, toIndex) {
    this.#assertLoaded()
    this.#slideManager.moveSlide(fromIndex, toIndex)
    return this
  }

  insertSlide(slideIndex, options = {}) {
    this.#assertLoaded()
    this.#slideManager.insertSlide(
      slideIndex,
      options,
      this.#relationshipManager,
      this.#mediaManager
    )
    return this
  }

  getSlides() {
    this.#assertLoaded()
    return this.#slideManager.getSlides()
  }

  // === Table Features ===
  getTableRows(tableId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    if (targetIndices.length === 0) {
      throw new PPTXError('No slides active/loaded')
    }
    const idx = targetIndices[0]
    return this.#tableManager.getTableRows(idx, tableId, options, this.#slideManager)
  }

  addTableRow(tableId, rowData, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.addTableRow(idx, tableId, rowData, this.#slideManager, options)
    }
    return this
  }

  removeTableRow(tableId, rowIndex) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.removeTableRow(idx, tableId, rowIndex, this.#slideManager)
    }
    return this
  }

  insertTableRow(tableId, rowIndex, rowData) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.insertTableRow(idx, tableId, rowIndex, rowData, this.#slideManager)
    }
    return this
  }

  cloneTableRow(tableId, sourceRowIndex, targetRowIndex) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.cloneTableRow(
        idx,
        tableId,
        sourceRowIndex,
        targetRowIndex,
        this.#slideManager
      )
    }
    return this
  }

  updateCell(tableId, rowIndex, colIndex, value, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.updateCell(
        idx,
        tableId,
        rowIndex,
        colIndex,
        value,
        options,
        this.#slideManager
      )
    }
    return this
  }

  mergeCells(tableIdOrOptions, startRow, startCol, endRow, endCol) {
    this.#assertLoaded()
    let tableId = tableIdOrOptions
    let sRow = startRow
    let sCol = startCol
    let eRow = endRow
    let eCol = endCol
    let targetIndices = this.#getTargetSlideIndices()

    if (tableIdOrOptions && typeof tableIdOrOptions === 'object') {
      const opt = tableIdOrOptions
      tableId = opt.tableId
      sRow = opt.startRow
      sCol = opt.startCol
      eRow = opt.endRow
      eCol = opt.endCol
      if (opt.slide !== undefined) {
        targetIndices = [opt.slide]
      }
    }

    for (const idx of targetIndices) {
      this.#tableManager.mergeCells(idx, tableId, sRow, sCol, eRow, eCol, this.#slideManager)
    }
    return this
  }

  unmergeCells(tableIdOrOptions, startRow, startCol, endRow, endCol) {
    this.#assertLoaded()
    let tableId = tableIdOrOptions
    let sRow = startRow
    let sCol = startCol
    let eRow = endRow
    let eCol = endCol
    let targetIndices = this.#getTargetSlideIndices()
    let isCellCoord = false
    let cellRow, cellCol

    if (tableIdOrOptions && typeof tableIdOrOptions === 'object') {
      const opt = tableIdOrOptions
      tableId = opt.tableId
      sRow = opt.startRow
      sCol = opt.startCol
      eRow = opt.endRow
      eCol = opt.endCol
      if (opt.slide !== undefined) {
        targetIndices = [opt.slide]
      }
      if (opt.row !== undefined && opt.col !== undefined) {
        isCellCoord = true
        cellRow = opt.row
        cellCol = opt.col
      }
    }

    for (const idx of targetIndices) {
      if (isCellCoord) {
        this.#tableManager.unmergeCells(idx, tableId, cellRow, cellCol, this.#slideManager)
      } else {
        this.#tableManager.unmergeCells(idx, tableId, sRow, sCol, eRow, eCol, this.#slideManager)
      }
    }
    return this
  }

  getMergedCells(tableId) {
    this.#assertLoaded()
    const slideIndex = this.#getTargetSlideIndices()[0] || 1
    return this.#tableManager.getMergedCells(slideIndex, tableId || 'first', this.#slideManager)
  }

  validateMergeRegion(tableId, startRow, startCol, endRow, endCol) {
    this.#assertLoaded()
    const slideIndex = this.#getTargetSlideIndices()[0] || 1
    return this.#tableManager.validateMergeRegion(
      slideIndex,
      tableId || 'first',
      startRow,
      startCol,
      endRow,
      endCol,
      this.#slideManager
    )
  }

  isMergedCell(tableId, row, col) {
    this.#assertLoaded()
    const slideIndex = this.#getTargetSlideIndices()[0] || 1
    return this.#tableManager.isMergedCell(
      slideIndex,
      tableId || 'first',
      row,
      col,
      this.#slideManager
    )
  }

  getMergeParent(tableId, row, col) {
    this.#assertLoaded()
    const slideIndex = this.#getTargetSlideIndices()[0] || 1
    return this.#tableManager.getMergeParent(
      slideIndex,
      tableId || 'first',
      row,
      col,
      this.#slideManager
    )
  }

  getMergeRegion(tableId, row, col) {
    this.#assertLoaded()
    const slideIndex = this.#getTargetSlideIndices()[0] || 1
    return this.#tableManager.getMergeRegion(
      slideIndex,
      tableId || 'first',
      row,
      col,
      this.#slideManager
    )
  }

  splitMergedRegion(tableId, row, col) {
    this.#assertLoaded()
    const slideIndex = this.#getTargetSlideIndices()[0] || 1
    this.#tableManager.splitMergedRegion(
      slideIndex,
      tableId || 'first',
      row,
      col,
      this.#slideManager
    )
    return this
  }

  cloneMergedRegion(tableId, row, col, targetRow, targetCol) {
    this.#assertLoaded()
    const slideIndex = this.#getTargetSlideIndices()[0] || 1
    this.#tableManager.cloneMergedRegion(
      slideIndex,
      tableId || 'first',
      row,
      col,
      targetRow,
      targetCol,
      this.#slideManager
    )
    return this
  }

  autoFitTable(tableId) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.autoFitTable(idx, tableId, this.#slideManager)
    }
    return this
  }

  resizeTable(tableId, width, height) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.resizeTable(idx, tableId, width, height, this.#slideManager)
    }
    return this
  }

  getTables() {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const tables = []
    for (const idx of targetIndices) {
      tables.push(...this.#tableManager.inspectTables(idx, this.#slideManager))
    }
    return tables
  }

  // === Chart Features ===
  updateChartData(chartId, data) {
    return this.updateChart(chartId, data)
  }

  replaceChartSeries(chartId, seriesIndex, newSeriesData) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#chartManager.replaceChartSeries(
        idx,
        chartId,
        seriesIndex,
        newSeriesData,
        this.#slideManager,
        this.#relationshipManager
      )
    }
    return this
  }

  updateChartTitle(chartId, title) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#chartManager.updateChartTitle(
        idx,
        chartId,
        title,
        this.#slideManager,
        this.#relationshipManager
      )
    }
    return this
  }

  updateChartCategories(chartId, categories) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#chartManager.updateChartCategories(
        idx,
        chartId,
        categories,
        this.#slideManager,
        this.#relationshipManager
      )
    }
    return this
  }

  updateDataLabels(chartId, options) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#chartManager.updateDataLabels(
        idx,
        chartId,
        options,
        this.#slideManager,
        this.#relationshipManager
      )
    }
    return this
  }

  async getDataLabels(chartId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    if (targetIndices.length === 0) return []
    const idx = targetIndices[0]
    return this.#chartManager.getDataLabels(
      idx,
      chartId,
      options,
      this.#slideManager,
      this.#relationshipManager
    )
  }

  async validateDataLabels(chartId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    if (targetIndices.length === 0) return { valid: true, errors: [] }
    const idx = targetIndices[0]
    return ValidationEngine.validateDataLabels(this, idx, chartId, options)
  }

  async validateChartLabels(chartId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    if (targetIndices.length === 0) return { valid: true, errors: [], warnings: [] }
    const idx = targetIndices[0]
    return this.#chartManager.validateChartLabels(
      idx,
      chartId,
      options,
      this.#slideManager,
      this.#relationshipManager
    )
  }

  async validateSeriesNameLabels(chartId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    if (targetIndices.length === 0) return { valid: true, errors: [], warnings: [] }
    const idx = targetIndices[0]
    return this.#chartManager.validateSeriesNameLabels(
      idx,
      chartId,
      options,
      this.#slideManager,
      this.#relationshipManager
    )
  }

  getCharts() {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const charts = []
    for (const idx of targetIndices) {
      charts.push(
        ...this.#chartManager.getChartsInSlide(idx, this.#slideManager, this.#relationshipManager)
      )
    }
    return charts
  }

  /**
   * Retrieves the exact coordinate positions of all data labels for a chart on the active slide.
   * Calculates absolute layout limits in EMUs (English Metric Units).
   *
   * @param {string} chartId The unique chart name/id in the template slide.
   * @returns {Promise<Array<{series: string, category: string, seriesIndex: number, categoryIndex: number, value: number, x: number, y: number, width: number, height: number}>>} An array of data label geometry objects.
   */
  async getChartLabelPositions(chartId) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    if (targetIndices.length === 0) return []
    const idx = targetIndices[0]
    return this.#chartManager.getChartLabelPositions(
      idx,
      chartId,
      this.#slideManager,
      this.#relationshipManager
    )
  }

  /**
   * Retrieves the exact coordinate positions of all bars/columns for a chart on the active slide.
   * Calculates absolute layout limits in EMUs (English Metric Units).
   *
   * @param {string} chartId The unique chart name/id in the template slide.
   * @returns {Promise<Array<{series: string, category: string, seriesIndex: number, categoryIndex: number, value: number, x: number, y: number, width: number, height: number}>>} An array of bar geometry objects.
   */
  async getChartBarPositions(chartId) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    if (targetIndices.length === 0) return []
    const idx = targetIndices[0]
    return this.#chartManager.getChartBarPositions(
      idx,
      chartId,
      this.#slideManager,
      this.#relationshipManager
    )
  }

  /**
   * Adds a textbox shape at a specific EMU coordinate position on targeted slides.
   * Supports custom font styling and alignment configuration.
   *
   * @param {Object} options Textbox positioning and style configuration.
   * @param {string} options.text Text content to insert in the textbox.
   * @param {number} options.x Bounding box X offset coordinate (in EMUs).
   * @param {number} options.y Bounding box Y offset coordinate (in EMUs).
   * @param {number} [options.width=1200000] Bounding box width (in EMUs).
   * @param {number} [options.height=300000] Bounding box height (in EMUs).
   * @param {Object} [options.style] Font formatting properties (fontSize, fontFamily, color, bold, italic, align).
   * @returns {this} The chainable presentation engine instance.
   */
  addTextAtPosition(options) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      const p = this.#chartManager.addTextAtPosition(idx, options, this.#slideManager)
      this.#zipManager.addPendingPromise(p)
    }
    return this
  }

  /**
   * Dynamically places textboxes next to a chart's data labels with vertical collision avoidance.
   * Textboxes are positioned either on the left or right of the chart area, vertically aligned with their corresponding label.
   *
   * @param {Object} options Text alignment, naming, and position configuration.
   * @param {string} options.chart The target chart name/id.
   * @param {string|Function} options.text Static label text or a callback function receiving `({ series, category, value })`.
   * @param {'left'|'right'} [options.position='left'] Alignment position relative to the chart boundaries.
   * @param {Object} [options.style] Text styling attributes (fontSize, fontFamily, color, bold, italic, align, autoFit).
   * @returns {this} The chainable presentation engine instance.
   */
  addTextNearChartLabel(options) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      const p = this.#chartManager.addTextNearChartLabel(
        idx,
        options,
        this.#slideManager,
        this.#relationshipManager
      )
      this.#zipManager.addPendingPromise(p)
    }
    return this
  }

  /**
   * Updates shape text or list content by placeholder tag or shape name/ID.
   * Supports bullet lists, numbered lists, nested lists, and custom styling.
   *
   * @param {string} tag - Placeholder tag (e.g. '{{name}}' or 'name') or shape name/ID.
   * @param {string|Object} data - String value or list configuration object.
   * @returns {PPTXTemplater} this (chainable)
   */
  updateText(tag, data) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.updateText(idx, tag, data, this.#slideManager, this.#templateEngine)
    }
    return this
  }

  /**
   * Retrieves list items from a shape or text box by name or placeholder tag.
   *
   * @param {string} tag - Shape name/ID or placeholder tag.
   * @returns {Array} Nested list structure of items.
   */
  getList(tag) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const idx = targetIndices.length > 0 ? targetIndices[0] : 1
    return this.#textManager.getList(idx, tag, this.#slideManager)
  }

  /**
   * Validates a list structure and values.
   *
   * @param {Object|Array} data - List config object or array of items.
   * @returns {Object} Report containing validation result.
   */
  validateList(data) {
    return ValidationEngine.validateList(data)
  }

  // === Text Features ===
  replaceTextByTag(tag, value, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.replaceTextByTag(
        idx,
        tag,
        value,
        options,
        this.#slideManager,
        this.#templateEngine
      )
    }
    return this
  }

  replaceMultiple(replacements, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.replaceMultiple(
        idx,
        replacements,
        options,
        this.#slideManager,
        this.#templateEngine
      )
    }
    return this
  }

  findText(text) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const matches = []
    for (const idx of targetIndices) {
      matches.push(...this.#textManager.findText(idx, text, this.#slideManager))
    }
    return matches
  }

  getTextElements() {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const elements = []
    for (const idx of targetIndices) {
      elements.push(...this.#textManager.getTextElements(idx, this.#slideManager))
    }
    return elements
  }

  // === Shape Features ===
  updateShapeText(shapeId, text) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#shapeManager.updateShapeText(idx, shapeId, text, this.#slideManager)
    }
    return this
  }

  /**
   * Updates the position and/or dimensions of an existing shape on targeted slides.
   *
   * @param {string} shapeId The unique shape name/id in the template slide.
   * @param {Object} options Positioning and styling dimensions config.
   * @param {number} [options.x] Absolute X offset coordinate (in EMUs).
   * @param {number} [options.y] Absolute Y offset coordinate (in EMUs).
   * @param {number} [options.width] Bounding box width (in EMUs).
   * @param {number} [options.height] Bounding box height (in EMUs).
   * @returns {this} The chainable presentation engine instance.
   */
  updateShapePosition(shapeId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#shapeManager.updateShapePosition(idx, shapeId, options, this.#slideManager)
    }
    return this
  }

  /**
   * Updates the position and/or dimensions of an existing textbox on targeted slides.
   *
   * @param {string} textBoxId The unique textbox shape name/id in the template slide.
   * @param {Object} options Positioning and styling dimensions config.
   * @param {number} [options.x] Absolute X offset coordinate (in EMUs).
   * @param {number} [options.y] Absolute Y offset coordinate (in EMUs).
   * @param {number} [options.width] Bounding box width (in EMUs).
   * @param {number} [options.height] Bounding box height (in EMUs).
   * @returns {this} The chainable presentation engine instance.
   */
  updateTextBoxPosition(textBoxId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#shapeManager.updateTextBoxPosition(idx, textBoxId, options, this.#slideManager)
    }
    return this
  }

  cloneShape(shapeId, newShapeId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#shapeManager.cloneShape(idx, shapeId, newShapeId, options, this.#slideManager)
    }
    return this
  }

  deleteShape(shapeId) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#shapeManager.deleteShape(idx, shapeId, this.#slideManager)
    }
    return this
  }

  getShapes() {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const shapes = []
    for (const idx of targetIndices) {
      shapes.push(...this.#shapeManager.getShapes(idx, this.#slideManager))
    }
    return shapes
  }

  /**
   * Validates shape options configuration.
   *
   * @param {Object} options Shape creation/update options.
   * @returns {string[]} List of validation error messages.
   */
  validateShape(options) {
    return this.#shapeManager.validateShape(options)
  }

  /**
   * Adds a new shape dynamically to the targeted slide(s).
   *
   * @param {Object} options Shape configuration options.
   * @returns {this} The chainable presentation templater instance.
   */
  async addShape(options) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#shapeManager.addShape(idx, options, this.#slideManager)
    }
    return this
  }

  /**
   * Updates an existing shape in-place.
   *
   * @param {string} shapeId Shape ID or template name to update.
   * @param {Object} options Configuration properties to update.
   * @returns {this} The chainable presentation templater instance.
   */
  async updateShape(shapeId, options) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#shapeManager.updateShape(idx, shapeId, options, this.#slideManager)
    }
    return this
  }

  /**
   * Removes a shape from the targeted slide(s).
   *
   * @param {string} shapeId Shape ID or template name to remove.
   * @returns {this} The chainable presentation templater instance.
   */
  async removeShape(shapeId) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#shapeManager.removeShape(idx, shapeId, this.#slideManager)
    }
    return this
  }

  /**
   * Discovers and retrieves details of an existing shape on the targeted slides.
   *
   * @param {string} shapeId Shape ID or template name to locate.
   * @returns {Object|null} Shape details object, or null if not found.
   */
  getShape(shapeId) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      const shape = this.#shapeManager.getShape(idx, shapeId, this.#slideManager)
      if (shape) return shape
    }
    return null
  }

  /**
   * Dynamically adds a shape inside a table cell based on cell coordinates.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} colIndex - 0-based column index.
   * @param {Object} options - Shape configuration options.
   * @returns {this} The chainable presentation templater instance.
   */
  async addCellShape(tableId, rowIndex, colIndex, options) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.addCellShape(
        idx,
        tableId,
        rowIndex,
        colIndex,
        options,
        this.#slideManager,
        this.#shapeManager
      )
    }
    return this
  }

  /**
   * Updates an existing shape inside a table cell.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} colIndex - 0-based column index.
   * @param {number} shapeIndex - 0-based shape index in the cell.
   * @param {Object} options - Shape configuration properties to update.
   * @returns {this} The chainable presentation templater instance.
   */
  async updateCellShape(tableId, rowIndex, colIndex, shapeIndex, options) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.updateCellShape(
        idx,
        tableId,
        rowIndex,
        colIndex,
        shapeIndex,
        options,
        this.#slideManager,
        this.#shapeManager
      )
    }
    return this
  }

  /**
   * Removes a shape from a table cell.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} colIndex - 0-based column index.
   * @param {number} shapeIndex - 0-based shape index in the cell.
   * @returns {this} The chainable presentation templater instance.
   */
  async removeCellShape(tableId, rowIndex, colIndex, shapeIndex) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.removeCellShape(
        idx,
        tableId,
        rowIndex,
        colIndex,
        shapeIndex,
        this.#slideManager,
        this.#shapeManager
      )
    }
    return this
  }

  /**
   * Discovers and retrieves details of an existing cell shape on the targeted slide.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} colIndex - 0-based column index.
   * @param {number} shapeIndex - 0-based shape index in the cell.
   * @returns {Object|null} Shape details object, or null if not found.
   */
  getCellShape(tableId, rowIndex, colIndex, shapeIndex) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      const shape = this.#tableManager.getCellShape(
        idx,
        tableId,
        rowIndex,
        colIndex,
        shapeIndex,
        this.#slideManager,
        this.#shapeManager
      )
      if (shape) return shape
    }
    return null
  }

  /**
   * Retrieves final rendered bounds of a table cell in pixels.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} colIndex - 0-based column index.
   * @returns {Object|null} Cell bounds { x, y, width, height } in pixels, or null.
   */
  getCellBounds(tableId, rowIndex, colIndex) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      try {
        const bounds = this.#tableManager.getCellBounds(
          idx,
          tableId,
          rowIndex,
          colIndex,
          this.#slideManager
        )
        if (bounds) return bounds
      } catch (err) {
        logger.debug(
          `Could not get cell bounds for table ${tableId} on slide ${idx}: ${err.message}`
        )
      }
    }
    return null
  }

  /**
   * Retrieves final rendered position of a table cell in pixels.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} colIndex - 0-based column index.
   * @returns {Object|null} Cell position { row, column, x, y } in pixels, or null.
   */
  getCellPosition(tableId, rowIndex, colIndex) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      try {
        const pos = this.#tableManager.getCellPosition(
          idx,
          tableId,
          rowIndex,
          colIndex,
          this.#slideManager
        )
        if (pos) return pos
      } catch (err) {
        logger.debug(
          `Could not get cell position for table ${tableId} on slide ${idx}: ${err.message}`
        )
      }
    }
    return null
  }

  // === Image Features ===
  async replaceImage(imageIdOrName, sourcePathOrBuffer) {
    this.#assertLoaded()
    const t0 = performance.now()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      await this.#imageManager.replaceImage(
        idx,
        imageIdOrName,
        sourcePathOrBuffer,
        this.#slideManager,
        this.#mediaManager,
        this.#relationshipManager
      )
    }
    this.#profiler.imageUpdateMs += performance.now() - t0
    return this
  }

  async addImage(sourcePathOrBuffer, options = {}) {
    this.#assertLoaded()
    const t0 = performance.now()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      await this.#imageManager.addImage(
        idx,
        sourcePathOrBuffer,
        options,
        this.#slideManager,
        this.#mediaManager,
        this.#relationshipManager
      )
    }
    this.#profiler.imageUpdateMs += performance.now() - t0
    return this
  }

  removeImage(imageIdOrName) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#imageManager.removeImage(
        idx,
        imageIdOrName,
        this.#slideManager,
        this.#relationshipManager
      )
    }
    return this
  }

  getImages() {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const images = []
    for (const idx of targetIndices) {
      images.push(
        ...this.#imageManager.getImages(idx, this.#slideManager, this.#relationshipManager)
      )
    }
    return images
  }

  // === Validation Features ===
  async validatePresentation() {
    this.#assertLoaded()
    return await ValidationEngine.validatePresentation(this)
  }

  /**
   * Performs validation specifically on PowerPoint XML folder contents/relationships.
   *
   * @returns {Promise<{valid: boolean, errors: string[], warnings: string[]}>} Validation report.
   */
  async validatePresentationXml() {
    this.#assertLoaded()
    const errors = []
    const warnings = []

    try {
      const presResult = await this.validatePresentation()
      errors.push(...presResult.errors)
      warnings.push(...presResult.warnings)
    } catch (err) {
      errors.push(`Presentation validation error: ${err.message}`)
    }

    try {
      await this.validateArchive()
    } catch (err) {
      errors.push(err.message)
    }

    if (this.#zipManager.hasFile('[Content_Types].xml')) {
      try {
        const ctXml = await this.#zipManager.readFile('[Content_Types].xml')
        const ctObj = this.#xmlParser.parse(ctXml, '[Content_Types].xml')
        const overrides = ctObj?.Types?.Override || []
        const overrideList = Array.isArray(overrides) ? overrides : [overrides]

        for (const override of overrideList) {
          const partName = override['@_PartName']
          const contentType = override['@_ContentType']
          if (partName && contentType) {
            const cleanPath = partName.startsWith('/') ? partName.substring(1) : partName
            if (!this.#zipManager.hasFile(cleanPath)) {
              errors.push(`Content types override refers to missing file: ${cleanPath}`)
            }
          }
        }
      } catch (err) {
        errors.push(`Invalid [Content_Types].xml structure: ${err.message}`)
      }
    } else {
      errors.push('Missing [Content_Types].xml')
    }

    const slideInfo = this.#slideManager.getAllSlideInfo()
    for (const slide of slideInfo) {
      const rels = this.#relationshipManager.getRelationships(slide.zipPath)
      const layoutRel = rels.find(
        r =>
          r.type ===
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout'
      )
      if (layoutRel) {
        const layoutPath = this.#relationshipManager.resolveTarget(slide.zipPath, layoutRel.target)
        if (!this.#zipManager.hasFile(layoutPath)) {
          errors.push(`Slide ${slide.index} refers to missing slideLayout: ${layoutPath}`)
        } else {
          const layoutRels = this.#relationshipManager.getRelationships(layoutPath)
          const masterRel = layoutRels.find(
            r =>
              r.type ===
              'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster'
          )
          if (masterRel) {
            const masterPath = this.#relationshipManager.resolveTarget(layoutPath, masterRel.target)
            if (!this.#zipManager.hasFile(masterPath)) {
              errors.push(`Slide layout ${layoutPath} refers to missing slideMaster: ${masterPath}`)
            }
          } else {
            warnings.push(`Slide layout ${layoutPath} has no slideMaster relationship`)
          }
        }
      } else {
        errors.push(`Slide ${slide.index} has no slideLayout relationship`)
      }
    }

    return {
      valid: errors.length === 0,
      errors,
      warnings,
    }
  }

  async validateSlide(slideIndex) {
    this.#assertLoaded()
    return await ValidationEngine.validateSlide(this, slideIndex)
  }

  async validateTable(tableId) {
    this.#assertLoaded()
    return await ValidationEngine.validateTable(
      this,
      this.#getTargetSlideIndices()[0] || 1,
      tableId
    )
  }

  async validateArchive() {
    this.#assertLoaded()
    await this.#zipManager.validateArchive()
    return this
  }

  enableDebugZip() {
    this.#outputWriter.debugZip = true
    return this
  }

  validateRelationships(partPath) {
    this.#assertLoaded()
    return ValidationEngine.validateRelationships(this, partPath)
  }

  /**
   * Returns the total number of slides in the loaded presentation.
   * @type {number}
   */
  get slideCount() {
    return this.#slideManager.slideCount
  }

  // --- Public Getters for Internal Managers ---
  get zipManager() {
    return this.#zipManager
  }
  get xmlParser() {
    return this.#xmlParser
  }
  get contentTypesManager() {
    return this.#contentTypesManager
  }
  get relationshipManager() {
    return this.#relationshipManager
  }
  get slideManager() {
    return this.#slideManager
  }
  get chartManager() {
    return this.#chartManager
  }
  get tableManager() {
    return this.#tableManager
  }
  get shapeManager() {
    return this.#shapeManager
  }
  get imageManager() {
    return this.#imageManager
  }
  get textManager() {
    return this.#textManager
  }
  get hyperlinkManager() {
    return this.#hyperlinkManager
  }
  get mediaManager() {
    return this.#mediaManager
  }

  // Z-Order / Layer Management APIs

  /**
   * Moves slide element one layer forward.
   */
  bringForward(optionsOrId) {
    this.#assertLoaded()
    let slideIndex, objectId
    if (typeof optionsOrId === 'object' && optionsOrId !== null) {
      slideIndex =
        optionsOrId.slide !== undefined ? optionsOrId.slide : this.#getTargetSlideIndices()[0] || 1
      objectId = optionsOrId.objectId
    } else {
      slideIndex = this.#getTargetSlideIndices()[0] || 1
      objectId = optionsOrId
    }
    this.#zOrderManager.bringForward(slideIndex, objectId, this.#slideManager)
    return this
  }

  /**
   * Moves slide element one layer backward.
   */
  sendBackward(optionsOrId) {
    this.#assertLoaded()
    let slideIndex, objectId
    if (typeof optionsOrId === 'object' && optionsOrId !== null) {
      slideIndex =
        optionsOrId.slide !== undefined ? optionsOrId.slide : this.#getTargetSlideIndices()[0] || 1
      objectId = optionsOrId.objectId
    } else {
      slideIndex = this.#getTargetSlideIndices()[0] || 1
      objectId = optionsOrId
    }
    this.#zOrderManager.sendBackward(slideIndex, objectId, this.#slideManager)
    return this
  }

  /**
   * Moves slide element above all other objects.
   */
  bringToFront(optionsOrId) {
    this.#assertLoaded()
    let slideIndex, objectId
    if (typeof optionsOrId === 'object' && optionsOrId !== null) {
      slideIndex =
        optionsOrId.slide !== undefined ? optionsOrId.slide : this.#getTargetSlideIndices()[0] || 1
      objectId = optionsOrId.objectId
    } else {
      slideIndex = this.#getTargetSlideIndices()[0] || 1
      objectId = optionsOrId
    }
    this.#zOrderManager.bringToFront(slideIndex, objectId, this.#slideManager)
    return this
  }

  /**
   * Moves slide element behind all other objects.
   */
  sendToBack(optionsOrId) {
    this.#assertLoaded()
    let slideIndex, objectId
    if (typeof optionsOrId === 'object' && optionsOrId !== null) {
      slideIndex =
        optionsOrId.slide !== undefined ? optionsOrId.slide : this.#getTargetSlideIndices()[0] || 1
      objectId = optionsOrId.objectId
    } else {
      slideIndex = this.#getTargetSlideIndices()[0] || 1
      objectId = optionsOrId
    }
    this.#zOrderManager.sendToBack(slideIndex, objectId, this.#slideManager)
    return this
  }

  /**
   * Moves slide element to the specific 1-based stacking position.
   */
  setZIndex(optionsOrId, zIndex) {
    this.#assertLoaded()
    let slideIndex, objectId, targetZIndex
    if (
      typeof optionsOrId === 'object' &&
      optionsOrId !== null &&
      optionsOrId.zIndex !== undefined
    ) {
      slideIndex =
        optionsOrId.slide !== undefined ? optionsOrId.slide : this.#getTargetSlideIndices()[0] || 1
      objectId = optionsOrId.objectId
      targetZIndex = optionsOrId.zIndex
    } else {
      slideIndex = this.#getTargetSlideIndices()[0] || 1
      objectId = optionsOrId
      targetZIndex = zIndex
    }
    this.#zOrderManager.setZIndex(slideIndex, objectId, targetZIndex, this.#slideManager)
    return this
  }

  /**
   * Moves slide element directly before (below) a target element.
   */
  moveObjectBefore(optionsOrId, targetId) {
    this.#assertLoaded()
    let slideIndex, objectId, target
    if (
      typeof optionsOrId === 'object' &&
      optionsOrId !== null &&
      optionsOrId.targetId !== undefined
    ) {
      slideIndex =
        optionsOrId.slide !== undefined ? optionsOrId.slide : this.#getTargetSlideIndices()[0] || 1
      objectId = optionsOrId.objectId
      target = optionsOrId.targetId
    } else {
      slideIndex = this.#getTargetSlideIndices()[0] || 1
      objectId = optionsOrId
      target = targetId
    }
    this.#zOrderManager.moveObjectBefore(slideIndex, objectId, target, this.#slideManager)
    return this
  }

  /**
   * Moves slide element directly after (above) a target element.
   */
  moveObjectAfter(optionsOrId, targetId) {
    this.#assertLoaded()
    let slideIndex, objectId, target
    if (
      typeof optionsOrId === 'object' &&
      optionsOrId !== null &&
      optionsOrId.targetId !== undefined
    ) {
      slideIndex =
        optionsOrId.slide !== undefined ? optionsOrId.slide : this.#getTargetSlideIndices()[0] || 1
      objectId = optionsOrId.objectId
      target = optionsOrId.targetId
    } else {
      slideIndex = this.#getTargetSlideIndices()[0] || 1
      objectId = optionsOrId
      target = targetId
    }
    this.#zOrderManager.moveObjectAfter(slideIndex, objectId, target, this.#slideManager)
    return this
  }

  /**
   * Reorders slide objects exactly as specified in the array.
   */
  reorderObjects(optionsOrOrder) {
    this.#assertLoaded()
    let slideIndex, order
    if (
      typeof optionsOrOrder === 'object' &&
      optionsOrOrder !== null &&
      Array.isArray(optionsOrOrder.order)
    ) {
      slideIndex =
        optionsOrOrder.slide !== undefined
          ? optionsOrOrder.slide
          : this.#getTargetSlideIndices()[0] || 1
      order = optionsOrOrder.order
    } else {
      slideIndex = this.#getTargetSlideIndices()[0] || 1
      order = optionsOrOrder
    }
    this.#zOrderManager.reorderObjects(slideIndex, order, this.#slideManager)
    return this
  }

  /**
   * Gets the ordered metadata of all objects on the slide.
   */
  getObjectOrder(slideIndex) {
    this.#assertLoaded()
    const targetIdx = slideIndex !== undefined ? slideIndex : this.#getTargetSlideIndices()[0] || 1
    return this.#zOrderManager.getObjectOrder(targetIdx, this.#slideManager)
  }

  /**
   * Applies bulk template configurations for slide elements stacking layers.
   */
  applyZOrder(slideOrConfigs, configsOption) {
    this.#assertLoaded()
    let slideIndex, configs
    if (Array.isArray(slideOrConfigs)) {
      slideIndex = this.#getTargetSlideIndices()[0] || 1
      configs = slideOrConfigs
    } else {
      slideIndex = slideOrConfigs
      configs = configsOption
    }
    this.#zOrderManager.applyZOrder(slideIndex, configs, this.#slideManager)
    return this
  }

  /**
   * Retrieves the info of the top-most object on the slide.
   */
  getTopMostObject(slideIndex) {
    this.#assertLoaded()
    const targetIdx = slideIndex !== undefined ? slideIndex : this.#getTargetSlideIndices()[0] || 1
    return this.#zOrderManager.getTopMostObject(targetIdx, this.#slideManager)
  }

  /**
   * Retrieves the info of the bottom-most object on the slide.
   */
  getBottomMostObject(slideIndex) {
    this.#assertLoaded()
    const targetIdx = slideIndex !== undefined ? slideIndex : this.#getTargetSlideIndices()[0] || 1
    return this.#zOrderManager.getBottomMostObject(targetIdx, this.#slideManager)
  }

  /**
   * Swaps stacking positions of two slide objects.
   */
  swapObjects(slideIndexOrId1, id1OrId2, id2) {
    this.#assertLoaded()
    let slideIndex, objectId1, objectId2
    if (id2 !== undefined) {
      slideIndex = slideIndexOrId1
      objectId1 = id1OrId2
      objectId2 = id2
    } else {
      slideIndex = this.#getTargetSlideIndices()[0] || 1
      objectId1 = slideIndexOrId1
      objectId2 = id1OrId2
    }
    this.#zOrderManager.swapObjects(slideIndex, objectId1, objectId2, this.#slideManager)
    return this
  }

  /**
   * Sorts stacking order using a custom comparison function.
   */
  sortObjects(slideIndexOrCompareFn, compareFnOption) {
    this.#assertLoaded()
    let slideIndex, compareFn
    if (typeof slideIndexOrCompareFn === 'function') {
      slideIndex = this.#getTargetSlideIndices()[0] || 1
      compareFn = slideIndexOrCompareFn
    } else {
      slideIndex = slideIndexOrCompareFn
      compareFn = compareFnOption
    }
    this.#zOrderManager.sortObjects(slideIndex, compareFn, this.#slideManager)
    return this
  }

  /**
   * Cleans up and normalizes stacking order consistency.
   */
  normalizeZOrder(slideIndex) {
    this.#assertLoaded()
    const targetIdx = slideIndex !== undefined ? slideIndex : this.#getTargetSlideIndices()[0] || 1
    this.#zOrderManager.normalizeZOrder(targetIdx, this.#slideManager)
    return this
  }
}

module.exports = { PPTXTemplater }
