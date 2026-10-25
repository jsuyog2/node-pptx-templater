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
const { CommentManager } = require('../managers/CommentManager.js')
const { ChartManager } = require('../elements/ChartManager.js')
const { ChartCacheGenerator } = require('../elements/charts/ChartCacheGenerator.js')
const { TableManager } = require('../elements/TableManager.js')
const { HyperlinkManager } = require('../managers/HyperlinkManager.js')
const { MediaManager } = require('../managers/MediaManager.js')
const { RelationshipManager } = require('../managers/RelationshipManager.js')
const { ShapeManager } = require('../elements/ShapeManager.js')
const { ImageManager } = require('../elements/ImageManager.js')
const { TextManager } = require('../elements/TextManager.js')
const { ZOrderManager } = require('../spatial/ZOrderManager.js')
const { ValidationEngine } = require('./ValidationEngine.js')
const { OutputWriter } = require('./OutputWriter.js')
const { TemplateEngine } = require('./TemplateEngine.js')
const { createLogger, setGlobalLogLevel } = require('../utils/logger.js')
const { PPTXError, ChartNotFoundError, TableNotFoundError } = require('../utils/errors.js')
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
   * @type {CommentManager}
   */
  #commentManager

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
    this.#commentManager = new CommentManager(
      this.#xmlParser,
      this.#relationshipManager,
      this.#contentTypesManager
    )
    this.#chartManager = new ChartManager(this.#xmlParser, this.#contentTypesManager)
    this.#shapeManager = new ShapeManager(this.#xmlParser)
    this.#tableManager = new TableManager(this.#xmlParser, this.#shapeManager)
    this.#hyperlinkManager = new HyperlinkManager(this.#xmlParser, this.#relationshipManager)
    this.#mediaManager = new MediaManager(this.#contentTypesManager)
    this.#imageManager = new ImageManager(this.#xmlParser)
    this.#textManager = new TextManager(this.#xmlParser)
    this.#templateEngine = new TemplateEngine(this.#xmlParser)
    this.#zOrderManager = new ZOrderManager(this.#xmlParser)
    this.#outputWriter = new OutputWriter(
      this.#zipManager,
      this.#contentTypesManager,
      this.#relationshipManager
    )

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
   * const buffer = fs.promises.readFile('./template.pptx');
   * const ppt = await PPTXTemplater.load(buffer);
   */
  static async load(source, options = {}) {
    if (options.logLevel) {
      setGlobalLogLevel(options.logLevel)
    }
    const engine = new PPTXTemplater()
    try {
      const { PackageLoader } = require('./PackageLoader.js')
      const vfs = await PackageLoader.load(source, options)
      await engine.#initialize(vfs)
      return engine
    } catch (err) {
      if (err instanceof PPTXError) throw err
      throw new PPTXError(err.message, err)
    }
  }

  /**
   * Creates a minimal valid OpenXML blank presentation deck without template files.
   * Supports custom aspect ratios ('16:9', '4:3', '16:10') and custom pixel/EMU dimensions.
   *
   * @static
   * @param {Object} [options={}] - Creation options (aspectRatio, width, height, unit).
   * @returns {Promise<PPTXTemplater>} Initialized engine instance with blank deck.
   *
   * @example
   * const ppt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' });
   */
  static async createBlank(options = {}) {
    if (options.logLevel) {
      setGlobalLogLevel(options.logLevel)
    }
    const engine = new PPTXTemplater()
    await engine.#initializeBlank(options)
    return engine
  }

  /**
   * Sets the global log level for all PPTXTemplater logger instances.
   * This is equivalent to setting the PPTX_LOG_LEVEL environment variable,
   * but works at runtime without restarting the process.
   *
   * @static
   * @param {'verbose'|'debug'|'info'|'warn'|'error'|'silent'} level - Log level.
   * @returns {void}
   *
   * @example
   * PPTXTemplater.setLogLevel('debug');  // Enable full debug output
   * PPTXTemplater.setLogLevel('silent'); // Suppress all output
   */
  static setLogLevel(level) {
    setGlobalLogLevel(level)
  }

  /**
   * Preloads a PPTX template into an in-memory cache for fast repeated generation.
   * Call this once at startup; subsequent `fromCache()` calls read from memory
   * with zero disk I/O.
   *
   * @static
   * @param {string|Buffer|Object} source - File path, Buffer, or folder object.
   * @param {Object} [options={}] - Options (e.g. { key: string, logLevel: string }).
   * @returns {Promise<Map>} The cache Map keyed by file paths.
   *
   * @example
   * await PPTXTemplater.preload('./template.pptx');
   * // Later, in request handlers:
   * const ppt = await PPTXTemplater.fromCache('./template.pptx');
   */
  static async preload(source, options = {}) {
    if (options && options.logLevel) {
      setGlobalLogLevel(options.logLevel)
    }
    let key = options && options.key ? options.key : source
    if (Buffer.isBuffer(key)) {
      const crypto = require('crypto')
      key = crypto.createHash('sha256').update(key).digest('hex')
    } else if (typeof key === 'object' && key !== null) {
      key = JSON.stringify(key)
    }

    if (PPTXTemplater.#templateCache.has(key)) {
      return PPTXTemplater.#templateCache.get(key)
    }

    const zipManager = new ZipManager()
    await zipManager.load(source, options)
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

  /**
   * Creates an engine instance from a previously preloaded cache entry.
   * Falls back to preloading if the source has not been cached yet.
   *
   * @static
   * @param {string|Buffer|Object} source - Same source used with `preload()`.
   * @returns {Promise<PPTXTemplater>} Initialized engine from cache.
   *
   * @example
   * await PPTXTemplater.preload('./template.pptx');
   * const ppt = await PPTXTemplater.fromCache('./template.pptx');
   * // ppt is ready — no disk I/O on the load
   */
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

  /**
   * Clears all preloaded templates from the in-memory cache.
   * Call this to free memory or force templates to be reloaded from disk.
   *
   * @static
   * @returns {void}
   *
   * @example
   * PPTXTemplater.clearCache(); // Force fresh reload on next fromCache()
   */
  static clearCache() {
    PPTXTemplater.#templateCache.clear()
  }

  /**
   * Enables internal performance profiling.
   * After calling this, use `getPerformanceMetrics()` to read timing data.
   *
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.enablePerformanceProfile();
   * // ... do work ...
   * console.log(ppt.getPerformanceMetrics());
   */
  enablePerformanceProfile() {
    this.#profiler.enabled = true
    this.#profiler.startTime = performance.now()
    return this
  }

  /**
   * Enables debug-level logging for this session.
   * Shortcut for `PPTXTemplater.setLogLevel('debug')`.
   *
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * const ppt = await PPTXTemplater.load('./template.pptx');
   * ppt.enableDebug(); // Shows all debug output
   */
  enableDebug() {
    setGlobalLogLevel('debug')
    return this
  }

  /**
   * Returns performance metrics collected since `enablePerformanceProfile()` was called.
   * Includes timing for template load, XML parse, chart update, image update,
   * ZIP generation, total elapsed time, and memory usage.
   *
   * @returns {Object} Metrics object with `templateLoadMs`, `parseMs`, `chartUpdateMs`,
   *   `imageUpdateMs`, `zipGenerationMs`, `totalMs`, `memoryUsedMB`.
   *
   * @example
   * ppt.enablePerformanceProfile();
   * await ppt.updateChart('chart1', data);
   * const metrics = ppt.getPerformanceMetrics();
   * console.log(`Total: ${metrics.totalMs}ms, Memory: ${metrics.memoryUsedMB}MB`);
   */
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

    if (!(await fs.pathExists(resolvedPptx))) {
      throw new PPTXError(`Source PPTX file not found: ${pptxPath}`)
    }

    if (await fs.pathExists(resolvedOut)) {
      const stats = await fs.stat(resolvedOut)
      if (stats.isFile()) {
        throw new PPTXError(`Destination is a file: ${outputPath}`)
      }
      const files = await fs.readdir(resolvedOut)
      if (files.length > 0) {
        if (!options.overwrite) {
          throw new PPTXError(
            `Destination directory "${outputPath}" is not empty. Set overwrite: true to overwrite.`
          )
        } else {
          await fs.emptyDir(resolvedOut)
        }
      }
    } else {
      await fs.ensureDir(resolvedOut)
    }

    const engine = await PPTXTemplater.load(resolvedPptx)
    await engine.#zipManager.toFolder(resolvedOut)

    // Validation — async pathExists replaces pathExists
    const criticalParts = ['ppt/presentation.xml', 'ppt/slides', 'ppt/_rels', '[Content_Types].xml']

    for (const part of criticalParts) {
      const p = path.join(resolvedOut, part)
      if (!(await fs.pathExists(p))) {
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
   * @param {Object} [options={}] - Options (e.g. { compressionLevel }).
   * @returns {Promise<void>}
   */
  static async buildPptx(folderPath, pptxPath, options = {}) {
    const fs = require('fs-extra')
    const path = require('path')

    const resolvedFolder = path.resolve(folderPath)
    const resolvedPptx = path.resolve(pptxPath)

    if (!(await fs.pathExists(resolvedFolder))) {
      throw new PPTXError(`Source folder not found: ${folderPath}`)
    }

    // Validation of the source folder — async pathExists replaces pathExists
    const criticalParts = ['ppt/presentation.xml', 'ppt/slides', 'ppt/_rels', '[Content_Types].xml']

    for (const part of criticalParts) {
      const p = path.join(resolvedFolder, part)
      if (!(await fs.pathExists(p))) {
        throw new PPTXError(`Source folder is missing critical OpenXML part: ${part}`)
      }
    }

    const engine = await PPTXTemplater.load(resolvedFolder)
    await fs.ensureDir(path.dirname(resolvedPptx))
    await engine.saveToFile(resolvedPptx, options)
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

    // Initialize comment manager
    this.#commentManager.initialize(this.#zipManager)

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

    // Initialize comment manager
    this.#commentManager.initialize(this.#zipManager)

    const t2 = performance.now()
    this.#profiler.parseMs = t2 - t1

    this.#loaded = true
  }

  /**
   * Initializes a blank PPTX structure from embedded template XML.
   * @private
   * @param {Object} [options={}]
   */
  async #initializeBlank(options = {}) {
    const { createBlankPresentationVfs } = require('../templates/blankPresentation.js')
    const vfs = createBlankPresentationVfs(options)
    await this.#initialize(vfs)
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
      this.#textManager.replaceText(
        slideIndex,
        replacements,
        this.#slideManager,
        this.#templateEngine
      )
    }

    logger.debug(`Replaced placeholder(s) in ${targetIndices.length} slide(s)`)
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
   * Retrieves all hyperlinks (external URLs, inter-slide links, navigation actions) on targeted slide.
   *
   * @param {number} [slideIndex] - 1-based slide index. If omitted, uses active slide.
   * @returns {Array<Object>} Array of hyperlink descriptor objects.
   *
   * @example
   * const links = ppt.getHyperlinks(1);
   * // [
   * //   { type: 'external', url: 'https://example.com', text: 'Visit site', rId: 'rId2' },
   * //   { type: 'slide', targetSlide: 3, shapeName: 'NextBtn' }
   * // ]
   */
  getHyperlinks(slideIndex) {
    this.#assertLoaded()
    const targetIdx = slideIndex !== undefined ? slideIndex : this.#getTargetSlideIndices()[0] || 1
    return this.#hyperlinkManager.getHyperlinks(
      targetIdx,
      this.#slideManager,
      this.#relationshipManager
    )
  }

  /**
   * Removes a hyperlink from text or a shape on the targeted slide(s).
   *
   * @param {string|Object} textOrOptions - Text string or options object { text, shapeName, slide }.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.removeHyperlink('Click here');
   * ppt.removeHyperlink({ text: 'Read more', slide: 2 });
   * ppt.removeHyperlink({ shapeName: 'Button 1' });
   */
  removeHyperlink(textOrOptions) {
    this.#assertLoaded()
    let slides = []

    if (
      typeof textOrOptions === 'object' &&
      textOrOptions !== null &&
      textOrOptions.slide !== undefined
    ) {
      slides = [textOrOptions.slide]
    } else {
      slides = this.#getTargetSlideIndices()
    }

    for (const slideIndex of slides) {
      this.#hyperlinkManager.removeHyperlink(
        slideIndex,
        textOrOptions,
        this.#slideManager,
        this.#relationshipManager
      )
    }

    return this
  }

  /**
   * Adds a new slide to the presentation.
   * Automatically generates required XML and relationship entries.
   *
   * @param {NewSlideOptions} [options={}] - Slide definition.
   * @param {string} [options.title] - Slide title text.
   * @param {string} [options.layout] - Layout name to use (default: 'blank').
   * @param {SlideElement[]} [options.elements] - Elements to add to the slide.
   * @param {number} [position] - Optional 1-based target slide position.
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
  addSlide(options = {}, position) {
    this.#assertLoaded()
    this.#slideManager.addNewSlide(options, position, this.#relationshipManager, this.#mediaManager)
    logger.debug(`Added new slide: "${options.title || 'Untitled'}"`)
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
   * @param {number} [atPosition] - 1-based target insertion index.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async importSlideFrom(sourceEngine, slideRef, atPosition) {
    this.#assertLoaded()
    await this.#slideManager.importSlide(sourceEngine, slideRef, this.#mediaManager, atPosition)
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
      dimensions: this.#slideManager.getSlideSize(),
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
    logger.info('=== Relationship Graph ===')
    for (const file of files) {
      logger.info(`\n${file}:`)
      const rels = this.#relationshipManager.getRelationships(
        file.replace('_rels/', '').replace('.rels', '')
      )
      rels.forEach(r => logger.info(`  - ${r.id} [${r.type.split('/').pop()}] -> ${r.target}`))
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

    logger.info(`=== Slide ${slideIndex} Inspection ===`)
    logger.info(`Path: ${info.zipPath}`)
    logger.info(`ID: ${info.slideId}`)
    logger.info(`rId: ${info.relationshipId}`)
    logger.info(`Title: ${info.title}`)
    logger.info(`XML Size: ${xml.length} characters`)
    logger.info(`Relationships (${rels.length}):`)
    rels.forEach(r => logger.info(`  - ${r.id} [${r.type.split('/').pop()}] -> ${r.target}`))

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
    logger.info(`=== XML Inspection: ${xmlPath} ===`)
    if (!xml) {
      logger.info('(File not found or empty)')
    } else {
      logger.info(xml.substring(0, 1500) + (xml.length > 1500 ? '...\n[Truncated]' : ''))
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
    if (options.validate || options.strict) {
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
    }
    if (options.validateArchive) {
      await this.validateArchive()
    }

    const t0 = performance.now()
    await this.#outputWriter.saveToFile(filePath, this.#slideManager, this.#zipManager, options)
    this.#profiler.zipGenerationMs += performance.now() - t0

    logger.info(`Saved PPTX to ${filePath}`)
  }

  /**
   * Unpacks all presentation parts into an uncompressed directory tree.
   *
   * @param {string} destinationDir - Directory path to write extracted files into.
   * @param {Object} [options={}] - Options (e.g. { formatXml: true } for version control formatting).
   * @returns {Promise<void>}
   */
  async extractToXml(destinationDir, options = {}) {
    this.#assertLoaded()
    await this.#outputWriter.extractToXml(
      this.#slideManager,
      this.#zipManager,
      destinationDir,
      options
    )
  }

  /**
   * Packages and dumps the presentation into a compressed ZIP file on disk.
   *
   * @param {string} destinationPath - Path to target ZIP file.
   * @param {Object} [options={}] - Save/compression options.
   * @returns {Promise<void>}
   */
  async extractToZip(destinationPath, options = {}) {
    this.#assertLoaded()
    return this.#outputWriter.extractToZip(
      this.#slideManager,
      this.#zipManager,
      destinationPath,
      options
    )
  }

  /**
   * Emits a standard Web API Response object for Next.js Route Handlers, Remix, SvelteKit, and Edge workers.
   *
   * @param {Object} [options={}] - Response options (status, headers, filename).
   * @returns {Promise<Response>} Web API Response instance.
   */
  async toResponse(options = {}) {
    this.#assertLoaded()
    if (options.validateArchive) {
      await this.validateArchive()
    }
    return this.#outputWriter.toResponse(this.#slideManager, this.#zipManager, options)
  }

  /**
   * Sends the presentation directly through a Fastify reply object.
   *
   * @param {Object} reply - Fastify reply object.
   * @param {string} [filename='presentation.pptx'] - Attachment filename.
   * @param {Object} [options={}] - Header/response options.
   * @returns {Promise<void>}
   */
  async toFastify(reply, filename = 'presentation.pptx', options = {}) {
    this.#assertLoaded()
    if (options.validateArchive) {
      await this.validateArchive()
    }
    await this.#outputWriter.toFastify(
      this.#slideManager,
      this.#zipManager,
      reply,
      filename,
      options
    )
  }

  /**
   * Returns the PPTX as an ArrayBuffer.
   * @param {Object} [options]
   * @returns {Promise<ArrayBuffer>}
   */
  async toArrayBuffer(options = {}) {
    return this.#outputWriter.toArrayBuffer(this.#slideManager, this.#zipManager, options)
  }

  /**
   * Returns the PPTX as a Base64 string.
   * @param {Object} [options] - Pass { dataUri: true } to prepend the Data URI scheme.
   * @returns {Promise<string>}
   */
  async toBase64(options = {}) {
    return this.#outputWriter.toBase64(this.#slideManager, this.#zipManager, options)
  }

  /**
   * Streams the PPTX directly to an Express.js response object.
   * @param {Object} res - Express HTTP Response object.
   * @param {string} [filename='presentation.pptx']
   * @param {Object} [options]
   */
  async toExpress(res, filename = 'presentation.pptx', options = {}) {
    await this.#outputWriter.toExpress(this.#slideManager, this.#zipManager, res, filename, options)
  }

  /**
   * Converts the entire presentation package into a single Microsoft Office Flat OPC XML file (<pkg:package>) or string.
   *
   * @param {string|Object} [destinationPathOrOptions={}] - Path to output file, or serialization options.
   * @param {Object} [options={}] - Additional serialization options when destinationPath is provided.
   * @returns {Promise<string>} Complete Flat OPC XML string.
   *
   * @example
   * // Convert to string
   * const flatXml = await ppt.toFlatXml();
   *
   * // Convert and save directly to file
   * await ppt.toFlatXml('./output/presentation.xml');
   */
  async toFlatXml(destinationPathOrOptions = {}, options = {}) {
    this.#assertLoaded()
    return this.#outputWriter.toFlatXml(
      this.#slideManager,
      this.#zipManager,
      destinationPathOrOptions,
      options
    )
  }

  /**
   * Returns the PPTX presentation as a Node.js Buffer.
   *
   * @param {Object} [options] - Options passed to the ZIP serializer.
   * @returns {Promise<Buffer>} Generated presentation buffer.
   */
  async toBuffer(options = {}) {
    this.#assertLoaded()
    if (options.validateArchive) {
      await this.validateArchive()
    }

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
    if (options.validateArchive) {
      await this.validateArchive()
    }

    const t0 = performance.now()
    const stream = await this.#outputWriter.toStream(this.#slideManager, this.#zipManager, options)
    this.#profiler.zipGenerationMs += performance.now() - t0

    return stream
  }

  // === Slide Features ===
  /**
   * Duplicates an existing slide and inserts the copy at the specified position.
   *
   * @param {number} slideIndex - 1-based index of the slide to duplicate.
   * @param {number} [atPosition] - 1-based position to insert the copy. Defaults to end.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.duplicateSlide(1, 2); // Copy slide 1 and insert it as slide 2
   */
  duplicateSlide(slideIndex, atPosition) {
    this.#assertLoaded()
    const promise = this.#slideManager.duplicateSlide(
      slideIndex,
      atPosition,
      this.#relationshipManager,
      this.#mediaManager
    )
    this.#zipManager.addPendingPromise(promise)
    return promise.then(() => this)
  }

  /**
   * Moves a slide from one position to another within the presentation.
   *
   * @param {number} fromIndex - 1-based source slide index.
   * @param {number} toIndex - 1-based target slide index.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.moveSlide(5, 1); // Move slide 5 to position 1 (make it first)
   */
  moveSlide(fromIndex, toIndex) {
    this.#assertLoaded()
    this.#slideManager.moveSlide(fromIndex, toIndex)
    return this
  }

  /**
   * Returns an array of all slides in the presentation with their metadata.
   *
   * @returns {Array<{index: number, slideId: string, title: string, zipPath: string}>}
   *
   * @example
   * const slides = ppt.getSlides();
   * slides.forEach(s => console.log(`Slide ${s.index}: ${s.title}`));
   */
  getSlides() {
    this.#assertLoaded()
    return this.#slideManager.getSlides()
  }

  /**
   * Retrieves detailed metadata for a specific slide by 1-based index or reference.
   *
   * @param {number|string} slideRef - 1-based slide index or custom tag/ID.
   * @returns {Object} Normalized slide metadata object.
   *
   * @example
   * const slide = ppt.getSlide(1);
   * console.log(`Slide ${slide.index} (ID: ${slide.slideId})`);
   */
  getSlide(slideRef) {
    this.#assertLoaded()
    return this.#slideManager.getSlide(slideRef)
  }

  /**
   * Selects slides belonging to a section for subsequent operations.
   *
   * @param {string|number} sectionRef - Section name, GUID, or 1-based section index.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.useSection('Appendix').replaceText({ '{{CONFIDENTIAL}}': 'TRUE' });
   */
  useSection(sectionRef) {
    this.#assertLoaded()
    const section = this.#slideManager.getSection(sectionRef)
    if (!section) {
      throw new PPTXError(`Section "${sectionRef}" not found`)
    }
    this.#selectedSlides = [...section.slideIndices]
    logger.debug(`Selected slides in section "${section.name}": ${this.#selectedSlides.join(', ')}`)
    return this
  }

  /**
   * Adds a new section to the presentation.
   *
   * @param {string} name - Section title.
   * @param {number} [startSlideIndex=1] - 1-based start slide index. Defaults to 1.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.addSection('Executive Summary', 1);
   */
  addSection(name, startSlideIndex = 1) {
    this.#assertLoaded()
    this.#slideManager.addSection(name, startSlideIndex)
    return this
  }

  /**
   * Updates an existing section's name or metadata.
   *
   * @param {string|number} sectionIdentifier - Name, GUID, or 1-based section index.
   * @param {Object} options - Update options.
   * @param {string} [options.name] - New section title.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.updateSection('Summary', { name: 'Executive Overview' });
   */
  updateSection(sectionIdentifier, options = {}) {
    this.#assertLoaded()
    this.#slideManager.updateSection(sectionIdentifier, options)
    return this
  }

  /**
   * Removes a section from the presentation, optionally deleting its member slides.
   *
   * @param {string|number} sectionIdentifier - Name, GUID, or 1-based section index.
   * @param {boolean} [removeSlides=false] - Whether to delete all member slides.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.removeSection('Appendix', false); // Keep slides, remove section grouping
   */
  removeSection(sectionIdentifier, removeSlides = false) {
    this.#assertLoaded()
    this.#slideManager.removeSection(sectionIdentifier, removeSlides)
    return this
  }

  /**
   * Retrieves section metadata by name, GUID, or 1-based index.
   *
   * @param {string|number} sectionRef - Name, GUID, or 1-based section index.
   * @returns {Object|null} Section metadata or null if not found.
   *
   * @example
   * const section = ppt.getSection('Financials');
   * console.log(`Section "${section.name}" has ${section.slideIndices.length} slides`);
   */
  getSection(sectionRef) {
    this.#assertLoaded()
    return this.#slideManager.getSection(sectionRef)
  }

  /**
   * Hides one or more slides so they are skipped during slide show presentation.
   *
   * @param {...(number|string)} slideRefs - 1-based slide indices, IDs, or tags.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.hideSlide(2, 3);
   */
  hideSlide(...slideRefs) {
    this.#assertLoaded()
    const targets =
      slideRefs.length > 0
        ? slideRefs.flatMap(ref =>
            typeof ref === 'number' ? [ref] : this.#slideManager.resolveSlideRef(ref)
          )
        : this.#getTargetSlideIndices()

    for (const idx of targets) {
      this.#slideManager.hideSlide(idx)
    }
    return this
  }

  /**
   * Unhides one or more previously hidden slides.
   *
   * @param {...(number|string)} slideRefs - 1-based slide indices, IDs, or tags.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.unhideSlide(2);
   */
  unhideSlide(...slideRefs) {
    this.#assertLoaded()
    const targets =
      slideRefs.length > 0
        ? slideRefs.flatMap(ref =>
            typeof ref === 'number' ? [ref] : this.#slideManager.resolveSlideRef(ref)
          )
        : this.#getTargetSlideIndices()

    for (const idx of targets) {
      this.#slideManager.unhideSlide(idx)
    }
    return this
  }

  /**
   * Resets the slide canvas by purging user elements while preserving group non-visual properties and placeholders.
   *
   * @param {number|string} [slideRef] - 1-based slide index or tag. Defaults to active/selected slide.
   * @param {Object} [options={ preservePlaceholders: true }] - Reset options.
   * @param {boolean} [options.preservePlaceholders=true] - Whether layout placeholders are preserved.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.resetSlide(1, { preservePlaceholders: true });
   */
  resetSlide(slideRef, options = { preservePlaceholders: true }) {
    this.#assertLoaded()
    let targets = []
    if (slideRef !== undefined && (typeof slideRef === 'number' || typeof slideRef === 'string')) {
      targets =
        typeof slideRef === 'number' ? [slideRef] : this.#slideManager.resolveSlideRef(slideRef)
    } else {
      targets = this.#getTargetSlideIndices()
      if (typeof slideRef === 'object' && slideRef !== null) {
        options = slideRef
      }
    }

    for (const idx of targets) {
      this.#slideManager.resetSlide(idx, options)
    }
    return this
  }

  /**
   * Adds a collaborative review comment to slide(s).
   *
   * @param {string} text - Comment text.
   * @param {Object} [options={}] - Author and coordinate options.
   * @param {string} [options.author='Author'] - Reviewer name.
   * @param {string} [options.initials='A'] - Reviewer initials.
   * @param {number|string} [options.x] - Horizontal coordinate.
   * @param {number|string} [options.y] - Vertical coordinate.
   * @param {string} [options.dateTime] - ISO timestamp.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   *
   * @example
   * await ppt.useSlide(1).addComment('Please verify Q3 ARR data', { author: 'Sarah', x: '1in', y: '1in' });
   */
  async addComment(text, options = {}) {
    this.#assertLoaded()
    const targets = this.#getTargetSlideIndices()
    await this.#commentManager.addComment(this.#slideManager, targets, text, options)
    return this
  }

  /**
   * Removes a comment by index (1-based), text match, or author filter from slide(s).
   *
   * @param {number|string} commentIdentifier - 1-based comment index or comment text search query.
   * @param {Object} [options={}] - Removal options (e.g. { author: '...' }).
   * @returns {Promise<PPTXTemplater>} this (chainable)
   *
   * @example
   * await ppt.useSlide(1).removeComment(1); // Remove first comment
   * await ppt.useSlide(1).removeComment('verify Q3'); // Remove comment containing text
   */
  async removeComment(commentIdentifier, options = {}) {
    this.#assertLoaded()
    const targets = this.#getTargetSlideIndices()
    await this.#commentManager.removeComment(
      this.#slideManager,
      targets,
      commentIdentifier,
      options
    )
    return this
  }

  /**
   * Retrieves the layout target path for a slide.
   *
   * @param {number} [slideIndex] - 1-based slide index. Defaults to active slide.
   * @returns {string|null} Layout target path or null.
   */
  getSlideLayout(slideIndex) {
    this.#assertLoaded()
    const targetIndex = slideIndex !== undefined ? slideIndex : this.activeSlideIndex
    return this.#slideManager.getSlideLayout(targetIndex)
  }

  /**
   * Retrieves presentation slide dimensions in EMUs.
   *
   * @returns {{width: number, height: number}} Slide dimensions in EMUs.
   *
   * @example
   * const dimensions = ppt.getSlideDimensions();
   * console.log(`Presentation size: ${dimensions.width} x ${dimensions.height} EMUs`);
   */
  getSlideDimensions() {
    this.#assertLoaded()
    return this.#slideManager.getSlideSize()
  }

  // ==========================================================================
  // === Tables (Strict 30 Whitelisted Methods) ===============================
  // ==========================================================================

  /**
   * Adds a new DrawingML table to the active or specified slide.
   *
   * @param {number|string|Object} [slideIdOrOptions] - Target slide number/ID, or table configuration object.
   * @param {number} [rows] - Number of table rows.
   * @param {number} [columns] - Number of table columns.
   * @param {Array<Array<*>>} [data] - 2D matrix of row and cell values.
   * @param {Object} [options={}] - Geometry, styling, and positioning options (x, y, width, height, colWidths, rowHeights, firstRow, bandRow, styleId).
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.addTable(1, 3, 3, [['A', 'B', 'C'], ['1', '2', '3']], { x: 100, y: 150, width: 500, height: 200 });
   * ppt.useSlide(1).addTable(2, 2, [['X', 'Y'], ['Z', 'W']]);
   * ppt.addTable({ id: 'SummaryTable', rows: 4, columns: 3, data: matrix });
   */
  addTable(slideIdOrOptions, rows, columns, data = [], options = {}) {
    this.#assertLoaded()
    let targetSlideId
    let numRows = rows
    let numCols = columns
    let tableData = data
    let tableOptions = options

    if (
      typeof slideIdOrOptions === 'object' &&
      slideIdOrOptions !== null &&
      !Array.isArray(slideIdOrOptions)
    ) {
      const opt = slideIdOrOptions
      targetSlideId = opt.slideId !== undefined ? opt.slideId : opt.slide
      numRows = opt.rows !== undefined ? opt.rows : Array.isArray(opt.data) ? opt.data.length : 2
      numCols =
        opt.columns !== undefined
          ? opt.columns
          : opt.cols !== undefined
            ? opt.cols
            : Array.isArray(opt.data) && opt.data[0]
              ? opt.data[0].length
              : 2
      tableData = opt.data || opt.rowsData || []
      tableOptions = opt
    } else if (Array.isArray(slideIdOrOptions)) {
      tableData = slideIdOrOptions
      tableOptions = typeof rows === 'object' && rows !== null ? rows : {}
      numRows = tableData.length
      numCols = tableData[0] ? tableData[0].length : 2
    } else if (
      typeof slideIdOrOptions === 'number' &&
      typeof rows === 'number' &&
      typeof columns === 'number'
    ) {
      targetSlideId = slideIdOrOptions
    } else if (
      typeof rows === 'number' &&
      (typeof columns === 'number' || Array.isArray(columns))
    ) {
      numRows = slideIdOrOptions
      numCols = rows
      tableData = Array.isArray(columns) ? columns : data
      tableOptions = typeof data === 'object' && !Array.isArray(data) ? data : options
    } else if (typeof slideIdOrOptions === 'string' && typeof rows === 'object') {
      tableOptions = { ...rows, id: slideIdOrOptions }
      numRows =
        tableOptions.rows || (Array.isArray(tableOptions.data) ? tableOptions.data.length : 2)
      numCols =
        tableOptions.columns ||
        tableOptions.cols ||
        (Array.isArray(tableOptions.data) && tableOptions.data[0] ? tableOptions.data[0].length : 2)
      tableData = tableOptions.data || []
    }

    const targetIndices =
      targetSlideId !== undefined
        ? [
            typeof targetSlideId === 'number'
              ? targetSlideId
              : this.#slideManager.resolveSlideIndex(targetSlideId),
          ]
        : this.#getTargetSlideIndices()

    const sIdx = targetIndices.length > 0 ? targetIndices[0] : 1
    this.#tableManager.addTable(sIdx, numRows, numCols, tableData, tableOptions, this.#slideManager)
    return this
  }

  /**
   * Resolves target slide indices for table operations.
   * Prioritizes explicitly selected slides; falls back to locating the table across all slides.
   *
   * @private
   * @param {string} tableId - Table name or shape ID.
   * @returns {Array<number>} Slide indices to target.
   */
  #resolveTableTargets(tableId) {
    if (this.#selectedSlides.length > 0) {
      return this.#selectedSlides
    }
    if (tableId) {
      const matching = []
      const allIndices = this.#slideManager.getAllSlideIndices()
      for (const idx of allIndices) {
        if (this.#slideManager.getSlideTable(idx, tableId)) {
          matching.push(idx)
        }
      }
      if (matching.length > 0) {
        return matching
      }
    }
    return [this.#slideManager.getAllSlideIndices()[0] || 1]
  }

  /**
   * Updates table geometry, dimensions, position, properties, or styling.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {Object} options - Updates for position (x, y, width, height), colWidths, rowHeights, or properties.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.updateTable('SalesTable', { x: 200, y: 150, width: 600, height: 300, firstRow: true });
   */
  updateTable(tableId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.updateTable(tableId, options, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Replaces matrix cell data while preserving existing column widths, cell styling, and row heights.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {Array<Array<*>>} data - 2D matrix of cell values.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.updateTableData('SalesTable', [['Metric', 'Value'], ['Revenue', '$1.2M']]);
   */
  updateTableData(tableId, data = []) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.updateTableData(tableId, data, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Removes a table and its container frame from the slide cleanly.
   *
   * @param {string} tableId - Table name or shape ID.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.removeTable('OldTable');
   */
  removeTable(tableId) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.removeTable(tableId, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Discovers and retrieves metadata for all tables on the specified or active slides.
   *
   * @param {number|string} [slideId] - Optional slide number or ID.
   * @returns {Array<Object>} List of table descriptors.
   *
   * @example
   * const tables = ppt.getTables(1);
   */
  getTables(slideId) {
    this.#assertLoaded()
    if (slideId !== undefined) {
      return this.#tableManager.getTables(slideId, this.#slideManager)
    }
    const targetIndices = this.#getTargetSlideIndices()
    const all = []
    for (const idx of targetIndices) {
      all.push(...this.#tableManager.getTables(idx, this.#slideManager))
    }
    return all
  }

  /**
   * Retrieves detailed metadata and metrics for a specific table.
   *
   * @param {string} tableId - Table name or shape ID.
   * @returns {Object|null} Normalized table metadata object, or null if not found.
   *
   * @example
   * const table = ppt.getTable('SalesTable');
   */
  getTable(tableId) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      const found = this.#tableManager.getTable(tableId, idx, this.#slideManager)
      if (found) return found
    }
    return null
  }

  /**
   * Deep-clones an existing table, assigning a new identifier and applying overrides.
   *
   * @param {string} tableId - Source table name or shape ID.
   * @param {string} clonedTableId - Unique identifier for the cloned table.
   * @param {Object} [clonedOptions={}] - Positioning and styling overrides.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.cloneTable('TemplateTable', 'ClonedTable', { y: 350 });
   */
  cloneTable(tableId, clonedTableId, clonedOptions = {}) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.cloneTable(tableId, clonedTableId, clonedOptions, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Inserts or appends a row to a table with cell values.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number|Array<*>} [rowIndexOrData] - 0-based insertion row index, or row data array if appending.
   * @param {Array<*>} [data] - Row cell data if rowIndex was specified.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.addTableRow('SalesTable', ['North', '$120K', '8%']);
   * ppt.addTableRow('SalesTable', 1, ['East', '$100K', '5%']);
   */
  addTableRow(tableId, rowIndexOrData, data) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.addTableRow(tableId, rowIndexOrData, data, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Inserts or appends a column to a table with cell values.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number|Array<*>} [columnIndexOrData] - 0-based insertion column index, or column data array if appending.
   * @param {Array<*>} [data] - Column cell data if columnIndex was specified.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.addTableColumn('SalesTable', ['Q4', '$45K', '$50K']);
   * ppt.addTableColumn('SalesTable', 2, ['Target', '$40K', '$48K']);
   */
  addTableColumn(tableId, columnIndexOrData, data) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.addTableColumn(tableId, columnIndexOrData, data, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Removes a row from a table by 0-based row index.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index to remove.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.removeTableRow('SalesTable', 2);
   */
  removeTableRow(tableId, rowIndex) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.removeTableRow(tableId, rowIndex, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Removes a column from a table by 0-based column index.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} columnIndex - 0-based column index to remove.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.removeTableColumn('SalesTable', 1);
   */
  removeTableColumn(tableId, columnIndex) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.removeTableColumn(tableId, columnIndex, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Updates cell text values along target row rowIndex.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @param {Array<*>} data - Row cell values.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.updateRowData('SalesTable', 1, ['Updated Region', '$150K', '12%']);
   */
  updateRowData(tableId, rowIndex, data) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.updateRowData(tableId, rowIndex, data, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Updates cell text values down target column columnIndex.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} columnIndex - 0-based column index.
   * @param {Array<*>} data - Column cell values down each row.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.updateColumnData('SalesTable', 1, ['$100', '$200', '$300']);
   */
  updateColumnData(tableId, columnIndex, data) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.updateColumnData(tableId, columnIndex, data, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Updates text, typography, alignments, and styling of a single table cell.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} columnIndex - 0-based column index.
   * @param {string|Object} options - New cell value or formatting object.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.updateTableCell('SalesTable', 1, 2, '$9,800');
   * ppt.updateTableCell('SalesTable', 0, 0, { text: 'Title', bold: true, fill: '1E3A8A', color: 'FFFFFF' });
   */
  updateTableCell(tableId, rowIndex, columnIndex, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.updateTableCell(
        tableId,
        rowIndex,
        columnIndex,
        options,
        idx,
        this.#slideManager
      )
    }
    return this
  }

  /**
   * Clears text content across rowIndex while preserving cell formatting.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.clearRowData('SalesTable', 2);
   */
  clearRowData(tableId, rowIndex) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.clearRowData(tableId, rowIndex, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Clears text content down columnIndex while preserving cell formatting.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} columnIndex - 0-based column index.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.clearColumnData('SalesTable', 1);
   */
  clearColumnData(tableId, columnIndex) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.clearColumnData(tableId, columnIndex, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Purges text content from a single table cell while preserving underlying formatting.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} columnIndex - 0-based column index.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.clearCellData('SalesTable', 0, 0);
   */
  clearCellData(tableId, rowIndex, columnIndex) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.clearCellData(tableId, rowIndex, columnIndex, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Extracts raw string data from all cells in the table as a 2D array.
   *
   * @param {string} tableId - Table name or shape ID.
   * @returns {Array<Array<string>>} 2D array matrix of cell data.
   *
   * @example
   * const matrix = ppt.getTableData('SalesTable');
   */
  getTableData(tableId) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      try {
        return this.#tableManager.getTableData(tableId, idx, this.#slideManager)
      } catch (err) {
        if (!(err instanceof TableNotFoundError)) throw err
      }
    }
    return []
  }

  /**
   * Returns cell text array down columnIndex.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} columnIndex - 0-based column index.
   * @returns {Array<string>} Column cell values.
   *
   * @example
   * const values = ppt.getColumnData('SalesTable', 1);
   */
  getColumnData(tableId, columnIndex) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      try {
        return this.#tableManager.getColumnData(tableId, columnIndex, idx, this.#slideManager)
      } catch (err) {
        if (!(err instanceof TableNotFoundError)) throw err
      }
    }
    return []
  }

  /**
   * Returns cell text array along rowIndex.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @returns {Array<string>} Row cell values.
   *
   * @example
   * const values = ppt.getRowData('SalesTable', 0);
   */
  getRowData(tableId, rowIndex) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      try {
        return this.#tableManager.getRowData(tableId, rowIndex, idx, this.#slideManager)
      } catch (err) {
        if (!(err instanceof TableNotFoundError)) throw err
      }
    }
    return []
  }

  /**
   * Returns column metadata: index, width, text data, and formatting.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} columnIndex - 0-based column index.
   * @returns {Object|null} Column descriptor object.
   *
   * @example
   * const col = ppt.getColumn('SalesTable', 0);
   */
  getColumn(tableId, columnIndex) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      try {
        return this.#tableManager.getColumn(tableId, columnIndex, idx, this.#slideManager)
      } catch (err) {
        if (!(err instanceof TableNotFoundError)) throw err
      }
    }
    return null
  }

  /**
   * Returns row metadata: index, height, text data, and formatting.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @returns {Object|null} Row descriptor object.
   *
   * @example
   * const row = ppt.getRow('SalesTable', 0);
   */
  getRow(tableId, rowIndex) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      try {
        return this.#tableManager.getRow(tableId, rowIndex, idx, this.#slideManager)
      } catch (err) {
        if (!(err instanceof TableNotFoundError)) throw err
      }
    }
    return null
  }

  /**
   * Returns detailed cell metadata, bounds, borders, fill, and merge state.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} columnIndex - 0-based column index.
   * @returns {Object|null} Cell descriptor object.
   *
   * @example
   * const cell = ppt.getCell('SalesTable', 0, 0);
   */
  getCell(tableId, rowIndex, columnIndex) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      try {
        const found = this.#tableManager.getCell(
          tableId,
          rowIndex,
          columnIndex,
          idx,
          this.#slideManager
        )
        if (found) return found
      } catch (err) {
        if (!(err instanceof TableNotFoundError)) throw err
      }
    }
    return null
  }

  /**
   * Extracts raw string text from cell at (rowIndex, columnIndex).
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} columnIndex - 0-based column index.
   * @returns {string} Cell string value.
   *
   * @example
   * const text = ppt.getCellData('SalesTable', 0, 0);
   */
  getCellData(tableId, rowIndex, columnIndex) {
    this.#assertLoaded()
    const cell = this.getCell(tableId, rowIndex, columnIndex)
    return cell ? cell.text : ''
  }

  /**
   * Merges a rectangular region of table cells.
   *
   * @param {string|Object} tableIdOrOptions - Table ID string, or merge options object.
   * @param {number} [startRow] - 0-based start row.
   * @param {number} [startColumn] - 0-based start column.
   * @param {number} [endRow] - 0-based end row.
   * @param {number} [endColumn] - 0-based end column.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.mergeCells('SalesTable', 0, 0, 0, 2);
   */
  mergeCells(tableIdOrOptions, startRow, startColumn, endRow, endColumn) {
    this.#assertLoaded()
    let tableId = tableIdOrOptions
    let sRow = startRow
    let sCol = startColumn
    let eRow = endRow
    let eCol = endColumn

    if (typeof tableIdOrOptions === 'object' && tableIdOrOptions !== null) {
      const opt = tableIdOrOptions
      tableId = opt.tableId
      sRow = opt.startRow
      sCol = opt.startColumn !== undefined ? opt.startColumn : opt.startCol
      eRow = opt.endRow
      eCol = opt.endColumn !== undefined ? opt.endColumn : opt.endCol
    }

    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.mergeCells(tableId, sRow, sCol, eRow, eCol, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Splits a previously merged cell region back into independent cells.
   *
   * @param {string|Object} tableIdOrOptions - Table ID string, or split options object.
   * @param {number} [rowIndex] - 0-based row index.
   * @param {number} [columnIndex] - 0-based column index.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.splitCell('SalesTable', 0, 0);
   */
  splitCell(tableIdOrOptions, rowIndex, columnIndex) {
    this.#assertLoaded()
    let tableId = tableIdOrOptions
    let r = rowIndex
    let c = columnIndex

    if (typeof tableIdOrOptions === 'object' && tableIdOrOptions !== null) {
      const opt = tableIdOrOptions
      tableId = opt.tableId
      r = opt.rowIndex !== undefined ? opt.rowIndex : opt.row
      c = opt.columnIndex !== undefined ? opt.columnIndex : opt.col
    }

    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.splitCell(tableId, r, c, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Applies table-wide properties, alternating banding fills, and borders.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {Object} options - Table styling options (bandRow, bandCol, fill, border, etc.).
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.styleTable('SalesTable', { bandRow: true, oddFill: 'F8FAFC', evenFill: 'EDF2F7' });
   */
  styleTable(tableId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.styleTable(tableId, options, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Applies fills, typography, and borders down columnIndex.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} columnIndex - 0-based column index.
   * @param {Object} options - Column styling options.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.styleColumn('SalesTable', 1, { fill: 'E0E7FF', bold: true });
   */
  styleColumn(tableId, columnIndex, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.styleColumn(tableId, columnIndex, options, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Applies fills, typography, height, and borders across rowIndex.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @param {Object} options - Row styling options.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.styleRow('SalesTable', 0, { fill: '1E3A8A', color: 'FFFFFF', bold: true });
   */
  styleRow(tableId, rowIndex, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.styleRow(tableId, rowIndex, options, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Applies fills, alignments, margins, and borders to cell at (rowIndex, columnIndex).
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} columnIndex - 0-based column index.
   * @param {Object} options - Cell styling options.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.styleCell('SalesTable', 0, 0, { fill: '3B82F6', align: 'ctr', verticalAlign: 'ctr' });
   */
  styleCell(tableId, rowIndex, columnIndex, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#resolveTableTargets(tableId)
    for (const idx of targetIndices) {
      this.#tableManager.styleCell(tableId, rowIndex, columnIndex, options, idx, this.#slideManager)
    }
    return this
  }

  /**
   * Retrieves the bounding box, coordinates, and rotation of a shape or element on the slide.
   *
   * @param {string} objectIdOrName - Shape or element ID/name.
   * @returns {{x: number, y: number, width: number, height: number, rotation: number, flipH: boolean, flipV: boolean}} Bounding box details in points.
   */
  getBounds(objectIdOrName) {
    this.#assertLoaded()
    const slideIndex = this.#getTargetSlideIndices()[0] || 1
    const shape = this.#shapeManager.getShape(slideIndex, objectIdOrName, this.#slideManager)
    if (shape) {
      return {
        x: shape.x,
        y: shape.y,
        width: shape.width,
        height: shape.height,
        rotation: shape.rotation || 0,
        flipH: shape.flipH || false,
        flipV: shape.flipV || false,
      }
    }
    const obj = this.#slideManager.findDrawableObject(slideIndex, objectIdOrName)
    if (!obj) {
      throw new Error(`Object "${objectIdOrName}" not found on slide ${slideIndex}`)
    }
    const { getBounds: calcBounds, emuToPx } = require('../spatial/units.js')
    const rawBounds = calcBounds(obj.frame || obj.element)
    return {
      x: Math.round(emuToPx(rawBounds.x)),
      y: Math.round(emuToPx(rawBounds.y)),
      width: Math.round(emuToPx(rawBounds.width)),
      height: Math.round(emuToPx(rawBounds.height)),
      rotation: rawBounds.rotation || 0,
      flipH: rawBounds.flipH || false,
      flipV: rawBounds.flipV || false,
    }
  }

  /**
   * Sets the bounding box, coordinates, dimensions, and rotation of a shape or element.
   *
   * @param {string} objectIdOrName - Shape or element ID/name.
   * @param {Object} options - Coordinates and dimensions config `{ x, y, width, height, rotation, flipH, flipV }`.
   * @returns {PPTXTemplater} this (chainable)
   */
  setBounds(objectIdOrName, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const { resolveCoordinate } = require('../utils/coordinateUtils.js')
    const dims = this.#slideManager.getSlideSize()

    const emuOptions = {}
    if (options.x !== undefined) emuOptions.x = resolveCoordinate(options.x, dims.width, options)
    if (options.y !== undefined) emuOptions.y = resolveCoordinate(options.y, dims.height, options)
    if (options.width !== undefined)
      emuOptions.width = resolveCoordinate(options.width, dims.width, options)
    if (options.height !== undefined)
      emuOptions.height = resolveCoordinate(options.height, dims.height, options)
    if (options.rotation !== undefined) emuOptions.rotation = options.rotation
    if (options.flipH !== undefined) emuOptions.flipH = options.flipH
    if (options.flipV !== undefined) emuOptions.flipV = options.flipV

    for (const idx of targetIndices) {
      this.#shapeManager.updateShape(idx, objectIdOrName, emuOptions, this.#slideManager)
    }
    return this
  }

  /**
   * Moves a shape to a new coordinate.
   *
   * @param {string} objectIdOrName - Shape or element ID/name.
   * @param {Object} options - Position options `{ x, y }`.
   * @returns {PPTXTemplater} this (chainable)
   */
  move(objectIdOrName, options = {}) {
    return this.setBounds(objectIdOrName, { x: options.x, y: options.y })
  }

  /**
   * Checks if two elements overlap or collide on the active slide.
   *
   * @param {string} idA - First element ID or name.
   * @param {string} idB - Second element ID or name.
   * @returns {boolean} True if bounding boxes overlap.
   */
  checkCollision(idA, idB) {
    this.#assertLoaded()
    const boundsA = this.getBounds(idA)
    const boundsB = this.getBounds(idB)
    const { checkCollision } = require('../spatial/units.js')
    return checkCollision(boundsA, boundsB)
  }

  /**
   * Checks if one element is completely contained within another.
   *
   * @param {string} innerId - Contained element ID or name.
   * @param {string} outerId - Container element ID or name.
   * @returns {boolean} True if inner element is completely inside outer element.
   */
  isContained(innerId, outerId) {
    this.#assertLoaded()
    const innerBounds = this.getBounds(innerId)
    const outerBounds = this.getBounds(outerId)
    const { isContained } = require('../spatial/units.js')
    return isContained(innerBounds, outerBounds)
  }

  /**
   * Retrieves all elements on targeted slides, optionally filtered by type, name, or spatial bounds.
   *
   * @param {Object} [filter] - Filter options: `{ type, name, bounds, within, intersects, slideIndex }`.
   * @returns {Array<Object>} List of element descriptors.
   */
  getElements(filter = {}) {
    this.#assertLoaded()
    const targetIndices =
      filter.slideIndex !== undefined ? [filter.slideIndex] : this.#getTargetSlideIndices()
    const all = []
    for (const idx of targetIndices) {
      const elems = this.#slideManager.getSlideElements(idx, filter)
      for (const el of elems) {
        all.push({ ...el, slideIndex: idx })
      }
    }
    return all
  }

  /**
   * Finds slide elements matching a filter object or custom predicate function.
   *
   * @param {Object|Function} filterOrPredicate - Filter criteria or predicate function.
   * @returns {Array<Object>} List of matched element descriptors.
   */
  findElements(filterOrPredicate) {
    if (typeof filterOrPredicate === 'function') {
      const elements = this.getElements()
      return elements.filter(filterOrPredicate)
    }
    return this.getElements(filterOrPredicate)
  }

  /**
   * Alias for bringForward. Moves element one layer forward.
   */
  stepForward(optionsOrId) {
    return this.bringForward(optionsOrId)
  }

  /**
   * Alias for sendBackward. Moves element one layer backward.
   */
  stepBackward(optionsOrId) {
    return this.sendBackward(optionsOrId)
  }

  /**
   * Replaces a placeholder shape with an image, locking its bounding box.
   * @param {string} shapeIdOrName
   * @param {string|Buffer} sourcePathOrBuffer
   */
  async replaceShape(shapeIdOrName, sourcePathOrBuffer) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()

    for (const idx of targetIndices) {
      await this.#imageManager.replaceShapeWithImage(
        idx,
        shapeIdOrName,
        sourcePathOrBuffer,
        this.#slideManager,
        this.#mediaManager,
        this.#relationshipManager
      )
    }
    return this
  }

  /**
   * Resizes any drawable object (shape, textbox, picture, chart, table, SVG, icon, placeholder, group) on active slides.
   * Supports absolute dimensions (width/height), relative delta (increase/decrease), and percentage scaling (scale).
   *
   * @param {string} objectIdOrName - Object name or shape ID.
   * @param {Object} options - Resize options.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.resize("Logo", { width: 4 });
   * ppt.resize("Chart", { height: 3 });
   * ppt.resize("Rectangle", { width: 5, height: 2 });
   * ppt.resize("Image", { increase: { width: 1 } });
   * ppt.resize("Chart", { decrease: { height: 0.5 } });
   * ppt.resize("Image", { scale: 1.2 });
   * ppt.resize("Chart", { scale: { width: 1.1, height: 0.9 } });
   */
  resize(objectIdOrName, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#slideManager.resizeObject(idx, objectIdOrName, options, this.#tableManager)
    }
    return this
  }

  // === Chart Features ===
  /**
   * Adds a new chart to the presentation supporting all 17 chart types.
   * Supports both explicit slide targets `ppt.addChart(1, 'bar', data, options)`
   * and active slide context `ppt.useSlide(1).addChart('bar', data, options)`
   * or direct invocation `ppt.addChart('bar', data, options)`.
   *
   * @param {number|string} [slideIdOrType] - Slide number/ID or chart type if slideId is omitted.
   * @param {string|Object} [chartTypeOrData] - Chart type or data payload.
   * @param {Object} [dataOrOptions] - Chart data payload or options.
   * @param {Object} [options={}] - Layout and styling options (x, y, width, height, title, legend).
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.useSlide(1).addChart('bar', { categories: ['A', 'B'], series: [{ name: 'S', values: [1, 2] }] });
   * ppt.addChart('column', { categories: ['Q1', 'Q2'], series: [{ name: 'S', values: [10, 20] }] });
   * ppt.addChart(1, 'pie', { categories: ['X', 'Y'], series: [{ name: 'S', values: [50, 50] }] });
   */
  addChart(slideIdOrType, chartTypeOrData, dataOrOptions, options = {}) {
    this.#assertLoaded()
    let targetSlideIndex
    let chartType
    let chartData
    let chartOptions

    const isFirstArgChartType =
      typeof slideIdOrType === 'string' &&
      (ChartCacheGenerator.isSupportedType(slideIdOrType) ||
        (typeof chartTypeOrData === 'object' && chartTypeOrData !== null))

    if (isFirstArgChartType) {
      const targets = this.#getTargetSlideIndices()
      targetSlideIndex = targets.length > 0 ? targets[0] : 1
      chartType = slideIdOrType
      chartData = chartTypeOrData || {}
      chartOptions = dataOrOptions || {}
    } else if (
      typeof slideIdOrType === 'number' ||
      (typeof slideIdOrType === 'string' && typeof chartTypeOrData === 'string')
    ) {
      targetSlideIndex =
        typeof slideIdOrType === 'number'
          ? slideIdOrType
          : this.#slideManager.resolveSlideIndex(slideIdOrType)
      chartType = chartTypeOrData
      chartData = dataOrOptions || {}
      chartOptions = options || {}
    } else {
      const targets = this.#getTargetSlideIndices()
      targetSlideIndex = targets.length > 0 ? targets[0] : 1
      chartType = slideIdOrType
      chartData = chartTypeOrData || {}
      chartOptions = dataOrOptions || {}
    }

    this.#chartManager.addChart(
      targetSlideIndex,
      chartType,
      chartData,
      chartOptions,
      this.#slideManager,
      this.#relationshipManager,
      this.#contentTypesManager
    )
    return this
  }

  /**
   * Removes a chart by ID, name, or relationship ID from targeted slides.
   *
   * @param {string} chartId - Chart name, shape ID, or relationship ID.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.useSlide(1).removeChart('sales-chart');
   */
  removeChart(chartId) {
    this.#assertLoaded()
    let removed = false
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      try {
        this.#chartManager.removeChart(
          idx,
          chartId,
          this.#slideManager,
          this.#relationshipManager,
          this.#contentTypesManager
        )
        removed = true
      } catch (e) {
        // Continue
      }
    }
    if (!removed) {
      for (const idx of this.#slideManager.getAllSlideIndices()) {
        try {
          this.#chartManager.removeChart(
            idx,
            chartId,
            this.#slideManager,
            this.#relationshipManager,
            this.#contentTypesManager
          )
          break
        } catch (e) {
          // Continue
        }
      }
    }
    return this
  }

  /**
   * Clears all charts from the selected slide(s) while preserving layout shapes, tables, and text frames.
   * Supports polymorphic arguments:
   * - No arguments: Clears all charts on the active slide context (or all slides).
   * - Slide index (number or numeric string): Clears all charts on the target slide.
   * - Chart ID (string): Gracefully removes that specific chart without throwing errors.
   *
   * @param {number|string} [target] - Optional 1-based slide index, slide tag, or specific chart ID.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.useSlide(1).clearCharts();
   * ppt.clearCharts(2);
   * ppt.clearCharts('StandaloneBarChart');
   */
  clearCharts(target) {
    this.#assertLoaded()
    if (target === undefined) {
      const targetIndices = this.#getTargetSlideIndices()
      for (const idx of targetIndices) {
        this.#chartManager.clearCharts(
          idx,
          this.#slideManager,
          this.#relationshipManager,
          this.#contentTypesManager
        )
      }
      return this
    }

    if (typeof target === 'number') {
      this.#chartManager.clearCharts(
        target,
        this.#slideManager,
        this.#relationshipManager,
        this.#contentTypesManager
      )
      return this
    }

    if (typeof target === 'string') {
      const num = Number(target)
      if (Number.isInteger(num) && num >= 1 && num <= this.#slideManager.slideCount) {
        this.#chartManager.clearCharts(
          num,
          this.#slideManager,
          this.#relationshipManager,
          this.#contentTypesManager
        )
        return this
      }

      // Check if target matches an existing chart on the slide/deck
      try {
        const chart = this.getChart(target)
        if (chart) {
          return this.removeChart(target)
        }
      } catch (e) {
        // Chart not found
      }

      // Check if target is a slide tag in slideManager
      try {
        const slideIndices = this.#slideManager.resolveSlideRef(target)
        if (slideIndices && slideIndices.length > 0) {
          for (const idx of slideIndices) {
            this.#chartManager.clearCharts(
              idx,
              this.#slideManager,
              this.#relationshipManager,
              this.#contentTypesManager
            )
          }
          return this
        }
      } catch (e) {
        // Not a slide tag
      }

      return this.removeChart(target)
    }

    return this
  }

  /**
   * Retrieves normalized descriptor and metadata for a specific chart.
   *
   * @param {string} chartId - Chart name, shape ID, or relationship ID.
   * @returns {Object} Normalized chart descriptor.
   *
   * @example
   * const chart = ppt.getChart('sales-chart');
   * console.log(chart.type, chart.series, chart.coordinates);
   */
  getChart(chartId) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      try {
        const chart = this.#chartManager.getChart(
          idx,
          chartId,
          this.#slideManager,
          this.#relationshipManager
        )
        if (chart) return chart
      } catch (e) {}
    }
    for (const idx of this.#slideManager.getAllSlideIndices()) {
      try {
        const chart = this.#chartManager.getChart(
          idx,
          chartId,
          this.#slideManager,
          this.#relationshipManager
        )
        if (chart) return chart
      } catch (e) {}
    }
    throw new ChartNotFoundError(`Chart "${chartId}" not found in presentation`)
  }

  /**
   * Clones a chart, duplicating its XML part, relationship, and embedded Excel workbook.
   *
   * @param {string} chartId - Chart name or ID to clone.
   * @param {string} [clonedChartId] - Unique name for the cloned chart.
   * @param {Object} [clonedOptions={}] - Optional layout/coordinate overrides (x, y, width, height).
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.useSlide(1).cloneChart('chart1', 'chart1_copy', { x: '5in', y: '2in' });
   */
  cloneChart(chartId, clonedChartId, clonedOptions = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      try {
        this.#chartManager.cloneChart(
          idx,
          chartId,
          clonedChartId,
          clonedOptions,
          this.#slideManager,
          this.#relationshipManager,
          this.#contentTypesManager
        )
        return this
      } catch (e) {
        if (!e.message.includes('not found')) {
          throw e
        }
      }
    }
    throw new ChartNotFoundError(`Chart "${chartId}" not found for cloning`)
  }

  /**
   * Alias for `updateChart()`. Updates chart data for a named chart.
   *
   * @param {string} chartId - Chart name or shape ID.
   * @param {Object} data - Chart data object with `categories` and `series`.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.updateChartData('revenue-chart', {
   *   categories: ['Q1', 'Q2', 'Q3'],
   *   series: [{ name: 'Revenue', values: [100, 150, 200] }]
   * });
   */
  updateChartData(chartId, data) {
    return this.updateChart(chartId, data)
  }

  /**
   * Replaces a specific data series in a chart.
   *
   * @param {string} chartId - Chart name or shape ID.
   * @param {number} seriesIndex - 0-based index of the series to replace.
   * @param {Object} newSeriesData - New series data `{ name, values }`.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.replaceChartSeries('revenue-chart', 0, { name: 'Revenue', values: [100, 200, 300] });
   */
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

  /**
   * Updates only the title text of a chart.
   *
   * @param {string} chartId - Chart name or shape ID.
   * @param {string} title - New chart title.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.updateChartTitle('revenue-chart', 'Q2 2026 Revenue');
   */
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

  /**
   * Updates only the category labels (X-axis) of a chart, keeping values unchanged.
   *
   * @param {string} chartId - Chart name or shape ID.
   * @param {string[]} categories - Array of category label strings.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.updateChartCategories('revenue-chart', ['Jan', 'Feb', 'Mar', 'Apr']);
   */
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

  /**
   * Updates data labels for a specific chart series.
   * Supports custom arrays, label maps, template strings, and cell references.
   *
   * @param {string} chartId - Chart name or shape ID.
   * @param {Object} options - Data label options.
   * @param {number} [options.series=0] - 0-based series index.
   * @param {string[]} [options.labels] - Array of custom label strings.
   * @param {Object} [options.labelMap] - Map of `{ categoryValue: label }`.
   * @param {string} [options.template] - Template string with `{value}`, `{category}`, `{percentage}` tokens.
   * @param {string} [options.labelsFromCells] - Excel cell range (e.g. `'Sheet1!$C$2:$C$6'`).
   * @param {boolean} [options.showSeriesNameInBar] - Prepend series name to bar chart labels.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.updateDataLabels('revenue-chart', {
   *   series: 0,
   *   template: '{value} ({percentage}%)',
   * });
   */
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

  /**
   * Retrieves the current data labels configuration for a specific chart series.
   *
   * @param {string} chartId - Chart name or shape ID.
   * @param {Object} [options] - Options.
   * @param {number} [options.series=0] - 0-based series index.
   * @returns {Promise<Object>} Current data label settings.
   *
   * @example
   * const labels = await ppt.getDataLabels('revenue-chart', { series: 0 });
   * console.log(labels.showValue, labels.position);
   */
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

  /**
   * Validates the data labels configuration for a chart series against the chart XML.
   *
   * @param {string} chartId - Chart name or shape ID.
   * @param {Object} [options] - Options (same as `updateDataLabels`).
   * @returns {Promise<{valid: boolean, errors: string[]}>}
   *
   * @example
   * const result = await ppt.validateDataLabels('revenue-chart');
   * if (!result.valid) console.error(result.errors);
   */
  async validateDataLabels(chartId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    if (targetIndices.length === 0) return { valid: true, errors: [] }
    const idx = targetIndices[0]
    return ValidationEngine.validateDataLabels(this, idx, chartId, options)
  }

  /**
   * Validates chart data labels across all series, including cell reference checks.
   *
   * @param {string} chartId - Chart name or shape ID.
   * @param {Object} [options] - Options.
   * @returns {Promise<{valid: boolean, errors: string[], warnings: string[]}>}
   *
   * @example
   * const result = await ppt.validateChartLabels('revenue-chart');
   */
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

  /**
   * Validates series name labels (the labels showing series names inside bar chart bars).
   *
   * @param {string} chartId - Chart name or shape ID.
   * @param {Object} [options] - Options.
   * @returns {Promise<{valid: boolean, errors: string[], warnings: string[]}>}
   *
   * @example
   * const result = await ppt.validateSeriesNameLabels('bar-chart');
   */
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

  /**
   * Returns an array of normalized chart metadata found on the targeted slide(s) or specified slide.
   *
   * @param {number|string} [slideId] - Optional slide number, ID, or tag.
   * @returns {Array<Object>} Chart descriptor objects.
   *
   * @example
   * const charts = ppt.getCharts();
   * const slide1Charts = ppt.getCharts(1);
   */
  getCharts(slideId) {
    this.#assertLoaded()
    const targetIndices =
      slideId !== undefined
        ? [typeof slideId === 'number' ? slideId : this.#slideManager.resolveSlideIndex(slideId)]
        : this.#getTargetSlideIndices()
    const charts = []
    for (const idx of targetIndices) {
      charts.push(
        ...this.#chartManager.getCharts(idx, this.#slideManager, this.#relationshipManager)
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
  // === Text & Placeholders Features (Strict 18-Method Whitelist) ===

  /**
   * Adds a new formatted text box shape to targeted slides or adds text to an existing shape.
   * Supports pixel, EMU, or percentage string coordinates, plus the 31-property style engine.
   *
   * @param {string} textContent - The text string to insert.
   * @param {Object} [options={}] - Coordinate and style formatting options.
   * @returns {this} The chainable presentation engine instance.
   *
   * @example
   * ppt.addText('Executive Summary', { x: '10%', y: 150, fontSize: 24, bold: true, color: '#1A73E8' });
   */
  addText(textContent, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.addText(idx, textContent, options, this.#slideManager)
    }
    return this
  }

  /**
   * Replaces the entire text content of a text box or shape identified by ID, name, or tag.
   * Clears existing text runs while preserving paragraph layout and character styling.
   *
   * @param {string|number} tagOrId - Text box ID (e.g. '2'), shape name (e.g. 'Title 1'), or semantic tag.
   * @param {string|number} newText - New text content to replace the full text box.
   * @param {Object} [options={}] - Optional formatting options or slide override ({ slide: 1 }).
   * @returns {this} The chainable presentation engine instance.
   *
   * @example
   * ppt.replaceTextByTag('2', 'New Executive Summary');
   * ppt.replaceTextByTag('TitleBox', 'First Line\nSecond Line', { bold: true });
   */
  replaceTextByTag(tagOrId, newText, options = {}) {
    this.#assertLoaded()
    const targetIndices =
      options.slide !== undefined ? [Number(options.slide)] : this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.replaceTextByTag(
        idx,
        tagOrId,
        newText,
        options,
        this.#slideManager,
        this.#templateEngine
      )
    }
    return this
  }

  /**
   * Updates/replaces text and typography in an existing shape or placeholder non-destructively.
   *
   * @param {string} tagOrId - Target shape name, ID, or template tag.
   * @param {string|Object} newText - New text content, or options object containing text.
   * @param {Object} [options={}] - Typography and style override options.
   * @returns {this} The chainable presentation engine instance.
   *
   * @example
   * ppt.updateText('TitleBox', 'Updated Header', { fontSize: 28, color: '#3B82F6' });
   */
  updateText(tagOrId, newText, options = {}) {
    this.#assertLoaded()
    let actualText = newText
    let actualOptions = options
    if (
      typeof newText === 'object' &&
      newText !== null &&
      (options === undefined || Object.keys(options).length === 0)
    ) {
      actualOptions = newText
      actualText = actualOptions.text
    }
    const targetIndices =
      actualOptions.slide !== undefined
        ? [Number(actualOptions.slide)]
        : this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.updateText(
        idx,
        tagOrId,
        actualText,
        actualOptions,
        this.#slideManager,
        this.#templateEngine
      )
    }
    return this
  }

  /**
   * Adds a new bulleted or numbered list shape to targeted slides.
   *
   * @param {Array<string|Object>} items - List items or nested hierarchy.
   * @param {Object} [options={}] - Coordinates, styling, and bullet formatting.
   * @returns {this} The chainable presentation engine instance.
   *
   * @example
   * ppt.addLists(['Point 1', 'Point 2'], { x: 50, y: 100, ordered: true });
   */
  addLists(items, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.addLists(idx, items, options, this.#slideManager)
    }
    return this
  }

  /**
   * Retrieves list items from a shape or text box by name or placeholder tag.
   *
   * @param {string} tag - Shape name/ID or placeholder tag.
   * @returns {Array<string|Object>} Nested list structure of items.
   *
   * @example
   * const items = ppt.getList('BulletBox');
   */
  getList(tag) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const idx = targetIndices.length > 0 ? targetIndices[0] : 1
    return this.#textManager.getList(idx, tag, this.#slideManager)
  }

  /**
   * Searches for all occurrences of a text query across targeted slides with AST run healing.
   *
   * @param {string} query - Text to search for.
   * @param {Object} [options={}] - Search options.
   * @returns {Array<{slideIndex: number, shapeId: string, shapeName: string, text: string, match: string}>} Array of matches.
   *
   * @example
   * const matches = ppt.findText('Revenue');
   */
  findText(query, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const matches = []
    for (const idx of targetIndices) {
      matches.push(...this.#textManager.findText(idx, query, this.#slideManager, options))
    }
    return matches
  }

  /**
   * Adds or merges style properties into a target shape, text box, or placeholder.
   *
   * @param {string} target - Target shape name/ID or placeholder tag.
   * @param {Object} style - Style properties to apply.
   * @returns {this} The chainable presentation engine instance.
   *
   * @example
   * ppt.addStyle('Title', { bold: true, color: '#FF0000', fontSize: 24 });
   */
  addStyle(target, style) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.addStyle(idx, target, style, this.#slideManager)
    }
    return this
  }

  /**
   * Removes specified style properties from a target shape, text box, or placeholder.
   *
   * @param {string} target - Target shape name/ID or placeholder tag.
   * @param {Array<string>|Object} styleProperties - Properties to remove.
   * @returns {this} The chainable presentation engine instance.
   *
   * @example
   * ppt.removeStyle('Title', ['bold', 'italic', 'color']);
   */
  removeStyle(target, styleProperties) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.removeStyle(idx, target, styleProperties, this.#slideManager)
    }
    return this
  }

  /**
   * Updates style properties on a target shape, text box, or placeholder.
   *
   * @param {string} target - Target shape name/ID or placeholder tag.
   * @param {Object} style - Updated style properties.
   * @returns {this} The chainable presentation engine instance.
   *
   * @example
   * ppt.updateStyle('Title', { fontSize: 32, alignment: 'center' });
   */
  updateStyle(target, style) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.updateStyle(idx, target, style, this.#slideManager)
    }
    return this
  }

  /**
   * Retrieves full concatenated plain text content from a target shape or placeholder.
   *
   * @param {string} tag - Target shape name/ID or placeholder tag.
   * @returns {string} Plain text content.
   *
   * @example
   * const text = ppt.getText('Title');
   */
  getText(tag) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const idx = targetIndices.length > 0 ? targetIndices[0] : 1
    return this.#textManager.getText(idx, tag, this.#slideManager)
  }

  /**
   * Retrieves all list structures found across shapes on targeted slides.
   *
   * @returns {Array<{shapeId: string, shapeName: string, items: Array}>} Array of discovered lists.
   *
   * @example
   * const lists = ppt.getLists();
   */
  getLists() {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const allLists = []
    for (const idx of targetIndices) {
      allLists.push(...this.#textManager.getLists(idx, this.#slideManager))
    }
    return allLists
  }

  /**
   * Finds lists across targeted slides matching a query (by shape name, ID, or item contents).
   *
   * @param {string} query - Query string to match.
   * @returns {Array<{shapeId: string, shapeName: string, items: Array}>} Array of matching lists.
   *
   * @example
   * const results = ppt.findLists('Action Items');
   */
  findLists(query) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const matches = []
    for (const idx of targetIndices) {
      matches.push(...this.#textManager.findLists(idx, query, this.#slideManager))
    }
    return matches
  }

  /**
   * Removes a text shape or text container matching the specified tag.
   *
   * @param {string} tag - Shape name/ID or placeholder tag.
   * @returns {this} The chainable presentation engine instance.
   *
   * @example
   * ppt.removeText('OldTitle');
   */
  removeText(tag) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.removeText(idx, tag, this.#slideManager)
    }
    return this
  }

  /**
   * Removes a list shape matching the specified tag from targeted slides.
   *
   * @param {string} tag - Shape name/ID or placeholder tag.
   * @returns {this} The chainable presentation engine instance.
   *
   * @example
   * ppt.removeList('DeprecatedList');
   */
  removeList(tag) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.removeList(idx, tag, this.#slideManager)
    }
    return this
  }

  /**
   * Clears text content inside a shape while keeping the shape container, geometry, and coordinates intact.
   *
   * @param {string} tag - Shape name/ID or placeholder tag.
   * @returns {this} The chainable presentation engine instance.
   *
   * @example
   * ppt.clearText('PlaceholderBox');
   */
  clearText(tag) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.clearText(idx, tag, this.#slideManager)
    }
    return this
  }

  /**
   * Clears list items from a shape while retaining shape position, dimensions, and bullet properties.
   *
   * @param {string} tag - Shape name/ID or placeholder tag.
   * @returns {this} The chainable presentation engine instance.
   *
   * @example
   * ppt.clearList('ActionList');
   */
  clearList(tag) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.clearList(idx, tag, this.#slideManager)
    }
    return this
  }

  /**
   * Replaces list items inside a shape or table cell matching tag.
   *
   * @param {string} tag - Shape name/ID or placeholder tag.
   * @param {Array<string|Object>|Object} newItems - New list items or configuration.
   * @param {Object} [options={}] - Additional list options.
   * @returns {this} The chainable presentation engine instance.
   *
   * @example
   * ppt.replaceList('Title', ['Item 1', 'Item 2'], { ordered: true });
   */
  replaceList(tag, newItems, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.replaceList(
        idx,
        tag,
        newItems,
        options,
        this.#slideManager,
        this.#templateEngine
      )
    }
    return this
  }

  /**
   * Returns all text elements (paragraphs) across targeted slides.
   *
   * @returns {Array<{slideIndex: number, shapeId: string, shapeName: string, paragraphIndex: number, text: string}>}
   *
   * @example
   * const elements = ppt.getTextElements();
   */
  getTextElements() {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const elements = []
    for (const idx of targetIndices) {
      elements.push(...this.#textManager.getTextElements(idx, this.#slideManager))
    }
    return elements
  }

  // === Shapes & Geometry Features (Strict 10-Method Whitelist) ===

  /**
   * Adds a new shape to the slide.
   *
   * @param {number|string|Object} slideIdOrTypeOrOptions - Target slide index/ID, shape type string, or options object.
   * @param {string|Object} [typeOrOptions] - Shape type string or options object.
   * @param {Object} [options={}] - Configuration options.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.addShape(1, 'rect', { x: 100, y: 100, width: 200, height: 100, fill: '#2563EB' });
   */
  addShape(slideIdOrTypeOrOptions, typeOrOptions, options = {}) {
    this.#assertLoaded()
    let targetSlideIndices
    let shapeType
    let shapeOptions = {}

    if (
      typeof slideIdOrTypeOrOptions === 'number' ||
      (typeof slideIdOrTypeOrOptions === 'string' &&
        /^\d+$/.test(slideIdOrTypeOrOptions) &&
        typeof typeOrOptions === 'string')
    ) {
      targetSlideIndices = [parseInt(slideIdOrTypeOrOptions, 10)]
      shapeType = typeOrOptions
      shapeOptions = { ...options }
    } else if (typeof slideIdOrTypeOrOptions === 'string') {
      targetSlideIndices = this.#getTargetSlideIndices()
      shapeType = slideIdOrTypeOrOptions
      shapeOptions =
        typeof typeOrOptions === 'object' && typeOrOptions !== null ? { ...typeOrOptions } : {}
    } else if (typeof slideIdOrTypeOrOptions === 'object' && slideIdOrTypeOrOptions !== null) {
      targetSlideIndices = this.#getTargetSlideIndices()
      shapeOptions = { ...slideIdOrTypeOrOptions }
      shapeType = shapeOptions.type
    } else {
      targetSlideIndices = this.#getTargetSlideIndices()
    }

    if (shapeType && !shapeOptions.type) {
      shapeOptions.type = shapeType
    }

    for (const idx of targetSlideIndices) {
      let finalOptions = shapeOptions
      if (shapeOptions.alignToCell) {
        finalOptions = this.#resolveAlignToCell(idx, shapeOptions)
      }
      this.#shapeManager.addShape(idx, shapeType, finalOptions, this.#slideManager)
    }
    return this
  }

  /**
   * Updates spatial coordinates, fills, borders, text, or visibility of an existing shape by ID or name.
   *
   * @param {string} targetId - Shape ID or name to update.
   * @param {Object} options - Configuration properties to update.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.updateShape('banner_box', { fill: '#10B981', width: 300 });
   */
  updateShape(targetId, options = {}, maybeSlideIndex) {
    this.#assertLoaded()
    let slideIndex = maybeSlideIndex
    let actualTargetId = targetId
    let actualOptions = options
    if (typeof targetId === 'number' && typeof options === 'string') {
      slideIndex = targetId
      actualTargetId = options
      actualOptions = maybeSlideIndex || {}
    }
    if (actualOptions && actualOptions.slide !== undefined) {
      slideIndex = actualOptions.slide
    }

    const targetIndices = slideIndex !== undefined ? [slideIndex] : this.#getTargetSlideIndices()
    if (this.#selectedSlides.length > 0 || slideIndex !== undefined) {
      for (const idx of targetIndices) {
        let resolvedOptions = actualOptions
        if (actualOptions.alignToCell) {
          resolvedOptions = this.#resolveAlignToCell(idx, actualOptions, actualTargetId)
        }
        this.#shapeManager.updateShape(idx, actualTargetId, resolvedOptions, this.#slideManager)
      }
    } else {
      let found = false
      for (const idx of targetIndices) {
        if (this.#slideManager.getSlideShape(idx, actualTargetId)) {
          found = true
          let resolvedOptions = actualOptions
          if (actualOptions.alignToCell) {
            resolvedOptions = this.#resolveAlignToCell(idx, actualOptions, actualTargetId)
          }
          this.#shapeManager.updateShape(idx, actualTargetId, resolvedOptions, this.#slideManager)
        }
      }
      if (!found) {
        this.#shapeManager.updateShape(
          targetIndices[0] || 1,
          actualTargetId,
          actualOptions,
          this.#slideManager
        )
      }
    }
    return this
  }

  /**
   * Removes a shape from the targeted slide(s) by shape ID or name.
   *
   * @param {string} shapeId - Shape ID or name to remove.
   * @param {number} [maybeSlideIndex] - Optional 1-based slide index.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.removeShape('draft_watermark');
   */
  removeShape(shapeId, maybeSlideIndex) {
    this.#assertLoaded()
    const slideIndex = maybeSlideIndex
    const targetIndices = slideIndex !== undefined ? [slideIndex] : this.#getTargetSlideIndices()
    if (this.#selectedSlides.length > 0 || slideIndex !== undefined) {
      for (const idx of targetIndices) {
        this.#shapeManager.removeShape(idx, shapeId, this.#slideManager)
      }
    } else {
      let found = false
      for (const idx of targetIndices) {
        if (this.#slideManager.getSlideShape(idx, shapeId)) {
          found = true
          this.#shapeManager.removeShape(idx, shapeId, this.#slideManager)
        }
      }
      if (!found) {
        this.#shapeManager.removeShape(targetIndices[0] || 1, shapeId, this.#slideManager)
      }
    }
    return this
  }

  /**
   * Clears all user-created shapes from targeted slide(s) while preserving layout placeholders.
   *
   * @param {number} [maybeSlideIndex] - Optional 1-based slide index.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.clearShapes();
   */
  clearShapes(maybeSlideIndex) {
    this.#assertLoaded()
    const targetIndices =
      maybeSlideIndex !== undefined ? [maybeSlideIndex] : this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#shapeManager.clearShapes(idx, this.#slideManager)
    }
    return this
  }

  /**
   * Returns metadata for all shapes on the targeted slide(s).
   *
   * @param {number} [maybeSlideIndex] - Optional 1-based slide index.
   * @returns {Array<Object>} List of shape descriptors.
   *
   * @example
   * const shapes = ppt.getShapes();
   */
  getShapes(maybeSlideIndex) {
    this.#assertLoaded()
    const targetIndices =
      maybeSlideIndex !== undefined ? [maybeSlideIndex] : this.#getTargetSlideIndices()
    const shapes = []
    for (const idx of targetIndices) {
      shapes.push(...this.#shapeManager.getShapes(idx, this.#slideManager))
    }
    return shapes
  }

  /**
   * Discovers and retrieves details of an existing shape on targeted slide(s).
   *
   * @param {string} shapeId - Shape ID or name.
   * @param {number} [maybeSlideIndex] - Optional 1-based slide index.
   * @returns {Object|null} Shape details object, or null if not found.
   *
   * @example
   * const shape = ppt.getShape('HeroCard');
   */
  getShape(shapeId, maybeSlideIndex) {
    this.#assertLoaded()
    const targetIndices =
      maybeSlideIndex !== undefined ? [maybeSlideIndex] : this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      const shape = this.#shapeManager.getShape(idx, shapeId, this.#slideManager)
      if (shape) return shape
    }
    return null
  }

  /**
   * Duplicates an existing shape with a new ID and optional overrides.
   *
   * @param {string} shapeId - Source shape name or ID.
   * @param {string} clonedShapeId - Identifier for cloned shape.
   * @param {Object} [clonedOptions={}] - Positioning and styling overrides.
   * @param {number} [maybeSlideIndex] - Optional 1-based slide index.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.cloneShape('logo', 'logo-copy', { x: 200, y: 50 });
   */
  cloneShape(shapeId, clonedShapeId, clonedOptions = {}, maybeSlideIndex) {
    this.#assertLoaded()
    const slideIndex = maybeSlideIndex || clonedOptions.slide
    const targetIndices = slideIndex !== undefined ? [slideIndex] : this.#getTargetSlideIndices()
    if (this.#selectedSlides.length > 0 || slideIndex !== undefined) {
      for (const idx of targetIndices) {
        this.#shapeManager.cloneShape(
          idx,
          shapeId,
          clonedShapeId,
          clonedOptions,
          this.#slideManager
        )
      }
    } else {
      let found = false
      for (const idx of targetIndices) {
        if (this.#slideManager.getSlideShape(idx, shapeId)) {
          found = true
          this.#shapeManager.cloneShape(
            idx,
            shapeId,
            clonedShapeId,
            clonedOptions,
            this.#slideManager
          )
        }
      }
      if (!found) {
        this.#shapeManager.cloneShape(
          targetIndices[0] || 1,
          shapeId,
          clonedShapeId,
          clonedOptions,
          this.#slideManager
        )
      }
    }
    return this
  }

  /**
   * Inspects a group shape and returns normalized child shape descriptors.
   *
   * @param {string} groupId - Group shape ID or name.
   * @param {number} [maybeSlideIndex] - Optional 1-based slide index.
   * @returns {Array<Object>} Child shape descriptors.
   *
   * @example
   * const children = ppt.getGroupShapes('HeaderGroup');
   */
  getGroupShapes(groupId, maybeSlideIndex) {
    this.#assertLoaded()
    const targetIndices =
      maybeSlideIndex !== undefined ? [maybeSlideIndex] : this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      const children = this.#shapeManager.getGroupShapes(idx, groupId, this.#slideManager)
      if (children && children.length > 0) return children
    }
    return []
  }

  /**
   * Groups multiple shapes into a composite group container (<p:grpSp>).
   *
   * @param {string[]} shapeIds - Array of shape IDs or names to group.
   * @param {Object} [groupOptions={}] - Group configuration options (e.g. name, id).
   * @param {number} [maybeSlideIndex] - Optional 1-based slide index.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.groupShapes(['shape1', 'shape2'], { name: 'CardGroup' });
   */
  groupShapes(shapeIds, groupOptions = {}, maybeSlideIndex) {
    this.#assertLoaded()
    const slideIndex = maybeSlideIndex || groupOptions.slide
    const targetIndices = slideIndex !== undefined ? [slideIndex] : this.#getTargetSlideIndices()
    if (this.#selectedSlides.length > 0 || slideIndex !== undefined) {
      for (const idx of targetIndices) {
        this.#shapeManager.groupShapes(idx, shapeIds, groupOptions, this.#slideManager)
      }
    } else {
      let found = false
      const arr = Array.isArray(shapeIds) ? shapeIds : [shapeIds]
      for (const idx of targetIndices) {
        const allExist = arr.every(id => Boolean(this.#slideManager.getSlideShape(idx, id)))
        if (allExist) {
          found = true
          this.#shapeManager.groupShapes(idx, shapeIds, groupOptions, this.#slideManager)
          break
        }
      }
      if (!found) {
        this.#shapeManager.groupShapes(
          targetIndices[0] || 1,
          shapeIds,
          groupOptions,
          this.#slideManager
        )
      }
    }
    return this
  }

  /**
   * Ungroups a composite group container, reprojecting child transforms back to absolute slide coordinates.
   *
   * @param {string} groupId - Group shape ID or name to dissolve.
   * @param {number} [maybeSlideIndex] - Optional 1-based slide index.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.ungroupShapes('CardGroup');
   */
  ungroupShapes(groupId, maybeSlideIndex) {
    this.#assertLoaded()
    const slideIndex = maybeSlideIndex
    const targetIndices = slideIndex !== undefined ? [slideIndex] : this.#getTargetSlideIndices()
    if (this.#selectedSlides.length > 0 || slideIndex !== undefined) {
      for (const idx of targetIndices) {
        this.#shapeManager.ungroupShapes(idx, groupId, this.#slideManager)
      }
    } else {
      let found = false
      for (const idx of targetIndices) {
        if (this.#slideManager.getSlideShape(idx, groupId)) {
          found = true
          this.#shapeManager.ungroupShapes(idx, groupId, this.#slideManager)
          break
        }
      }
      if (!found) {
        this.#shapeManager.ungroupShapes(targetIndices[0] || 1, groupId, this.#slideManager)
      }
    }
    return this
  }

  // === Images & Media ===

  /**
   * Helper to normalize arguments for addImage, addVideo, and addAudio.
   * Supports (slideId, source, options) and (source, options) using active slide context.
   * @private
   */
  #resolveMediaAddArgs(arg1, arg2, arg3) {
    const isOptionsObj = obj =>
      typeof obj === 'object' &&
      obj !== null &&
      !Buffer.isBuffer(obj) &&
      !(obj instanceof Uint8Array) &&
      typeof obj.pipe !== 'function'

    // If arg2 is an options object, it's (source, options) using active slide context
    if (isOptionsObj(arg2)) {
      const activeIdx = this.#getTargetSlideIndices()[0] || 1
      return { slideIndex: activeIdx, source: arg1, options: arg2 }
    }

    // If 3 arguments provided: (slideId, source, options)
    if (arg3 !== undefined) {
      const slideInfo = this.#slideManager.resolveSlideInfo(arg1)
      const slideIndex = slideInfo ? slideInfo.index : typeof arg1 === 'number' ? arg1 : 1
      return { slideIndex, source: arg2, options: arg3 || {} }
    }

    // If 2 arguments provided: (slideId, source)
    if (arg2 !== undefined) {
      const slideInfo = this.#slideManager.resolveSlideInfo(arg1)
      const slideIndex = slideInfo ? slideInfo.index : typeof arg1 === 'number' ? arg1 : 1
      return { slideIndex, source: arg2, options: {} }
    }

    // If 1 argument provided: (source)
    const activeIdx = this.#getTargetSlideIndices()[0] || 1
    return { slideIndex: activeIdx, source: arg1, options: {} }
  }

  /**
   * Adds an image to a slide.
   *
   * @param {number|string|Buffer|Uint8Array} slideIdOrSource - 1-based slide index, tag, or image source.
   * @param {string|Buffer|Uint8Array|Object} [sourceOrOptions] - Image source or options when slide is inferred.
   * @param {Object} [options={}] - Bounds, coordinates, crop, opacity, rotation, and border styling.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async addImage(slideIdOrSource, sourceOrOptions, options = {}) {
    this.#assertLoaded()
    const t0 = performance.now()
    const {
      slideIndex,
      source,
      options: opts,
    } = this.#resolveMediaAddArgs(slideIdOrSource, sourceOrOptions, options)
    await this.#imageManager.addImage(
      slideIndex,
      source,
      opts,
      this.#slideManager,
      this.#mediaManager,
      this.#relationshipManager
    )
    this.#profiler.imageUpdateMs += performance.now() - t0
    return this
  }

  /**
   * Alias for addImage.
   *
   * @param {number|string|Buffer|Uint8Array} slideIdOrSource - 1-based slide index, tag, or image source.
   * @param {string|Buffer|Uint8Array|Object} [sourceOrOptions] - Image source or options.
   * @param {Object} [options={}] - Options object.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async AddImage(...args) {
    return this.addImage(...args)
  }

  /**
   * Updates an existing image on the targeted slide(s).
   *
   * @param {string|number} imageId - Shape ID or name of the image to update.
   * @param {Object|string|Buffer} [options={}] - Updated bounds, crop, opacity, rotation, or new source.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async updateImage(imageId, options = {}) {
    this.#assertLoaded()
    const t0 = performance.now()
    const targetIndices = this.#getTargetSlideIndices()
    const opts =
      typeof options === 'string' || Buffer.isBuffer(options) || options instanceof Uint8Array
        ? { source: options }
        : { ...options }

    let updated = false
    for (const idx of targetIndices) {
      updated = await this.#imageManager.updateImage(
        idx,
        imageId,
        opts,
        this.#slideManager,
        this.#mediaManager,
        this.#relationshipManager
      )
      if (updated) break
    }
    if (!updated) {
      for (const idx of this.#slideManager.getAllSlideIndices()) {
        if (targetIndices.includes(idx)) continue
        updated = await this.#imageManager.updateImage(
          idx,
          imageId,
          opts,
          this.#slideManager,
          this.#mediaManager,
          this.#relationshipManager
        )
        if (updated) break
      }
    }
    this.#profiler.imageUpdateMs += performance.now() - t0
    return this
  }

  /**
   * Alias for updateImage.
   *
   * @param {string|number} imageId - Shape ID or name of the image to update.
   * @param {Object|string|Buffer} [options={}] - Updated bounds, crop, opacity, rotation, or source.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async UpdateImage(...args) {
    return this.updateImage(...args)
  }

  /**
   * Removes an image from the targeted slide(s) by shape ID, name, or relationship ID.
   *
   * @param {string|number} imageId - Shape ID, name, or relationship ID of the image to remove.
   * @returns {PPTXTemplater} this (chainable)
   */
  removeImage(imageId) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    let removed = false
    for (const idx of targetIndices) {
      removed = this.#imageManager.removeImage(
        idx,
        imageId,
        this.#slideManager,
        this.#relationshipManager,
        this.#mediaManager
      )
      if (removed) break
    }
    if (!removed) {
      for (const idx of this.#slideManager.getAllSlideIndices()) {
        if (targetIndices.includes(idx)) continue
        removed = this.#imageManager.removeImage(
          idx,
          imageId,
          this.#slideManager,
          this.#relationshipManager,
          this.#mediaManager
        )
        if (removed) break
      }
    }
    return this
  }

  /**
   * Alias for removeImage.
   *
   * @param {string|number} imageId - Shape ID, name, or relationship ID of the image to remove.
   * @returns {PPTXTemplater} this (chainable)
   */
  RemoveImage(...args) {
    return this.removeImage(...args)
  }

  /**
   * Adds a video element to a slide.
   *
   * @param {number|string|Buffer|Uint8Array} slideIdOrSource - 1-based slide index, tag, or video source.
   * @param {string|Buffer|Uint8Array|Object} [sourceOrOptions] - Video source or options when slide is inferred.
   * @param {Object} [options={}] - Video positioning, poster image, and playback options.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async addVideo(slideIdOrSource, sourceOrOptions, options = {}) {
    this.#assertLoaded()
    const {
      slideIndex,
      source,
      options: opts,
    } = this.#resolveMediaAddArgs(slideIdOrSource, sourceOrOptions, options)
    await this.#mediaManager.addVideo(
      slideIndex,
      source,
      opts,
      this.#slideManager,
      this.#relationshipManager,
      this.#xmlParser
    )
    return this
  }

  /**
   * Alias for addVideo.
   *
   * @param {number|string|Buffer|Uint8Array} slideIdOrSource - 1-based slide index, tag, or video source.
   * @param {string|Buffer|Uint8Array|Object} [sourceOrOptions] - Video source or options.
   * @param {Object} [options={}] - Options object.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async AddVideo(...args) {
    return this.addVideo(...args)
  }

  /**
   * Updates an existing video shape and playback properties.
   *
   * @param {string|number} videoId - Shape ID or name of the video.
   * @param {Object} [options={}] - Updated bounds, poster, loop, mute, timing, or new source.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async updateVideo(videoId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    let updated = null
    for (const idx of targetIndices) {
      updated = await this.#mediaManager.updateVideo(
        idx,
        videoId,
        options,
        this.#slideManager,
        this.#relationshipManager
      )
      if (updated) break
    }
    if (!updated) {
      for (const idx of this.#slideManager.getAllSlideIndices()) {
        if (targetIndices.includes(idx)) continue
        updated = await this.#mediaManager.updateVideo(
          idx,
          videoId,
          options,
          this.#slideManager,
          this.#relationshipManager
        )
        if (updated) break
      }
    }
    return this
  }

  /**
   * Alias for updateVideo.
   *
   * @param {string|number} videoId - Shape ID or name of the video.
   * @param {Object} [options={}] - Updated options.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async UpdateVideo(...args) {
    return this.updateVideo(...args)
  }

  /**
   * Removes a video shape and associated timing/relationships.
   *
   * @param {string|number} videoId - Shape ID or name of the video to remove.
   * @returns {PPTXTemplater} this (chainable)
   */
  removeVideo(videoId) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    let removed = false
    for (const idx of targetIndices) {
      removed = this.#mediaManager.removeVideo(
        idx,
        videoId,
        this.#slideManager,
        this.#relationshipManager
      )
      if (removed) break
    }
    if (!removed) {
      for (const idx of this.#slideManager.getAllSlideIndices()) {
        if (targetIndices.includes(idx)) continue
        removed = this.#mediaManager.removeVideo(
          idx,
          videoId,
          this.#slideManager,
          this.#relationshipManager
        )
        if (removed) break
      }
    }
    return this
  }

  /**
   * Alias for removeVideo.
   *
   * @param {string|number} videoId - Shape ID or name of the video to remove.
   * @returns {PPTXTemplater} this (chainable)
   */
  RemoveVideo(...args) {
    return this.removeVideo(...args)
  }

  /**
   * Adds an audio element to a slide.
   *
   * @param {number|string|Buffer|Uint8Array} slideIdOrSource - 1-based slide index, tag, or audio source.
   * @param {string|Buffer|Uint8Array|Object} [sourceOrOptions] - Audio source or options when slide is inferred.
   * @param {Object} [options={}] - Positioning, volume, loop, autoplay, or visibility options.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async addAudio(slideIdOrSource, sourceOrOptions, options = {}) {
    this.#assertLoaded()
    const {
      slideIndex,
      source,
      options: opts,
    } = this.#resolveMediaAddArgs(slideIdOrSource, sourceOrOptions, options)
    await this.#mediaManager.addAudio(
      slideIndex,
      source,
      opts,
      this.#slideManager,
      this.#relationshipManager,
      this.#xmlParser
    )
    return this
  }

  /**
   * Alias for addAudio.
   *
   * @param {number|string|Buffer|Uint8Array} slideIdOrSource - 1-based slide index, tag, or audio source.
   * @param {string|Buffer|Uint8Array|Object} [sourceOrOptions] - Audio source or options.
   * @param {Object} [options={}] - Options object.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async AddAudio(...args) {
    return this.addAudio(...args)
  }

  /**
   * Updates an existing audio element and playback properties.
   *
   * @param {string|number} audioId - Shape ID or name of the audio.
   * @param {Object} [options={}] - Updated bounds, volume, loop, autoplay, visibility settings, or source.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async updateAudio(audioId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    let updated = null
    for (const idx of targetIndices) {
      updated = await this.#mediaManager.updateAudio(
        idx,
        audioId,
        options,
        this.#slideManager,
        this.#relationshipManager
      )
      if (updated) break
    }
    if (!updated) {
      for (const idx of this.#slideManager.getAllSlideIndices()) {
        if (targetIndices.includes(idx)) continue
        updated = await this.#mediaManager.updateAudio(
          idx,
          audioId,
          options,
          this.#slideManager,
          this.#relationshipManager
        )
        if (updated) break
      }
    }
    return this
  }

  /**
   * Alias for updateAudio.
   *
   * @param {string|number} audioId - Shape ID or name of the audio.
   * @param {Object} [options={}] - Updated options.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   */
  async UpdateAudio(...args) {
    return this.updateAudio(...args)
  }

  /**
   * Removes an audio shape and associated timing/relationships.
   *
   * @param {string|number} audioId - Shape ID or name of the audio to remove.
   * @returns {PPTXTemplater} this (chainable)
   */
  removeAudio(audioId) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    let removed = false
    for (const idx of targetIndices) {
      removed = this.#mediaManager.removeAudio(
        idx,
        audioId,
        this.#slideManager,
        this.#relationshipManager
      )
      if (removed) break
    }
    if (!removed) {
      for (const idx of this.#slideManager.getAllSlideIndices()) {
        if (targetIndices.includes(idx)) continue
        removed = this.#mediaManager.removeAudio(
          idx,
          audioId,
          this.#slideManager,
          this.#relationshipManager
        )
        if (removed) break
      }
    }
    return this
  }

  /**
   * Alias for removeAudio.
   *
   * @param {string|number} audioId - Shape ID or name of the audio to remove.
   * @returns {PPTXTemplater} this (chainable)
   */
  RemoveAudio(...args) {
    return this.removeAudio(...args)
  }

  /**
   * Retrieves metadata for all media items (images, videos, audio) on a slide or presentation.
   *
   * @param {number|string} [slideId] - Optional slide index, ID, or tag.
   * @returns {Array<Object>} List of media metadata items.
   */
  getMedias(slideId) {
    this.#assertLoaded()
    if (slideId !== undefined) {
      const slideInfo = this.#slideManager.resolveSlideInfo(slideId)
      const slideIndex = slideInfo ? slideInfo.index : typeof slideId === 'number' ? slideId : 1
      return this.#mediaManager.getMedias(slideIndex, this.#slideManager, this.#relationshipManager)
    }
    const targetIndices = this.#getTargetSlideIndices()
    const all = []
    for (const idx of targetIndices) {
      all.push(...this.#mediaManager.getMedias(idx, this.#slideManager, this.#relationshipManager))
    }
    return all
  }

  /**
   * Alias for getMedias.
   *
   * @param {number|string} [slideId] - Optional slide index, ID, or tag.
   * @returns {Array<Object>} List of media metadata items.
   */
  GetMedias(...args) {
    return this.getMedias(...args)
  }

  /**
   * Retrieves detailed metadata for a specific media element by ID across slides.
   *
   * @param {string|number} mediaId - Shape ID, name, or relationship ID of the media item.
   * @returns {Object|null} Media metadata object or null if not found.
   */
  getMedia(mediaId) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      const found = this.#mediaManager.getMedia(
        mediaId,
        idx,
        this.#slideManager,
        this.#relationshipManager
      )
      if (found) return found
    }
    for (const idx of this.#slideManager.getAllSlideIndices()) {
      if (targetIndices.includes(idx)) continue
      const found = this.#mediaManager.getMedia(
        mediaId,
        idx,
        this.#slideManager,
        this.#relationshipManager
      )
      if (found) return found
    }
    return null
  }

  /**
   * Alias for getMedia.
   *
   * @param {string|number} mediaId - Shape ID, name, or relationship ID of the media item.
   * @returns {Object|null} Media metadata object or null if not found.
   */
  GetMedia(...args) {
    return this.getMedia(...args)
  }

  /**
   * Clones an existing media item and applies property overrides while deduplicating assets.
   *
   * @param {string|number} mediaId - Shape ID or name of the media item to clone.
   * @param {string|number} [clonedMediaId] - Name or ID for the cloned shape.
   * @param {Object} [clonedOptions={}] - Property overrides for position, dimensions, or styling.
   * @returns {PPTXTemplater} this (chainable)
   */
  cloneMedia(mediaId, clonedMediaId, clonedOptions = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    let cloned = null
    for (const idx of targetIndices) {
      cloned = this.#mediaManager.cloneMedia(
        mediaId,
        clonedMediaId,
        clonedOptions,
        idx,
        this.#slideManager,
        this.#relationshipManager,
        this.#xmlParser
      )
      if (cloned) break
    }
    if (!cloned) {
      for (const idx of this.#slideManager.getAllSlideIndices()) {
        if (targetIndices.includes(idx)) continue
        cloned = this.#mediaManager.cloneMedia(
          mediaId,
          clonedMediaId,
          clonedOptions,
          idx,
          this.#slideManager,
          this.#relationshipManager,
          this.#xmlParser
        )
        if (cloned) break
      }
    }
    return this
  }

  /**
   * Alias for cloneMedia.
   *
   * @param {string|number} mediaId - Shape ID or name of the media item to clone.
   * @param {string|number} [clonedMediaId] - Name or ID for the cloned shape.
   * @param {Object} [clonedOptions={}] - Property overrides for position, dimensions, or styling.
   * @returns {PPTXTemplater} this (chainable)
   */
  CloneMedia(...args) {
    return this.cloneMedia(...args)
  }

  /**
   * Performs a comprehensive validation of the entire PPTX structure.
   * Checks slide XML, relationships, content types, slide masters, and layouts.
   *
   * @returns {Promise<{valid: boolean, errors: string[], warnings: string[]}>} Validation report.
   *
   * @example
   * const result = await ppt.validatePresentation();
   * if (!result.valid) {
   *   console.error('Errors:', result.errors);
   *   console.warn('Warnings:', result.warnings);
   * }
   */
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

  /**
   * Validates the XML structure of a specific slide.
   *
   * @param {number} slideIndex - 1-based slide index to validate.
   * @returns {Promise<{valid: boolean, errors: string[], warnings: string[]}>}
   *
   * @example
   * const result = await ppt.validateSlide(1);
   * if (!result.valid) console.error(result.errors);
   */
  async validateSlide(slideIndex) {
    this.#assertLoaded()
    return await ValidationEngine.validateSlide(this, slideIndex)
  }

  /**
   * Validates the XML structure of a specific table on the active slide.
   *
   * @param {string} tableId - Table name or shape ID.
   * @returns {Promise<{valid: boolean, errors: string[], warnings: string[]}>}
   *
   * @example
   * const result = await ppt.validateTable('SalesTable');
   * if (!result.valid) console.error(result.errors);
   */
  async validateTable(tableId) {
    this.#assertLoaded()
    return await ValidationEngine.validateTable(
      this,
      this.#getTargetSlideIndices()[0] || 1,
      tableId
    )
  }

  /**
   * Validates the internal ZIP archive structure of the PPTX file.
   * Checks that all files referenced in the archive are accessible and uncorrupted.
   * Throws if critical structural issues are found.
   *
   * @returns {Promise<PPTXTemplater>} this (chainable)
   *
   * @example
   * await ppt.validateArchive(); // Throws PPTXError if the ZIP is corrupt
   */
  async validateArchive() {
    this.#assertLoaded()
    await this.#zipManager.validateArchive()
    return this
  }

  /**
   * Enables ZIP debug output. When enabled, every call to `toBuffer()` or `toStream()`
   * will log all ZIP entries (name, compression method, sizes, CRC).
   *
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.enableDebugZip();
   * const buffer = await ppt.toBuffer(); // Logs ZIP entries to debug output
   */
  enableDebugZip() {
    this.#outputWriter.debugZip = true
    return this
  }

  /**
   * Validates relationships for a specific part path inside the ZIP.
   *
   * @param {string} partPath - ZIP path to validate (e.g. `'ppt/slides/slide1.xml'`).
   * @returns {Object} Validation result with `valid`, `errors`, and `warnings`.
   *
   * @example
   * const result = ppt.validateRelationships('ppt/slides/slide1.xml');
   */
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
  /**
   * Returns an ordered array of all slide objects (shapes, images, charts, tables)
   * from bottom to top of the stacking order.
   *
   * @param {number} [slideIndex] - 1-based slide index. Defaults to the active slide.
   * @returns {Array<{id: string, name: string, type: string, zIndex: number}>}
   *
   * @example
   * const order = ppt.getObjectOrder(1);
   * order.forEach(o => console.log(`[${o.zIndex}] ${o.name}`));
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
  /**
   * Returns the top-most (front) object on the slide.
   *
   * @param {number} [slideIndex] - 1-based slide index. Defaults to the active slide.
   * @returns {{id: string, name: string, type: string, zIndex: number}|null}
   *
   * @example
   * const top = ppt.getTopMostObject(1);
   * console.log('Front-most shape:', top.name);
   */
  getTopMostObject(slideIndex) {
    this.#assertLoaded()
    const targetIdx = slideIndex !== undefined ? slideIndex : this.#getTargetSlideIndices()[0] || 1
    return this.#zOrderManager.getTopMostObject(targetIdx, this.#slideManager)
  }

  /**
   * Retrieves the info of the bottom-most object on the slide.
   */
  /**
   * Returns the bottom-most (back) object on the slide.
   *
   * @param {number} [slideIndex] - 1-based slide index. Defaults to the active slide.
   * @returns {{id: string, name: string, type: string, zIndex: number}|null}
   *
   * @example
   * const bottom = ppt.getBottomMostObject(1);
   * console.log('Back-most shape:', bottom.name);
   */
  getBottomMostObject(slideIndex) {
    this.#assertLoaded()
    const targetIdx = slideIndex !== undefined ? slideIndex : this.#getTargetSlideIndices()[0] || 1
    return this.#zOrderManager.getBottomMostObject(targetIdx, this.#slideManager)
  }

  /**
   * Swaps stacking positions of two slide objects.
   */
  /**
   * Swaps the stacking positions of two slide objects.
   *
   * @param {number|string} slideIndexOrId1 - Slide index (if 3 args) or first object ID (if 2 args).
   * @param {string} id1OrId2 - First object ID (if 3 args) or second object ID (if 2 args).
   * @param {string} [id2] - Second object ID (only if slide index is provided as first arg).
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.swapObjects(1, 'logo', 'background'); // On slide 1, swap 'logo' and 'background'
   * ppt.swapObjects('logo', 'background');     // On active slide
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
  /**
   * Sorts all slide objects using a custom comparison function.
   *
   * @param {number|Function} slideIndexOrCompareFn - Slide index (if 2 args) or compare function (if 1 arg).
   * @param {Function} [compareFnOption] - Compare function when slide index is provided.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * // Sort alphabetically by name on the active slide
   * ppt.sortObjects((a, b) => a.name.localeCompare(b.name));
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
  /**
   * Normalizes the stacking order of all objects on a slide, removing gaps
   * and ensuring Z-index values are sequential (1, 2, 3, ...).
   *
   * @param {number} [slideIndex] - 1-based slide index. Defaults to the active slide.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.normalizeZOrder(1); // Clean up stacking order on slide 1
   */
  normalizeZOrder(slideIndex) {
    this.#assertLoaded()
    const targetIdx = slideIndex !== undefined ? slideIndex : this.#getTargetSlideIndices()[0] || 1
    this.#zOrderManager.normalizeZOrder(targetIdx, this.#slideManager)
    return this
  }

  /**
   * Aligns an existing shape to a table cell's position.
   *
   * @param {string} shapeId - Unique shape name/id in the template.
   * @param {string|Object} tableIdOrObj - Table ID string, or table object.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} colIndex - 0-based column index.
   * @param {Object} [options] - Alignment options.
   * @param {'left'|'center'|'right'} [options.horizontal='center'] - Horizontal alignment.
   * @param {'top'|'middle'|'bottom'} [options.vertical='middle'] - Vertical alignment.
   * @returns {this} The chainable presentation templater instance.
   */
  alignShapeToCell(shapeId, tableIdOrObj, rowIndex, colIndex, options = {}) {
    const tableId =
      typeof tableIdOrObj === 'object'
        ? tableIdOrObj.id || tableIdOrObj.name || tableIdOrObj.tableId
        : tableIdOrObj
    this.updateShape(shapeId, {
      alignToCell: {
        table: tableId,
        row: rowIndex,
        col: colIndex,
        horizontal: options.horizontal || 'center',
        vertical: options.vertical || 'middle',
      },
    })
    return this
  }

  #resolveAlignToCell(slideIndex, options, shapeId, convertToEmus = false) {
    const align = options.alignToCell
    if (!align || !align.table) return options

    const tableId =
      typeof align.table === 'object'
        ? align.table.id || align.table.name || align.table.tableId
        : align.table
    const row = align.row !== undefined ? align.row : 0
    const col = align.col !== undefined ? align.col : 0

    // Get cell bounds
    const bounds = this.#tableManager.getCellBounds(
      slideIndex,
      tableId,
      row,
      col,
      this.#slideManager
    )

    if (!bounds) return options

    // Determine shape dimensions
    let shapeWidth = options.width
    let shapeHeight = options.height

    if (convertToEmus) {
      if (shapeWidth !== undefined) shapeWidth = Math.round(shapeWidth / 9525)
      if (shapeHeight !== undefined) shapeHeight = Math.round(shapeHeight / 9525)
    }

    if (shapeWidth === undefined || shapeHeight === undefined) {
      if (options.type === 'square' && options.size !== undefined) {
        shapeWidth = options.size
        shapeHeight = options.size
      } else if (options.type === 'circle' && options.radius !== undefined) {
        shapeWidth = options.radius * 2
        shapeHeight = options.radius * 2
      } else if (shapeId) {
        // Try getting existing shape dimensions
        const existing = this.#shapeManager.getShape(slideIndex, shapeId, this.#slideManager)
        if (existing) {
          shapeWidth = existing.width
          shapeHeight = existing.height
        }
      }
    }

    // Default to fallback dimensions if still undefined
    if (shapeWidth === undefined) shapeWidth = 100
    if (shapeHeight === undefined) shapeHeight = 100

    // Align horizontally
    let horiz = align.horizontal || align.alignX || 'center'
    horiz = String(horiz).toLowerCase()
    if (horiz === 'middle') horiz = 'center'

    let x = bounds.x
    if (horiz === 'center') {
      x = bounds.x + (bounds.width - shapeWidth) / 2
    } else if (horiz === 'right') {
      x = bounds.x + bounds.width - shapeWidth
    }

    // Align vertically
    let vert = align.vertical || align.alignY || 'middle'
    vert = String(vert).toLowerCase()
    if (vert === 'center') vert = 'middle'

    let y = bounds.y
    if (vert === 'middle') {
      y = bounds.y + (bounds.height - shapeHeight) / 2
    } else if (vert === 'bottom') {
      y = bounds.y + bounds.height - shapeHeight
    }

    const resolved = { ...options }
    if (convertToEmus) {
      resolved.x = Math.round(x * 9525)
      resolved.y = Math.round(y * 9525)
    } else {
      resolved.x = Math.round(x)
      resolved.y = Math.round(y)
    }

    // Remove alignToCell to prevent it polluting lower levels
    delete resolved.alignToCell

    return resolved
  }
}

module.exports = { PPTXTemplater }
