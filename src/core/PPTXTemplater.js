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
    this.#tableManager = new TableManager(this.#xmlParser)
    this.#hyperlinkManager = new HyperlinkManager(this.#xmlParser, this.#relationshipManager)
    this.#mediaManager = new MediaManager(this.#contentTypesManager)
    this.#shapeManager = new ShapeManager(this.#xmlParser)
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
   * Adds a new table to the active or specified slide(s).
   *
   * @param {string|Object} tableIdOrOptions - Table name/ID or configuration options.
   * @param {Object} [options={}] - Table configuration options if first argument is string.
   * @param {string} [options.id] - Table name or identifier.
   * @param {number} [options.x=50] - Horizontal position.
   * @param {number} [options.y=50] - Vertical position.
   * @param {number} [options.width=500] - Total table width.
   * @param {number} [options.height] - Total table height.
   * @param {number} [options.rows] - Number of rows.
   * @param {number} [options.columns] - Number of columns.
   * @param {Array<Array<string|Object>>} [options.data] - 2D array of rows and cells.
   * @param {number[]} [options.colWidths] - Array of column widths.
   * @param {number[]} [options.rowHeights] - Array of row heights.
   * @param {boolean} [options.header=true] - Whether row 0 has header formatting.
   * @param {boolean} [options.firstRow=true] - Format first row via DrawingML table properties.
   * @param {boolean} [options.bandRow=true] - Alternate row banding via DrawingML table properties.
   * @param {string} [options.styleId] - DrawingML table style GUID.
   * @returns {string} The created table name or ID.
   *
   * @example
   * ppt.addTable({
   *   id: 'summary-table',
   *   x: 50,
   *   y: 100,
   *   width: 600,
   *   data: [
   *     ['Metric', 'Q1', 'Q2', 'Q3'],
   *     ['Revenue', '$10M', '$12M', '$15M'],
   *   ]
   * });
   */
  addTable(tableIdOrOptions, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    if (targetIndices.length === 0) {
      throw new PPTXError('No slide selected to add table')
    }

    let finalOptions = {}
    if (typeof tableIdOrOptions === 'string') {
      finalOptions = { ...options, id: tableIdOrOptions }
    } else if (Array.isArray(tableIdOrOptions)) {
      finalOptions = { ...options, rows: tableIdOrOptions }
    } else if (typeof tableIdOrOptions === 'object' && tableIdOrOptions !== null) {
      finalOptions = { ...tableIdOrOptions }
    }

    let createdName = ''
    for (const slideIndex of targetIndices) {
      createdName = this.#tableManager.addTable(slideIndex, finalOptions, this.#slideManager)
    }

    return createdName
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
    logger.info(`=== Chart Inspection: ${chartId} ===`)
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
          logger.info(`Found on Slide ${i}`)
          logger.info(`ZIP Path: ${chart.zipPath}`)
          logger.info(`Relationship ID: ${chart.rId}`)
          found = true
          break
        }
      } catch (e) {}
    }
    if (!found) logger.info('Chart not found.')
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
    logger.info('=== Chart Relationships ===')
    const chartFiles = this.#zipManager.listFiles('ppt/charts/').filter(f => {
      const name = f.split('/').pop()
      return name.startsWith('chart') && name.endsWith('.xml') && !f.includes('_rels')
    })
    for (const chartPath of chartFiles) {
      logger.info(`\n${chartPath}:`)
      const rels = this.#relationshipManager.getRelationships(chartPath)
      rels.forEach(r => logger.info(`  - ${r.id} [${r.type.split('/').pop()}] -> ${r.target}`))
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
   * @param {number} [startSlideIndex] - 1-based start slide index. Defaults to 1 (or next unassigned).
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.addSection('Executive Summary', 1);
   */
  addSection(name, startSlideIndex) {
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

  // === Table Features ===
  /**
   * Extracts table data from the active slide as structured JSON.
   * The first row of the table is treated as the header row.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {Object} [options] - Extraction options.
   * @param {boolean} [options.raw=false] - Return `string[][]` instead of object array.
   * @param {boolean} [options.includeMetadata=false] - Return `{rows, rowCount, columnCount, mergedCells}`.
   * @returns {Array<Object>|Array<Array<string>>|Object} Extracted table data.
   *
   * @example
   * // Object mode (default)
   * const rows = await ppt.getTableRows('SalesTable');
   * // → [{ region: 'North', sales: '1200' }, ...]
   *
   * // Raw mode
   * const raw = await ppt.getTableRows('SalesTable', { raw: true });
   * // → [['North', '1200'], ...]
   *
   * // Metadata mode
   * const meta = await ppt.getTableRows('SalesTable', { includeMetadata: true });
   * // → { rows: [...], rowCount: 5, columnCount: 3, mergedCells: [] }
   */
  getTableRows(tableId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    if (targetIndices.length === 0) {
      throw new PPTXError('No slides active/loaded')
    }
    const idx = targetIndices[0]
    return this.#tableManager.getTableRows(idx, tableId, options, this.#slideManager)
  }

  /**
   * Appends one or more rows to a table. Supports flat arrays and nested arrays
   * for rowspan-merged cells.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {Array<string|Array<string>>} rowData - Row data. Nested arrays create rowspan cells.
   * @param {Object} [options] - Row insertion options.
   * @param {'rowspan'|'auto'|'none'} [options.mergeStrategy='rowspan'] - How to handle nested arrays.
   *   `'rowspan'` creates OpenXML vertical spans, `'auto'` merges identical adjacent values,
   *   `'none'` expands nested arrays into multiple flat rows.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * // Simple flat row
   * ppt.addTableRow('SalesTable', ['North', '1200', '15%']);
   *
   * // Nested row with rowspan
   * ppt.addTableRow('SalesTable', ['Region', ['Q1', 'Q2'], '$5K'], { mergeStrategy: 'rowspan' });
   */
  addTableRow(tableId, rowData, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.addTableRow(
        idx,
        tableId,
        rowData,
        this.#slideManager,
        this.#shapeManager,
        options
      )
    }
    return this
  }

  /**
   * Removes a row from a table by its 0-based row index.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index to remove.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.removeTableRow('SalesTable', 2); // Remove the third row (0-based)
   */
  removeTableRow(tableId, rowIndex) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.removeTableRow(
        idx,
        tableId,
        rowIndex,
        this.#slideManager,
        this.#shapeManager
      )
    }
    return this
  }

  /**
   * Inserts a new row at the specified 0-based index, shifting existing rows down.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based index at which to insert the new row.
   * @param {Array<string>} rowData - Cell values for the new row.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.insertTableRow('SalesTable', 1, ['East', '980', '8%']); // Insert at row 1
   */
  insertTableRow(tableId, rowIndex, rowData) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.insertTableRow(
        idx,
        tableId,
        rowIndex,
        rowData,
        this.#slideManager,
        this.#shapeManager
      )
    }
    return this
  }

  /**
   * Clones a row and inserts the copy at a target position.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} sourceRowIndex - 0-based index of the row to clone.
   * @param {number} targetRowIndex - 0-based index where the clone is inserted.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.cloneTableRow('SalesTable', 0, 3); // Clone row 0 to position 3
   */
  cloneTableRow(tableId, sourceRowIndex, targetRowIndex) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.cloneTableRow(
        idx,
        tableId,
        sourceRowIndex,
        targetRowIndex,
        this.#slideManager,
        this.#shapeManager
      )
    }
    return this
  }

  /**
   * Updates the text and optional formatting of a single table cell.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} colIndex - 0-based column index.
   * @param {string} value - New cell text content.
   * @param {Object} [options] - Cell formatting options.
   * @param {boolean} [options.bold] - Bold text.
   * @param {boolean} [options.italic] - Italic text.
   * @param {number} [options.fontSize] - Font size in points.
   * @param {'left'|'center'|'right'} [options.align] - Text alignment.
   * @param {string} [options.fill] - Cell background color (hex, e.g. '#FF0000').
   * @param {string} [options.color] - Text color (hex).
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.updateCell('SalesTable', 1, 2, '$9,800', { bold: true, color: '#10B981' });
   */
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
        this.#slideManager,
        this.#shapeManager
      )
    }
    return this
  }

  /**
   * Adds or appends a column to a table.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number|Array} colIndexOrData - Column index (0-based) or cell data array.
   * @param {Array|Object} [colDataOrOptions] - Cell data array or options.
   * @param {Object} [options={}] - Options if 2nd arg was colIndex.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.addColumn('SalesTable', ['Q4', '1500', '20%']);
   * ppt.addColumn('SalesTable', 1, ['Target', '1000', '10%'], { width: 120 });
   */
  addColumn(tableId, colIndexOrData, colDataOrOptions, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()

    for (const slideIndex of targetIndices) {
      this.#tableManager.addColumn(
        slideIndex,
        tableId,
        colIndexOrData,
        colDataOrOptions,
        options,
        this.#slideManager,
        this.#shapeManager
      )
    }

    return this
  }

  /**
   * Deletes a column from a table by 0-based index.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} colIndex - 0-based column index.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.deleteColumn('SalesTable', 2);
   */
  deleteColumn(tableId, colIndex) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()

    for (const slideIndex of targetIndices) {
      this.#tableManager.deleteColumn(
        slideIndex,
        tableId,
        colIndex,
        this.#slideManager,
        this.#shapeManager
      )
    }

    return this
  }

  /**
   * Applies alternating row striping to a table.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {Object} [options={}] - Striping options.
   * @param {string} [options.evenFill='F2F2F2'] - Background color for even rows.
   * @param {string} [options.oddFill] - Background color for odd rows.
   * @param {string} [options.headerFill] - Background color for row 0.
   * @param {number} [options.startRow=1] - 0-based starting row for striping.
   * @param {boolean} [options.bandRow] - Enable DrawingML bandRow table property.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.stripeRows('SalesTable', { evenFill: '#F8F9FA', oddFill: '#FFFFFF' });
   */
  stripeRows(tableId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()

    for (const slideIndex of targetIndices) {
      this.#tableManager.stripeRows(slideIndex, tableId, options, this.#slideManager)
    }

    return this
  }

  /**
   * Merges a rectangular region of table cells into a single merged cell.
   * Supports both positional arguments and an options object.
   *
   * @param {string|Object} tableIdOrOptions - Table ID string, or options object with all fields.
   * @param {number} [startRow] - 0-based start row.
   * @param {number} [startCol] - 0-based start column.
   * @param {number} [endRow] - 0-based end row (inclusive).
   * @param {number} [endCol] - 0-based end column (inclusive).
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.mergeCells('SalesTable', 0, 0, 0, 2); // Merge first 3 cells of row 0
   * // Or with options object:
   * ppt.mergeCells({ tableId: 'SalesTable', startRow: 0, startCol: 0, endRow: 0, endCol: 2 });
   */
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
      this.#tableManager.mergeCells(
        idx,
        tableId,
        sRow,
        sCol,
        eRow,
        eCol,
        this.#slideManager,
        this.#shapeManager
      )
    }
    return this
  }

  /**
   * Unmerges (splits) a previously merged cell region.
   * Supports both positional arguments and an options object.
   *
   * @param {string|Object} tableIdOrOptions - Table ID string, or options object.
   * @param {number} [startRow] - 0-based start row of the merged region.
   * @param {number} [startCol] - 0-based start column.
   * @param {number} [endRow] - 0-based end row.
   * @param {number} [endCol] - 0-based end column.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.unmergeCells('SalesTable', 0, 0, 0, 2); // Unmerge region
   * // Via cell coordinate:
   * ppt.unmergeCells({ tableId: 'SalesTable', row: 0, col: 0 });
   */
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
        this.#tableManager.unmergeCells(
          idx,
          tableId,
          cellRow,
          cellCol,
          this.#slideManager,
          this.#shapeManager
        )
      } else {
        this.#tableManager.unmergeCells(
          idx,
          tableId,
          sRow,
          sCol,
          eRow,
          eCol,
          this.#slideManager,
          this.#shapeManager
        )
      }
    }
    return this
  }

  /**
   * Returns an array of all merged cell regions in a table.
   *
   * @param {string} [tableId] - Table name or shape ID. Defaults to the first table found.
   * @returns {Array<{startRow: number, startCol: number, endRow: number, endCol: number}>}
   *
   * @example
   * const merges = ppt.getMergedCells('SalesTable');
   * merges.forEach(m => console.log(`Merged: row ${m.startRow}-${m.endRow}, col ${m.startCol}-${m.endCol}`));
   */
  getMergedCells(tableId) {
    this.#assertLoaded()
    const slideIndex = this.#getTargetSlideIndices()[0] || 1
    return this.#tableManager.getMergedCells(slideIndex, tableId || 'first', this.#slideManager)
  }

  /**
   * Validates whether a merge region is valid for the given table dimensions.
   * Checks for overlapping merges, out-of-bounds coordinates, etc.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} startRow - 0-based start row.
   * @param {number} startCol - 0-based start column.
   * @param {number} endRow - 0-based end row.
   * @param {number} endCol - 0-based end column.
   * @returns {{valid: boolean, errors: string[]}} Validation result.
   *
   * @example
   * const result = ppt.validateMergeRegion('SalesTable', 0, 0, 1, 2);
   * if (!result.valid) console.error(result.errors);
   */
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

  /**
   * Checks whether a specific table cell is part of a merged region.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} row - 0-based row index.
   * @param {number} col - 0-based column index.
   * @returns {boolean} `true` if the cell is merged (parent or continuation).
   *
   * @example
   * if (ppt.isMergedCell('SalesTable', 0, 0)) {
   *   console.log('Cell is part of a merged region');
   * }
   */
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

  /**
   * Returns the anchor (parent) cell coordinates of a merged region
   * that contains the given cell.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} row - 0-based row index of any cell in the merged region.
   * @param {number} col - 0-based column index.
   * @returns {{row: number, col: number}|null} Anchor cell coordinates, or null if not merged.
   *
   * @example
   * const parent = ppt.getMergeParent('SalesTable', 0, 1);
   * // → { row: 0, col: 0 } if cells (0,0)-(0,2) are merged
   */
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

  /**
   * Returns the full extent of the merged region containing a given cell.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} row - 0-based row index of any cell in the merged region.
   * @param {number} col - 0-based column index.
   * @returns {{startRow: number, startCol: number, endRow: number, endCol: number}|null}
   *
   * @example
   * const region = ppt.getMergeRegion('SalesTable', 0, 1);
   * // → { startRow: 0, startCol: 0, endRow: 0, endCol: 2 }
   */
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

  /**
   * Splits a previously merged cell region back into individual cells.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} row - 0-based row of the merged region anchor.
   * @param {number} col - 0-based column of the merged region anchor.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.splitMergedRegion('SalesTable', 0, 0); // Split merge starting at row 0, col 0
   */
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

  /**
   * Clones an existing merged region to a new anchor position in the table.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} row - 0-based row of the source merged region.
   * @param {number} col - 0-based column of the source merged region.
   * @param {number} targetRow - 0-based target row.
   * @param {number} targetCol - 0-based target column.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.cloneMergedRegion('SalesTable', 0, 0, 4, 0);
   */
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

  /**
   * Automatically adjusts column widths to fit the content of each cell.
   *
   * @param {string} tableId - Table name or shape ID.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.autoFitTable('SalesTable');
   */
  autoFitTable(tableId) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.autoFitTable(idx, tableId, this.#slideManager)
    }
    return this
  }

  /**
   * Resizes a table to the specified width and height in EMUs.
   * 1 inch = 914,400 EMUs.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} width - New width in EMUs.
   * @param {number} height - New height in EMUs.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.resizeTable('SalesTable', 6858000, 1371600); // 7.5" wide × 1.5" tall
   */
  resizeTable(tableId, width, height) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.resizeTable(idx, tableId, width, height, this.#slideManager)
    }
    return this
  }

  /**
   * Adjusts table row height(s). Supports direct height values or auto-calculation based on cell content.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {Object|Array} options - Row adjustment options.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.adjustTableRow("Table", { row: 2, height: 35 });
   * ppt.adjustTableRow("Table", { row: 2, auto: true });
   */
  adjustTableRow(tableId, options) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.adjustTableRow(
        idx,
        tableId,
        options,
        this.#slideManager,
        this.#shapeManager
      )
    }
    return this
  }

  /**
   * Adjusts table column width(s). Supports direct width values or auto-calculation based on cell content.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {Object|Array} options - Column adjustment options.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.adjustTableColumn("Table", { column: 3, width: 2.4 });
   * ppt.adjustTableColumn("Table", { column: 3, auto: true });
   */
  adjustTableColumn(tableId, options) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.adjustTableColumn(
        idx,
        tableId,
        options,
        this.#slideManager,
        this.#shapeManager
      )
    }
    return this
  }

  /**
   * Automatically splits a table across consecutive slides if its height exceeds a specified vertical threshold.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {Object} [options={}] - Options.
   * @param {number} [options.bottomMargin] - Distance from the bottom of the slide (in inches, pts, px, or EMUs) to use as the limit. Default is 0.5 inches (457200 EMUs).
   * @param {number} [options.repeatHeaders] - Number of header rows to repeat on duplicated slides. Default is 1.
   * @param {number} [options.summaryRows] - Number of totals/summary rows to preserve on the final generated slide.
   * @param {boolean} [options.keepSummaryOnLastSlide] - If true, keeps 1 summary row on the last generated slide.
   * @returns {Promise<this>}
   */
  async splitTableOnOverflow(tableId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    if (targetIndices.length === 0) return this

    const units = require('../utils/units.js')
    const bottomMargin =
      options.bottomMargin !== undefined ? units.toEmu(options.bottomMargin) : 457200
    const repeatHeaders = options.repeatHeaders !== undefined ? options.repeatHeaders : 1
    const summaryRows =
      options.summaryRows !== undefined
        ? options.summaryRows
        : options.keepSummaryOnLastSlide
          ? 1
          : 0

    const slideSize = this.#slideManager.getSlideSize()
    const limitY = slideSize.height - bottomMargin

    // Process from last to first to avoid index shifting issues when duplicating slides
    for (let i = targetIndices.length - 1; i >= 0; i--) {
      const slideIndex = targetIndices[i]
      await this.#splitTableRecursive(slideIndex, tableId, limitY, repeatHeaders, summaryRows)
    }

    return this
  }

  async #splitTableRecursive(slideIndex, tableId, limitY, repeatHeaders, summaryRows = 0) {
    const tableRes = this.#slideManager.getSlideTable(slideIndex, tableId)
    if (!tableRes) return // Table not found on this slide

    const tblObj = tableRes.table
    const graphicFrame = tableRes.frame
    const xfrm = graphicFrame?.['p:xfrm']
    if (!xfrm || !xfrm['a:off'] || !xfrm['a:ext']) return

    const tableY = parseInt(xfrm['a:off']['@_y'] || 0, 10)

    const trs = tblObj['a:tr'] || []
    const rowHeights = []

    for (let r = 0; r < trs.length; r++) {
      const bounds = this.#tableManager.getCellBounds(slideIndex, tableId, r, 0, this.#slideManager)
      const h = bounds ? Math.round(bounds.height * 9525) : parseInt(trs[r]['@_h'] || 0, 10)
      rowHeights.push(h)
    }

    let splitRowIndex = -1
    let currentY = tableY

    for (let r = 0; r < rowHeights.length; r++) {
      currentY += rowHeights[r]
      if (currentY > limitY) {
        // If even the first row (or headers) exceeds the limit, we must split after it to avoid infinite loops
        splitRowIndex = Math.max(r, repeatHeaders)
        break
      }
    }

    if (summaryRows > 0 && splitRowIndex !== -1) {
      // Summary rows must only appear on the final generated slide.
      // Ensure splitRowIndex does not cut into or include summary rows on this intermediate slide.
      const maxAllowed = trs.length - summaryRows
      if (splitRowIndex > maxAllowed) {
        splitRowIndex = Math.max(maxAllowed, repeatHeaders)
      }
    }

    if (splitRowIndex !== -1 && splitRowIndex < trs.length) {
      // Split required
      // 1. Deep clone the slide directly after current slide
      const newSlideIndex = slideIndex + 1
      await this.#slideManager.duplicateSlide(
        slideIndex,
        newSlideIndex,
        this.#relationshipManager,
        this.#mediaManager
      )

      // 2. Original slide: remove rows from splitRowIndex to end
      const rowsToRemove = trs.length - splitRowIndex
      for (let k = 0; k < rowsToRemove; k++) {
        this.#tableManager.removeTableRow(slideIndex, tableId, splitRowIndex, this.#slideManager)
      }

      // 3. New slide: remove rows between headers and splitRowIndex
      const numToRemoveFromNew = splitRowIndex - repeatHeaders
      for (let k = 0; k < numToRemoveFromNew; k++) {
        this.#tableManager.removeTableRow(newSlideIndex, tableId, repeatHeaders, this.#slideManager)
      }

      // 4. Recurse on the new slide in case it still overflows
      await this.#splitTableRecursive(newSlideIndex, tableId, limitY, repeatHeaders, summaryRows)
    }
  }

  /**
   * Auto-adjusts or sets table row heights, column widths, and cell layout on targeted slides.
   *
   * @param {string} tableId - Name or ID of the table to adjust.
   * @param {Object} [options] - Adjustment options.
   * @returns {PPTXTemplater} this (chainable)
   */
  adjustTable(tableId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.adjustTable(idx, tableId, options, this.#slideManager)
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
      this.#shapeManager.updateShapePosition(idx, objectIdOrName, emuOptions, this.#slideManager)
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
   * Aligns multiple shapes on the active slide.
   *
   * @param {string[]} shapeIds - Array of shape IDs or names to align.
   * @param {'left'|'center'|'right'|'top'|'middle'|'bottom'} alignment - Target alignment edge or axis.
   * @returns {PPTXTemplater} this (chainable)
   */
  alignShapes(shapeIds, alignment) {
    this.#assertLoaded()
    if (!Array.isArray(shapeIds) || shapeIds.length < 2) return this

    const boxes = shapeIds.map(id => ({ id, bounds: this.getBounds(id) }))
    const validBoxes = boxes.filter(b => b.bounds)
    if (validBoxes.length < 2) return this

    const align = String(alignment).toLowerCase()
    if (align === 'left') {
      const minX = Math.min(...validBoxes.map(b => b.bounds.x))
      for (const b of validBoxes) {
        this.setBounds(b.id, { x: minX })
      }
    } else if (align === 'right') {
      const maxRight = Math.max(...validBoxes.map(b => b.bounds.x + b.bounds.width))
      for (const b of validBoxes) {
        this.setBounds(b.id, { x: maxRight - b.bounds.width })
      }
    } else if (align === 'center') {
      const minX = Math.min(...validBoxes.map(b => b.bounds.x))
      const maxRight = Math.max(...validBoxes.map(b => b.bounds.x + b.bounds.width))
      const center = (minX + maxRight) / 2
      for (const b of validBoxes) {
        this.setBounds(b.id, { x: center - b.bounds.width / 2 })
      }
    } else if (align === 'top') {
      const minY = Math.min(...validBoxes.map(b => b.bounds.y))
      for (const b of validBoxes) {
        this.setBounds(b.id, { y: minY })
      }
    } else if (align === 'bottom') {
      const maxBottom = Math.max(...validBoxes.map(b => b.bounds.y + b.bounds.height))
      for (const b of validBoxes) {
        this.setBounds(b.id, { y: maxBottom - b.bounds.height })
      }
    } else if (align === 'middle') {
      const minY = Math.min(...validBoxes.map(b => b.bounds.y))
      const maxBottom = Math.max(...validBoxes.map(b => b.bounds.y + b.bounds.height))
      const middle = (minY + maxBottom) / 2
      for (const b of validBoxes) {
        this.setBounds(b.id, { y: middle - b.bounds.height / 2 })
      }
    }

    return this
  }

  /**
   * Distributes multiple shapes evenly along a direction.
   *
   * @param {string[]} shapeIds - Array of shape IDs or names to distribute.
   * @param {'horizontal'|'vertical'} [direction='horizontal'] - Distribution direction.
   * @returns {PPTXTemplater} this (chainable)
   */
  distributeShapes(shapeIds, direction = 'horizontal') {
    this.#assertLoaded()
    if (!Array.isArray(shapeIds) || shapeIds.length < 3) return this

    const boxes = shapeIds.map(id => ({ id, bounds: this.getBounds(id) }))
    const validBoxes = boxes.filter(b => b.bounds)
    if (validBoxes.length < 3) return this

    const dir = String(direction).toLowerCase()
    if (dir === 'horizontal') {
      validBoxes.sort((a, b) => a.bounds.x - b.bounds.x)
      const first = validBoxes[0]
      const last = validBoxes[validBoxes.length - 1]
      const totalWidth = validBoxes.reduce((acc, b) => acc + b.bounds.width, 0)
      const span = last.bounds.x + last.bounds.width - first.bounds.x
      const gap = (span - totalWidth) / (validBoxes.length - 1)

      let currentX = first.bounds.x + first.bounds.width + gap
      for (let i = 1; i < validBoxes.length - 1; i++) {
        this.setBounds(validBoxes[i].id, { x: currentX })
        currentX += validBoxes[i].bounds.width + gap
      }
    } else if (dir === 'vertical') {
      validBoxes.sort((a, b) => a.bounds.y - b.bounds.y)
      const first = validBoxes[0]
      const last = validBoxes[validBoxes.length - 1]
      const totalHeight = validBoxes.reduce((acc, b) => acc + b.bounds.height, 0)
      const span = last.bounds.y + last.bounds.height - first.bounds.y
      const gap = (span - totalHeight) / (validBoxes.length - 1)

      let currentY = first.bounds.y + first.bounds.height + gap
      for (let i = 1; i < validBoxes.length - 1; i++) {
        this.setBounds(validBoxes[i].id, { y: currentY })
        currentY += validBoxes[i].bounds.height + gap
      }
    }

    return this
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
      this.#shapeManager.resizeObject(
        idx,
        objectIdOrName,
        options,
        this.#slideManager,
        this.#tableManager
      )
    }
    return this
  }

  /**
   * Returns metadata for all tables on the targeted slide(s).
   *
   * @returns {Array<{name: string, rows: number, cols: number, zipPath: string}>}
   *
   * @example
   * const tables = ppt.getTables();
   * tables.forEach(t => console.log(`${t.name}: ${t.rows}×${t.cols}`));
   */
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
   * Returns an array of all charts found on the targeted slide(s).
   *
   * @returns {Array<{rId: string, zipPath: string}>} Chart info objects.
   *
   * @example
   * const charts = ppt.getCharts();
   * charts.forEach(c => console.log(c.zipPath));
   */
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
  addTextAtPosition(textOrOptions, maybeOptions = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    let config = {}
    if (typeof textOrOptions === 'string') {
      config = { ...maybeOptions, text: textOrOptions }
    } else if (typeof textOrOptions === 'object' && textOrOptions !== null) {
      config = { ...textOrOptions }
    }
    for (const idx of targetIndices) {
      // Ensure dirty XML from prior synchronous operations (like addTable) is flushed to zip
      this.#slideManager.getSlideXml(idx)
      const p = this.#chartManager.addTextAtPosition(idx, config, this.#slideManager)
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
  /**
   * Updates text content or list items in a named shape or text box.
   * Supports plain strings, bullet lists, numbered lists, and nested lists.
   *
   * @param {string} tag - Shape name/ID or placeholder tag.
   * @param {string|Object} data - Text string, or list configuration object.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * // Plain text
   * ppt.updateText('SubtitleBox', 'Updated subtitle');
   *
   * // Bullet list
   * ppt.updateText('BulletBox', { items: ['Item A', 'Item B', 'Item C'] });
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
  /**
   * Retrieves list items from a shape or text box.
   *
   * @param {string} tag - Shape name/ID or placeholder tag.
   * @returns {Array} Nested list structure of items.
   *
   * @example
   * const items = ppt.getList('BulletBox');
   * console.log(items); // ['Item A', ['Nested', 'Sub-item'], 'Item B']
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
  /**
   * Replaces a text placeholder tag across all targeted shapes on selected slides.
   * Finds shapes containing `{{tag}}` or `tag` and replaces the placeholder value.
   *
   * @param {string} tag - Placeholder tag name (e.g. `'{{name}}'` or `'name'`).
   * @param {string} value - Replacement value.
   * @param {Object} [options] - Options.
   * @param {number} [options.slide] - Target a specific slide index (overrides `useSlide`).
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.replaceTextByTag('{{company}}', 'Acme Corp');
   * ppt.replaceTextByTag('{{date}}', '2026-06-12');
   */
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

  /**
   * Replaces multiple text placeholder tags in a single pass.
   * More efficient than calling `replaceTextByTag()` repeatedly.
   *
   * @param {Object<string, string>} replacements - Map of `{ tag: value }` pairs.
   * @param {Object} [options] - Options (same as `replaceTextByTag`).
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.replaceMultiple({
   *   '{{title}}': 'Q2 Report',
   *   '{{date}}': 'June 2026',
   *   '{{company}}': 'Acme Corp'
   * });
   */
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

  /**
   * Searches for all occurrences of a text string across the targeted slides.
   *
   * @param {string} text - Text to search for.
   * @returns {Array<{slideIndex: number, shapeName: string, text: string}>} Array of matches.
   *
   * @example
   * const matches = ppt.findText('Revenue');
   * matches.forEach(m => console.log(`Found on slide ${m.slideIndex} in "${m.shapeName}"`));
   */
  findText(text) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const matches = []
    for (const idx of targetIndices) {
      matches.push(...this.#textManager.findText(idx, text, this.#slideManager))
    }
    return matches
  }

  /**
   * Extracts all unique placeholders (e.g., {{placeholder}}) across targeted slides.
   *
   * @param {RegExp} [pattern] - Optional custom RegExp pattern. Defaults to `/\{\{([^{}]+)\}\}/g`.
   * @returns {string[]} Array of unique placeholder token strings.
   *
   * @example
   * const placeholders = ppt.extractPlaceholders();
   * console.log('Placeholders found:', placeholders); // ['{{title}}', '{{company}}']
   */
  extractPlaceholders(pattern) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const placeholders = new Set()
    for (const idx of targetIndices) {
      const slideXml = this.#slideManager.getSlideXml(idx)
      const matches = this.#templateEngine.extractPlaceholders(slideXml, pattern)
      for (const m of matches) {
        placeholders.add(m)
      }
    }
    return Array.from(placeholders)
  }

  /**
   * Returns all text elements (paragraphs) across the targeted slides.
   *
   * @returns {Array<{slideIndex: number, shapeName: string, text: string}>}
   *
   * @example
   * const elements = ppt.getTextElements();
   * elements.forEach(el => console.log(`Slide ${el.slideIndex}: ${el.text}`))
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

  /**
   * Appends a formatted text run to a shape or text box on targeted slides.
   *
   * @param {string} shapeId - Shape name or ID.
   * @param {string} text - Text to append.
   * @param {Object} [options] - Formatting options `{ style, paragraphIndex, inheritStyle }`.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.appendTextRun('titleShape', ' (Updated)', { style: { bold: true, color: '#FF0000' } });
   */
  appendTextRun(shapeId, text, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.appendTextRun(idx, shapeId, text, options, this.#slideManager)
    }
    return this
  }

  /**
   * Prepends a formatted text run to a shape or text box on targeted slides.
   *
   * @param {string} shapeId - Shape name or ID.
   * @param {string} text - Text to prepend.
   * @param {Object} [options] - Formatting options `{ style, paragraphIndex, inheritStyle }`.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.prependTextRun('titleShape', 'Draft: ', { style: { italic: true } });
   */
  prependTextRun(shapeId, text, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#textManager.prependTextRun(idx, shapeId, text, options, this.#slideManager)
    }
    return this
  }

  // === Shape Features ===
  /**
   * Sets the text content of an existing shape by name or ID.
   *
   * @param {string} shapeId - Shape name or ID.
   * @param {string} text - New text content.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.updateShapeText('CalloutBox', 'Important note here');
   */
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
      let resolvedOptions = options
      if (options.alignToCell) {
        resolvedOptions = this.#resolveAlignToCell(idx, options, shapeId, true)
      }
      this.#shapeManager.updateShapePosition(idx, shapeId, resolvedOptions, this.#slideManager)
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

  /**
   * Duplicates an existing shape with a new ID, optionally at a different position.
   *
   * @param {string} shapeId - Source shape name or ID.
   * @param {string} newShapeId - Name/ID for the cloned shape.
   * @param {Object} [options] - Position overrides for the clone.
   * @param {number} [options.x] - X offset for the clone (EMUs).
   * @param {number} [options.y] - Y offset for the clone (EMUs).
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.cloneShape('logo', 'logo-copy', { x: 2000000, y: 500000 });
   */
  cloneShape(shapeId, newShapeId, options = {}) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#shapeManager.cloneShape(idx, shapeId, newShapeId, options, this.#slideManager)
    }
    return this
  }

  /**
   * Removes a shape from the targeted slide(s). Alias for `removeShape()`.
   *
   * @param {string} shapeId - Shape name or ID to delete.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.deleteShape('temp-banner');
   */
  deleteShape(shapeId) {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#shapeManager.deleteShape(idx, shapeId, this.#slideManager)
    }
    return this
  }

  /**
   * Returns metadata for all shapes on the targeted slide(s).
   *
   * @returns {Array<{id: string, name: string, type: string, x: number, y: number, width: number, height: number}>}
   *
   * @example
   * const shapes = ppt.getShapes();
   * shapes.forEach(s => console.log(`${s.name}: ${s.type} at (${s.x}, ${s.y})`));
   */
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
   * Adds a new shape to the targeted slide(s).
   *
   * @param {string|Object} typeOrOptions Either the shape type string (e.g., 'rect', 'ellipse') or a full options object.
   * @param {Object} [options={}] Additional configuration options if type was specified as a string.
   * @returns {Promise<PPTXTemplater>} The templater instance for chaining.
   *
   * @example
   * await ppt.addShape('rect', { id: 'MyShape', x: 100, y: 100, width: 200, height: 100 });
   */
  async addShape(typeOrOptions, options = {}) {
    this.#assertLoaded()
    let resolvedOptions = {}
    if (typeof typeOrOptions === 'string') {
      resolvedOptions = { ...options, type: typeOrOptions }
    } else {
      resolvedOptions = { ...typeOrOptions }
    }

    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      let finalOptions = resolvedOptions
      if (resolvedOptions.alignToCell) {
        finalOptions = this.#resolveAlignToCell(idx, resolvedOptions)
      }
      this.#shapeManager.addShape(idx, finalOptions, this.#slideManager)
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
      let resolvedOptions = options
      if (options.alignToCell) {
        resolvedOptions = this.#resolveAlignToCell(idx, options, shapeId)
      }
      this.#shapeManager.updateShape(idx, shapeId, resolvedOptions, this.#slideManager)
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
   * Cell shapes are overlay graphics anchored independently of the table layout,
   * and adding a cell shape never modifies row heights, column widths, or table dimensions.
   *
   * @param {string} tableId - Table name or shape ID.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} colIndex - 0-based column index.
   * @param {Object} options - Shape configuration options.
   * @returns {this} The chainable presentation templater instance.
   */
  async addCellShape(tableId, rowIndex, colIndex, options) {
    this.#assertLoaded()
    let row = rowIndex
    let col = colIndex
    let opts = options
    if (rowIndex && typeof rowIndex === 'object') {
      row = rowIndex.row !== undefined ? rowIndex.row : rowIndex.rowIndex
      col =
        rowIndex.column !== undefined
          ? rowIndex.column
          : rowIndex.col !== undefined
            ? rowIndex.col
            : rowIndex.colIndex
      opts = rowIndex
    }

    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.addCellShape(
        idx,
        tableId,
        row,
        col,
        opts,
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
    let row = rowIndex
    let col = colIndex
    let shpIdx = shapeIndex
    let opts = options
    if (rowIndex && typeof rowIndex === 'object') {
      row = rowIndex.row !== undefined ? rowIndex.row : rowIndex.rowIndex
      col =
        rowIndex.column !== undefined
          ? rowIndex.column
          : rowIndex.col !== undefined
            ? rowIndex.col
            : rowIndex.colIndex
      shpIdx =
        rowIndex.shapeIndex !== undefined
          ? rowIndex.shapeIndex
          : rowIndex.shape !== undefined
            ? rowIndex.shape
            : colIndex
      opts = options !== undefined ? options : rowIndex
    }

    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.updateCellShape(
        idx,
        tableId,
        row,
        col,
        shpIdx,
        opts,
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
    let row = rowIndex
    let col = colIndex
    let shpIdx = shapeIndex
    if (rowIndex && typeof rowIndex === 'object') {
      row = rowIndex.row !== undefined ? rowIndex.row : rowIndex.rowIndex
      col =
        rowIndex.column !== undefined
          ? rowIndex.column
          : rowIndex.col !== undefined
            ? rowIndex.col
            : rowIndex.colIndex
      shpIdx =
        rowIndex.shapeIndex !== undefined
          ? rowIndex.shapeIndex
          : rowIndex.shape !== undefined
            ? rowIndex.shape
            : colIndex
    }

    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      this.#tableManager.removeCellShape(
        idx,
        tableId,
        row,
        col,
        shpIdx,
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
    let row = rowIndex
    let col = colIndex
    let shpIdx = shapeIndex
    if (rowIndex && typeof rowIndex === 'object') {
      row = rowIndex.row !== undefined ? rowIndex.row : rowIndex.rowIndex
      col =
        rowIndex.column !== undefined
          ? rowIndex.column
          : rowIndex.col !== undefined
            ? rowIndex.col
            : rowIndex.colIndex
      shpIdx =
        rowIndex.shapeIndex !== undefined
          ? rowIndex.shapeIndex
          : rowIndex.shape !== undefined
            ? rowIndex.shape
            : colIndex
    }

    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      const shape = this.#tableManager.getCellShape(
        idx,
        tableId,
        row,
        col,
        shpIdx,
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
   * @param {string|Object} tableIdOrObj - Table name, shape ID, or table object.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} colIndex - 0-based column index.
   * @returns {Object|null} Cell bounds { x, y, width, height } in pixels, or null.
   */
  getCellBounds(tableIdOrObj, rowIndex, colIndex) {
    this.#assertLoaded()
    const tableId =
      typeof tableIdOrObj === 'object'
        ? tableIdOrObj.id || tableIdOrObj.name || tableIdOrObj.tableId
        : tableIdOrObj
    let row = rowIndex
    let col = colIndex
    if (rowIndex && typeof rowIndex === 'object') {
      row = rowIndex.row !== undefined ? rowIndex.row : rowIndex.rowIndex
      col =
        rowIndex.column !== undefined
          ? rowIndex.column
          : rowIndex.col !== undefined
            ? rowIndex.col
            : rowIndex.colIndex
    }

    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      try {
        const bounds = this.#tableManager.getCellBounds(idx, tableId, row, col, this.#slideManager)
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
   * Optionally calculates centered top-left coordinates for a shape of given dimensions.
   *
   * @param {string|Object} tableIdOrObj - Table name, shape ID, or table object.
   * @param {number} rowIndex - 0-based row index.
   * @param {number} colIndex - 0-based column index.
   * @param {number|Object} [shapeWidthOrOptions] - Width of the shape in pixels, or options object.
   * @param {number} [shapeHeight] - Height of the shape in pixels.
   * @returns {Object|null} Cell position { row, column, x, y, width, height } in pixels, or null.
   */
  getCellPosition(tableIdOrObj, rowIndex, colIndex, shapeWidthOrOptions, shapeHeight) {
    this.#assertLoaded()
    const tableId =
      typeof tableIdOrObj === 'object'
        ? tableIdOrObj.id || tableIdOrObj.name || tableIdOrObj.tableId
        : tableIdOrObj
    let row = rowIndex
    let col = colIndex
    let widthOrOpts = shapeWidthOrOptions
    let height = shapeHeight
    if (rowIndex && typeof rowIndex === 'object') {
      row = rowIndex.row !== undefined ? rowIndex.row : rowIndex.rowIndex
      col =
        rowIndex.column !== undefined
          ? rowIndex.column
          : rowIndex.col !== undefined
            ? rowIndex.col
            : rowIndex.colIndex
      widthOrOpts = colIndex
      height = shapeWidthOrOptions
    }

    const targetIndices = this.#getTargetSlideIndices()
    for (const idx of targetIndices) {
      try {
        const pos = this.#tableManager.getCellPosition(
          idx,
          tableId,
          row,
          col,
          this.#slideManager,
          widthOrOpts,
          height
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
  /**
   * Replaces an existing image in the presentation by shape name or relationship ID.
   *
   * @param {string} imageIdOrName - Shape name, alt text, or relationship ID of the image.
   * @param {string|Buffer} sourcePathOrBuffer - Path to the replacement image file, or a Buffer.
   * @returns {Promise<PPTXTemplater>} this (chainable)
   *
   * @example
   * await ppt.replaceImage('company-logo', './new-logo.png');
   */
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

  /**
   * Adds a new image to the targeted slide(s) at the specified position.
   *
   * @param {string|Buffer} sourcePathOrBuffer - Path to the image file, or a Buffer.
   * @param {Object} [options] - Positioning and display options.
   * @param {number} [options.x] - X offset in EMUs (1 inch = 914,400 EMUs).
   * @param {number} [options.y] - Y offset in EMUs.
   * @param {number} [options.width] - Width in EMUs.
   * @param {number} [options.height] - Height in EMUs.
   * @param {number} [options.rotation] - Rotation in degrees (0–360).
   * @param {number} [options.opacity] - Opacity (0–100).
   * @param {string} [options.name] - Shape name for the image.
   * @param {Object} [options.cropTo] - Crop percentages `{ l, r, t, b }` (0–100000).
   * @returns {Promise<PPTXTemplater>} this (chainable)
   *
   * @example
   * await ppt.addImage('./photo.jpg', {
   *   x: 914400,       // 1 inch
   *   y: 914400,       // 1 inch
   *   width: 3657600,  // 4 inches
   *   height: 2743200, // 3 inches
   *   name: 'hero-image'
   * });
   */
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

  /**
   * Removes an image from the targeted slide(s) by shape name or relationship ID.
   *
   * @param {string} imageIdOrName - Shape name, alt text, or relationship ID of the image.
   * @returns {PPTXTemplater} this (chainable)
   *
   * @example
   * ppt.removeImage('old-logo');
   */
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

  /**
   * Returns metadata for all images found on the targeted slide(s).
   * Extracts format, MIME type, byte size, dimensions, position, and relationship info.
   *
   * @returns {Array<{rId: string, name: string, x: number, y: number, width: number, height: number, mediaPath: string, format: string, mimeType: string, size: number}>}
   *
   * @example
   * const images = ppt.getImages();
   * images.forEach(img => console.log(`${img.name}: ${img.width}×${img.height} EMUs, ${img.format}, ${img.size} bytes`));
   */
  getImages() {
    this.#assertLoaded()
    const targetIndices = this.#getTargetSlideIndices()
    const images = []
    for (const idx of targetIndices) {
      const rawImages = this.#imageManager.getImages(
        idx,
        this.#slideManager,
        this.#relationshipManager
      )
      for (const img of rawImages) {
        const mediaInfo = img.targetPath ? this.#mediaManager.getMediaInfo(img.targetPath) : null
        const ext = img.targetPath ? img.targetPath.split('.').pop().toLowerCase() : ''
        images.push({
          ...img,
          slideIndex: idx,
          format: ext,
          mimeType: mediaInfo?.mimeType || (ext ? `image/${ext}` : undefined),
          size: mediaInfo?.size,
          width: img.position ? img.position.cx : 0,
          height: img.position ? img.position.cy : 0,
          x: img.position ? img.position.x : 0,
          y: img.position ? img.position.y : 0,
          mediaPath: img.targetPath,
          rId: img.relationshipId,
        })
      }
    }
    return images
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
    this.updateShapePosition(shapeId, {
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
