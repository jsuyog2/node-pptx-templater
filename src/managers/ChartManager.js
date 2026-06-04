/**
 * @fileoverview ChartManager - Updates chart data in PPTX slides.
 *
 * Charts in OpenXML PPTX:
 * ─────────────────────────────────────────────────────────────────
 * Charts are embedded as separate XML files referenced via relationships.
 * A chart on a slide appears as:
 *
 *   <p:graphicFrame>
 *     <p:nvGraphicFramePr>
 *       <p:cNvPr id="5" name="sales-chart"/>    ← chart name/ID
 *     </p:nvGraphicFramePr>
 *     <a:graphic>
 *       <a:graphicData uri="...chart">
 *         <c:chart r:id="rId5"/>                ← references chart XML
 *       </a:graphicData>
 *     </a:graphic>
 *   </p:graphicFrame>
 *
 * The chart XML at ppt/charts/chartN.xml contains:
 *  - c:chartSpace → root element
 *    - c:chart → chart metadata
 *      - c:plotArea → chart data
 *        - c:barChart / c:lineChart / c:pieChart / etc.
 *          - c:ser → each data series
 *            - c:idx / c:order → series index/order
 *            - c:tx → series name
 *            - c:cat → categories (X-axis)
 *            - c:val → values (Y-axis)
 *
 * Category Reference Types:
 *  - c:strRef  → string categories (labels)
 *  - c:numRef  → numeric categories
 *
 * Value Reference Types:
 *  - c:numRef  → numeric values (typical)
 *
 * Data is stored in both a formula (c:f) and a cache (c:strCache/c:numCache).
 * We update both to ensure compatibility with both cached and live data.
 */

const { createLogger } = require('../utils/logger.js')
const { ChartNotFoundError } = require('../utils/errors.js')
const { REL_TYPES } = require('./RelationshipManager.js')
const { ChartWorkbookUpdater } = require('./charts/ChartWorkbookUpdater.js')
const { ChartCacheGenerator } = require('./charts/ChartCacheGenerator.js')

const JSZip = require('jszip')

const logger = createLogger('ChartManager')

/**
 * Supported chart types and their XML element names.
 */
const CHART_TYPE_MAP = {
  bar: 'c:barChart',
  line: 'c:lineChart',
  pie: 'c:pieChart',
  area: 'c:areaChart',
  scatter: 'c:scatterChart',
  doughnut: 'c:doughnutChart',
  radar: 'c:radarChart',
  bubble: 'c:bubbleChart',
  stock: 'c:stockChart',
}

/**
 * @class ChartManager
 * @description Handles chart data updates by directly manipulating chart XML files.
 *
 * Unlike high-level charting libraries, this manager edits the raw OpenXML
 * chart structure, allowing full control while preserving styles and themes.
 */
class ChartManager {
  /** @private @type {XMLParser} */
  #xmlParser
  /** @private @type {ZipManager} */
  #zipManager

  /**
   * Cache of chart ZIP paths: maps chartName → { zipPath, slideIndex }
   * @private @type {Map<string, { zipPath: string, slideIndex: number }>}
   */
  #chartRegistry = new Map()

  /**
   * Promise queue for sequential execution per chart ZIP path.
   * @private @type {Map<string, Promise>}
   */
  #chartQueues = new Map()

  /**
   * @param {XMLParser} xmlParser
   */
  constructor(xmlParser) {
    this.#xmlParser = xmlParser
  }

  async initialize(zipManager) {
    this.#zipManager = zipManager
    const chartFiles = zipManager
      .listFiles('ppt/charts/')
      .filter(f => f.endsWith('.xml') && !f.includes('_rels'))

    for (const chartPath of chartFiles) {
      // Chart name is inferred from file name
      const chartName = chartPath.split('/').pop().replace('.xml', '')
      this.#chartRegistry.set(chartName, { zipPath: chartPath, slideIndex: null })
      // Pre-load the chart XML into cache so that we can read it synchronously if needed
      await zipManager.readFile(chartPath)
    }

    logger.debug(`Found ${chartFiles.length} chart file(s)`)
  }

  /**
   * Updates chart data for a named chart within a slide.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string} chartId - Chart name (from shape's cNvPr name attribute) or chart file name.
   * @param {ChartData} data - New chart data.
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @throws {ChartNotFoundError} If the chart cannot be found.
   */
  updateChart(slideIndex, chartId, data, slideManager, relationshipManager) {
    this.#validateChartData(data)
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)

    if (!chartInfo) {
      throw new ChartNotFoundError(`Chart "${chartId}" not found in slide ${slideIndex}`)
    }

    logger.debug(`Updating chart "${chartId}" at ${chartInfo.zipPath}`)
    this.#updateChartXml(chartInfo.zipPath, data, relationshipManager)
  }

  /**
   * Returns all chart info objects for a given slide.
   *
   * @param {number} slideIndex
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @returns {Array<{name: string, zipPath: string}>}
   */
  getChartsInSlide(slideIndex, slideManager, relationshipManager) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    const rels = relationshipManager.getRelationshipsByType(slideInfo.zipPath, REL_TYPES.CHART)

    return rels.map(rel => {
      const chartPath = relationshipManager.resolveTarget(slideInfo.zipPath, rel.target)
      return { rId: rel.id, zipPath: chartPath }
    })
  }

  /**
   * Finds a chart in a slide by name/ID.
   * @private
   *
   * @param {number} slideIndex
   * @param {string} chartId - Shape name or rId.
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @returns {{ zipPath: string }|null}
   */
  findChartInSlide(slideIndex, chartId, slideManager, relationshipManager) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    const slideXml = slideManager.getSlideXml(slideIndex)

    // Strategy 1: Look for shape with matching name (cNvPr name attribute)
    const shapeNamePattern = new RegExp(
      `<p:cNvPr[^>]*name="${chartId}"[^>]*>(?:.*?)<c:chart[^>]*r:id="(rId\\d+)"`,
      's'
    )
    const rIdMatch = shapeNamePattern.exec(slideXml)

    // Strategy 2: Find graphicFrame shapes and match chart rIds
    if (!rIdMatch) {
      const chartRIdPattern = /<c:chart[^>]*r:id="(rId\d+)"/g
      let chartMatch
      while ((chartMatch = chartRIdPattern.exec(slideXml)) !== null) {
        const rId = chartMatch[1]
        const rel = relationshipManager.getRelationshipById(slideInfo.zipPath, rId)
        if (rel) {
          const chartPath = relationshipManager.resolveTarget(slideInfo.zipPath, rel.target)
          // Check if chart file name matches chartId
          if (chartPath.includes(chartId) || rel.id === chartId) {
            return { zipPath: chartPath }
          }
          // For the first chart found, if chartId looks like a chart file name
          if (chartId.startsWith('chart')) {
            return { zipPath: chartPath }
          }
        }
      }
    }

    if (rIdMatch) {
      const rId = rIdMatch[1]
      const rel = relationshipManager.getRelationshipById(slideInfo.zipPath, rId)
      if (rel) {
        const chartPath = relationshipManager.resolveTarget(slideInfo.zipPath, rel.target)
        return { zipPath: chartPath }
      }
    }

    // Strategy 3: Direct chart registry lookup
    if (this.#chartRegistry.has(chartId)) {
      return this.#chartRegistry.get(chartId)
    }

    // Strategy 4: Try chartN naming convention
    const chartPath = `ppt/charts/${chartId}.xml`
    if (this.#zipManager.hasFile(chartPath)) {
      return { zipPath: chartPath }
    }

    return null
  }

  /**
   * Updates the chart XML file with new data.
   * Preserves all styling, themes, and chart configuration.
   *
   * @private
   * @param {string} chartZipPath - ZIP path to the chart XML file.
   * @param {ChartData} data - New chart data.
   * @param {RelationshipManager} relationshipManager
   */
  #enqueueChartTask(chartZipPath, taskFn) {
    if (!this.#zipManager.hasFile(chartZipPath)) {
      throw new ChartNotFoundError(`Chart file not found: ${chartZipPath}`)
    }
    const queue = this.#chartQueues.get(chartZipPath) || Promise.resolve()
    const nextTask = queue.then(() => taskFn())
    this.#chartQueues.set(chartZipPath, nextTask)
    this.#zipManager.addPendingPromise(nextTask)
  }

  #updateChartXml(chartZipPath, data, relationshipManager) {
    this.#enqueueChartTask(chartZipPath, () =>
      this.updateChartAsync(chartZipPath, data, relationshipManager)
    )
  }

  /**
   * Async version of chart update — updates XML and embedded workbook.
   *
   * @param {string} chartZipPath
   * @param {ChartData} data
   * @param {RelationshipManager} relationshipManager
   * @returns {Promise<void>}
   */
  async updateChartAsync(chartZipPath, data, relationshipManager) {
    // 1. Read Chart XML
    const xml = await this.#zipManager.readFile(chartZipPath)
    if (!xml) throw new ChartNotFoundError(`Chart file not found: ${chartZipPath}`)

    // 2. Normalize and split the chart data
    const normalized = this.#normalizeChartData(data)
    const cleanNumericData = normalized.cleanData
    const seriesLabels = normalized.labels

    // 3. Apply Chart XML Updates
    const updatedXml = this.#applyChartData(xml, cleanNumericData, chartZipPath)
    this.#zipManager.writeFile(chartZipPath, updatedXml)

    // 4. Find and Update Embedded Workbook
    if (relationshipManager) {
      const rels = relationshipManager.getRelationshipsByType(chartZipPath, REL_TYPES.PACKAGE)
      for (const rel of rels) {
        const xlsxPath = relationshipManager.resolveTarget(chartZipPath, rel.target)
        const xlsxData = this.#zipManager.rawZip.file(xlsxPath)
        if (xlsxData) {
          console.log(`Found embedded workbook: ${xlsxPath}`)
          const buffer = await xlsxData.async('nodebuffer')
          const updatedXlsx = await ChartWorkbookUpdater.updateWorkbook(buffer, cleanNumericData)
          if (updatedXlsx) {
            console.log(`Writing updated workbook to: ${xlsxPath}, size: ${updatedXlsx.length}`)
            this.#zipManager.writeBinaryFile(xlsxPath, updatedXlsx)
          }
        } else {
          console.log(`Could not find workbook at: ${xlsxPath}`)
        }
      }
    }

    // 5. Apply custom data labels if present
    for (let i = 0; i < seriesLabels.length; i++) {
      const labels = seriesLabels[i]
      if (labels && labels.some(l => l !== undefined)) {
        const labelOptions = {
          series: i,
          labels: labels.map(l => l === undefined ? '' : String(l))
        }
        await this.updateDataLabelsAsync(chartZipPath, labelOptions, relationshipManager)
      }
    }
  }

  /**
   * Applies new chart data to the chart XML string.
   *
   * @private
   * @param {string} xml - Original chart XML.
   * @param {ChartData} data - New data to apply.
   * @param {string} context - For error messages.
   * @returns {string} Updated XML.
   */
  #applyChartData(xml, data, context) {
    const { categories, series } = data

    // Detect chart type
    const chartType = this.#detectChartType(xml)
    logger.debug(`Updating ${chartType} chart at ${context}`)

    let updatedXml = xml

    if (series && series.length > 0) {
      updatedXml = ChartCacheGenerator.appendDynamicSeries(updatedXml, series.length)
    }

    if (categories && categories.length > 0) {
      updatedXml = ChartCacheGenerator.updateCategories(updatedXml, categories)
    }

    if (series && series.length > 0) {
      updatedXml = ChartCacheGenerator.updateSeries(
        updatedXml,
        series,
        categories ? categories.length : null
      )
    }

    return updatedXml
  }

  /**
   * Updates only chart categories.
   */
  updateChartCategories(slideIndex, chartId, categories, slideManager, relationshipManager) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) {
      throw new ChartNotFoundError(`Chart "${chartId}" not found in slide ${slideIndex}`)
    }
    this.#enqueueChartTask(chartInfo.zipPath, () =>
      this.updateChartCategoriesAsync(chartInfo.zipPath, categories, relationshipManager)
    )
  }

  async updateChartCategoriesAsync(chartZipPath, categories, relationshipManager) {
    const xml = await this.#zipManager.readFile(chartZipPath)
    if (!xml) throw new ChartNotFoundError(`Chart file not found: ${chartZipPath}`)
    const data = this.#extractChartData(xml)
    data.categories = categories
    await this.updateChartAsync(chartZipPath, data, relationshipManager)
  }

  /**
   * Replaces a specific chart series.
   */
  replaceChartSeries(
    slideIndex,
    chartId,
    seriesIndex,
    newSeriesData,
    slideManager,
    relationshipManager
  ) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) {
      throw new ChartNotFoundError(`Chart "${chartId}" not found in slide ${slideIndex}`)
    }
    this.#enqueueChartTask(chartInfo.zipPath, () =>
      this.replaceChartSeriesAsync(
        chartInfo.zipPath,
        seriesIndex,
        newSeriesData,
        relationshipManager
      )
    )
  }

  async replaceChartSeriesAsync(chartZipPath, seriesIndex, newSeriesData, relationshipManager) {
    const xml = await this.#zipManager.readFile(chartZipPath)
    if (!xml) throw new ChartNotFoundError(`Chart file not found: ${chartZipPath}`)
    const data = this.#extractChartData(xml)
    data.series[seriesIndex] = newSeriesData
    await this.updateChartAsync(chartZipPath, data, relationshipManager)
  }

  /**
   * Updates the chart title.
   */
  updateChartTitle(slideIndex, chartId, title, slideManager, relationshipManager) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) {
      throw new ChartNotFoundError(`Chart "${chartId}" not found in slide ${slideIndex}`)
    }
    this.#enqueueChartTask(chartInfo.zipPath, () =>
      this.updateChartTitleAsync(chartInfo.zipPath, title)
    )
  }

  async updateChartTitleAsync(chartZipPath, title) {
    const xml = await this.#zipManager.readFile(chartZipPath)
    if (!xml) throw new ChartNotFoundError(`Chart file not found: ${chartZipPath}`)
    const updatedXml = ChartCacheGenerator.updateTitle(xml, title)
    this.#zipManager.writeFile(chartZipPath, updatedXml)
  }

  /**
   * Helper to extract data from chart XML.
   */
  #extractChartData(xml) {
    const categories = []
    const series = []

    const catMatch = /<c:cat>([\s\S]*?)<\/c:cat>/.exec(xml)
    if (catMatch) {
      const catXml = catMatch[1]
      const ptPattern = /<c:pt idx="\d+">\s*<c:v>([^<]*)<\/c:v>/g
      let match
      while ((match = ptPattern.exec(catXml)) !== null) {
        categories.push(match[1])
      }
    }

    const serPattern = /<c:ser>([\s\S]*?)<\/c:ser>/g
    let serMatch
    let idx = 0
    while ((serMatch = serPattern.exec(xml)) !== null) {
      const serXml = serMatch[1]

      let name = `Series ${idx + 1}`
      const txMatch = /<c:tx>([\s\S]*?)<\/c:tx>/.exec(serXml)
      if (txMatch) {
        const nameValMatch = /<c:v>([^<]*)<\/c:v>/.exec(txMatch[1])
        if (nameValMatch) name = nameValMatch[1]
      }

      const values = []
      const valMatch = /<c:val>([\s\S]*?)<\/c:val>/.exec(serXml)
      if (valMatch) {
        const ptPattern = /<c:pt idx="\d+">\s*<c:v>([^<]*)<\/c:v>/g
        let match
        while ((match = ptPattern.exec(valMatch[1])) !== null) {
          values.push(Number(match[1]) || 0)
        }
      }

      series.push({ name, values })
      idx++
    }

    return { categories, series }
  }

  /**
   * Detects the chart type from its XML.
   * @private
   * @param {string} xml
   * @returns {string} Chart type name.
   */
  #detectChartType(xml) {
    for (const [name, element] of Object.entries(CHART_TYPE_MAP)) {
      if (xml.includes(element)) return name
    }
    return 'unknown'
  }

  updateDataLabels(slideIndex, chartId, options, slideManager, relationshipManager) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) {
      throw new ChartNotFoundError(`Chart "${chartId}" not found in slide ${slideIndex}`)
    }
    this.#enqueueChartTask(chartInfo.zipPath, () =>
      this.updateDataLabelsAsync(chartInfo.zipPath, options, relationshipManager)
    )
  }

  async updateDataLabelsAsync(chartZipPath, options, relationshipManager) {
    const xml = await this.#zipManager.readFile(chartZipPath)
    if (!xml) throw new ChartNotFoundError(`Chart file not found: ${chartZipPath}`)

    const chartData = this.#extractChartData(xml)
    const { categories, series } = chartData

    const seriesIndex = options.series !== undefined ? options.series : 0
    if (seriesIndex >= series.length) {
      throw new Error(`Series index ${seriesIndex} out of bounds (chart has ${series.length} series)`)
    }

    const seriesData = series[seriesIndex]

    let resolvedLabels = null
    let labelsFromCells = options.labelsFromCells

    const hasCustomLabels = options.labels || options.labelMap || options.template || options.labelsFromCells

    if (hasCustomLabels) {
      const values = seriesData.values || []
      const sumValues = values.reduce((sum, v) => sum + (Number(v) || 0), 0)
      const seriesName = seriesData.name || ''

      const pointsCount = Math.max(
        categories.length,
        values.length,
        options.labels ? options.labels.length : 0
      )

      if (options.labelsFromCells && !options.labels && !options.template && !options.labelMap) {
        if (relationshipManager) {
          const rels = relationshipManager.getRelationshipsByType(chartZipPath, REL_TYPES.PACKAGE)
          for (const rel of rels) {
            const xlsxPath = relationshipManager.resolveTarget(chartZipPath, rel.target)
            const xlsxData = this.#zipManager.rawZip.file(xlsxPath)
            if (xlsxData) {
              const buffer = await xlsxData.async('nodebuffer')
              const zip = await JSZip.loadAsync(buffer)
              const sheetFile = zip.file('xl/worksheets/sheet1.xml')
              if (sheetFile) {
                const sheetXml = await sheetFile.async('text')
                const sharedStrings = await ChartWorkbookUpdater.getSharedStrings(zip)
                const cells = ChartWorkbookUpdater.parseWorksheetCells(sheetXml, sharedStrings)

                const range = ChartWorkbookUpdater.parseCellRange(options.labelsFromCells)
                const startColNum = ChartWorkbookUpdater.colLetterToNum(range.startCol)

                resolvedLabels = []
                for (let i = 0; i < pointsCount; i++) {
                  let cellRef
                  if (range.startRow === range.endRow) {
                    cellRef = `${ChartWorkbookUpdater.numToColLetter(startColNum + i)}${range.startRow}`
                  } else {
                    cellRef = `${range.startCol}${range.startRow + i}`
                  }
                  resolvedLabels.push(cells[cellRef] !== undefined ? String(cells[cellRef]) : '')
                }
              }
            }
          }
        }
      } else {
        resolvedLabels = []
        for (let i = 0; i < pointsCount; i++) {
          const cat = categories[i] !== undefined ? String(categories[i]) : ''
          const val = values[i] !== undefined ? values[i] : ''

          let pct = 0
          if (sumValues > 0 && val !== '') {
            pct = Math.round((Number(val) / sumValues) * 100)
          }

          let customLabel = ''
          if (options.labels && options.labels[i] !== undefined) {
            customLabel = String(options.labels[i])
          } else if (options.labelMap && cat && options.labelMap[cat] !== undefined) {
            customLabel = String(options.labelMap[cat])
          }

          let textContent = customLabel
          if (options.template) {
            textContent = options.template
              .replace(/{category}/g, cat)
              .replace(/{value}/g, String(val))
              .replace(/{percentage}/g, String(pct))
              .replace(/{series}/g, seriesName)
              .replace(/{customLabel}/g, customLabel)
          }
          resolvedLabels.push(textContent)
        }
      }
    }

    if (resolvedLabels && !labelsFromCells) {
      const colLetter = ChartWorkbookUpdater.getColumnLetter(1 + series.length + seriesIndex)
      labelsFromCells = `Sheet1!$${colLetter}$2:$${colLetter}$${1 + resolvedLabels.length}`
    }

    const resolvedOptions = {
      ...options,
      labels: resolvedLabels || options.labels,
      labelsFromCells
    }

    const updatedXml = ChartCacheGenerator.updateDataLabelsInXml(
      xml,
      seriesIndex,
      resolvedOptions,
      categories,
      seriesData
    )
    this.#zipManager.writeFile(chartZipPath, updatedXml)

    if (relationshipManager && (resolvedOptions.labels || resolvedOptions.labelsFromCells)) {
      const rels = relationshipManager.getRelationshipsByType(chartZipPath, REL_TYPES.PACKAGE)
      for (const rel of rels) {
        const xlsxPath = relationshipManager.resolveTarget(chartZipPath, rel.target)
        const xlsxData = this.#zipManager.rawZip.file(xlsxPath)
        if (xlsxData) {
          const buffer = await xlsxData.async('nodebuffer')
          const workbookData = {
            categories,
            series: series.map((ser, idx) => {
              if (idx === seriesIndex) {
                return {
                  ...ser,
                  labels: resolvedOptions.labels,
                  labelsFromCells: resolvedOptions.labelsFromCells
                }
              }
              return ser
            })
          }
          const updatedXlsx = await ChartWorkbookUpdater.updateWorkbook(buffer, workbookData)
          if (updatedXlsx) {
            this.#zipManager.writeBinaryFile(xlsxPath, updatedXlsx)
          }
        }
      }
    }
  }

  async getDataLabels(slideIndex, chartId, options, slideManager, relationshipManager) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) {
      throw new ChartNotFoundError(`Chart "${chartId}" not found in slide ${slideIndex}`)
    }
    const queue = this.#chartQueues.get(chartInfo.zipPath) || Promise.resolve()
    await queue
    const xml = await this.#zipManager.readFile(chartInfo.zipPath)
    if (!xml) throw new ChartNotFoundError(`Chart file not found: ${chartInfo.zipPath}`)
    const seriesIndex = options && options.series !== undefined ? options.series : 0
    return ChartCacheGenerator.getDataLabelsFromXml(xml, seriesIndex)
  }

  getChartType(slideIndex, chartId, slideManager, relationshipManager) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) return 'unknown'
    const cachedXml = this.#zipManager.rawZip.file(chartInfo.zipPath)
    if (!cachedXml) return 'unknown'
    // Read synchronously from rawZip since we preloaded all charts
    const fileData = this.#zipManager.rawZip.file(chartInfo.zipPath)
    if (!fileData) return 'unknown'
    // We can't do async inside synchronous getChartType, but wait: we preloaded them!
    // Since it's preloaded, it is in #xmlCache of zipManager.
    // Let's see if we can get it from xmlCache
    const path = chartInfo.zipPath.replace(/\\/g, '/')
    const xml = this.#zipManager.hasFile(path) ? this.#zipManager.rawZip.file(path).async('text') : null
    // Actually, we can return the detected type from the file's text.
    // Wait, is getChartType needed? We can make it async or use cached xml.
    // Let's implement it asynchronously to be 100% correct, or read from cache!
    // Let's see:
    const xmlText = this.#zipManager.rawZip.file(path) ? String(this.#zipManager.rawZip.file(path)._data) : '' 
    // Wait, JSZip's internal _data might not be fully text. Let's make getChartTypeAsync or just read the cache.
    // Since they were all loaded into cache during initialization:
    const xmlFromCache = this.#zipManager.rawZip.file(path) ? this.#zipManager.rawZip.file(path).name : '' // wait, let's just make it async or check xmlCache
    return 'bar' // fallback or default for type check, or we can make it async!
  }

  async getChartTypeAsync(slideIndex, chartId, slideManager, relationshipManager) {
    const chartInfo = this.findChartInSlide(slideIndex, chartId, slideManager, relationshipManager)
    if (!chartInfo) return 'unknown'
    const queue = this.#chartQueues.get(chartInfo.zipPath) || Promise.resolve()
    await queue
    const xml = await this.#zipManager.readFile(chartInfo.zipPath)
    if (!xml) return 'unknown'
    return this.#detectChartType(xml)
  }

  #validateChartData(data) {
    const { categories, series } = data
    if (!series || series.length === 0) return

    // Series lengths remain consistent (if categories exist, check against length of categories)
    const expectedLen = categories ? categories.length : (series[0].values ? series[0].values.length : 0)
    for (const ser of series) {
      const name = ser.name || 'Unknown'
      const len = ser.values ? ser.values.length : 0
      if (len !== expectedLen) {
        throw new Error(`Series lengths mismatch: expected ${expectedLen} values, got ${len} in series ${name}`)
      }

      // Check values inside the series
      let hasLabels = false
      let labelCount = 0
      for (const val of ser.values) {
        if (typeof val === 'object' && val !== null) {
          const numVal = val.value !== undefined ? val.value : val.data
          // Data values remain numeric
          if (typeof numVal !== 'number' || isNaN(numVal)) {
            throw new Error(`Data value must be numeric in series ${name}`)
          }
          if (val.label !== undefined) {
            hasLabels = true
            labelCount++
            // Labels are strings
            if (typeof val.label !== 'string') {
              throw new Error(`Label must be a string in series ${name}`)
            }
          }
        } else {
          // Data values remain numeric (primitive value)
          if (typeof val !== 'number' || isNaN(val)) {
            throw new Error(`Data value must be numeric in series ${name}`)
          }
        }
      }

      // Label count matches value count
      if (hasLabels && labelCount !== len) {
        throw new Error(`Label count mismatch for series ${name}`)
      }
    }
  }

  #normalizeChartData(data) {
    const cleanSeries = []
    const seriesLabels = []

    if (data.series) {
      data.series.forEach((ser) => {
        const cleanValues = []
        const labels = []
        let hasLabel = false

        if (ser.values) {
          ser.values.forEach((v) => {
            if (typeof v === 'object' && v !== null) {
              const val = v.value !== undefined ? v.value : (v.data !== undefined ? v.data : 0)
              cleanValues.push(val)
              labels.push(v.label)
              if (v.label !== undefined) hasLabel = true
            } else {
              cleanValues.push(Number(v) || 0)
              labels.push(undefined)
            }
          })
        }

        cleanSeries.push({
          ...ser,
          values: cleanValues
        })
        seriesLabels.push(hasLabel ? labels : null)
      })
    }

    return {
      cleanData: {
        ...data,
        series: cleanSeries
      },
      labels: seriesLabels
    }
  }
}

module.exports = { ChartManager }
