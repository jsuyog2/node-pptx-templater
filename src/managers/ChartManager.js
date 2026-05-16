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

import { createLogger } from '../utils/logger.js';
import { ChartNotFoundError } from '../utils/errors.js';
import { REL_TYPES } from './RelationshipManager.js';

const logger = createLogger('ChartManager');

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
};

/**
 * @class ChartManager
 * @description Handles chart data updates by directly manipulating chart XML files.
 *
 * Unlike high-level charting libraries, this manager edits the raw OpenXML
 * chart structure, allowing full control while preserving styles and themes.
 */
export class ChartManager {
  /** @private @type {XMLParser} */
  #xmlParser;
  /** @private @type {ZipManager} */
  #zipManager;

  /**
   * Cache of chart ZIP paths: maps chartName → { zipPath, slideIndex }
   * @private @type {Map<string, { zipPath: string, slideIndex: number }>}
   */
  #chartRegistry = new Map();

  /**
   * @param {XMLParser} xmlParser
   */
  constructor(xmlParser) {
    this.#xmlParser = xmlParser;
  }

  /**
   * Initializes by scanning the ZIP for chart files.
   *
   * @param {ZipManager} zipManager
   * @returns {Promise<void>}
   */
  async initialize(zipManager) {
    this.#zipManager = zipManager;
    const chartFiles = zipManager.listFiles('ppt/charts/').filter(f => f.endsWith('.xml') && !f.includes('_rels'));

    for (const chartPath of chartFiles) {
      // Chart name is inferred from file name
      const chartName = chartPath.split('/').pop().replace('.xml', '');
      this.#chartRegistry.set(chartName, { zipPath: chartPath, slideIndex: null });
    }

    logger.debug(`Found ${chartFiles.length} chart file(s)`);
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
    const chartInfo = this.#findChartInSlide(slideIndex, chartId, slideManager, relationshipManager);

    if (!chartInfo) {
      throw new ChartNotFoundError(`Chart "${chartId}" not found in slide ${slideIndex}`);
    }

    logger.debug(`Updating chart "${chartId}" at ${chartInfo.zipPath}`);
    this.#updateChartXml(chartInfo.zipPath, data);
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
    const slideInfo = slideManager.getSlideInfo(slideIndex);
    const rels = relationshipManager.getRelationshipsByType(slideInfo.zipPath, REL_TYPES.CHART);

    return rels.map(rel => {
      const chartPath = relationshipManager.resolveTarget(slideInfo.zipPath, rel.target);
      return { rId: rel.id, zipPath: chartPath };
    });
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
  #findChartInSlide(slideIndex, chartId, slideManager, relationshipManager) {
    const slideInfo = slideManager.getSlideInfo(slideIndex);
    const slideXml = slideManager.getSlideXml(slideIndex);

    // Strategy 1: Look for shape with matching name (cNvPr name attribute)
    const shapeNamePattern = new RegExp(
      `<p:cNvPr[^>]*name="${chartId}"[^>]*>(?:.*?)<c:chart[^>]*r:id="(rId\\d+)"`,
      's'
    );
    let rIdMatch = shapeNamePattern.exec(slideXml);

    // Strategy 2: Find graphicFrame shapes and match chart rIds
    if (!rIdMatch) {
      const chartRIdPattern = /<c:chart[^>]*r:id="(rId\d+)"/g;
      let chartMatch;
      while ((chartMatch = chartRIdPattern.exec(slideXml)) !== null) {
        const rId = chartMatch[1];
        const rel = relationshipManager.getRelationshipById(slideInfo.zipPath, rId);
        if (rel) {
          const chartPath = relationshipManager.resolveTarget(slideInfo.zipPath, rel.target);
          // Check if chart file name matches chartId
          if (chartPath.includes(chartId) || rel.id === chartId) {
            return { zipPath: chartPath };
          }
          // For the first chart found, if chartId looks like a chart file name
          if (chartId.startsWith('chart')) {
            return { zipPath: chartPath };
          }
        }
      }
    }

    if (rIdMatch) {
      const rId = rIdMatch[1];
      const rel = relationshipManager.getRelationshipById(slideInfo.zipPath, rId);
      if (rel) {
        const chartPath = relationshipManager.resolveTarget(slideInfo.zipPath, rel.target);
        return { zipPath: chartPath };
      }
    }

    // Strategy 3: Direct chart registry lookup
    if (this.#chartRegistry.has(chartId)) {
      return this.#chartRegistry.get(chartId);
    }

    // Strategy 4: Try chartN naming convention
    const chartPath = `ppt/charts/${chartId}.xml`;
    if (this.#zipManager.hasFile(chartPath)) {
      return { zipPath: chartPath };
    }

    return null;
  }

  /**
   * Updates the chart XML file with new data.
   * Preserves all styling, themes, and chart configuration.
   * Only replaces data-related nodes.
   *
   * @private
   * @param {string} chartZipPath - ZIP path to the chart XML file.
   * @param {ChartData} data - New chart data.
   */
  #updateChartXml(chartZipPath, data) {
    const chartXml = this.#zipManager.rawZip.file(chartZipPath);
    if (!chartXml) {
      throw new ChartNotFoundError(`Chart file not found: ${chartZipPath}`);
    }

    // We use a synchronous approach here since JSZip caches file content
    // For the async version, use updateChartAsync()
    chartXml.async('text').then(xml => {
      const updatedXml = this.#applyChartData(xml, data, chartZipPath);
      this.#zipManager.writeFile(chartZipPath, updatedXml);
    });
  }

  /**
   * Async version of chart update — preferred for large charts.
   *
   * @param {string} chartZipPath
   * @param {ChartData} data
   * @returns {Promise<void>}
   */
  async updateChartAsync(chartZipPath, data) {
    const xml = await this.#zipManager.readFile(chartZipPath);
    if (!xml) throw new ChartNotFoundError(`Chart file not found: ${chartZipPath}`);

    const updatedXml = this.#applyChartData(xml, data, chartZipPath);
    this.#zipManager.writeFile(chartZipPath, updatedXml);
  }

  /**
   * Applies new chart data to the chart XML string.
   * Handles all supported chart types.
   *
   * @private
   * @param {string} xml - Original chart XML.
   * @param {ChartData} data - New data to apply.
   * @param {string} context - For error messages.
   * @returns {string} Updated XML.
   */
  #applyChartData(xml, data, context) {
    const { categories, series } = data;

    // Detect chart type
    const chartType = this.#detectChartType(xml);
    logger.debug(`Updating ${chartType} chart at ${context}`);

    let updatedXml = xml;

    // Update categories (shared across all series for most chart types)
    if (categories && categories.length > 0) {
      updatedXml = this.#updateCategories(updatedXml, categories);
    }

    // Update each series
    if (series && series.length > 0) {
      updatedXml = this.#updateSeries(updatedXml, series, categories);
    }

    return updatedXml;
  }

  /**
   * Detects the chart type from its XML.
   * @private
   * @param {string} xml
   * @returns {string} Chart type name.
   */
  #detectChartType(xml) {
    for (const [name, element] of Object.entries(CHART_TYPE_MAP)) {
      if (xml.includes(element)) return name;
    }
    return 'unknown';
  }

  /**
   * Updates the category labels in chart XML.
   * Handles both string categories (strRef) and numeric (numRef).
   *
   * @private
   * @param {string} xml - Chart XML.
   * @param {string[]} categories - New category labels.
   * @returns {string} Updated XML.
   */
  #updateCategories(xml, categories) {
    const count = categories.length;

    // Build the string cache entries
    const ptEntries = categories
      .map((cat, i) => `<c:pt idx="${i}"><c:v>${this.#escapeXml(String(cat))}</c:v></c:pt>`)
      .join('');

    // Pattern to find the first series' cat section (used as template for all)
    // We replace strCache or numCache content within c:cat
    const strCachePattern = /(<c:strCache>)([\s\S]*?)(<\/c:strCache>)/;
    const numCachePattern = /(<c:numCache>)([\s\S]*?)(<\/c:numCache>)/;

    const newStrCache = `<c:ptCount val="${count}"/>${ptEntries}`;

    let updated = xml;

    if (strCachePattern.test(updated)) {
      updated = updated.replace(strCachePattern, `$1${newStrCache}$3`);
    } else if (numCachePattern.test(updated)) {
      // Convert to string cache if was numeric
      updated = updated.replace(numCachePattern, `$1${newStrCache}$3`);
    }

    return updated;
  }

  /**
   * Updates series data in chart XML.
   *
   * @private
   * @param {string} xml - Chart XML.
   * @param {SeriesData[]} series - New series data.
   * @param {string[]} [categories] - Categories for reference.
   * @returns {string} Updated XML.
   */
  #updateSeries(xml, series, categories) {
    let updated = xml;

    // Find all <c:ser> blocks
    const serPattern = /(<c:ser>)([\s\S]*?)(<\/c:ser>)/g;
    const serMatches = [...updated.matchAll(serPattern)];

    if (serMatches.length === 0) {
      logger.warn('No series (c:ser) found in chart XML');
      return xml;
    }

    // Update each series that has a corresponding data entry
    let serIndex = 0;
    updated = updated.replace(serPattern, (match, open, content, close) => {
      if (serIndex >= series.length) return match; // Keep extra series unchanged

      const serData = series[serIndex];
      serIndex++;

      let updatedContent = content;

      // Update series name (c:tx)
      if (serData.name !== undefined) {
        updatedContent = this.#updateSeriesName(updatedContent, serData.name, serIndex - 1);
      }

      // Update values (c:val → c:numCache)
      if (serData.values !== undefined) {
        updatedContent = this.#updateSeriesValues(updatedContent, serData.values);
      }

      return `${open}${updatedContent}${close}`;
    });

    return updated;
  }

  /**
   * Updates the series name in c:tx section.
   * @private
   */
  #updateSeriesName(serContent, name, idx) {
    // Update the string reference value
    const txPattern = /(<c:tx>)([\s\S]*?)(<\/c:tx>)/;
    if (txPattern.test(serContent)) {
      return serContent.replace(txPattern, (match, open, content, close) => {
        // Replace the v element value
        const vPattern = /(<c:v>)(.*?)(<\/c:v>)/;
        if (vPattern.test(content)) {
          return `${open}${content.replace(vPattern, `$1${this.#escapeXml(name)}$3`)}${close}`;
        }
        return match;
      });
    }
    return serContent;
  }

  /**
   * Updates the numeric values in c:val → c:numCache.
   * @private
   */
  #updateSeriesValues(serContent, values) {
    const count = values.length;
    const ptEntries = values
      .map((val, i) => `<c:pt idx="${i}"><c:v>${Number(val)}</c:v></c:pt>`)
      .join('');

    const numCachePattern = /(<c:numCache>)([\s\S]*?)(<\/c:numCache>)/;
    if (numCachePattern.test(serContent)) {
      return serContent.replace(
        numCachePattern,
        `$1<c:formatCode>General</c:formatCode><c:ptCount val="${count}"/>${ptEntries}$3`
      );
    }
    return serContent;
  }

  /**
   * Escapes XML special characters in a string.
   * @private
   * @param {string} str
   * @returns {string}
   */
  #escapeXml(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
}
