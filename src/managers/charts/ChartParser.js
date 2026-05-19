const { createLogger } = require('../../utils/logger.js');

const logger = createLogger('ChartParser');

class ChartParser {
  /**
   * Finds a chart's relationship ID and type in a slide's XML based on shape name/id.
   *
   * @param {string} slideXml
   * @param {string} chartId
   * @returns {{ rId: string } | null}
   */
  static findChartRIdInSlide(slideXml, chartId) {
    // Strategy 1: Look for shape with matching name (cNvPr name attribute)
    const shapeNamePattern = new RegExp(
      `<p:cNvPr[^>]*name="${chartId}"[^>]*>(?:.*?)<c:chart[^>]*r:id="(rId\\d+)"`,
      's'
    );
    const rIdMatch = shapeNamePattern.exec(slideXml);
    if (rIdMatch) {
      return { rId: rIdMatch[1] };
    }

    // Strategy 2: Find all chart graphicFrames and we will match later in manager
    return null;
  }

  /**
   * Parses the chart XML to extract series and categories for validation.
   *
   * @param {string} xml
   * @returns {Object} Data about the chart configuration
   */
  static parseChartData(xml) {
    // This could be used for validation and extracting current chart cache
    const ptCountMatch = xml.match(/<c:ptCount val="(\d+)"\/>/);
    const pointCount = ptCountMatch ? parseInt(ptCountMatch[1], 10) : 0;

    return {
      pointCount
    };
  }
}

module.exports = { ChartParser };
