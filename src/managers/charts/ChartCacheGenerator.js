import { createLogger } from '../../utils/logger.js';
import { ChartWorkbookUpdater } from './ChartWorkbookUpdater.js';

const logger = createLogger('ChartCacheGenerator');

export class ChartCacheGenerator {
  /**
   * Generates a string cache XML string (used for categories or series names).
   */
  static generateStrCache(values) {
    const ptEntries = values
      .map((val, i) => `<c:pt idx="${i}"><c:v>${this.#escapeXml(String(val))}</c:v></c:pt>`)
      .join('');
    return `<c:strCache><c:ptCount val="${values.length}"/>${ptEntries}</c:strCache>`;
  }

  /**
   * Generates a numeric cache XML string.
   */
  static generateNumCache(values) {
    const ptEntries = values
      .map((val, i) => `<c:pt idx="${i}"><c:v>${Number(val) || 0}</c:v></c:pt>`)
      .join('');
    return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${ptEntries}</c:numCache>`;
  }

  /**
   * Updates category cache and formulas in a chart XML.
   *
   * @param {string} xml - Raw chart XML.
   * @param {string[]} categories - Array of categories.
   * @param {string} sheetName - Target worksheet name.
   */
  static updateCategories(xml, categories, sheetName = 'Sheet1') {
    const count = categories.length;

    // Formula for categories: Sheet1!$A$2:$A$N
    const formula = ChartWorkbookUpdater.getFormulaRange(sheetName, 2, 0, count + 1, 0);
    const newStrCache = this.generateStrCache(categories);

    // Replace the entire <c:cat> block to ensure correct formula and cache
    const catPattern = /(<c:cat>)([\s\S]*?)(<\/c:cat>)/g;

    return xml.replace(catPattern, (match, open, content, close) => {
      // Reconstruct the cat block
      // Try to determine if it used strRef or numRef originally
      let refTag = content.includes('<c:numRef>') ? 'numRef' : 'strRef';
      // But typically categories are strings. Let's use strRef.
      refTag = 'strRef';

      return `${open}<c:${refTag}><c:f>${formula}</c:f>${newStrCache}</c:${refTag}>${close}`;
    });
  }

  /**
   * Updates series names and values in chart XML.
   */
  static updateSeries(xml, series, categoriesLength, sheetName = 'Sheet1') {
    let updated = xml;
    const serPattern = /(<c:ser>)([\s\S]*?)(<\/c:ser>)/g;
    const serMatches = [...updated.matchAll(serPattern)];

    if (serMatches.length === 0) return xml;

    let serIndex = 0;
    updated = updated.replace(serPattern, (match, open, content, close) => {
      if (serIndex >= series.length) {
        // If there are more series templates than data, we could drop them,
        // but replacing with an empty string might break the XML layout if we're not careful.
        // Actually, removing extra series is requested: "Allow removing old series".
        return '';
      }

      const serData = series[serIndex];
      const colIndex = serIndex + 1; // Series data starts in column B (1)
      serIndex++;

      let updatedContent = content;

      // 1. Update Series Name (c:tx)
      if (serData.name !== undefined) {
        const nameFormula = ChartWorkbookUpdater.getFormulaSingleCell(sheetName, 1, colIndex);
        const nameCache = `<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${this.#escapeXml(serData.name)}</c:v></c:pt></c:strCache>`;

        const txPattern = /(<c:tx>)([\s\S]*?)(<\/c:tx>)/;
        if (txPattern.test(updatedContent)) {
          updatedContent = updatedContent.replace(txPattern, (match, p1, p2, p3) => {
            return `${p1}<c:strRef><c:f>${nameFormula}</c:f>${nameCache}</c:strRef>${p3}`;
          });
        } else {
          // Some charts don't have <c:tx>, we prepend it after <c:order> or <c:idx>
          const insertAfter = /(<c:order[^>]*>)/;
          if (insertAfter.test(updatedContent)) {
            updatedContent = updatedContent.replace(insertAfter, (match, p1) => {
              return `${p1}<c:tx><c:strRef><c:f>${nameFormula}</c:f>${nameCache}</c:strRef></c:tx>`;
            });
          }
        }
      }

      // 2. Update Series Values (c:val)
      if (serData.values !== undefined) {
        const valuesCount = categoriesLength || serData.values.length;
        const valFormula = ChartWorkbookUpdater.getFormulaRange(sheetName, 2, colIndex, valuesCount + 1, colIndex);
        const valCache = this.generateNumCache(serData.values);

        const valPattern = /(<c:val>)([\s\S]*?)(<\/c:val>)/;
        if (valPattern.test(updatedContent)) {
          updatedContent = updatedContent.replace(valPattern, (match, p1, p2, p3) => {
            return `${p1}<c:numRef><c:f>${valFormula}</c:f>${valCache}</c:numRef>${p3}`;
          });
        }
      }

      return `${open}${updatedContent}${close}`;
    });

    return updated;
  }

  /**
   * Clones a series template to support dynamic series addition.
   */
  static appendDynamicSeries(xml, targetCount) {
    const serPattern = /(<c:ser>)([\s\S]*?)(<\/c:ser>)/g;
    const matches = [...xml.matchAll(serPattern)];
    if (matches.length === 0 || matches.length >= targetCount) return xml;

    // Use the last series as a template to clone
    const templateMatch = matches[matches.length - 1];
    const template = templateMatch[0];

    // Find the end of the last series
    const lastIndex = templateMatch.index + template.length;

    let newSeriesBlocks = '';
    for (let i = matches.length; i < targetCount; i++) {
      let clone = template;
      // Update c:idx and c:order
      clone = clone.replace(/(<c:idx val=")\d+("\/>)/g, `$1${i}$2`);
      clone = clone.replace(/(<c:order val=")\d+("\/>)/g, `$1${i}$2`);
      newSeriesBlocks += clone;
    }

    return xml.substring(0, lastIndex) + newSeriesBlocks + xml.substring(lastIndex);
  }

  static #escapeXml(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
}
