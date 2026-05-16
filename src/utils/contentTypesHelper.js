/**
 * @fileoverview Content Types helper - manages [Content_Types].xml.
 *
 * The [Content_Types].xml file tells the ZIP package parser what
 * MIME type each file inside is. Every new part (slide, chart, image)
 * must be registered here.
 *
 * Structure:
 *   <Types xmlns="...">
 *     <!-- Default: applies to all files with matching extension -->
 *     <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
 *     <Default Extension="xml" ContentType="application/xml"/>
 *
 *     <!-- Override: specific part override -->
 *     <Override PartName="/ppt/slides/slide1.xml"
 *               ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
 *   </Types>
 */

import { createLogger } from './logger.js';

const logger = createLogger('ContentTypes');

/** MIME type for PPTX slide parts. */
const SLIDE_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';

/** MIME type for chart parts. */
const CHART_CONTENT_TYPE = 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml';

/**
 * Singleton helper for [Content_Types].xml manipulation.
 */
class ContentTypesHelper {
  /**
   * Adds a slide Override entry to [Content_Types].xml.
   *
   * @param {ZipManager} zipManager
   * @param {string} slideFileName - e.g., 'slide5.xml'
   */
  addSlideContentType(zipManager, slideFileName) {
    this.#addOverride(
      zipManager,
      `/ppt/slides/${slideFileName}`,
      SLIDE_CONTENT_TYPE
    );
  }

  /**
   * Adds a chart Override entry to [Content_Types].xml.
   *
   * @param {ZipManager} zipManager
   * @param {string} chartFileName - e.g., 'chart3.xml'
   */
  addChartContentType(zipManager, chartFileName) {
    this.#addOverride(
      zipManager,
      `/ppt/charts/${chartFileName}`,
      CHART_CONTENT_TYPE
    );
  }

  /**
   * Adds a Default entry for a media extension.
   *
   * @param {ZipManager} zipManager
   * @param {string} extension - File extension (e.g., 'png').
   * @param {string} mimeType - MIME type.
   */
  addMediaDefault(zipManager, extension, mimeType) {
    const xml = zipManager.rawZip.file('[Content_Types].xml');
    if (!xml) return;

    zipManager.addPendingPromise(
      xml.async('text').then(content => {
        const entry = `Extension="${extension}" ContentType="${mimeType}"`;
        if (!content.includes(entry)) {
          const updated = content.replace(
            '</Types>',
            `  <Default ${entry}/>\n</Types>`
          );
          zipManager.writeFile('[Content_Types].xml', updated);
        }
      })
    );
  }

  /**
   * Adds an Override entry to [Content_Types].xml.
   * @private
   */
  #addOverride(zipManager, partName, contentType) {
    const xmlFile = zipManager.rawZip.file('[Content_Types].xml');
    if (!xmlFile) {
      logger.warn('[Content_Types].xml not found');
      return;
    }

    zipManager.addPendingPromise(
      xmlFile.async('text').then(content => {
        const entry = `PartName="${partName}"`;
        if (!content.includes(entry)) {
          const override = `<Override PartName="${partName}" ContentType="${contentType}"/>`;
          const updated = content.replace('</Types>', `  ${override}\n</Types>`);
          zipManager.writeFile('[Content_Types].xml', updated);
          logger.debug(`Registered content type for ${partName}`);
        }
      })
    );
  }
}

export const contentTypesHelper = new ContentTypesHelper();
