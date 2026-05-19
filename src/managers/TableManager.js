/**
 * @fileoverview TableManager - Updates table data in PPTX slides.
 *
 * Tables in OpenXML PPTX:
 * ─────────────────────────────────────────────────────────────────
 * Tables are stored as a:tbl inside a p:graphicFrame shape.
 *
 *   <p:graphicFrame>
 *     <p:nvGraphicFramePr>
 *       <p:cNvPr id="6" name="employees-table"/>   ← table name
 *     </p:nvGraphicFramePr>
 *     <a:graphic>
 *       <a:graphicData uri="...table">
 *         <a:tbl>
 *           <a:tblGrid>
 *             <a:gridCol w="2286000"/>    ← column widths
 *           </a:tblGrid>
 *           <a:tr h="370840">             ← row (height in EMUs)
 *             <a:tc>                      ← table cell
 *               <a:txBody>
 *                 <a:p>
 *                   <a:r><a:t>Cell text</a:t></a:r>
 *                 </a:p>
 *               </a:txBody>
 *               <a:tcPr/>                 ← cell properties (borders, fills)
 *             </a:tc>
 *           </a:tr>
 *         </a:tbl>
 *       </a:graphicData>
 *     </a:graphic>
 *   </p:graphicFrame>
 *
 * EMU (English Metric Units):
 *  - 1 inch = 914400 EMU
 *  - 1 pt = 12700 EMU
 *  - 1 cm = 360000 EMU
 *
 * Key challenge: Preserving cell formatting while replacing text.
 * We ONLY modify the <a:t> text nodes, keeping all <a:tcPr>, <a:rPr>, etc.
 *
 * Merged cells use:
 *  - <a:tc gridSpan="2"> → horizontal merge (spans 2 columns)
 *  - <a:tc rowSpan="2"> → vertical merge (spans 2 rows)
 *  - <a:tc hMerge="1"> → continuation of horizontal merge
 *  - <a:tc vMerge="1"> → continuation of vertical merge
 */

const { createLogger } = require('../utils/logger.js');
const { TableNotFoundError } = require('../utils/errors.js');

const logger = createLogger('TableManager');

/**
 * @class TableManager
 * @description Handles table data replacement in PPTX slides.
 *
 * The key design principle is "preserve formatting, replace content".
 * We never touch table styles, borders, or fonts — only the text.
 */
class TableManager {
  /** @private @type {XMLParser} */
  #xmlParser;

  /**
   * @param {XMLParser} xmlParser
   */
  constructor(xmlParser) {
    this.#xmlParser = xmlParser;
  }

  /**
   * Updates a table with new row data.
   * Finds the table by name/ID and replaces its row content.
   * Preserves all formatting properties.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {string} tableId - Table name (from shape's cNvPr name attribute).
   * @param {string[][]} rows - 2D array of cell values [row][col].
   * @param {SlideManager} slideManager
   * @throws {TableNotFoundError} If the table is not found.
   */
  updateTable(slideIndex, tableId, rows, slideManager) {
    const slideXml = slideManager.getSlideXml(slideIndex);
    const updatedXml = this.#updateTableInXml(slideXml, tableId, rows);

    if (updatedXml === null) {
      throw new TableNotFoundError(`Table "${tableId}" not found in slide ${slideIndex}`);
    }

    slideManager.setSlideXml(slideIndex, updatedXml);
    logger.debug(`Updated table "${tableId}" with ${rows.length} rows`);
  }

  /**
   * Updates a table within raw XML.
   * Returns null if the table was not found.
   *
   * @private
   * @param {string} slideXml - Slide XML content.
   * @param {string} tableId - Table name to find.
   * @param {string[][]} rows - New row data.
   * @returns {string|null} Updated XML or null if not found.
   */
  #updateTableInXml(slideXml, tableId, rows) {
    // Step 1: Find the graphicFrame containing our table
    // We look for the cNvPr name attribute matching tableId
    const framePattern = new RegExp(
      `(<p:graphicFrame>(?:(?!<\\/p:graphicFrame>)[\\s\\S])*?<p:cNvPr[^>]*name="${this.#escapeRegex(tableId)}"[^>]*>[\\s\\S]*?<\\/p:graphicFrame>)`,
      'g'
    );

    // Alternative: more robust approach — find graphicFrames with table data
    let found = false;
    let updatedXml = slideXml;

    // Strategy 1: Find by name attribute
    const namePattern = new RegExp(`name="${this.#escapeRegex(tableId)}"`, 'g');
    const nameMatch = namePattern.exec(slideXml);

    if (nameMatch) {
      // Find the graphicFrame containing this name
      const frameStart = slideXml.lastIndexOf('<p:graphicFrame>', nameMatch.index);
      const frameEnd = this.#findClosingTag(slideXml, '</p:graphicFrame>', nameMatch.index);

      if (frameStart !== -1 && frameEnd !== -1) {
        const frameXml = slideXml.substring(frameStart, frameEnd + '</p:graphicFrame>'.length);
        const updatedFrame = this.#updateTableRows(frameXml, rows);

        updatedXml = slideXml.substring(0, frameStart) + updatedFrame +
          slideXml.substring(frameEnd + '</p:graphicFrame>'.length);
        found = true;
      }
    }

    // Strategy 2: Find by shape ID (tableId is numeric)
    if (!found && /^\d+$/.test(tableId)) {
      const idPattern = new RegExp(`id="${tableId}"`, 'g');
      const idMatch = idPattern.exec(slideXml);
      if (idMatch) {
        const frameStart = slideXml.lastIndexOf('<p:graphicFrame>', idMatch.index);
        const frameEnd = this.#findClosingTag(slideXml, '</p:graphicFrame>', idMatch.index);

        if (frameStart !== -1 && frameEnd !== -1) {
          const frameXml = slideXml.substring(frameStart, frameEnd + '</p:graphicFrame>'.length);
          const updatedFrame = this.#updateTableRows(frameXml, rows);
          updatedXml = slideXml.substring(0, frameStart) + updatedFrame +
            slideXml.substring(frameEnd + '</p:graphicFrame>'.length);
          found = true;
        }
      }
    }

    // Strategy 3: Find first table in slide
    if (!found && tableId === 'first') {
      const tableStart = slideXml.indexOf('<a:tbl>');
      const frameStart = slideXml.lastIndexOf('<p:graphicFrame>', tableStart);
      const frameEnd = this.#findClosingTag(slideXml, '</p:graphicFrame>', tableStart);

      if (frameStart !== -1 && frameEnd !== -1) {
        const frameXml = slideXml.substring(frameStart, frameEnd + '</p:graphicFrame>'.length);
        const updatedFrame = this.#updateTableRows(frameXml, rows);
        updatedXml = slideXml.substring(0, frameStart) + updatedFrame +
          slideXml.substring(frameEnd + '</p:graphicFrame>'.length);
        found = true;
      }
    }

    return found ? updatedXml : null;
  }

  /**
   * Replaces the table rows within a graphicFrame XML snippet.
   * Preserves the first row's styling as a template for new rows.
   *
   * @private
   * @param {string} frameXml - XML of the graphicFrame containing the table.
   * @param {string[][]} rows - New data rows.
   * @returns {string} Updated frame XML.
   */
  #updateTableRows(frameXml, rows) {
    // Extract all existing rows
    const existingRows = this.#extractAllRows(frameXml);

    if (existingRows.length === 0) {
      logger.warn('No rows found in table');
      return frameXml;
    }

    // Use first row as header template, second as data row template (if available)
    const headerTemplate = existingRows[0];
    const dataTemplate = existingRows[1] || existingRows[0];

    // Build new rows
    const newRowsXml = rows.map((rowData, rowIdx) => {
      const template = rowIdx === 0 ? headerTemplate : dataTemplate;
      return this.#buildRow(template, rowData);
    }).join('');

    // Replace all existing rows with new rows
    const tblStart = frameXml.indexOf('<a:tbl>');
    const tblEnd = frameXml.lastIndexOf('</a:tbl>') + '</a:tbl>'.length;
    const tblXml = frameXml.substring(tblStart, tblEnd);

    // Find where rows begin (after tblGrid) and end (before </a:tbl>)
    const firstRowStart = tblXml.indexOf('<a:tr ') !== -1
      ? tblXml.indexOf('<a:tr ')
      : tblXml.indexOf('<a:tr>');
    const lastRowEnd = tblXml.lastIndexOf('</a:tr>') + '</a:tr>'.length;

    if (firstRowStart === -1 || lastRowEnd === -1) {
      return frameXml;
    }

    const newTblXml =
      tblXml.substring(0, firstRowStart) +
      newRowsXml +
      tblXml.substring(lastRowEnd);

    return frameXml.substring(0, tblStart) + newTblXml + frameXml.substring(tblEnd);
  }

  /**
   * Extracts all <a:tr> row XML strings from a table.
   * @private
   * @param {string} tableXml
   * @returns {string[]}
   */
  #extractAllRows(tableXml) {
    const rows = [];
    let searchFrom = 0;

    while (true) {
      const rowStart = tableXml.indexOf('<a:tr', searchFrom);
      if (rowStart === -1) break;

      const rowEnd = tableXml.indexOf('</a:tr>', rowStart);
      if (rowEnd === -1) break;

      rows.push(tableXml.substring(rowStart, rowEnd + '</a:tr>'.length));
      searchFrom = rowEnd + '</a:tr>'.length;
    }

    return rows;
  }

  /**
   * Builds a new row XML by cloning a template row and replacing cell text.
   *
   * @private
   * @param {string} templateRow - Template row XML to clone.
   * @param {string[]} cellValues - Text values for each cell.
   * @returns {string} New row XML.
   */
  #buildRow(templateRow, cellValues) {
    // Extract template cells
    const cells = this.#extractCells(templateRow);

    // Build new cells by replacing text in templates
    const newCells = cellValues.map((value, colIdx) => {
      const template = cells[colIdx] || cells[cells.length - 1] || '<a:tc><a:txBody><a:p><a:r><a:t/></a:r></a:p></a:txBody><a:tcPr/></a:tc>';
      return this.#setCellText(template, value);
    });

    // Replace cells in template row
    const firstCellStart = templateRow.indexOf('<a:tc>') !== -1
      ? templateRow.indexOf('<a:tc>')
      : templateRow.indexOf('<a:tc ');
    const lastCellEnd = templateRow.lastIndexOf('</a:tc>') + '</a:tc>'.length;

    if (firstCellStart === -1 || lastCellEnd === -1) {
      return templateRow;
    }

    return (
      templateRow.substring(0, firstCellStart) +
      newCells.join('') +
      templateRow.substring(lastCellEnd)
    );
  }

  /**
   * Extracts all <a:tc> cell XML strings from a row.
   * @private
   * @param {string} rowXml
   * @returns {string[]}
   */
  #extractCells(rowXml) {
    const cells = [];
    let searchFrom = 0;

    while (true) {
      let cellStart = rowXml.indexOf('<a:tc>', searchFrom);
      if (cellStart === -1) {
        cellStart = rowXml.indexOf('<a:tc ', searchFrom);
      }
      if (cellStart === -1) break;

      const cellEnd = rowXml.indexOf('</a:tc>', cellStart);
      if (cellEnd === -1) break;

      cells.push(rowXml.substring(cellStart, cellEnd + '</a:tc>'.length));
      searchFrom = cellEnd + '</a:tc>'.length;
    }

    return cells;
  }

  /**
   * Replaces the text content of a table cell.
   * Preserves all formatting (borders, colors, fonts) and only changes <a:t> content.
   *
   * @private
   * @param {string} cellXml - Cell XML template.
   * @param {string} text - New text content.
   * @returns {string} Updated cell XML.
   */
  #setCellText(cellXml, text) {
    const escapedText = this.#escapeXml(String(text));

    // If there's an existing <a:t> element, replace its content
    const tPattern = /(<a:t>)(.*?)(<\/a:t>)/;
    if (tPattern.test(cellXml)) {
      // Only replace the first <a:t> to avoid touching other text runs
      return cellXml.replace(tPattern, `$1${escapedText}$3`);
    }

    // If no <a:t> exists, inject one inside the first text run
    const rPattern = /(<a:r[^>]*>)([\s\S]*?)(<\/a:r>)/;
    if (rPattern.test(cellXml)) {
      return cellXml.replace(rPattern, `$1$2<a:t>${escapedText}</a:t>$3`);
    }

    // Fallback: inject a basic text paragraph
    const txBodyPattern = /(<a:txBody>)([\s\S]*?)(<\/a:txBody>)/;
    if (txBodyPattern.test(cellXml)) {
      return cellXml.replace(
        txBodyPattern,
        `$1<a:p><a:r><a:t>${escapedText}</a:t></a:r></a:p>$3`
      );
    }

    return cellXml;
  }

  /**
   * Finds the position of the closing tag that matches an opening at searchFrom.
   * @private
   * @param {string} xml
   * @param {string} closingTag
   * @param {number} searchFrom
   * @returns {number} Index of closing tag or -1.
   */
  #findClosingTag(xml, closingTag, searchFrom) {
    return xml.indexOf(closingTag, searchFrom);
  }

  /**
   * Escapes regex special characters.
   * @private
   * @param {string} str
   * @returns {string}
   */
  #escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  /**
   * Escapes XML special characters.
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

  /**
   * Inspects a slide's table structure (for debugging).
   *
   * @param {number} slideIndex
   * @param {SlideManager} slideManager
   * @returns {Array<{name: string, id: string, rows: number, cols: number}>}
   */
  inspectTables(slideIndex, slideManager) {
    const slideXml = slideManager.getSlideXml(slideIndex);
    const tables = [];

    // Find all graphicFrames with table data
    const framePattern = /<p:graphicFrame>([\s\S]*?)<\/p:graphicFrame>/g;
    let match;

    while ((match = framePattern.exec(slideXml)) !== null) {
      const frameXml = match[1];
      if (!frameXml.includes('<a:tbl>')) continue;

      const nameMatch = /name="([^"]*)"/.exec(frameXml);
      const idMatch = /id="([^"]*)"/.exec(frameXml);
      const rows = this.#extractAllRows(frameXml);
      const cols = rows[0] ? this.#extractCells(rows[0]).length : 0;

      tables.push({
        name: nameMatch ? nameMatch[1] : 'unnamed',
        id: idMatch ? idMatch[1] : 'unknown',
        rows: rows.length,
        cols,
      });
    }

    return tables;
  }
}

module.exports = { TableManager };
