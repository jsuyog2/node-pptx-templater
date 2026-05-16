import JSZip from 'jszip';
import { createLogger } from '../../utils/logger.js';

const logger = createLogger('ChartWorkbookUpdater');

export class ChartWorkbookUpdater {
  /**
   * Updates the embedded Excel workbook for a chart.
   *
   * @param {Buffer|Uint8Array} workbookData - The raw XLSX buffer.
   * @param {Object} data - Chart data (categories, series).
   * @returns {Promise<Buffer>} - The updated XLSX buffer.
   */
  static async updateWorkbook(workbookData, data) {
    if (!workbookData) return null;
    
    try {
      const zip = await JSZip.loadAsync(workbookData);
      
      // Look for sheet1.xml
      const sheetPath = 'xl/worksheets/sheet1.xml';
      if (!zip.file(sheetPath)) {
        logger.warn('sheet1.xml not found in embedded workbook, trying to find first sheet');
        // fallback to finding the first sheet
      }
      
      const newSheetXml = this.#generateSheetXml(data);
      zip.file(sheetPath, newSheetXml);
      
      // Clean up any existing Excel tables, as our new sheet data might not align with them
      const tableFiles = Object.keys(zip.files).filter(f => f.startsWith('xl/tables/'));
      tableFiles.forEach(f => zip.remove(f));
      
      const sheetRels = Object.keys(zip.files).filter(f => f.startsWith('xl/worksheets/_rels/'));
      sheetRels.forEach(f => zip.remove(f));
      
      const contentTypesFile = zip.file('[Content_Types].xml');
      if (contentTypesFile) {
        const contentTypesXml = await contentTypesFile.async('text');
        const updatedContentTypes = contentTypesXml.replace(/<Override[^>]*PartName="\/xl\/tables\/[^"]*"[^>]*\/>/g, '');
        zip.file('[Content_Types].xml', updatedContentTypes);
      }
      
      return await zip.generateAsync({
        type: 'nodebuffer',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
      });
    } catch (err) {
      console.error('Failed to update embedded workbook', err);
      logger.error('Failed to update embedded workbook', err);
      return workbookData; // Return original if failed
    }
  }

  static #generateSheetXml(data) {
    const { categories = [], series = [] } = data;
    
    // Column count = 1 (categories) + series.length
    const numCols = 1 + series.length;
    const numRows = 1 + categories.length; // Row 1 = headers
    
    const lastColLetter = this.getColumnLetter(numCols - 1);
    const dimensionRef = `A1:${lastColLetter}${numRows}`;
    
    let sheetData = '<sheetData>';
    
    // Row 1: Headers (empty cell A1, then series names)
    sheetData += '<row r="1">';
    sheetData += `<c r="A1" t="inlineStr"><is><t></t></is></c>`;
    series.forEach((ser, i) => {
      const colLetter = this.getColumnLetter(i + 1);
      sheetData += `<c r="${colLetter}1" t="inlineStr"><is><t>${this.#escapeXml(ser.name || '')}</t></is></c>`;
    });
    sheetData += '</row>';
    
    // Rows 2..N: Data (category name in A, then values)
    categories.forEach((cat, rowIndex) => {
      const r = rowIndex + 2; // +1 for 1-based, +1 for header row
      sheetData += `<row r="${r}">`;
      sheetData += `<c r="A${r}" t="inlineStr"><is><t>${this.#escapeXml(String(cat))}</t></is></c>`;
      
      series.forEach((ser, colIndex) => {
        const colLetter = this.getColumnLetter(colIndex + 1);
        const val = ser.values && ser.values[rowIndex] !== undefined ? ser.values[rowIndex] : 0;
        sheetData += `<c r="${colLetter}${r}"><v>${Number(val)}</v></c>`;
      });
      sheetData += '</row>';
    });
    
    sheetData += '</sheetData>';
    
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="${dimensionRef}"/>
  <sheetViews><sheetView workbookViewId="0"/></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  ${sheetData}
</worksheet>`;
  }

  static getColumnLetter(colIndex) {
    let letter = '';
    while (colIndex >= 0) {
      letter = String.fromCharCode(65 + (colIndex % 26)) + letter;
      colIndex = Math.floor(colIndex / 26) - 1;
    }
    return letter;
  }

  static getFormulaRange(sheetName, startRow, startCol, endRow, endCol) {
    const startLetter = this.getColumnLetter(startCol);
    const endLetter = this.getColumnLetter(endCol);
    return `${sheetName}!$${startLetter}$${startRow}:$${endLetter}$${endRow}`;
  }

  static getFormulaSingleCell(sheetName, row, col) {
    const letter = this.getColumnLetter(col);
    return `${sheetName}!$${letter}$${row}`;
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
