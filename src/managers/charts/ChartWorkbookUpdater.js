const JSZip = require('jszip')
const { createLogger } = require('../../utils/logger.js')

const logger = createLogger('ChartWorkbookUpdater')

class ChartWorkbookUpdater {
  /**
   * Updates the embedded Excel workbook for a chart.
   *
   * @param {Buffer|Uint8Array} workbookData - The raw XLSX buffer.
   * @param {Object} data - Chart data (categories, series).
   * @returns {Promise<Buffer>} - The updated XLSX buffer.
   */
  static async updateWorkbook(workbookData, data) {
    if (!workbookData) return null

    try {
      const zip = await JSZip.loadAsync(workbookData)

      // Look for sheet1.xml
      const sheetPath = 'xl/worksheets/sheet1.xml'
      if (!zip.file(sheetPath)) {
        logger.warn('sheet1.xml not found in embedded workbook')
        return workbookData
      }

      const sheetXml = await zip.file(sheetPath).async('text')
      const sharedStrings = await this.getSharedStrings(zip)
      const cells = this.parseWorksheetCells(sheetXml, sharedStrings)

      // Update categories, series values, and custom labels in the cell grid
      this.#updateCellGrid(cells, data)

      // Serialize cells to sheet XML
      const updatedSheetXml = this.#serializeSheetXml(sheetXml, cells)
      zip.file(sheetPath, updatedSheetXml)

      // Clean up any existing Excel tables
      const tableFiles = Object.keys(zip.files).filter(f => f.startsWith('xl/tables/'))
      tableFiles.forEach(f => zip.remove(f))

      const sheetRels = Object.keys(zip.files).filter(f => f.startsWith('xl/worksheets/_rels/'))
      sheetRels.forEach(f => zip.remove(f))

      const contentTypesFile = zip.file('[Content_Types].xml')
      if (contentTypesFile) {
        const contentTypesXml = await contentTypesFile.async('text')
        const updatedContentTypes = contentTypesXml.replace(
          /<Override[^>]*PartName="\/xl\/tables\/[^"]*"[^>]*\/>/g,
          ''
        )
        zip.file('[Content_Types].xml', updatedContentTypes)
      }

      return await zip.generateAsync({
        type: 'nodebuffer',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 },
      })
    } catch (err) {
      logger.error('Failed to update embedded workbook', err)
      return workbookData // Return original if failed
    }
  }

  static async getSharedStrings(zip) {
    const sstFile = zip.file('xl/sharedStrings.xml')
    if (!sstFile) return []
    const xml = await sstFile.async('text')
    const strings = []
    const pattern = /<si>([\s\S]*?)<\/si>/g
    let match
    while ((match = pattern.exec(xml)) !== null) {
      const siContent = match[1]
      const tPattern = /<t\b[^>]*>([^<]*)<\/t>/g
      let tMatch
      let textVal = ''
      while ((tMatch = tPattern.exec(siContent)) !== null) {
        textVal += tMatch[1]
      }
      strings.push(textVal)
    }
    return strings
  }

  static parseWorksheetCells(sheetXml, sharedStrings) {
    const cells = {}
    const cellPattern = /<c r="([A-Z]+\d+)"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g
    let match
    while ((match = cellPattern.exec(sheetXml)) !== null) {
      const ref = match[1]
      const attrs = match[2]
      const content = match[3] || ''

      const tMatch = /t="([^"]*)"/.exec(attrs)
      const t = tMatch ? tMatch[1] : null

      let val = ''
      if (t === 'inlineStr') {
        const tValMatch = /<t\b[^>]*>([^<]*)<\/t>/.exec(content)
        val = tValMatch ? tValMatch[1] : ''
      } else if (t === 's') {
        const vMatch = /<v>(\d+)<\/v>/.exec(content)
        if (vMatch) {
          const idx = parseInt(vMatch[1], 10)
          val = sharedStrings[idx] !== undefined ? sharedStrings[idx] : ''
        }
      } else if (content) {
        const vMatch = /<v>([^<]*)<\/v>/.exec(content)
        if (vMatch) {
          val = vMatch[1]
          if (val !== '' && !isNaN(val)) {
            val = Number(val)
          }
        }
      }
      cells[ref] = val
    }
    return cells
  }

  static #updateCellGrid(cells, data) {
    const { categories = [], series = [] } = data

    // Clear cells that are outside the new category/series grid
    const maxRow = categories.length + 1
    const maxCol = series.length // Column A is 0, Column B is 1, etc.

    for (const ref of Object.keys(cells)) {
      const match = /^([A-Z]+)(\d+)$/.exec(ref)
      if (match) {
        const colLetter = match[1]
        const row = parseInt(match[2], 10)
        const col = this.colLetterToNum(colLetter)
        if (row > maxRow || col > maxCol) {
          delete cells[ref]
        }
      }
    }

    // 1. Write Header A1 as empty
    cells['A1'] = ''

    // 2. Write series titles in Row 1 (B1, C1, etc.)
    series.forEach((ser, i) => {
      const colLetter = this.getColumnLetter(i + 1)
      cells[`${colLetter}1`] = ser.name || ''
    })

    // 3. Write categories in column A (A2, A3, etc.)
    categories.forEach((cat, rowIndex) => {
      cells[`A${rowIndex + 2}`] = String(cat)
    })

    // 4. Write series values
    series.forEach((ser, colIndex) => {
      const colLetter = this.getColumnLetter(colIndex + 1)
      if (ser.values) {
        ser.values.forEach((val, rowIndex) => {
          cells[`${colLetter}${rowIndex + 2}`] = val !== undefined ? val : null
        })
      }
    })

    // 5. Write custom data labels
    series.forEach(ser => {
      if (ser.labels && ser.labelsFromCells) {
        const range = this.parseCellRange(ser.labelsFromCells)
        const startColNum = this.colLetterToNum(range.startCol)
        ser.labels.forEach((lbl, i) => {
          let cellRef
          if (range.startRow === range.endRow) {
            cellRef = `${this.numToColLetter(startColNum + i)}${range.startRow}`
          } else {
            cellRef = `${range.startCol}${range.startRow + i}`
          }
          cells[cellRef] = lbl
        })
      }
    })
  }

  static #serializeSheetXml(sheetXml, cells) {
    // Group cells by row
    const rows = {}
    for (const ref of Object.keys(cells)) {
      const rowMatch = /\d+$/.exec(ref)
      if (!rowMatch) continue
      const r = parseInt(rowMatch[0], 10)
      if (!rows[r]) rows[r] = []
      rows[r].push(ref)
    }

    // Sort rows ascending
    const sortedRowKeys = Object.keys(rows)
      .map(Number)
      .sort((a, b) => a - b)

    let maxRow = 1
    let maxColNum = 0

    let sheetData = '<sheetData>'
    for (const r of sortedRowKeys) {
      if (r > maxRow) maxRow = r
      sheetData += `<row r="${r}">`

      // Sort cells in this row by column letters
      const sortedRefs = rows[r].sort((a, b) => {
        const colA = /^[A-Z]+/.exec(a)[0]
        const colB = /^[A-Z]+/.exec(b)[0]
        if (colA.length !== colB.length) return colA.length - colB.length
        return colA.localeCompare(colB)
      })

      for (const ref of sortedRefs) {
        const colLetter = /^[A-Z]+/.exec(ref)[0]
        const colNum = this.colLetterToNum(colLetter)
        if (colNum > maxColNum) maxColNum = colNum

        const val = cells[ref]
        if (val === null || val === undefined) {
          sheetData += `<c r="${ref}" t="inlineStr"><is><t></t></is></c>`
        } else if (typeof val === 'number') {
          sheetData += `<c r="${ref}"><v>${val}</v></c>`
        } else {
          sheetData += `<c r="${ref}" t="inlineStr"><is><t>${this.#escapeXml(String(val))}</t></is></c>`
        }
      }
      sheetData += '</row>'
    }
    sheetData += '</sheetData>'

    const maxColLetter = this.numToColLetter(maxColNum)
    const newDimension = `<dimension ref="A1:${maxColLetter}${maxRow}"/>`

    let updatedXml = sheetXml.replace(/<sheetData>[\s\S]*?<\/sheetData>/, sheetData)
    updatedXml = updatedXml.replace(/<dimension ref="[^"]*"\/>/, newDimension)

    return updatedXml
  }

  static getColumnLetter(colIndex) {
    let letter = ''
    while (colIndex >= 0) {
      letter = String.fromCharCode(65 + (colIndex % 26)) + letter
      colIndex = Math.floor(colIndex / 26) - 1
    }
    return letter
  }

  static colLetterToNum(letter) {
    let num = 0
    for (let i = 0; i < letter.length; i++) {
      num = num * 26 + (letter.charCodeAt(i) - 64)
    }
    return num - 1
  }

  static numToColLetter(num) {
    return this.getColumnLetter(num)
  }

  static parseCellRange(rangeStr) {
    const parts = rangeStr.split('!')
    const range = parts.length > 1 ? parts[1] : parts[0]
    const cleanRange = range.replace(/\$/g, '')
    const [start, end] = cleanRange.split(':')

    const startCol = /^[A-Z]+/.exec(start)[0]
    const startRow = parseInt(/\d+$/.exec(start)[0], 10)

    const endCol = end ? /^[A-Z]+/.exec(end)[0] : startCol
    const endRow = end ? parseInt(/\d+$/.exec(end)[0], 10) : startRow

    return {
      sheetName: parts.length > 1 ? parts[0] : 'Sheet1',
      startCol,
      startRow,
      endCol,
      endRow,
    }
  }

  static getFormulaRange(sheetName, startRow, startCol, endRow, endCol) {
    const startLetter = this.getColumnLetter(startCol)
    const endLetter = this.getColumnLetter(endCol)
    return `${sheetName}!$${startLetter}$${startRow}:$${endLetter}$${endRow}`
  }

  static getFormulaSingleCell(sheetName, row, col) {
    const letter = this.getColumnLetter(col)
    return `${sheetName}!$${letter}$${row}`
  }

  static #escapeXml(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;')
  }
}

module.exports = { ChartWorkbookUpdater }
