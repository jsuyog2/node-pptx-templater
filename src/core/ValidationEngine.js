/**
 * @fileoverview ValidationEngine - Complete validation and checking engine for PowerPoint presentations.
 */

/**
 * @class ValidationEngine
 * @description Runs audits and verification checks on PPTX files.
 */
class ValidationEngine {
  /**
   * Validates the entire presentation.
   *
   * @param {PPTXTemplater} ppt - The presentation templater instance.
   * @returns {Promise<Object>} Structured report.
   */
  static async validatePresentation(ppt) {
    const errors = []
    const warnings = []

    // 1. Validate slides and references
    const slides = ppt.slideManager.getAllSlideInfo()
    for (const slide of slides) {
      // Validate slide relationships
      const relResult = this.validateRelationships(ppt, slide.zipPath)
      errors.push(...relResult.errors.map(e => `Slide ${slide.index} relationship error: ${e}`))
      warnings.push(
        ...relResult.warnings.map(w => `Slide ${slide.index} relationship warning: ${w}`)
      )

      // Validate slide XML and elements
      const slideResult = await this.validateSlide(ppt, slide.index)
      errors.push(...slideResult.errors)
      warnings.push(...slideResult.warnings)
    }

    // 2. Validate presentation level relationships
    const presRelResult = this.validateRelationships(ppt, 'ppt/presentation.xml')
    errors.push(...presRelResult.errors.map(e => `Presentation relationship error: ${e}`))
    warnings.push(...presRelResult.warnings.map(w => `Presentation relationship warning: ${w}`))

    return {
      valid: errors.length === 0,
      errors,
      warnings,
    }
  }

  /**
   * Validates a single slide's XML and elements (like tables, shapes, charts).
   *
   * @param {PPTXTemplater} ppt
   * @param {number} slideIndex
   * @returns {Promise<Object>}
   */
  static async validateSlide(ppt, slideIndex) {
    const errors = []
    const warnings = []

    try {
      const slideXml = await ppt.slideManager.getSlideXmlAsync(slideIndex)

      // Verify well-formed XML
      const xmlCheck = ppt.xmlParser.validate(slideXml)
      if (!xmlCheck.valid) {
        errors.push(`Slide ${slideIndex} XML syntax error: ${xmlCheck.error}`)
        return { valid: false, errors, warnings }
      }

      // Verify tables on the slide
      const tables = ppt.tableManager.inspectTables(slideIndex, ppt.slideManager)
      for (const table of tables) {
        const tableResult = await this.validateTable(ppt, slideIndex, table.id)
        errors.push(
          ...tableResult.errors.map(e => `Slide ${slideIndex} Table "${table.name}": ${e}`)
        )
        warnings.push(
          ...tableResult.warnings.map(w => `Slide ${slideIndex} Table "${table.name}": ${w}`)
        )
      }

      // Verify charts on the slide
      const charts = ppt.chartManager.getChartsInSlide(
        slideIndex,
        ppt.slideManager,
        ppt.relationshipManager
      )
      for (const chart of charts) {
        if (!ppt.zipManager.hasFile(chart.zipPath)) {
          errors.push(
            `Slide ${slideIndex} referenced chart file does not exist at ${chart.zipPath}`
          )
        }
      }

      // Verify image references on the slide
      const images = ppt.imageManager.getImages(
        slideIndex,
        ppt.slideManager,
        ppt.relationshipManager
      )
      for (const image of images) {
        if (image.targetPath && !ppt.zipManager.hasFile(image.targetPath)) {
          errors.push(
            `Slide ${slideIndex} referenced image file does not exist at ${image.targetPath}`
          )
        }
      }

      // Verify Z-order and duplicate IDs
      const zOrderResult = this.validateObjectOrder(ppt, slideIndex)
      errors.push(...zOrderResult.errors)
      warnings.push(...zOrderResult.warnings)
    } catch (err) {
      errors.push(`Slide ${slideIndex} validation error: ${err.message}`)
    }

    return {
      valid: errors.length === 0,
      errors,
      warnings,
    }
  }

  /**
   * Validates a table's XML and OpenXML conformity.
   *
   * @param {PPTXTemplater} ppt
   * @param {number} slideIndex
   * @param {string} tableId
   * @returns {Promise<Object>}
   */
  static async validateTable(ppt, slideIndex, tableId) {
    const errors = []
    const warnings = []

    try {
      const slideXml = await ppt.slideManager.getSlideXmlAsync(slideIndex)
      const slideObj = ppt.xmlParser.parse(slideXml, `slide${slideIndex}.xml`)

      const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']
      if (!spTree) {
        errors.push('Slide shape tree not found')
        return { valid: false, errors, warnings }
      }

      let frames = spTree['p:graphicFrame'] || []
      if (!Array.isArray(frames)) frames = [frames]

      let tbl = null
      for (const frame of frames) {
        const t = frame?.['a:graphic']?.['a:graphicData']?.['a:tbl']
        if (!t) continue
        const cNvPr = frame?.['p:nvGraphicFramePr']?.['p:cNvPr']
        if (cNvPr && (cNvPr['@_name'] === tableId || String(cNvPr['@_id']) === tableId)) {
          tbl = t
          break
        }
      }

      if (!tbl) {
        errors.push(`Table "${tableId}" not found in slide object`)
        return { valid: false, errors, warnings }
      }

      // Check cols against gridcols
      const cols = tbl['a:tblGrid']?.['a:gridCol'] || []
      const trs = tbl['a:tr'] || []

      if (cols.length === 0) {
        errors.push('Table column definitions (tblGrid) are missing')
      }

      // Check rowIds duplicate values
      const rowIds = new Set()
      trs.forEach((tr, rIdx) => {
        const tcs = tr['a:tc'] || []
        if (tcs.length !== cols.length) {
          warnings.push(
            `Row ${rIdx} cell count (${tcs.length}) does not match grid columns count (${cols.length})`
          )
        }

        // Check for rowId
        const ext = tr['a:extLst']?.['a:ext']
        const exts = Array.isArray(ext) ? ext : [ext]
        for (const e of exts) {
          if (e?.['a16:rowId']) {
            const val = e['a16:rowId']['@_val']
            if (val) {
              if (rowIds.has(val)) {
                errors.push(`Duplicate a16:rowId "${val}" found at row index ${rIdx}`)
              }
              rowIds.add(val)
            }
          }
        }
      })
    } catch (err) {
      errors.push(`Table validation error: ${err.message}`)
    }

    return {
      valid: errors.length === 0,
      errors,
      warnings,
    }
  }

  /**
   * Validates relationship mappings for a specific part.
   *
   * @param {PPTXTemplater} ppt
   * @param {string} partPath
   * @returns {Object}
   */
  static validateRelationships(ppt, partPath) {
    const errors = []
    const warnings = []

    const relsPath = ppt.relationshipManager.getRelsPath(partPath)
    if (!ppt.zipManager.hasFile(relsPath)) {
      warnings.push(`Relationship file missing at ${relsPath}`)
      return { valid: true, errors, warnings }
    }

    const rels = ppt.relationshipManager.getRelationships(partPath)
    const relIds = new Set()

    for (const rel of rels) {
      if (relIds.has(rel.id)) {
        errors.push(`Duplicate relationship ID "${rel.id}" inside ${relsPath}`)
      }
      relIds.add(rel.id)

      if (rel.targetMode !== 'External') {
        const resolved = ppt.relationshipManager.resolveTarget(partPath, rel.target)
        if (!ppt.zipManager.hasFile(resolved)) {
          errors.push(`Relationship ${rel.id} points to non-existent file: ${resolved}`)
        }
      }
    }

    return {
      valid: errors.length === 0,
      errors,
      warnings,
    }
  }

  /**
   * Audits the shape tree structure for duplicate drawing element IDs.
   */
  static validateObjectOrder(ppt, slideIndex) {
    const errors = []
    const warnings = []

    try {
      const slideXml = ppt.slideManager.getSlideXml(slideIndex)
      const slideObj = ppt.xmlParser.parse(slideXml, `slide${slideIndex}.xml`)
      const spTree = slideObj?.['p:sld']?.['p:cSld']?.['p:spTree']

      if (!spTree) {
        errors.push(`Slide ${slideIndex} shape tree not found`)
        return { valid: false, errors, warnings }
      }

      const idToTags = new Map()

      const checkIdsRecursive = container => {
        if (!container) return

        for (const tag of ['p:sp', 'p:pic', 'p:graphicFrame', 'p:grpSp', 'p:cxnSp']) {
          let items = container[tag] || []
          if (!Array.isArray(items)) items = [items]
          for (const item of items) {
            let id = null
            let name = null
            if (tag === 'p:sp') {
              id = item?.['p:nvSpPr']?.['p:cNvPr']?.['@_id']
              name = item?.['p:nvSpPr']?.['p:cNvPr']?.['@_name']
            } else if (tag === 'p:pic') {
              id = item?.['p:nvPicPr']?.['p:cNvPr']?.['@_id']
              name = item?.['p:nvPicPr']?.['p:cNvPr']?.['@_name']
            } else if (tag === 'p:graphicFrame') {
              id = item?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_id']
              name = item?.['p:nvGraphicFramePr']?.['p:cNvPr']?.['@_name']
            } else if (tag === 'p:grpSp') {
              id = item?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_id']
              name = item?.['p:nvGrpSpPr']?.['p:cNvPr']?.['@_name']
              checkIdsRecursive(item)
            } else if (tag === 'p:cxnSp') {
              id = item?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_id']
              name = item?.['p:nvCxnSpPr']?.['p:cNvPr']?.['@_name']
            }

            if (id !== undefined && id !== null) {
              const strId = String(id)
              if (idToTags.has(strId)) {
                errors.push(
                  `Duplicate drawing object ID "${strId}" found in slide ${slideIndex} (name: "${name}")`
                )
              } else {
                idToTags.set(strId, tag)
              }
            }
          }
        }
      }

      checkIdsRecursive(spTree)
    } catch (err) {
      errors.push(`Slide ${slideIndex} shape tree validation error: ${err.message}`)
    }

    return {
      valid: errors.length === 0,
      errors,
      warnings,
    }
  }

  /**
   * Validates data label configurations for a chart.
   *
   * @param {PPTXTemplater} ppt
   * @param {number} slideIndex
   * @param {string} chartId
   * @param {Object} options
   * @returns {Promise<Object>} report
   */
  static async validateDataLabels(ppt, slideIndex, chartId, options = {}) {
    const errors = []
    const warnings = []

    try {
      const chartInfo = ppt.chartManager.findChartInSlide(
        slideIndex,
        chartId,
        ppt.slideManager,
        ppt.relationshipManager
      )

      if (!chartInfo) {
        errors.push(`Chart "${chartId}" not found in slide ${slideIndex}`)
        return { valid: false, errors, warnings }
      }

      const chartType = await ppt.chartManager.getChartTypeAsync(
        slideIndex,
        chartId,
        ppt.slideManager,
        ppt.relationshipManager
      )

      const supportedTypes = [
        'bar',
        'column',
        'line',
        'pie',
        'doughnut',
        'area',
        'scatter',
        'combo',
        'unknown',
      ]
      if (!supportedTypes.includes(chartType)) {
        errors.push(`Unsupported chart type "${chartType}" for data labels`)
      }

      const xml = await ppt.zipManager.readFile(chartInfo.zipPath)
      let ptsCount = 0
      const catMatch = /<c:cat>([\s\S]*?)<\/c:cat>/.exec(xml)
      const valMatch = /<c:val>([\s\S]*?)<\/c:val>/.exec(xml)
      const targetBlock = catMatch ? catMatch[1] : valMatch ? valMatch[1] : ''
      const ptCountMatch = /<c:ptCount val="(\d+)"\/>/.exec(targetBlock)
      if (ptCountMatch) {
        ptsCount = parseInt(ptCountMatch[1], 10)
      }

      if (options.labels) {
        if (ptsCount > 0 && options.labels.length !== ptsCount) {
          errors.push(
            `Label count (${options.labels.length}) does not match chart data points count (${ptsCount})`
          )
        }

        options.labels.forEach((lbl, i) => {
          if (lbl === null || lbl === undefined || String(lbl).trim() === '') {
            warnings.push(`Label at index ${i} is empty`)
          }
        })
      }

      if (options.labelsFromCells) {
        const range = options.labelsFromCells
        const parts = range.split('!')
        const rangePart = parts.length > 1 ? parts[1] : parts[0]

        const rangeRegex = /^\$?[A-Z]+\$?\d+(?::\$?[A-Z]+\$?\d+)?$/i
        if (!rangeRegex.test(rangePart)) {
          errors.push(`Invalid range format: "${options.labelsFromCells}"`)
        }
      }
    } catch (err) {
      errors.push(`Data labels validation error: ${err.message}`)
    }

    return {
      valid: errors.length === 0,
      errors,
      warnings,
    }
  }
}

module.exports = { ValidationEngine }
