const { ChartWorkbookUpdater } = require('./ChartWorkbookUpdater.js')

class ChartCacheGenerator {
  /**
   * Generates a string cache XML string (used for categories or series names).
   */
  static generateStrCache(values) {
    const ptEntries = values
      .map((val, i) => `<c:pt idx="${i}"><c:v>${this.#escapeXml(String(val))}</c:v></c:pt>`)
      .join('')
    return `<c:strCache><c:ptCount val="${values.length}"/>${ptEntries}</c:strCache>`
  }

  /**
   * Generates a numeric cache XML string.
   */
  static generateNumCache(values) {
    const ptEntries = values
      .map((val, i) => {
        if (val === null || val === undefined) {
          return `<c:pt idx="${i}"/>`
        }
        return `<c:pt idx="${i}"><c:v>${val}</c:v></c:pt>`
      })
      .join('')
    return `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${ptEntries}</c:numCache>`
  }

  /**
   * Updates category cache and formulas in a chart XML.
   *
   * @param {string} xml - Raw chart XML.
   * @param {string[]} categories - Array of categories.
   * @param {string} sheetName - Target worksheet name.
   */
  static updateCategories(xml, categories, sheetName = 'Sheet1') {
    const count = categories.length

    // Formula for categories: Sheet1!$A$2:$A$N
    const formula = ChartWorkbookUpdater.getFormulaRange(sheetName, 2, 0, count + 1, 0)
    const newStrCache = this.generateStrCache(categories)

    // Replace the entire <c:cat> block to ensure correct formula and cache
    const catPattern = /(<c:cat>)([\s\S]*?)(<\/c:cat>)/g

    return xml.replace(catPattern, (match, open, content, close) => {
      // Reconstruct the cat block
      // Try to determine if it used strRef or numRef originally
      let refTag = content.includes('<c:numRef>') ? 'numRef' : 'strRef'
      // But typically categories are strings. Let's use strRef.
      refTag = 'strRef'

      return `${open}<c:${refTag}><c:f>${formula}</c:f>${newStrCache}</c:${refTag}>${close}`
    })
  }

  /**
   * Updates series names and values in chart XML.
   */
  static updateSeries(xml, series, categoriesLength, sheetName = 'Sheet1') {
    let updated = xml
    const serPattern = /(<c:ser>)([\s\S]*?)(<\/c:ser>)/g
    const serMatches = [...updated.matchAll(serPattern)]

    if (serMatches.length === 0) return xml

    let serIndex = 0
    updated = updated.replace(serPattern, (match, open, content, close) => {
      if (serIndex >= series.length) {
        // If there are more series templates than data, we could drop them,
        // but replacing with an empty string might break the XML layout if we're not careful.
        // Actually, removing extra series is requested: "Allow removing old series".
        return ''
      }

      const serData = series[serIndex]
      const colIndex = serIndex + 1 // Series data starts in column B (1)
      serIndex++

      let updatedContent = content

      // 1. Update Series Name (c:tx)
      if (serData.name !== undefined) {
        const nameFormula = ChartWorkbookUpdater.getFormulaSingleCell(sheetName, 1, colIndex)
        const nameCache = `<c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${this.#escapeXml(serData.name)}</c:v></c:pt></c:strCache>`

        const txPattern = /(<c:tx>)([\s\S]*?)(<\/c:tx>)/
        if (txPattern.test(updatedContent)) {
          updatedContent = updatedContent.replace(txPattern, (match, p1, p2, p3) => {
            return `${p1}<c:strRef><c:f>${nameFormula}</c:f>${nameCache}</c:strRef>${p3}`
          })
        } else {
          // Some charts don't have <c:tx>, we prepend it after <c:order> or <c:idx>
          const insertAfter = /(<c:order[^>]*>)/
          if (insertAfter.test(updatedContent)) {
            updatedContent = updatedContent.replace(insertAfter, (match, p1) => {
              return `${p1}<c:tx><c:strRef><c:f>${nameFormula}</c:f>${nameCache}</c:strRef></c:tx>`
            })
          }
        }
      }

      // 2. Update Series Values (c:val)
      if (serData.values !== undefined) {
        const valuesCount = categoriesLength || serData.values.length
        const valFormula = ChartWorkbookUpdater.getFormulaRange(
          sheetName,
          2,
          colIndex,
          valuesCount + 1,
          colIndex
        )
        const valCache = this.generateNumCache(serData.values)

        const valPattern = /(<c:val>)([\s\S]*?)(<\/c:val>)/
        if (valPattern.test(updatedContent)) {
          updatedContent = updatedContent.replace(valPattern, (match, p1, p2, p3) => {
            return `${p1}<c:numRef><c:f>${valFormula}</c:f>${valCache}</c:numRef>${p3}`
          })
        }
      }

      return `${open}${updatedContent}${close}`
    })

    return updated
  }

  /**
   * Clones a series template to support dynamic series addition.
   */
  static appendDynamicSeries(xml, targetCount) {
    const serPattern = /(<c:ser>)([\s\S]*?)(<\/c:ser>)/g
    const matches = [...xml.matchAll(serPattern)]
    if (matches.length === 0 || matches.length >= targetCount) return xml

    // Use the last series as a template to clone
    const templateMatch = matches[matches.length - 1]
    const template = templateMatch[0]

    // Find the end of the last series
    const lastIndex = templateMatch.index + template.length

    let newSeriesBlocks = ''
    for (let i = matches.length; i < targetCount; i++) {
      let clone = template
      // Update c:idx and c:order
      clone = clone.replace(/(<c:idx val=")\d+("\/>)/g, `$1${i}$2`)
      clone = clone.replace(/(<c:order val=")\d+("\/>)/g, `$1${i}$2`)
      newSeriesBlocks += clone
    }

    return xml.substring(0, lastIndex) + newSeriesBlocks + xml.substring(lastIndex)
  }

  /**
   * Updates the chart title text while fully preserving all existing styling.
   *
   * Three strategies in priority order:
   *
   * 1. If <c:tx><c:rich> already exists (title was previously set via PowerPoint or this
   *    library): ONLY replace <a:t> text values in-place, mapped by \n-split lines to
   *    existing runs in document order. This preserves alignment, bold/italic/underline,
   *    font sizes, paragraph structure, and any other per-run or per-paragraph properties.
   *
   * 2. If <c:title> exists but has no <c:tx> (title text comes from <c:txPr> default
   *    properties): Build a new <c:tx><c:rich> by extracting <a:bodyPr>, <a:pPr> (including
   *    algn), and <a:defRPr>→<a:rPr> from <c:txPr>. Supports multi-line via \n splitting
   *    into separate <a:p> paragraphs, each inheriting the same pPr and rPr.
   *
   * 3. If no <c:title> block exists at all: create a minimal one.
   */
  /**
   * Extracts parts of <c:txPr> block (bodyPr, pPr, and converts defRPr to rPr)
   * to use when creating custom rich text components.
   *
   * @param {string} txPrXml - The <c:txPr> XML string or block content.
   * @returns {Object} { bodyPr, pPrXml, rPrXml }
   */
  static extractTxPrParts(txPrXml) {
    let bodyPr = '<a:bodyPr/>'
    let pPrXml = ''
    let rPrXml = ''

    if (!txPrXml) {
      return { bodyPr, pPrXml, rPrXml }
    }

    let txPrContent = txPrXml
    const txPrInnerMatch = /<c:txPr>([\s\S]*?)<\/c:txPr>/.exec(txPrXml)
    if (txPrInnerMatch) {
      txPrContent = txPrInnerMatch[1]
    }

    // Extract <a:bodyPr>
    const bodyPrMatch = /(<a:bodyPr[^>]*\/>|<a:bodyPr[^>]*>[\s\S]*?<\/a:bodyPr>)/.exec(txPrContent)
    if (bodyPrMatch) bodyPr = bodyPrMatch[1]

    // Extract <a:defRPr> → convert to <a:rPr> for the run
    const defRPrMatch = /(<a:defRPr[\s\S]*?<\/a:defRPr>|<a:defRPr[^>]*\/>)/.exec(txPrContent)
    if (defRPrMatch) {
      rPrXml = defRPrMatch[1].replace(/^<a:defRPr/, '<a:rPr').replace(/<\/a:defRPr>$/, '</a:rPr>')
    }

    // Extract <a:pPr> (keeps algn, indent, etc.) but strip <a:defRPr> from it
    const pPrBlockMatch = /(<a:pPr[^>]*>)([\s\S]*?)(<\/a:pPr>)/.exec(txPrContent)
    if (pPrBlockMatch) {
      const innerContent = pPrBlockMatch[2]
        .replace(/<a:defRPr(?:[^>]*\/>|[\s\S]*?<\/a:defRPr>)/g, '')
        .trim()
      const attrs = pPrBlockMatch[1].slice(7, -1).trim()
      if (attrs || innerContent) {
        pPrXml = innerContent
          ? `${pPrBlockMatch[1]}${innerContent}${pPrBlockMatch[3]}`
          : `<a:pPr ${attrs}/>`
      }
    } else {
      const scPPrMatch = /(<a:pPr[^>]*\/>)/.exec(txPrContent)
      if (scPPrMatch) pPrXml = scPPrMatch[1]
    }

    return { bodyPr, pPrXml, rPrXml }
  }

  static updateTitle(xml, title) {
    // Split by \n so callers can drive multi-paragraph titles
    const titleLines = title.split('\n')

    if (!xml.includes('<c:title>')) {
      // Strategy 3 – no title block yet, create minimal
      const escapedText = this.#escapeXml(titleLines[0])
      const titleBlock = `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${escapedText}</a:t></a:r></a:p></c:rich></c:tx><c:layout/></c:title>`
      return xml.replace(/(<c:chart>)/, `$1${titleBlock}`)
    }

    return xml.replace(/<c:title>([\s\S]*?)<\/c:title>/, (match, titleContent) => {
      // ── Strategy 1 ──────────────────────────────────────────────────────────────
      // Existing <c:tx><c:rich> is present: preserve every element, just swap text.
      if (titleContent.includes('<c:tx>')) {
        const txMatch = /<c:tx>([\s\S]*?)<\/c:tx>/.exec(titleContent)
        if (txMatch && txMatch[1].includes('<c:rich>')) {
          let lineIndex = 0
          // Replace each <a:t>…</a:t> in document order with the next line.
          // Any run that maps beyond the supplied lines gets an empty string.
          const updatedTx = txMatch[1].replace(/<a:t>[^<]*<\/a:t>/g, () => {
            const text = lineIndex < titleLines.length ? this.#escapeXml(titleLines[lineIndex]) : ''
            lineIndex++
            return `<a:t>${text}</a:t>`
          })
          const updatedContent = titleContent.replace(
            /<c:tx>[\s\S]*?<\/c:tx>/,
            `<c:tx>${updatedTx}</c:tx>`
          )
          return `<c:title>${updatedContent}</c:title>`
        }
      }

      // ── Strategy 2 ──────────────────────────────────────────────────────────────
      // No <c:tx> yet – build one from <c:txPr> styles.
      const { bodyPr, pPrXml, rPrXml: rPr } = this.extractTxPrParts(titleContent)

      // Build one <a:p> per title line, each with the same pPr + rPr
      const paragraphs = titleLines
        .map(line => {
          const escapedLine = this.#escapeXml(line)
          return `<a:p>${pPrXml}<a:r>${rPr}<a:t>${escapedLine}</a:t></a:r></a:p>`
        })
        .join('')

      const newTxBlock = `<c:tx><c:rich>${bodyPr}<a:lstStyle/>${paragraphs}</c:rich></c:tx>`

      // Prepend <c:tx> before the first existing sibling (overlay, spPr, txPr, etc.)
      return `<c:title>${newTxBlock}${titleContent}</c:title>`
    })
  }

  static updateDataLabelsInXml(xml, seriesIndex, options, categories = [], seriesData = {}) {
    let serIndex = 0
    const serPattern = /(<c:ser>)([\s\S]*?)(<\/c:ser>)/g

    return xml.replace(serPattern, (match, open, content, close) => {
      if (serIndex !== seriesIndex) {
        serIndex++
        return match
      }
      serIndex++

      let pointsCount = categories.length
      if (pointsCount === 0) {
        const valMatch = /<c:val>([\s\S]*?)<\/c:val>/.exec(content)
        if (valMatch) {
          const countMatch = /<c:ptCount val="(\d+)"\/>/.exec(valMatch[1])
          if (countMatch) pointsCount = parseInt(countMatch[1], 10)
        }
      }

      let existingTxPr = ''
      let existingDLblPos = ''
      let existingNumFmt = ''
      let existingSpPr = ''
      let existingExtLst = ''
      let existingShowLeaderLines = ''
      const existingDLblSpPrs = {}
      const existingDLblTxPrs = {}
      const existingDLblLayouts = {}
      const existingDLblPositions = {}
      const existingShowTags = {}

      const dLblsMatch = /<c:dLbls>([\s\S]*?)<\/c:dLbls>/.exec(content)
      if (dLblsMatch) {
        const dLblsContent = dLblsMatch[1]
        const txPrMatch = /(<c:txPr>[\s\S]*?<\/c:txPr>)/.exec(dLblsContent)
        if (txPrMatch) {
          existingTxPr = txPrMatch[1]
        }
        const dLblPosMatch = /(<c:dLblPos\s+[^>]*\/>)/.exec(dLblsContent)
        if (dLblPosMatch) {
          existingDLblPos = dLblPosMatch[1]
        }
        const numFmtMatch = /(<c:numFmt\s+[^>]*\/>)/.exec(dLblsContent)
        if (numFmtMatch) {
          existingNumFmt = numFmtMatch[1]
        }
        const spPrMatch = /(<c:spPr>[\s\S]*?<\/c:spPr>)/.exec(dLblsContent)
        if (spPrMatch) {
          existingSpPr = spPrMatch[1]
        }
        const extLstMatch = /(<c:extLst>[\s\S]*?<\/c:extLst>)/.exec(dLblsContent)
        if (extLstMatch) {
          existingExtLst = extLstMatch[1]
        }
        const showLeaderMatch = /(<c:showLeaderLines\s+[^>]*\/>)/.exec(dLblsContent)
        if (showLeaderMatch) {
          existingShowLeaderLines = showLeaderMatch[1]
        }

        // Parse individual <c:dLbl> properties: fills, text styling, layouts, and positions
        const dLblPattern = /<c:dLbl>([\s\S]*?)<\/c:dLbl>/g
        let dLblMatch
        while ((dLblMatch = dLblPattern.exec(dLblsContent)) !== null) {
          const dLblContent = dLblMatch[1]
          const idxMatch = /<c:idx val="(\d+)"\/>/.exec(dLblContent)
          if (idxMatch) {
            const idx = parseInt(idxMatch[1], 10)
            const dLblSpPrMatch = /(<c:spPr>[\s\S]*?<\/c:spPr>)/.exec(dLblContent)
            if (dLblSpPrMatch) {
              existingDLblSpPrs[idx] = dLblSpPrMatch[1]
            }
            const dLblTxPrMatch = /(<c:txPr>[\s\S]*?<\/c:txPr>)/.exec(dLblContent)
            if (dLblTxPrMatch) {
              existingDLblTxPrs[idx] = dLblTxPrMatch[1]
            }
            const dLblLayoutMatch = /(<c:layout>[\s\S]*?<\/c:layout>|<c:layout\/>)/.exec(dLblContent)
            if (dLblLayoutMatch) {
              existingDLblLayouts[idx] = dLblLayoutMatch[1]
            }
            const dLblPosMatch = /(<c:dLblPos\s+[^>]*\/>)/.exec(dLblContent)
            if (dLblPosMatch) {
              existingDLblPositions[idx] = dLblPosMatch[1]
            }
          }
        }

        const showTagsList = [
          'showLegendKey',
          'showVal',
          'showCatName',
          'showSerName',
          'showPercent',
          'showBubbleSize',
        ]
        showTagsList.forEach(tag => {
          const tagPattern = new RegExp(`(<c:${tag}\\s+val="([^"]*)"\\s*\\/>)`)
          const tagMatch = tagPattern.exec(dLblsContent)
          if (tagMatch) {
            existingShowTags[tag] = tagMatch[1]
          }
        })
      }

      const dLblsXml = this.generateDLblsXml(
        pointsCount,
        options,
        categories,
        seriesData,
        existingTxPr,
        existingDLblPos,
        existingNumFmt,
        existingShowTags,
        existingSpPr,
        existingDLblSpPrs,
        existingDLblTxPrs,
        existingDLblLayouts,
        existingDLblPositions,
        existingExtLst,
        existingShowLeaderLines
      )

      let updatedContent = content
      const dLblsPattern = /(<c:dLbls>[\s\S]*?<\/c:dLbls>)/
      if (dLblsPattern.test(updatedContent)) {
        updatedContent = updatedContent.replace(dLblsPattern, dLblsXml)
      } else {
        const insertBefore = /(<c:cat>|<c:val>|<c:extLst>)/
        if (insertBefore.test(updatedContent)) {
          updatedContent = updatedContent.replace(insertBefore, `${dLblsXml}$1`)
        } else {
          updatedContent += dLblsXml
        }
      }

      return `${open}${updatedContent}${close}`
    })
  }

  static generateDLblsXml(
    pointsCount,
    options,
    categories = [],
    seriesData = {},
    existingTxPr = '',
    existingDLblPos = '',
    existingNumFmt = '',
    existingShowTags = {},
    existingSpPr = '',
    existingDLblSpPrs = {},
    existingDLblTxPrs = {},
    existingDLblLayouts = {},
    existingDLblPositions = {},
    existingExtLst = '',
    existingShowLeaderLines = ''
  ) {
    const {
      labels,
      labelsFromCells,
      template,
      position,
      labelStyle,
      labelMap,
      showSeriesNameInBar,
    } = options

    let xml = '<c:dLbls>'

    const posMap = {
      center: 'ctr',
      insideEnd: 'inEnd',
      insideBase: 'inBase',
      outsideEnd: 'outEnd',
      bestFit: 'bestFit',
      left: 'l',
      right: 'r',
      top: 't',
      bottom: 'b',
    }
    let openxmlPos = position ? posMap[position] : null
    if (!openxmlPos && showSeriesNameInBar) {
      openxmlPos = 'ctr'
    }

    const values = seriesData.values || []
    const sumValues = values.reduce((sum, v) => sum + (Number(v) || 0), 0)
    const seriesName = seriesData.name || ''

    const hasCustomLabels = labels || labelsFromCells || template || labelMap

    if (hasCustomLabels && pointsCount > 0) {
      for (let i = 0; i < pointsCount; i++) {
        const cat = categories[i] !== undefined ? String(categories[i]) : ''
        const val = values[i] !== undefined ? values[i] : ''

        let pct = 0
        if (sumValues > 0 && val !== '') {
          pct = Math.round((Number(val) / sumValues) * 100)
        }

        let customLabel = ''
        if (labels && labels[i] !== undefined) {
          customLabel = String(labels[i])
        } else if (labelMap && cat && labelMap[cat] !== undefined) {
          customLabel = String(labelMap[cat])
        }

        let textContent = customLabel
        if (template) {
          textContent = template
            .replace(/{category}/g, cat)
            .replace(/{value}/g, String(val))
            .replace(/{percentage}/g, String(pct))
            .replace(/{series}/g, seriesName)
            .replace(/{customLabel}/g, customLabel)
        }

        xml += `<c:dLbl>`
        xml += `<c:idx val="${i}"/>`

        // Restore per-point layout (manual position offsets) from template
        if (existingDLblLayouts && existingDLblLayouts[i]) {
          xml += existingDLblLayouts[i]
        }

        if (labelsFromCells && !template) {
          const range = ChartWorkbookUpdater.parseCellRange(labelsFromCells)
          const startColNum = ChartWorkbookUpdater.colLetterToNum(range.startCol)

          let cellRef
          if (range.startRow === range.endRow) {
            cellRef = `${ChartWorkbookUpdater.numToColLetter(startColNum + i)}${range.startRow}`
          } else {
            cellRef = `${range.startCol}${range.startRow + i}`
          }
          const fullCellRef = `${range.sheetName}!$${cellRef.replace(/(\d+)/, '$$$1')}`
          const displayVal = textContent || customLabel || ''

          xml += `<c:tx>`
          xml += `<c:strRef>`
          xml += `<c:f>${fullCellRef}</c:f>`
          xml += `<c:strCache>`
          xml += `<c:ptCount val="1"/>`
          xml += `<c:pt idx="0"><c:v>${this.#escapeXml(displayVal)}</c:v></c:pt>`
          xml += `</c:strCache>`
          xml += `</c:strRef>`
          xml += `</c:tx>`
        } else if (textContent) {
          if (labelsFromCells) {
            const range = ChartWorkbookUpdater.parseCellRange(labelsFromCells)
            const startColNum = ChartWorkbookUpdater.colLetterToNum(range.startCol)
            let cellRef
            if (range.startRow === range.endRow) {
              cellRef = `${ChartWorkbookUpdater.numToColLetter(startColNum + i)}${range.startRow}`
            } else {
              cellRef = `${range.startCol}${range.startRow + i}`
            }
            const fullCellRef = `${range.sheetName}!$${cellRef.replace(/(\d+)/, '$$$1')}`

            xml += `<c:tx>`
            xml += `<c:strRef>`
            xml += `<c:f>${fullCellRef}</c:f>`
            xml += `<c:strCache>`
            xml += `<c:ptCount val="1"/>`
            xml += `<c:pt idx="0"><c:v>${this.#escapeXml(textContent)}</c:v></c:pt>`
            xml += `</c:strCache>`
            xml += `</c:strRef>`
            xml += `</c:tx>`
          } else {
            let bodyPr = '<a:bodyPr/>'
            let pPrXml = ''
            let rPrXml = ''
            const dLblTxPr = (existingDLblTxPrs && existingDLblTxPrs[i]) || existingTxPr
            if (dLblTxPr) {
              const extracted = this.extractTxPrParts(dLblTxPr)
              bodyPr = extracted.bodyPr
              pPrXml = extracted.pPrXml
              rPrXml = extracted.rPrXml
            }

            const labelLines = String(textContent).split('\n')
            const paragraphs = labelLines
              .map(line => {
                const escapedLine = this.#escapeXml(line)
                return `<a:p>${pPrXml}<a:r>${rPrXml}<a:t>${escapedLine}</a:t></a:r></a:p>`
              })
              .join('')

            xml += `<c:tx>`
            xml += `<c:rich>`
            xml += bodyPr
            xml += `<a:lstStyle/>`
            xml += paragraphs
            xml += `</c:rich>`
            xml += `</c:tx>`
          }
        }

        // Restore individual point-level spPr (background fill, border, line) from template.
        // Priority: individual dLbl spPr → first available individual dLbl spPr → series-level spPr.
        // The series-level spPr fallback ensures fills are preserved even when the template stores
        // styling only at the dLbls level (not per individual dLbl override).
        if (existingDLblSpPrs && existingDLblSpPrs[i]) {
          xml += existingDLblSpPrs[i]
        } else {
          const firstIdx =
            existingDLblSpPrs && Object.keys(existingDLblSpPrs).length > 0
              ? Object.keys(existingDLblSpPrs)[0]
              : undefined
          if (firstIdx !== undefined && existingDLblSpPrs[firstIdx]) {
            xml += existingDLblSpPrs[firstIdx]
          } else if (existingSpPr) {
            // Fall back to series-level spPr so fills are preserved on each label
            xml += existingSpPr
          }
        }

        if (labelStyle) {
          xml += this.generateTxPrXml(labelStyle)
        } else if (existingDLblTxPrs && existingDLblTxPrs[i]) {
          xml += existingDLblTxPrs[i]
        } else if (existingTxPr) {
          xml += existingTxPr
        }

        if (openxmlPos) {
          xml += `<c:dLblPos val="${openxmlPos}"/>`
        } else if (existingDLblPositions && existingDLblPositions[i]) {
          xml += existingDLblPositions[i]
        } else if (existingDLblPos) {
          xml += existingDLblPos
        }

        xml += `<c:showLegendKey val="0"/>`
        xml += `<c:showVal val="0"/>`
        xml += `<c:showCatName val="0"/>`
        xml += `<c:showSerName val="0"/>`
        xml += `<c:showPercent val="0"/>`
        xml += `<c:showBubbleSize val="0"/>`

        xml += `</c:dLbl>`
      }
    }

    if (existingNumFmt) {
      xml += existingNumFmt
    }

    // Restore series-level spPr (background fill, border, line) from template
    if (existingSpPr) {
      xml += existingSpPr
    }

    if (labelStyle) {
      xml += this.generateTxPrXml(labelStyle)
    } else if (existingTxPr) {
      xml += existingTxPr
    }

    if (openxmlPos) {
      xml += `<c:dLblPos val="${openxmlPos}"/>`
    } else if (existingDLblPos) {
      xml += existingDLblPos
    }

    // showLegendKey
    if (existingShowTags['showLegendKey']) {
      xml += existingShowTags['showLegendKey']
    } else {
      xml += `<c:showLegendKey val="0"/>`
    }

    // showVal
    const defaultShowVal = hasCustomLabels ? '0' : '1'
    if (showSeriesNameInBar) {
      xml += `<c:showVal val="0"/>`
    } else if (existingShowTags['showVal'] && !hasCustomLabels) {
      xml += existingShowTags['showVal']
    } else {
      xml += `<c:showVal val="${defaultShowVal}"/>`
    }

    // showCatName
    if (existingShowTags['showCatName']) {
      xml += existingShowTags['showCatName']
    } else {
      xml += `<c:showCatName val="0"/>`
    }

    // showSerName
    if (showSeriesNameInBar) {
      xml += `<c:showSerName val="1"/>`
    } else if (existingShowTags['showSerName']) {
      xml += existingShowTags['showSerName']
    } else {
      xml += `<c:showSerName val="0"/>`
    }

    // showPercent
    const defaultShowPercent = hasCustomLabels || !options.showPercent ? '0' : '1'
    if (options.showPercent !== undefined) {
      xml += `<c:showPercent val="${options.showPercent ? '1' : '0'}"/>`
    } else if (existingShowTags['showPercent'] && !hasCustomLabels) {
      xml += existingShowTags['showPercent']
    } else {
      xml += `<c:showPercent val="${defaultShowPercent}"/>`
    }

    // showBubbleSize
    if (existingShowTags['showBubbleSize']) {
      xml += existingShowTags['showBubbleSize']
    } else {
      xml += `<c:showBubbleSize val="0"/>`
    }

    // Restore showLeaderLines and extension list from template
    if (existingShowLeaderLines) {
      xml += existingShowLeaderLines
    }
    if (existingExtLst) {
      xml += existingExtLst
    }

    xml += '</c:dLbls>'
    return xml
  }

  static generateTxPrXml(style) {
    const { fontFamily, fontSize, bold, italic, underline, color } = style

    const sz = fontSize ? ` sz="${fontSize * 100}"` : ''
    const b = bold !== undefined ? ` b="${bold ? '1' : '0'}"` : ''
    const i = italic !== undefined ? ` i="${italic ? '1' : '0'}"` : ''
    const u = underline !== undefined ? ` u="${underline ? 'sng' : 'none'}"` : ''

    let fillXml = ''
    if (color) {
      const cleanColor = color.replace('#', '')
      fillXml = `<a:solidFill><a:srgbClr val="${cleanColor}"/></a:solidFill>`
    }

    let latinXml = ''
    if (fontFamily) {
      latinXml = `<a:latin typeface="${fontFamily}"/><a:cs typeface="${fontFamily}"/>`
    }

    return `<c:txPr>
      <a:bodyPr/>
      <a:lstStyle/>
      <a:p>
        <a:pPr>
          <a:defRPr${sz}${b}${i}${u}>
            ${fillXml}
            ${latinXml}
          </a:defRPr>
        </a:pPr>
        <a:endParaRPr lang="en-US"/>
      </a:p>
    </c:txPr>`
  }

  static getDataLabelsFromXml(xml, seriesIndex) {
    const serPattern = /<c:ser>([\s\S]*?)<\/c:ser>/g
    const matches = [...xml.matchAll(serPattern)]
    if (seriesIndex >= matches.length) return []

    const serXml = matches[seriesIndex][1]
    const dLblsMatch = /<c:dLbls>([\s\S]*?)<\/c:dLbls>/.exec(serXml)
    if (!dLblsMatch) return []

    const dLblsXml = dLblsMatch[1]
    const dLblPattern = /<c:dLbl>([\s\S]*?)<\/c:dLbl>/g
    const result = []

    let dLblMatch
    while ((dLblMatch = dLblPattern.exec(dLblsXml)) !== null) {
      const dLblXml = dLblMatch[1]
      const idxMatch = /<c:idx val="(\d+)"\/>/.exec(dLblXml)
      if (!idxMatch) continue
      const point = parseInt(idxMatch[1], 10)

      let value = ''
      const strCacheMatch = /<c:strCache>([\s\S]*?)<\/c:strCache>/.exec(dLblXml)
      if (strCacheMatch) {
        const vMatch = /<c:v>([^<]*)<\/c:v>/.exec(strCacheMatch[1])
        if (vMatch) value = vMatch[1]
      } else {
        const tPattern = /<a:t>([^<]*)<\/a:t>/g
        let tMatch
        while ((tMatch = tPattern.exec(dLblXml)) !== null) {
          value += tMatch[1]
        }
      }

      result.push({ point, value })
    }
    return result
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

module.exports = { ChartCacheGenerator }
