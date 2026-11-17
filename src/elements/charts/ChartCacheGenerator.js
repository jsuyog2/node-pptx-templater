const { ChartWorkbookUpdater } = require('./ChartWorkbookUpdater.js')
const { ExtendedChartGenerator } = require('./ExtendedChartGenerator.js')

const CANONICAL_CHART_TYPES = {
  column: 'column',
  col: 'column',
  bar: 'bar',
  line: 'line',
  pie: 'pie',
  area: 'area',
  scatter: 'scatter',
  xy: 'scatter',
  xyscatter: 'scatter',
  stock: 'stock',
  surface: 'surface',
  radar: 'radar',
  treemap: 'treemap',
  sunburst: 'sunburst',
  histogram: 'histogram',
  boxwhisker: 'boxWhisker',
  boxandwhisker: 'boxWhisker',
  box_and_whisker: 'boxWhisker',
  waterfall: 'waterfall',
  funnel: 'funnel',
  combo: 'combo',
  map: 'map',
  regionmap: 'map',
  geochart: 'map',
}

const EXTENDED_CHART_TYPES = new Set([
  'treemap',
  'sunburst',
  'histogram',
  'boxWhisker',
  'waterfall',
  'funnel',
  'map',
])

class ChartCacheGenerator {
  /**
   * Normalizes chart type and aliases to canonical lowercase names.
   */
  static normalizeChartType(type) {
    if (!type) return 'column'
    const key = String(type).trim().toLowerCase()
    return CANONICAL_CHART_TYPES[key] || type
  }

  /**
   * Checks whether the chart type or alias is supported.
   */
  static isSupportedType(type) {
    if (!type || typeof type !== 'string') return false
    return CANONICAL_CHART_TYPES[type.trim().toLowerCase()] !== undefined
  }

  /**
   * Checks whether the chart type requires Office 2016+ extended chart parts (<cx:chartSpace>).
   */
  static isExtendedChart(type) {
    const canonical = this.normalizeChartType(type)
    return EXTENDED_CHART_TYPES.has(canonical)
  }

  /**
   * Generates a complete OpenXML chart XML string for all 17 supported types.
   */
  static generateChartSpaceXml({ type = 'column', data = {}, title, options = {}, rId = 'rId1' }) {
    const canonicalType = this.normalizeChartType(type)

    if (this.isExtendedChart(canonicalType)) {
      return ExtendedChartGenerator.generateChartSpaceXml({
        type: canonicalType,
        data,
        title,
        options,
        rId,
      })
    }

    const categories = data.categories || ['Category 1', 'Category 2', 'Category 3', 'Category 4']
    let series =
      data.series && data.series.length > 0
        ? data.series
        : [{ name: 'Series 1', values: [10, 20, 30, 40] }]

    // Stock charts require at least 3 series (High, Low, Close)
    if (canonicalType === 'stock' && series.length < 3) {
      const baseVals = series[0]?.values || [100, 105, 110, 115]
      series = [
        { name: 'High', values: baseVals.map(v => v + 5) },
        { name: 'Low', values: baseVals.map(v => Math.max(0, v - 5)) },
        { name: 'Close', values: baseVals },
      ]
    }

    const catStrCache = this.generateStrCache(categories)
    const catCount = categories.length

    // Title XML
    let titleXml = ''
    if (title || (options && options.title)) {
      const text = this.#escapeXml(title || options.title)
      titleXml = `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${text}</a:t></a:r></a:p></c:rich></c:tx><c:layout/></c:title>`
    }

    // Legend XML
    let legendXml = ''
    if (!options || options.legend !== false) {
      const pos = (options && options.legend && options.legend.position) || 'r'
      legendXml = `<c:legend><c:legendPos val="${pos}"/><c:layout/></c:legend>`
    }

    // View3D XML (required for 3D charts like surface)
    let view3DXml = ''
    if (canonicalType === 'surface') {
      view3DXml = `<c:view3D><c:rotX val="30"/><c:rotY val="30"/><c:rAngAx val="0"/><c:perspective val="30"/></c:view3D>`
    }

    // Helper to generate a standard series block
    const buildSerXml = (ser, idx, isScatter = false) => {
      const colLetter = ChartWorkbookUpdater.getColumnLetter(idx + 1)
      const values = ser.values || []
      const name = ser.name || `Series ${idx + 1}`
      const numCache = this.generateNumCache(values)
      const valCount = values.length

      const nameBlock = `<c:tx><c:strRef><c:f>Sheet1!$${colLetter}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>${this.#escapeXml(name)}</c:v></c:pt></c:strCache></c:strRef></c:tx>`

      if (isScatter) {
        // Scatter uses numeric xVal and yVal
        const xValues = categories.map((cat, i) => {
          const num = Number(cat)
          return isNaN(num) ? i + 1 : num
        })
        const xNumCache = this.generateNumCache(xValues)
        return `<c:ser>
          <c:idx val="${idx}"/>
          <c:order val="${idx}"/>
          ${nameBlock}
          <c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>
          <c:xVal><c:numRef><c:f>Sheet1!$A$2:$A$${catCount + 1}</c:f>${xNumCache}</c:numRef></c:xVal>
          <c:yVal><c:numRef><c:f>Sheet1!$${colLetter}$2:$${colLetter}$${valCount + 1}</c:f>${numCache}</c:numRef></c:yVal>
          <c:smooth val="0"/>
        </c:ser>`
      }

      const invertIfNeg =
        canonicalType === 'column' || canonicalType === 'bar' ? '<c:invertIfNegative val="0"/>' : ''

      return `<c:ser>
        <c:idx val="${idx}"/>
        <c:order val="${idx}"/>
        ${nameBlock}
        ${invertIfNeg}
        <c:cat><c:strRef><c:f>Sheet1!$A$2:$A$${catCount + 1}</c:f>${catStrCache}</c:strRef></c:cat>
        <c:val><c:numRef><c:f>Sheet1!$${colLetter}$2:$${colLetter}$${valCount + 1}</c:f>${numCache}</c:numRef></c:val>
      </c:ser>`
    }

    let plotXml = ''
    let axesXml = ''
    const axIdCat = '123456'
    const axIdVal = '123457'

    // Standard schema-compliant Axis XML builders (ISO/IEC 29500-1 §21.2 strictly ordered)
    const buildCatAx = (id, crossId, pos = 'b') => `<c:catAx>
      <c:axId val="${id}"/>
      <c:scaling><c:orientation val="minMax"/></c:scaling>
      <c:delete val="0"/>
      <c:axPos val="${pos}"/>
      <c:majorTickMark val="none"/>
      <c:minorTickMark val="none"/>
      <c:tickLblPos val="nextTo"/>
      <c:crossAx val="${crossId}"/>
      <c:crosses val="autoZero"/>
      <c:auto val="1"/>
      <c:lblAlgn val="ctr"/>
      <c:lblOffset val="100"/>
    </c:catAx>`

    const buildValAx = (id, crossId, pos = 'l', crossBetween = 'between') => `<c:valAx>
      <c:axId val="${id}"/>
      <c:scaling><c:orientation val="minMax"/></c:scaling>
      <c:delete val="0"/>
      <c:axPos val="${pos}"/>
      <c:majorGridlines/>
      <c:numFmt formatCode="General" sourceLinked="1"/>
      <c:majorTickMark val="none"/>
      <c:minorTickMark val="none"/>
      <c:tickLblPos val="nextTo"/>
      <c:crossAx val="${crossId}"/>
      <c:crosses val="autoZero"/>
      <c:crossBetween val="${crossBetween}"/>
    </c:valAx>`

    switch (canonicalType) {
      case 'column': {
        const serList = series.map((s, i) => buildSerXml(s, i)).join('')
        plotXml = `<c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <c:varyColors val="0"/>
          ${serList}
          <c:gapWidth val="150"/>
          <c:overlap val="0"/>
          <c:axId val="${axIdCat}"/>
          <c:axId val="${axIdVal}"/>
        </c:barChart>`
        axesXml = `${buildCatAx(axIdCat, axIdVal, 'b')}\n        ${buildValAx(axIdVal, axIdCat, 'l', 'between')}`
        break
      }

      case 'bar': {
        const serList = series.map((s, i) => buildSerXml(s, i)).join('')
        plotXml = `<c:barChart>
          <c:barDir val="bar"/>
          <c:grouping val="clustered"/>
          <c:varyColors val="0"/>
          ${serList}
          <c:gapWidth val="150"/>
          <c:overlap val="0"/>
          <c:axId val="${axIdCat}"/>
          <c:axId val="${axIdVal}"/>
        </c:barChart>`
        axesXml = `${buildCatAx(axIdCat, axIdVal, 'l')}\n        ${buildValAx(axIdVal, axIdCat, 'b', 'between')}`
        break
      }

      case 'line': {
        const serList = series.map((s, i) => buildSerXml(s, i)).join('')
        plotXml = `<c:lineChart>
          <c:grouping val="standard"/>
          <c:varyColors val="0"/>
          ${serList}
          <c:marker val="1"/>
          <c:smooth val="0"/>
          <c:axId val="${axIdCat}"/>
          <c:axId val="${axIdVal}"/>
        </c:lineChart>`
        axesXml = `${buildCatAx(axIdCat, axIdVal, 'b')}\n        ${buildValAx(axIdVal, axIdCat, 'l', 'between')}`
        break
      }

      case 'pie': {
        const serList = series
          .slice(0, 1)
          .map((s, i) => buildSerXml(s, i))
          .join('')
        plotXml = `<c:pieChart>
          <c:varyColors val="1"/>
          ${serList}
          <c:firstSliceAng val="0"/>
        </c:pieChart>`
        axesXml = '' // Pie charts do not use catAx / valAx
        break
      }

      case 'area': {
        const serList = series.map((s, i) => buildSerXml(s, i)).join('')
        plotXml = `<c:areaChart>
          <c:grouping val="standard"/>
          <c:varyColors val="0"/>
          ${serList}
          <c:axId val="${axIdCat}"/>
          <c:axId val="${axIdVal}"/>
        </c:areaChart>`
        axesXml = `${buildCatAx(axIdCat, axIdVal, 'b')}\n        ${buildValAx(axIdVal, axIdCat, 'l', 'midCat')}`
        break
      }

      case 'scatter': {
        const serList = series.map((s, i) => buildSerXml(s, i, true)).join('')
        plotXml = `<c:scatterChart>
          <c:scatterStyle val="lineMarker"/>
          <c:varyColors val="0"/>
          ${serList}
          <c:axId val="${axIdCat}"/>
          <c:axId val="${axIdVal}"/>
        </c:scatterChart>`
        // Scatter uses two valAx elements (bottom and left)
        axesXml = `${buildValAx(axIdCat, axIdVal, 'b', 'between')}\n        ${buildValAx(axIdVal, axIdCat, 'l', 'between')}`
        break
      }

      case 'radar': {
        const serList = series.map((s, i) => buildSerXml(s, i)).join('')
        plotXml = `<c:radarChart>
          <c:radarStyle val="standard"/>
          <c:varyColors val="0"/>
          ${serList}
          <c:axId val="${axIdCat}"/>
          <c:axId val="${axIdVal}"/>
        </c:radarChart>`
        axesXml = `${buildCatAx(axIdCat, axIdVal, 'b')}\n        ${buildValAx(axIdVal, axIdCat, 'l', 'between')}`
        break
      }

      case 'surface': {
        const axIdSer = '123458'
        const serList = series.map((s, i) => buildSerXml(s, i)).join('')
        plotXml = `<c:surfaceChart>
          <c:wireframe val="0"/>
          ${serList}
          <c:axId val="${axIdCat}"/>
          <c:axId val="${axIdVal}"/>
          <c:axId val="${axIdSer}"/>
        </c:surfaceChart>`
        const buildSerAx = (id, crossId) => `<c:serAx>
          <c:axId val="${id}"/>
          <c:scaling><c:orientation val="minMax"/></c:scaling>
          <c:delete val="0"/>
          <c:axPos val="b"/>
          <c:majorTickMark val="none"/>
          <c:minorTickMark val="none"/>
          <c:tickLblPos val="nextTo"/>
          <c:crossAx val="${crossId}"/>
          <c:crosses val="autoZero"/>
        </c:serAx>`
        axesXml = `${buildCatAx(axIdCat, axIdVal, 'b')}\n        ${buildValAx(axIdVal, axIdCat, 'l', 'between')}\n        ${buildSerAx(axIdSer, axIdVal)}`
        break
      }

      case 'stock': {
        const serList = series.map((s, i) => buildSerXml(s, i)).join('')
        plotXml = `<c:stockChart>
          ${serList}
          <c:hiLowLines/>
          <c:upDownBars><c:gapWidth val="150"/></c:upDownBars>
          <c:axId val="${axIdCat}"/>
          <c:axId val="${axIdVal}"/>
        </c:stockChart>`
        axesXml = `${buildCatAx(axIdCat, axIdVal, 'b')}\n        ${buildValAx(axIdVal, axIdCat, 'l', 'between')}`
        break
      }

      case 'map': {
        const serList = series.map((s, i) => buildSerXml(s, i)).join('')
        plotXml = `<c:areaChart>
          <c:grouping val="standard"/>
          <c:varyColors val="0"/>
          ${serList}
          <c:axId val="${axIdCat}"/>
          <c:axId val="${axIdVal}"/>
        </c:areaChart>`
        axesXml = `${buildCatAx(axIdCat, axIdVal, 'b')}\n        ${buildValAx(axIdVal, axIdCat, 'l', 'midCat')}`
        break
      }

      case 'combo': {
        // Multi-plot combining bar and line charts
        const barSeries = series.filter(s => s.type === 'bar' || s.type === 'column')
        const lineSeries = series.filter(s => s.type === 'line' || s.type === 'area')

        // If no explicit types, split half-and-half
        const effectiveBar = barSeries.length > 0 || lineSeries.length > 0 ? barSeries : [series[0]]
        const effectiveLine =
          barSeries.length > 0 || lineSeries.length > 0
            ? lineSeries
            : series.length > 1
              ? series.slice(1)
              : [series[0]]

        const barSerList = effectiveBar.map((s, i) => buildSerXml(s, i)).join('')
        const lineSerList = effectiveLine
          .map((s, i) => buildSerXml(s, effectiveBar.length + i))
          .join('')

        plotXml = `<c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <c:varyColors val="0"/>
          ${barSerList}
          <c:gapWidth val="150"/>
          <c:overlap val="0"/>
          <c:axId val="${axIdCat}"/>
          <c:axId val="${axIdVal}"/>
        </c:barChart>
        <c:lineChart>
          <c:grouping val="standard"/>
          <c:varyColors val="0"/>
          ${lineSerList}
          <c:marker val="1"/>
          <c:smooth val="0"/>
          <c:axId val="${axIdCat}"/>
          <c:axId val="${axIdVal}"/>
        </c:lineChart>`

        axesXml = `${buildCatAx(axIdCat, axIdVal, 'b')}\n        ${buildValAx(axIdVal, axIdCat, 'l', 'between')}`
        break
      }

      default: {
        const serList = series.map((s, i) => buildSerXml(s, i)).join('')
        plotXml = `<c:barChart>
          <c:barDir val="col"/>
          <c:grouping val="clustered"/>
          <c:varyColors val="0"/>
          ${serList}
          <c:gapWidth val="150"/>
          <c:overlap val="0"/>
          <c:axId val="${axIdCat}"/>
          <c:axId val="${axIdVal}"/>
        </c:barChart>`
        axesXml = `${buildCatAx(axIdCat, axIdVal, 'b')}\n        ${buildValAx(axIdVal, axIdCat, 'l', 'between')}`
      }
    }

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <c:date1904 val="0"/>
  <c:lang val="en-US"/>
  <c:roundedCorners val="0"/>
  <c:chart>
    ${titleXml}
    <c:autoTitleDeleted val="0"/>
    ${view3DXml}
    <c:plotArea>
      <c:layout/>
      ${plotXml}
      ${axesXml}
    </c:plotArea>
    ${legendXml}
    <c:plotVisOnly val="1"/>
    <c:dispBlanksAs val="gap"/>
    <c:showDLblsOverMax val="0"/>
  </c:chart>
  <c:externalData r:id="${rId}">
    <c:autoUpdate val="0"/>
  </c:externalData>
</c:chartSpace>`
  }
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
   * Updates an existing chart XML string (standard or extended) with new data and options.
   */
  static updateChartSpaceXml(xml, data = {}, options = {}) {
    if (xml.includes('<cx:chartSpace')) {
      return ExtendedChartGenerator.updateChartSpaceXml(xml, data, options)
    }

    let updated = xml
    if (data.categories && data.categories.length > 0) {
      updated = this.updateCategories(updated, data.categories)
    }
    if (data.series && data.series.length > 0) {
      const catLen = data.categories ? data.categories.length : 0
      updated = this.updateSeries(updated, data.series, catLen)
    }
    if (options.title !== undefined || (data && data.title !== undefined)) {
      const title = options.title !== undefined ? options.title : data.title
      updated = this.updateTitle(updated, title)
    }
    return updated
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
            const dLblLayoutMatch = /(<c:layout>[\s\S]*?<\/c:layout>|<c:layout\/>)/.exec(
              dLblContent
            )
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
      ctr: 'ctr',
      insideEnd: 'inEnd',
      inEnd: 'inEnd',
      insideBase: 'inBase',
      inBase: 'inBase',
      outsideEnd: 'outEnd',
      outEnd: 'outEnd',
      bestFit: 'bestFit',
      left: 'l',
      l: 'l',
      right: 'r',
      r: 'r',
      top: 't',
      t: 't',
      bottom: 'b',
      b: 'b',
    }
    let openxmlPos = position ? posMap[position] || position : null
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
    if (options.showLegendKey !== undefined) {
      xml += `<c:showLegendKey val="${options.showLegendKey ? '1' : '0'}"/>`
    } else if (existingShowTags['showLegendKey']) {
      xml += existingShowTags['showLegendKey']
    } else {
      xml += `<c:showLegendKey val="0"/>`
    }

    // showVal
    const defaultShowVal = hasCustomLabels ? '0' : '1'
    if (options.showVal !== undefined) {
      xml += `<c:showVal val="${options.showVal ? '1' : '0'}"/>`
    } else if (showSeriesNameInBar) {
      xml += `<c:showVal val="0"/>`
    } else if (existingShowTags['showVal'] && !hasCustomLabels) {
      xml += existingShowTags['showVal']
    } else {
      xml += `<c:showVal val="${defaultShowVal}"/>`
    }

    // showCatName
    if (options.showCatName !== undefined) {
      xml += `<c:showCatName val="${options.showCatName ? '1' : '0'}"/>`
    } else if (existingShowTags['showCatName']) {
      xml += existingShowTags['showCatName']
    } else {
      xml += `<c:showCatName val="0"/>`
    }

    // showSerName
    if (options.showSerName !== undefined) {
      xml += `<c:showSerName val="${options.showSerName ? '1' : '0'}"/>`
    } else if (showSeriesNameInBar) {
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
