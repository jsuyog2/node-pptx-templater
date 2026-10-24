/**
 * @fileoverview ExtendedChartGenerator - Generates Office 2016+ extended chart XML (<cx:chartSpace>).
 * Supports Treemap, Sunburst, Histogram, Box & Whisker, Waterfall, and Funnel chart types.
 *
 * Conforms strictly to Microsoft Office 2016+ DrawingML ChartEx schema ([MS-ODRAWXML]):
 * Reverse-engineered from ChartPowerPointSample.pptx (Slides 2-5).
 * Namespace: xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"
 */

'use strict'

const { generateGuid } = require('../../utils/idUtils.js')

class ExtendedChartGenerator {
  static #escapeXml(str) {
    if (str === null || str === undefined) return ''
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;')
  }

  static getColumnLetter(colIndex) {
    let letter = ''
    let temp = colIndex
    while (temp >= 0) {
      letter = String.fromCharCode((temp % 26) + 65) + letter
      temp = Math.floor(temp / 26) - 1
    }
    return letter
  }

  /**
   * Maps canonical chart type name to Office 2016 series layoutId.
   */
  static getLayoutId(chartType) {
    const type = String(chartType).toLowerCase().trim()
    switch (type) {
      case 'treemap':
        return 'treemap'
      case 'sunburst':
        return 'sunburst'
      case 'histogram':
        return 'clusteredColumn'
      case 'boxwhisker':
      case 'boxandwhisker':
      case 'box_and_whisker':
      case 'box-and-whisker':
        return 'boxWhisker'
      case 'waterfall':
        return 'waterfall'
      case 'funnel':
        return 'funnel'
      case 'map':
      case 'regionmap':
      case 'geochart':
        return 'regionMap'
      default:
        return 'clusteredColumn'
    }
  }

  /**
   * Generates a complete <cx:chartSpace> XML string for Office 2016 extended charts.
   *
   * @param {Object} params
   * @param {string} params.type - Chart type (treemap, sunburst, histogram, boxWhisker, waterfall, funnel).
   * @param {Object} params.data - { categories: string[], series: Array<{ name: string, values: number[] }>, hierarchyLevels?: string[][] }
   * @param {string} [params.title] - Chart title.
   * @param {Object} [params.options] - Styling and layout options.
   * @param {string} [params.rId] - Embedded workbook relationship ID.
   * @returns {string} XML string for ppt/charts/chartExN.xml.
   */
  static generateChartSpaceXml({ type, data = {}, title, options = {}, rId = 'rId1' }) {
    const layoutId = this.getLayoutId(type)
    const isHistogram = layoutId === 'clusteredColumn'
    const isTreemap = layoutId === 'treemap'
    const isSunburst = layoutId === 'sunburst'
    const isBoxWhisker = layoutId === 'boxWhisker'
    const isWaterfall = layoutId === 'waterfall'
    const isFunnel = layoutId === 'funnel'
    const isMap = layoutId === 'regionMap'

    const series =
      data.series && data.series.length > 0
        ? data.series
        : [{ name: 'Series 1', values: [10, 20, 30] }]
    const firstSeries = series[0]
    const values = firstSeries.values || []
    const valCount = values.length

    let categories = data.categories
    if (!categories || categories.length === 0) {
      categories = isHistogram
        ? []
        : Array.from({ length: valCount || 4 }, (_, i) => `Category ${i + 1}`)
    }
    const catCount = categories.length

    // Category points & hierarchy levels
    let catLevelsXml = ''
    let catColsCount = 1
    if (
      data.hierarchyLevels &&
      Array.isArray(data.hierarchyLevels) &&
      data.hierarchyLevels.length > 0
    ) {
      catColsCount = data.hierarchyLevels.length
      catLevelsXml = data.hierarchyLevels
        .map(lvlArr => {
          const pts = lvlArr
            .map((cat, i) => `<cx:pt idx="${i}">${this.#escapeXml(cat)}</cx:pt>`)
            .join('')
          return `<cx:lvl ptCount="${lvlArr.length}">${pts}</cx:lvl>`
        })
        .join('')
    } else {
      const catPts = categories
        .map((cat, i) => `<cx:pt idx="${i}">${this.#escapeXml(cat)}</cx:pt>`)
        .join('')
      catLevelsXml = `<cx:lvl ptCount="${catCount}">${catPts}</cx:lvl>`
    }

    // Title XML
    let titleXml = ''
    const titleText = this.#escapeXml(title || options.title || (type ? String(type) : ''))
    if (titleText) {
      titleXml = `<cx:title pos="t" align="ctr" overlay="0">
      <cx:tx>
        <cx:rich>
          <a:bodyPr spcFirstLastPara="1" vertOverflow="ellipsis" horzOverflow="overflow" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr" anchorCtr="1"/>
          <a:lstStyle/>
          <a:p>
            <a:pPr algn="ctr" rtl="0"><a:defRPr/></a:pPr>
            <a:r>
              <a:rPr lang="en-US" sz="1862" b="0" i="0" u="none" strike="noStrike" baseline="0" dirty="0">
                <a:solidFill><a:prstClr val="black"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:prstClr></a:solidFill>
                <a:latin typeface="Calibri"/>
              </a:rPr>
              <a:t>${titleText}</a:t>
            </a:r>
            <a:endParaRPr lang="en-US" sz="1862" b="0" i="0" u="none" strike="noStrike" baseline="0" dirty="0">
              <a:solidFill><a:prstClr val="black"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:prstClr></a:solidFill>
              <a:latin typeface="Calibri"/>
            </a:endParaRPr>
          </a:p>
        </cx:rich>
      </cx:tx>
    </cx:title>`
    }

    // Build Data and Series XML
    let dataXml = ''
    let seriesXml = ''
    let axesXml = ''
    let legendXml = ''

    if (isHistogram) {
      // 1. Histogram (clusteredColumn layoutId with continuous numeric values in Col A)
      const valPts = values
        .map((val, i) => `<cx:pt idx="${i}">${val !== null && val !== undefined ? val : 0}</cx:pt>`)
        .join('')
      dataXml = `<cx:data id="0">
      <cx:numDim type="val">
        <cx:f>Sheet1!$A$2:$A$${valCount + 1}</cx:f>
        <cx:lvl ptCount="${valCount}" formatCode="General">${valPts}</cx:lvl>
      </cx:numDim>
    </cx:data>`

      let binningXml = '<cx:binning intervalClosed="r"/>'
      if (options.binSize !== undefined || options.interval !== undefined) {
        const size = options.binSize !== undefined ? options.binSize : options.interval
        binningXml = `<cx:binning intervalClosed="r"><cx:binSize>${size}</cx:binSize></cx:binning>`
      } else if (options.binCount !== undefined) {
        binningXml = `<cx:binning intervalClosed="r"><cx:binCount>${options.binCount}</cx:binCount></cx:binning>`
      }

      seriesXml = `<cx:series layoutId="clusteredColumn" uniqueId="${generateGuid()}">
          <cx:tx>
            <cx:txData>
              <cx:f>Sheet1!$A$1</cx:f>
              <cx:v>${this.#escapeXml(firstSeries.name || 'Series1')}</cx:v>
            </cx:txData>
          </cx:tx>
          <cx:dataId val="0"/>
          <cx:layoutPr>${binningXml}</cx:layoutPr>
        </cx:series>`

      axesXml = `
      <cx:axis id="0"><cx:catScaling gapWidth="0"/><cx:tickLabels/></cx:axis>
      <cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis>`
    } else if (isTreemap || isSunburst) {
      // 2. Treemap & Sunburst (categorical hierarchy in A..X, size dimension in next column)
      const lastCatCol = this.getColumnLetter(catColsCount - 1)
      const valCol = this.getColumnLetter(catColsCount)
      const valPts = values
        .map((val, i) => `<cx:pt idx="${i}">${val !== null && val !== undefined ? val : 0}</cx:pt>`)
        .join('')

      dataXml = `<cx:data id="0">
      <cx:strDim type="cat">
        <cx:f>Sheet1!$A$2:$${lastCatCol}$${catCount + 1}</cx:f>
        ${catLevelsXml}
      </cx:strDim>
      <cx:numDim type="size">
        <cx:f>Sheet1!$${valCol}$2:$${valCol}$${valCount + 1}</cx:f>
        <cx:lvl ptCount="${valCount}" formatCode="General">${valPts}</cx:lvl>
      </cx:numDim>
    </cx:data>`

      if (isTreemap) {
        seriesXml = `<cx:series layoutId="treemap" uniqueId="${generateGuid()}">
          <cx:tx>
            <cx:txData>
              <cx:f>Sheet1!$${valCol}$1</cx:f>
              <cx:v>${this.#escapeXml(firstSeries.name || 'Series1')}</cx:v>
            </cx:txData>
          </cx:tx>
          <cx:dataLabels pos="inEnd"><cx:visibility seriesName="0" categoryName="1" value="0"/></cx:dataLabels>
          <cx:dataId val="0"/>
          <cx:layoutPr><cx:parentLabelLayout val="overlapping"/></cx:layoutPr>
        </cx:series>`
        legendXml = `<cx:legend pos="t" align="ctr" overlay="0"/>`
      } else {
        seriesXml = `<cx:series layoutId="sunburst" uniqueId="${generateGuid()}">
          <cx:tx>
            <cx:txData>
              <cx:f>Sheet1!$${valCol}$1</cx:f>
              <cx:v>${this.#escapeXml(firstSeries.name || 'Series1')}</cx:v>
            </cx:txData>
          </cx:tx>
          <cx:dataLabels pos="ctr"><cx:visibility seriesName="0" categoryName="1" value="0"/></cx:dataLabels>
          <cx:dataId val="0"/>
        </cx:series>`
      }
    } else if (isWaterfall) {
      // 3. Waterfall
      const valPts = values
        .map((val, i) => `<cx:pt idx="${i}">${val !== null && val !== undefined ? val : 0}</cx:pt>`)
        .join('')

      dataXml = `<cx:data id="0">
      <cx:strDim type="cat">
        <cx:f>Sheet1!$A$2:$A$${catCount + 1}</cx:f>
        ${catLevelsXml}
      </cx:strDim>
      <cx:numDim type="val">
        <cx:f>Sheet1!$B$2:$B$${valCount + 1}</cx:f>
        <cx:lvl ptCount="${valCount}" formatCode="General">${valPts}</cx:lvl>
      </cx:numDim>
    </cx:data>`

      let subtotalsXml = ''
      if (options.subtotals && Array.isArray(options.subtotals)) {
        subtotalsXml = `<cx:subtotals>${options.subtotals.map(idx => `<cx:idx val="${idx}"/>`).join('')}</cx:subtotals>`
      } else {
        const subtotalIndices = []
        categories.forEach((cat, idx) => {
          if (/total|ending|net|summary/i.test(String(cat))) {
            subtotalIndices.push(idx)
          }
        })
        if (subtotalIndices.length > 0) {
          subtotalsXml = `<cx:subtotals>${subtotalIndices.map(i => `<cx:idx val="${i}"/>`).join('')}</cx:subtotals>`
        }
      }

      seriesXml = `<cx:series layoutId="waterfall" uniqueId="${generateGuid()}">
          <cx:tx>
            <cx:txData>
              <cx:f>Sheet1!$B$1</cx:f>
              <cx:v>${this.#escapeXml(firstSeries.name || 'Series1')}</cx:v>
            </cx:txData>
          </cx:tx>
          <cx:dataLabels pos="outEnd"><cx:visibility seriesName="0" categoryName="0" value="1"/></cx:dataLabels>
          <cx:dataId val="0"/>
          <cx:layoutPr>${subtotalsXml}</cx:layoutPr>
        </cx:series>`

      axesXml = `
      <cx:axis id="0"><cx:catScaling gapWidth="0.5"/><cx:tickLabels/></cx:axis>
      <cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis>`
      legendXml = `<cx:legend pos="t" align="ctr" overlay="0"/>`
    } else if (isFunnel) {
      // 4. Funnel (Single cat scaling axis id="1", no axis 0)
      const valPts = values
        .map((val, i) => `<cx:pt idx="${i}">${val !== null && val !== undefined ? val : 0}</cx:pt>`)
        .join('')

      dataXml = `<cx:data id="0">
      <cx:strDim type="cat">
        <cx:f>Sheet1!$A$2:$A$${catCount + 1}</cx:f>
        ${catLevelsXml}
      </cx:strDim>
      <cx:numDim type="val">
        <cx:f>Sheet1!$B$2:$B$${valCount + 1}</cx:f>
        <cx:lvl ptCount="${valCount}" formatCode="General">${valPts}</cx:lvl>
      </cx:numDim>
    </cx:data>`

      seriesXml = `<cx:series layoutId="funnel" uniqueId="${generateGuid()}">
          <cx:tx>
            <cx:txData>
              <cx:f>Sheet1!$B$1</cx:f>
              <cx:v>${this.#escapeXml(firstSeries.name || 'Series1')}</cx:v>
            </cx:txData>
          </cx:tx>
          <cx:dataLabels><cx:visibility seriesName="0" categoryName="0" value="1"/></cx:dataLabels>
          <cx:dataId val="0"/>
        </cx:series>`

      axesXml = `
      <cx:axis id="1"><cx:catScaling gapWidth="0.0599999987"/><cx:tickLabels/></cx:axis>`
    } else if (isBoxWhisker) {
      // 5. Box and Whisker (Can have 1 or more series, each with own data block)
      const dataBlocks = []
      const seriesBlocks = []

      series.forEach((ser, sIdx) => {
        const colLetter = this.getColumnLetter(sIdx + 1)
        const serValues = ser.values || []
        const sValPts = serValues
          .map(
            (val, i) => `<cx:pt idx="${i}">${val !== null && val !== undefined ? val : 0}</cx:pt>`
          )
          .join('')

        dataBlocks.push(`<cx:data id="${sIdx}">
      <cx:strDim type="cat">
        <cx:f>Sheet1!$A$2:$A$${catCount + 1}</cx:f>
        ${catLevelsXml}
      </cx:strDim>
      <cx:numDim type="val">
        <cx:f>Sheet1!$${colLetter}$2:$${colLetter}$${serValues.length + 1}</cx:f>
        <cx:lvl ptCount="${serValues.length}" formatCode="General">${sValPts}</cx:lvl>
      </cx:numDim>
    </cx:data>`)

        seriesBlocks.push(`<cx:series layoutId="boxWhisker" uniqueId="${generateGuid()}">
          <cx:tx>
            <cx:txData>
              <cx:f>Sheet1!$${colLetter}$1</cx:f>
              <cx:v>${this.#escapeXml(ser.name || `Series${sIdx + 1}`)}</cx:v>
            </cx:txData>
          </cx:tx>
          <cx:dataId val="${sIdx}"/>
          <cx:layoutPr>
            <cx:visibility meanLine="0" meanMarker="1" nonoutliers="0" outliers="1"/>
            <cx:statistics quartileMethod="exclusive"/>
          </cx:layoutPr>
        </cx:series>`)
      })

      dataXml = dataBlocks.join('\n    ')
      seriesXml = seriesBlocks.join('\n        ')
      axesXml = `
      <cx:axis id="0"><cx:catScaling gapWidth="1"/><cx:tickLabels/></cx:axis>
      <cx:axis id="1"><cx:valScaling/><cx:majorGridlines/><cx:tickLabels/></cx:axis>`
      if (series.length > 1) {
        legendXml = `<cx:legend pos="t" align="ctr" overlay="0"/>`
      }
    } else if (isMap) {
      // 6. Map / Geographic Region Map
      const valPts = values
        .map((val, i) => `<cx:pt idx="${i}">${val !== null && val !== undefined ? val : 0}</cx:pt>`)
        .join('')

      dataXml = `<cx:data id="0">
      <cx:strDim type="cat">
        <cx:f>Sheet1!$A$2:$A$${catCount + 1}</cx:f>
        ${catLevelsXml}
      </cx:strDim>
      <cx:numDim type="colorVal">
        <cx:f>Sheet1!$B$2:$B$${valCount + 1}</cx:f>
        <cx:lvl ptCount="${valCount}" formatCode="General">${valPts}</cx:lvl>
      </cx:numDim>
    </cx:data>`

      const projection = options.projection || options.projectionType || 'mercator'
      const viewedRegion = options.viewedRegion || options.viewedRegionType || 'dataOnly'
      const cultureLang = options.cultureLanguage || 'en-US'
      const cultureRegion = options.cultureRegion || 'US'

      seriesXml = `<cx:series layoutId="regionMap" uniqueId="${generateGuid()}">
          <cx:tx>
            <cx:txData>
              <cx:f>Sheet1!$B$1</cx:f>
              <cx:v>${this.#escapeXml(firstSeries.name || 'Series1')}</cx:v>
            </cx:txData>
          </cx:tx>
          <cx:dataId val="0"/>
          <cx:layoutPr>
            <cx:geography projectionType="${projection}" viewedRegionType="${viewedRegion}" cultureLanguage="${cultureLang}" cultureRegion="${cultureRegion}" attribution="Powered by Bing"/>
          </cx:layoutPr>
        </cx:series>`

      axesXml = ''
      legendXml = `<cx:legend pos="r" align="min" overlay="0"/>`
    } else {
      // Fallback extended chart layout
      const valPts = values
        .map((val, i) => `<cx:pt idx="${i}">${val !== null && val !== undefined ? val : 0}</cx:pt>`)
        .join('')
      dataXml = `<cx:data id="0">
      <cx:strDim type="cat">
        <cx:f>Sheet1!$A$2:$A$${catCount + 1}</cx:f>
        ${catLevelsXml}
      </cx:strDim>
      <cx:numDim type="val">
        <cx:f>Sheet1!$B$2:$B$${valCount + 1}</cx:f>
        <cx:lvl ptCount="${valCount}" formatCode="General">${valPts}</cx:lvl>
      </cx:numDim>
    </cx:data>`
      seriesXml = `<cx:series layoutId="${layoutId}" uniqueId="${generateGuid()}">
          <cx:tx>
            <cx:txData>
              <cx:f>Sheet1!$B$1</cx:f>
              <cx:v>${this.#escapeXml(firstSeries.name || 'Series1')}</cx:v>
            </cx:txData>
          </cx:tx>
          <cx:dataId val="0"/>
        </cx:series>`
    }

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <cx:chartData>
    <cx:externalData r:id="${rId || 'rId1'}" cx:autoUpdate="0"/>
    ${dataXml}
  </cx:chartData>
  <cx:chart>
    ${titleXml}
    <cx:plotArea>
      <cx:plotAreaRegion>
        ${seriesXml}
      </cx:plotAreaRegion>${axesXml}
    </cx:plotArea>
    ${legendXml}
  </cx:chart>
</cx:chartSpace>`
  }

  /**
   * Updates an existing <cx:chartSpace> XML string with new data and options.
   *
   * @param {string} xml - Raw <cx:chartSpace> XML.
   * @param {Object} data - Updated chart data.
   * @param {Object} [options] - Additional update options.
   * @returns {string} Updated XML.
   */
  static updateChartSpaceXml(xml, data = {}, options = {}) {
    let updated = xml
    const { categories, series } = data

    // Check if histogram
    const isHistogram = xml.includes('layoutId="clusteredColumn"') && xml.includes('<cx:binning')

    // Update categories in <cx:strDim type="cat"> (Histogram has no category dimension)
    if (!isHistogram && categories && Array.isArray(categories)) {
      const catCount = categories.length
      const catPts = categories
        .map((cat, i) => `<cx:pt idx="${i}">${this.#escapeXml(cat)}</cx:pt>`)
        .join('')
      const strDimRegex = /<cx:strDim\b[^>]*type="cat"[^>]*>[\s\S]*?<\/cx:strDim>/
      if (strDimRegex.test(updated)) {
        updated = updated.replace(
          strDimRegex,
          `<cx:strDim type="cat"><cx:f>Sheet1!$A$2:$A$${catCount + 1}</cx:f><cx:lvl ptCount="${catCount}">${catPts}</cx:lvl></cx:strDim>`
        )
      }
    }

    // Update values
    if (series && Array.isArray(series) && series.length > 0) {
      const ser = series[0]
      const values = ser.values || []
      const valCount = values.length
      const valPts = values
        .map((val, i) => `<cx:pt idx="${i}">${val !== null && val !== undefined ? val : 0}</cx:pt>`)
        .join('')

      if (isHistogram) {
        // Histogram values are in Col A
        const numDimRegex = /<cx:numDim\b[^>]*type="val"[^>]*>[\s\S]*?<\/cx:numDim>/
        if (numDimRegex.test(updated)) {
          updated = updated.replace(
            numDimRegex,
            `<cx:numDim type="val"><cx:f>Sheet1!$A$2:$A$${valCount + 1}</cx:f><cx:lvl ptCount="${valCount}" formatCode="General">${valPts}</cx:lvl></cx:numDim>`
          )
        }
      } else {
        // Other types have values in Col B (or size dimension for treemap/sunburst, colorVal for map)
        const numDimRegex =
          /<cx:numDim\b[^>]*type="(?:val|size|colorVal)"[^>]*>[\s\S]*?<\/cx:numDim>/
        const dimTypeMatch = updated.match(/<cx:numDim\b[^>]*type="([^"]+)"/)
        const dimType = dimTypeMatch ? dimTypeMatch[1] : 'val'
        if (numDimRegex.test(updated)) {
          updated = updated.replace(
            numDimRegex,
            `<cx:numDim type="${dimType}"><cx:f>Sheet1!$B$2:$B$${valCount + 1}</cx:f><cx:lvl ptCount="${valCount}" formatCode="General">${valPts}</cx:lvl></cx:numDim>`
          )
        }
      }

      // Update series title in <cx:series> ... <cx:tx>
      if (ser.name !== undefined) {
        const colLetter = isHistogram ? 'A' : 'B'
        const serBlockRegex = /(<cx:series\b[^>]*>)([\s\S]*?)(<\/cx:series>)/
        if (serBlockRegex.test(updated)) {
          updated = updated.replace(serBlockRegex, (m, open, content, close) => {
            const newTx = `<cx:tx><cx:txData><cx:f>Sheet1!$${colLetter}$1</cx:f><cx:v>${this.#escapeXml(ser.name)}</cx:v></cx:txData></cx:tx>`
            const txRegex = /<cx:tx>[\s\S]*?<\/cx:tx>/
            if (txRegex.test(content)) {
              content = content.replace(txRegex, newTx)
            } else {
              content = newTx + content
            }
            return `${open}${content}${close}`
          })
        }
      }
    }

    // Update title if provided
    if (options.title !== undefined || data.title !== undefined) {
      const titleText = this.#escapeXml(options.title || data.title)
      const titleRegex = /<cx:title\b[^>]*>[\s\S]*?<\/cx:title>/
      const newTitle = `<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="1862" b="0"/><a:t>${titleText}</a:t></a:r></a:p></cx:rich></cx:tx></cx:title>`
      if (titleRegex.test(updated)) {
        updated = updated.replace(titleRegex, newTitle)
      } else {
        updated = updated.replace(/<cx:chart>/, `<cx:chart>\n    ${newTitle}`)
      }
    }

    return updated
  }
}

module.exports = {
  ExtendedChartGenerator,
}
