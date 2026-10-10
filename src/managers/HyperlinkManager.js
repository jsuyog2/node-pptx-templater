/**
 * @fileoverview HyperlinkManager - Manages hyperlinks in PPTX slides.
 *
 * Hyperlinks in OpenXML PPTX:
 * ─────────────────────────────────────────────────────────────────
 * Hyperlinks are defined as relationships and referenced in slide XML.
 *
 * External Hyperlink (text):
 *   1. Add relationship to slide .rels file:
 *      <Relationship Id="rId9" Type=".../hyperlink"
 *                   Target="https://example.com" TargetMode="External"/>
 *
 *   2. Reference in text run:
 *      <a:r>
 *        <a:rPr>
 *          <a:hlinkClick r:id="rId9"/>    ← click action
 *        </a:rPr>
 *        <a:t>Open Website</a:t>
 *      </a:r>
 *
 * Shape Hyperlink:
 *   Applied to a:sp via p:sp > p:nvSpPr > p:cNvSpPr:
 *      <a:hlinkClick r:id="rId9"/>
 *
 * Slide-to-Slide Navigation:
 *   Uses a special action setting:
 *      <a:hlinkClick r:id="rId9" action="ppaction://hlinksldjump"/>
 *   Where rId9 points to another slide (Target="../slides/slide3.xml")
 *
 * Mouse-Over Hyperlinks:
 *   Use <a:hlinkMouseOver r:id="..."/> instead of hlinkClick.
 */

const { createLogger } = require('../utils/logger.js')
const { PPTXError } = require('../utils/errors.js')
const { REL_TYPES } = require('./RelationshipManager.js')

const logger = createLogger('HyperlinkManager')

/**
 * @class HyperlinkManager
 * @description Manages hyperlink creation and modification in PPTX slides.
 */
class HyperlinkManager {
  /** @private @type {XMLParser} */
  #xmlParser
  /** @private @type {RelationshipManager} */
  #relationshipManager

  /**
   * @param {XMLParser} xmlParser
   * @param {RelationshipManager} relationshipManager
   */
  constructor(xmlParser, relationshipManager) {
    this.#xmlParser = xmlParser
    this.#relationshipManager = relationshipManager
  }

  /**
   * Adds an external hyperlink (URL) to text matching the given text in a slide.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {HyperlinkOptions} options
   * @param {string} options.text - Text to make clickable.
   * @param {string} options.url - Target URL.
   * @param {string} [options.tooltip] - Optional screen tip tooltip.
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   */
  addExternalHyperlink(slideIndex, options, slideManager, relationshipManager) {
    const { text, url, tooltip } = options
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    const slideXml = slideManager.getSlideXml(slideIndex)

    // Add the hyperlink relationship to the slide's .rels file
    const rId = relationshipManager.addRelationship(
      slideInfo.zipPath,
      REL_TYPES.HYPERLINK,
      url,
      'External'
    )

    // Update the slide XML to reference the new rId
    const updatedXml = this.#injectHyperlinkOnText(slideXml, text, rId, tooltip)
    slideManager.setSlideXml(slideIndex, updatedXml)

    logger.debug(`Added hyperlink to "${text}" → ${url} (${rId}) in slide ${slideIndex}`)
  }

  /**
   * Adds an inter-slide hyperlink (navigates to another slide).
   *
   * @param {number} sourceSlideIndex - Source slide (1-based).
   * @param {number} targetSlideIndex - Destination slide (1-based).
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   */
  addSlideHyperlink(sourceSlideIndex, targetSlideIndex, slideManager, relationshipManager) {
    const sourceInfo = slideManager.getSlideInfo(sourceSlideIndex)
    const targetInfo = slideManager.getSlideInfo(targetSlideIndex)

    // Build relative target path from source to target slide
    const relativePath = targetInfo.zipPath.split('/').pop()

    // Add relationship pointing to the target slide
    const rId = relationshipManager.addRelationship(
      sourceInfo.zipPath,
      REL_TYPES.SLIDE,
      relativePath
    )

    // Add hlinkClick with slide jump action to the slide number placeholder
    const slideXml = slideManager.getSlideXml(sourceSlideIndex)
    const updatedXml = this.#injectSlideJumpHyperlink(slideXml, rId)
    slideManager.setSlideXml(sourceSlideIndex, updatedXml)

    logger.debug(`Linked slide ${sourceSlideIndex} → slide ${targetSlideIndex} (${rId})`)
  }

  /**
   * Adds an inter-slide hyperlink to specific text.
   *
   * @param {number} sourceSlideIndex - Source slide (1-based).
   * @param {string} text - Text to make clickable.
   * @param {number} targetSlideIndex - Destination slide (1-based).
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   */
  addTextSlideLink(sourceSlideIndex, text, targetSlideIndex, slideManager, relationshipManager) {
    const sourceInfo = slideManager.getSlideInfo(sourceSlideIndex)
    const targetInfo = slideManager.getSlideInfo(targetSlideIndex)

    const relativePath = targetInfo.zipPath.split('/').pop()
    const rId = relationshipManager.addRelationship(
      sourceInfo.zipPath,
      REL_TYPES.SLIDE,
      relativePath
    )

    const slideXml = slideManager.getSlideXml(sourceSlideIndex)
    // Use injectHyperlinkOnText but append action attribute
    const actionAttr = 'action="ppaction://hlinksldjump"'
    const updatedXml = this.#injectHyperlinkOnText(slideXml, text, rId, null, actionAttr)
    slideManager.setSlideXml(sourceSlideIndex, updatedXml)

    logger.debug(
      `Linked text "${text}" on slide ${sourceSlideIndex} → slide ${targetSlideIndex} (${rId})`
    )
  }

  /**
   * Adds an inter-slide hyperlink to a shape/image.
   *
   * @param {number} sourceSlideIndex - Source slide (1-based).
   * @param {string} shapeName - Name/id of the shape.
   * @param {number} targetSlideIndex - Destination slide (1-based).
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   */
  addShapeSlideLink(
    sourceSlideIndex,
    shapeName,
    targetSlideIndex,
    slideManager,
    relationshipManager
  ) {
    const sourceInfo = slideManager.getSlideInfo(sourceSlideIndex)
    const targetInfo = slideManager.getSlideInfo(targetSlideIndex)

    const relativePath = targetInfo.zipPath.split('/').pop()
    const rId = relationshipManager.addRelationship(
      sourceInfo.zipPath,
      REL_TYPES.SLIDE,
      relativePath
    )

    const slideXml = slideManager.getSlideXml(sourceSlideIndex)
    const actionAttr = 'action="ppaction://hlinksldjump"'
    const updatedXml = this.#injectHyperlinkOnShape(slideXml, shapeName, rId, actionAttr)
    slideManager.setSlideXml(sourceSlideIndex, updatedXml)

    logger.debug(
      `Linked shape "${shapeName}" on slide ${sourceSlideIndex} → slide ${targetSlideIndex} (${rId})`
    )
  }

  /**
   * Adds a hyperlink to a shape by name.
   *
   * @param {number} slideIndex
   * @param {string} shapeName - cNvPr name attribute of the shape.
   * @param {string} url - Target URL.
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   */
  addShapeHyperlink(slideIndex, shapeName, url, slideManager, relationshipManager) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    const slideXml = slideManager.getSlideXml(slideIndex)

    const rId = relationshipManager.addRelationship(
      slideInfo.zipPath,
      REL_TYPES.HYPERLINK,
      url,
      'External'
    )

    const updatedXml = this.#injectHyperlinkOnShape(slideXml, shapeName, rId)
    slideManager.setSlideXml(slideIndex, updatedXml)
    logger.debug(`Added shape hyperlink on "${shapeName}" → ${url}`)
  }

  /**
   * Retrieves all hyperlinks present on a slide.
   *
   * @param {number} slideIndex - 1-based slide index.
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   * @returns {Array<Object>} List of hyperlinks with targets, types, and element references.
   */
  getHyperlinks(slideIndex, slideManager, relationshipManager) {
    const slideInfo = slideManager.getSlideInfo(slideIndex)
    if (!slideInfo) return []

    const slideXml = slideManager.getSlideXml(slideIndex)
    const rels = relationshipManager.getRelationships(slideInfo.zipPath)
    const relMap = new Map()
    for (const rel of rels) {
      relMap.set(rel.id, rel)
    }

    const allSlides = slideManager.getAllSlideInfo()
    const resolveSlideIndex = target => {
      if (!target) return undefined
      const targetFilename = target.split('/').pop()
      const match = allSlides.find(s => s.zipPath && s.zipPath.endsWith(targetFilename))
      return match ? match.index : undefined
    }

    const results = []
    const seenRIds = new Set()

    // Match all <a:hlinkClick .../> and <a:hlinkMouseOver .../> tags
    const hlinkRegex = /<a:hlink(?:Click|MouseOver)\b([^>]*)\/?>/g
    let match

    while ((match = hlinkRegex.exec(slideXml)) !== null) {
      const attrsStr = match[1]
      const matchPos = match.index

      const rIdMatch = /r:id="([^"]+)"/.exec(attrsStr) || /id="([^"]+)"/.exec(attrsStr)
      const actionMatch = /action="([^"]+)"/.exec(attrsStr)
      const tooltipMatch = /tooltip="([^"]+)"/.exec(attrsStr)

      const rId = rIdMatch ? rIdMatch[1] : null
      const action = actionMatch ? actionMatch[1] : null
      const tooltip = tooltipMatch ? tooltipMatch[1] : null

      if (rId) seenRIds.add(rId)

      let type = 'unknown'
      let target = null
      let url = null
      let targetSlide = undefined

      const rel = rId ? relMap.get(rId) : null
      if (rel) {
        if (rel.type === REL_TYPES.HYPERLINK || rel.targetMode === 'External') {
          type = 'external'
          url = rel.target
          target = rel.target
        } else if (rel.type === REL_TYPES.SLIDE || (action && action.includes('hlinksldjump'))) {
          type = 'slide'
          targetSlide = resolveSlideIndex(rel.target)
          target = targetSlide !== undefined ? targetSlide : rel.target
        } else {
          type = 'external'
          target = rel.target
          url = rel.target
        }
      } else if (action) {
        if (action.startsWith('ppaction://hlinkshowjump?s=')) {
          type = 'navigation'
          target = action.split('s=').pop()
        } else if (action.includes('hlinksldjump')) {
          type = 'slide'
          target = action
        } else {
          type = 'action'
          target = action
        }
      }

      // Determine enclosing element context
      let text = null
      let shapeName = null
      let elementType = 'unknown'

      // Check if inside text run <a:r>
      const beforeStr = slideXml.substring(Math.max(0, matchPos - 500), matchPos)
      const rStart = beforeStr.lastIndexOf('<a:r>')
      const rEndTag = beforeStr.lastIndexOf('</a:r>')

      if (rStart !== -1 && rStart > rEndTag) {
        elementType = 'text'
        const afterStr = slideXml.substring(matchPos, Math.min(slideXml.length, matchPos + 500))
        const rEnd = afterStr.indexOf('</a:r>')
        if (rEnd !== -1) {
          const runBlock = beforeStr.substring(rStart) + afterStr.substring(0, rEnd + 6)
          const tMatch = /<a:t>([^<]*)<\/a:t>/.exec(runBlock)
          if (tMatch) {
            text = tMatch[1]
          }
        }
      } else {
        // Check if inside shape <p:sp> or picture <p:pic>
        const shapeBefore = slideXml.substring(Math.max(0, matchPos - 2000), matchPos)
        const allNameMatches = [...shapeBefore.matchAll(/<p:cNvPr[^>]*\bname="([^"]+)"/g)]
        if (allNameMatches.length > 0) {
          shapeName = allNameMatches[allNameMatches.length - 1][1]
          elementType =
            shapeBefore.includes('<p:pic') &&
            shapeBefore.lastIndexOf('<p:pic') > shapeBefore.lastIndexOf('<p:sp')
              ? 'image'
              : 'shape'
        }
      }

      results.push({
        type,
        target,
        url: url || undefined,
        targetSlide,
        action: action || undefined,
        tooltip: tooltip || undefined,
        text: text || undefined,
        shapeName: shapeName || undefined,
        rId: rId || undefined,
        elementType,
      })
    }

    // Also include any standalone hyperlink relationships from .rels
    for (const rel of rels) {
      if (
        (rel.type === REL_TYPES.HYPERLINK || rel.targetMode === 'External') &&
        !seenRIds.has(rel.id)
      ) {
        results.push({
          type: 'external',
          target: rel.target,
          url: rel.target,
          rId: rel.id,
          elementType: 'relationship',
        })
      }
    }

    return results
  }

  /**
   * Removes a hyperlink from text or a shape in a slide.
   *
   * @param {number} slideIndex
   * @param {string|Object} textOrOptions - Text string or options with { text, shapeName }.
   * @param {SlideManager} slideManager
   * @param {RelationshipManager} relationshipManager
   */
  removeHyperlink(slideIndex, textOrOptions, slideManager, relationshipManager) {
    const slideXml = slideManager.getSlideXml(slideIndex)
    const slideInfo = slideManager.getSlideInfo(slideIndex)

    const text = typeof textOrOptions === 'string' ? textOrOptions : textOrOptions?.text || ''
    const shapeName = typeof textOrOptions === 'object' ? textOrOptions?.shapeName : null

    if (text) {
      const hlinkPattern = new RegExp(
        `<a:hlinkClick[^>]*r:id="(rId\\d+)"[^/]*/>[\\s\\S]*?<a:t>${this.#escapeRegex(text)}</a:t>`
      )
      const match = hlinkPattern.exec(slideXml)

      if (match) {
        const rId = match[1]
        const updatedXml = slideXml.replace(
          new RegExp(`<a:hlinkClick[^>]*r:id="${rId}"[^/]*/>`, 'g'),
          ''
        )
        slideManager.setSlideXml(slideIndex, updatedXml)
        relationshipManager.removeRelationship(slideInfo.zipPath, rId)
      }
    } else if (shapeName) {
      const namePattern = new RegExp(`name="${this.#escapeRegex(shapeName)}"`)
      const nameMatch = namePattern.exec(slideXml)
      if (nameMatch) {
        const spStart = slideXml.lastIndexOf('<p:sp>', nameMatch.index)
        const spEnd = slideXml.indexOf('</p:sp>', nameMatch.index)
        if (spStart !== -1 && spEnd !== -1) {
          const spXml = slideXml.substring(spStart, spEnd + '</p:sp>'.length)
          const rIdMatch = /r:id="(rId\d+)"/.exec(spXml)
          const updatedSpXml = spXml.replace(/<a:hlinkClick[^>]*\/>/g, '')
          const updatedXml =
            slideXml.substring(0, spStart) +
            updatedSpXml +
            slideXml.substring(spEnd + '</p:sp>'.length)
          slideManager.setSlideXml(slideIndex, updatedXml)
          if (rIdMatch) {
            relationshipManager.removeRelationship(slideInfo.zipPath, rIdMatch[1])
          }
        }
      }
    }
  }

  /**
   * Injects an hlinkClick reference on a text run matching the given text.
   *
   * @private
   * @param {string} slideXml - Slide XML.
   * @param {string} text - Target text to find.
   * @param {string} rId - Relationship ID.
   * @param {string} [tooltip] - Optional tooltip.
   * @param {string} [actionAttr] - Optional action attribute (e.g. for slide jump).
   * @returns {string} Updated slide XML.
   */
  #injectHyperlinkOnText(slideXml, text, rId, tooltip, actionAttr = '') {
    const escapedText = this.#escapeXml(text)
    const exactPattern = new RegExp(`(<a:t>)(${this.#escapeRegex(escapedText)})(<\/a:t>)`, 'g')

    if (!exactPattern.test(slideXml)) {
      // Check if text is a substring of an <a:t> node and split the containing run
      const subPattern = new RegExp(
        `(<a:t[^>]*>)([^<]*?)(${this.#escapeRegex(escapedText)})([^<]*?)(<\/a:t>)`
      )
      const subMatch = subPattern.exec(slideXml)

      if (subMatch) {
        const beforeText = subMatch[2]
        const matchedText = subMatch[3]
        const afterText = subMatch[4]

        // Find enclosing <a:r>
        const tIndex = subMatch.index
        const rStart = slideXml.lastIndexOf('<a:r>', tIndex)
        const rEnd = slideXml.indexOf('</a:r>', tIndex)

        if (rStart !== -1 && rEnd !== -1) {
          const runXml = slideXml.substring(rStart, rEnd + '</a:r>'.length)
          let rPrXml = ''
          const rPrMatch = /<a:rPr[^>]*>[\s\S]*?<\/a:rPr>|<a:rPr[^>]*\/>/.exec(runXml)
          if (rPrMatch) {
            rPrXml = rPrMatch[0]
          }

          const tipAttr = tooltip ? ` tooltip="${this.#escapeXml(tooltip)}"` : ''
          const actAttr = actionAttr ? ` ${actionAttr}` : ''
          const rIdAttr = rId ? ` r:id="${rId}"` : ''
          const hlinkXml = `<a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"${rIdAttr}${tipAttr}${actAttr}/>`

          let linkedRPrXml = ''
          if (rPrXml) {
            if (rPrXml.endsWith('/>')) {
              linkedRPrXml = rPrXml.slice(0, -2) + `>${hlinkXml}</a:rPr>`
            } else {
              linkedRPrXml = rPrXml.replace('</a:rPr>', `${hlinkXml}</a:rPr>`)
            }
          } else {
            linkedRPrXml = `<a:rPr lang="en-US" dirty="0">${hlinkXml}</a:rPr>`
          }

          let replacementRuns = ''
          if (beforeText) {
            replacementRuns += `<a:r>${rPrXml}<a:t xml:space="preserve">${beforeText}</a:t></a:r>`
          }
          replacementRuns += `<a:r>${linkedRPrXml}<a:t>${matchedText}</a:t></a:r>`
          if (afterText) {
            replacementRuns += `<a:r>${rPrXml}<a:t xml:space="preserve">${afterText}</a:t></a:r>`
          }

          return (
            slideXml.substring(0, rStart) +
            replacementRuns +
            slideXml.substring(rEnd + '</a:r>'.length)
          )
        }
      }

      logger.warn(`Text "${text}" not found in slide XML`)
      return slideXml
    }

    const tipAttr = tooltip ? ` tooltip="${this.#escapeXml(tooltip)}"` : ''
    const actAttr = actionAttr ? ` ${actionAttr}` : ''
    const rIdAttr = rId ? ` r:id="${rId}"` : ''
    const hlinkXml = `<a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"${rIdAttr}${tipAttr}${actAttr}/>`

    // We need to add the hlinkClick INSIDE the a:rPr of the text run containing our text
    let updated = slideXml

    // Find the text node
    const tStart = slideXml.indexOf(`<a:t>${escapedText}</a:t>`)
    if (tStart === -1) {
      const tStartPlain = slideXml.indexOf(`<a:t>${text}</a:t>`)
      if (tStartPlain === -1) {
        logger.warn(`Could not locate text "${text}" in slide XML`)
        return slideXml
      }
    }

    // Find the containing <a:r> tag
    const rStart = updated.lastIndexOf('<a:r>', tStart)
    const rEnd = updated.indexOf('</a:r>', tStart)

    if (rStart === -1 || rEnd === -1) {
      return slideXml
    }

    const runXml = updated.substring(rStart, rEnd + '</a:r>'.length)

    // Check if rPr exists
    if (runXml.includes('<a:rPr')) {
      const rPrEnd = runXml.indexOf('>', runXml.indexOf('<a:rPr'))
      const rPrIsSelfClosing = runXml[rPrEnd - 1] === '/'

      let newRunXml
      if (rPrIsSelfClosing) {
        const rPrStart = runXml.indexOf('<a:rPr')
        const rPrFull = runXml.substring(rPrStart, rPrEnd + 1)
        const rPrAttribs = rPrFull.replace('/>', '')
        newRunXml = runXml.replace(rPrFull, `${rPrAttribs}>${hlinkXml}</a:rPr>`)
      } else {
        const rPrClose = runXml.indexOf('</a:rPr>')
        newRunXml = runXml.substring(0, rPrClose) + hlinkXml + runXml.substring(rPrClose)
      }

      updated = updated.substring(0, rStart) + newRunXml + updated.substring(rEnd + '</a:r>'.length)
    } else {
      const tTagStart = runXml.indexOf('<a:t>')
      const newRunXml =
        runXml.substring(0, tTagStart) +
        `<a:rPr lang="en-US" dirty="0">${hlinkXml}</a:rPr>` +
        runXml.substring(tTagStart)
      updated = updated.substring(0, rStart) + newRunXml + updated.substring(rEnd + '</a:r>'.length)
    }

    return updated
  }

  /**
   * Injects a slide-jump hyperlink action on the slide number placeholder.
   * @private
   */
  #injectSlideJumpHyperlink(slideXml, rId) {
    const hlinkXml = `<a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rId}" action="ppaction://hlinksldjump"/>`

    // Look for slide number field (<a:fld type="slidenum">) or first text run
    const fldPattern = /<a:fld[^>]*type="slidenum"[^>]*>/
    if (fldPattern.test(slideXml)) {
      // Add action to the fld element
      return slideXml.replace(fldPattern, match => match.replace('>', `>${hlinkXml}`))
    }

    // Fallback: add to first text in slide
    return this.#injectHyperlinkOnFirstText(slideXml, rId)
  }

  /**
   * Injects hyperlink on the first text run in a slide.
   * @private
   */
  #injectHyperlinkOnFirstText(slideXml, rId) {
    const firstT = slideXml.indexOf('<a:t>')
    if (firstT === -1) return slideXml

    const text = slideXml.substring(firstT + 5, slideXml.indexOf('</a:t>', firstT))
    return this.#injectHyperlinkOnText(slideXml, text, rId)
  }

  /**
   * Injects hyperlink on a shape by adding hlinkClick to the cNvSpPr.
   * @private
   */
  #injectHyperlinkOnShape(slideXml, shapeName, rId, actionAttr = '') {
    const actAttr = actionAttr ? ` ${actionAttr}` : ''
    const rIdAttr = rId ? ` r:id="${rId}"` : ''
    const hlinkXml = `<a:hlinkClick xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"${rIdAttr}${actAttr}/>`

    // Find the shape by name
    const namePattern = new RegExp(`name="${this.#escapeRegex(shapeName)}"`)
    const nameMatch = namePattern.exec(slideXml)
    if (!nameMatch) {
      logger.warn(`Shape "${shapeName}" not found`)
      return slideXml
    }

    // Find the p:sp containing this shape and inject into nvSpPr
    const spStart = slideXml.lastIndexOf('<p:sp>', nameMatch.index)
    const spEnd = slideXml.indexOf('</p:sp>', nameMatch.index)

    if (spStart === -1 || spEnd === -1) return slideXml

    const spXml = slideXml.substring(spStart, spEnd + '</p:sp>'.length)
    const cNvSpPrEnd = spXml.indexOf('</p:cNvSpPr>')

    if (cNvSpPrEnd === -1) return slideXml

    const newSpXml = spXml.substring(0, cNvSpPrEnd) + hlinkXml + spXml.substring(cNvSpPrEnd)

    return slideXml.substring(0, spStart) + newSpXml + slideXml.substring(spEnd + '</p:sp>'.length)
  }

  /**
   * @private
   */
  #escapeXml(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;')
  }

  /**
   * Adds a special navigation action (next slide, prev slide, etc.) to a text element.
   *
   * @param {number} slideIndex
   * @param {string} text
   * @param {'next'|'previous'|'first'|'last'} navType
   * @param {SlideManager} slideManager
   */
  addTextNavigationLink(slideIndex, text, navType, slideManager) {
    const action = this.#getNavigationAction(navType)
    const slideXml = slideManager.getSlideXml(slideIndex)
    const updatedXml = this.#injectHyperlinkOnText(slideXml, text, '', null, `action="${action}"`)
    slideManager.setSlideXml(slideIndex, updatedXml)
    logger.debug(`Added navigation link (${navType}) to "${text}" in slide ${slideIndex}`)
  }

  /**
   * Adds a special navigation action (next slide, prev slide, etc.) to a shape.
   *
   * @param {number} slideIndex
   * @param {string} shapeName
   * @param {'next'|'previous'|'first'|'last'} navType
   * @param {SlideManager} slideManager
   */
  addShapeNavigationLink(slideIndex, shapeName, navType, slideManager) {
    const action = this.#getNavigationAction(navType)
    const slideXml = slideManager.getSlideXml(slideIndex)
    const updatedXml = this.#injectHyperlinkOnShape(slideXml, shapeName, '', `action="${action}"`)
    slideManager.setSlideXml(slideIndex, updatedXml)
    logger.debug(
      `Added navigation link (${navType}) to shape "${shapeName}" in slide ${slideIndex}`
    )
  }

  /**
   * Resolves a navigation type string to its ppaction target.
   * @private
   */
  #getNavigationAction(navType) {
    const type = String(navType).toLowerCase()
    switch (type) {
      case 'next':
        return 'ppaction://hlinkshowjump?s=nextslide'
      case 'previous':
      case 'prev':
        return 'ppaction://hlinkshowjump?s=prevslide'
      case 'first':
        return 'ppaction://hlinkshowjump?s=firstslide'
      case 'last':
        return 'ppaction://hlinkshowjump?s=lastslide'
      default:
        throw new PPTXError(`Invalid navigation type: ${navType}`)
    }
  }

  /**
   * @private
   */
  #escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
  }
}

module.exports = { HyperlinkManager }
