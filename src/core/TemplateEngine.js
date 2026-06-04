/**
 * @fileoverview TemplateEngine - Text placeholder replacement engine.
 *
 * Handles {{placeholder}} replacement in slide XML.
 *
 * Challenge with OpenXML text replacement:
 * ─────────────────────────────────────────────────────────────────
 * A simple string "{{title}}" in PowerPoint might be split across
 * MULTIPLE text run elements in the XML:
 *
 *   <a:r><a:t>{{ti</a:t></a:r>
 *   <a:r><a:t>tle}}</a:t></a:r>
 *
 * This happens because PowerPoint splits text runs when:
 *  - Spell-check marks parts of the text
 *  - Font formatting changes mid-word
 *  - The text was typed in multiple sessions
 *
 * Solution Strategy:
 * 1. Extract all <a:r><a:t>...</a:t></a:r> run sequences within each <a:p>
 * 2. Concatenate run text to find the full combined text
 * 3. Check if the combined text contains any placeholder
 * 4. If yes: merge all runs into a single run with replaced text,
 *    preserving the formatting of the FIRST run as the template
 *
 * This "text normalization" approach correctly handles fragmented placeholders.
 */

const { createLogger } = require('../utils/logger.js')

const logger = createLogger('TemplateEngine')

/**
 * Default placeholder pattern: {{key}}
 * Can be overridden per-call.
 */
const DEFAULT_PLACEHOLDER_PATTERN = /\{\{([^{}]+)\}\}/g

/**
 * @class TemplateEngine
 * @description Handles text placeholder replacement in OpenXML slide XML.
 *
 * Implements the text-normalization strategy to handle fragmented placeholders.
 */
class TemplateEngine {
  /** @private @type {XMLParser} */
  #xmlParser

  /**
   * @param {XMLParser} xmlParser
   */
  constructor(xmlParser) {
    this.#xmlParser = xmlParser
  }

  /**
   * Replaces all placeholders in a slide XML string.
   * Uses the text-normalization strategy to handle fragmented text runs.
   *
   * @param {string} slideXml - Raw slide XML.
   * @param {Object.<string, string>} replacements - Placeholder → value map.
   * @param {RegExp} [pattern] - Custom placeholder pattern. Defaults to {{key}}.
   * @returns {string} Modified slide XML with placeholders replaced.
   *
   * @example
   * const updated = engine.replaceTextInXml(slideXml, {
   *   '{{name}}': 'John Doe',
   *   '{{date}}': '2026-01-01'
   * });
   */
  replaceTextInXml(slideXml, replacements) {
    if (!replacements || Object.keys(replacements).length === 0) {
      return slideXml
    }

    logger.debug(`Replacing ${Object.keys(replacements).length} placeholder(s)`)

    // Step 1: Process paragraph by paragraph to handle fragmented runs
    let updated = this.#processParagraphs(slideXml, replacements)

    // Step 2: Simple direct replacement for any remaining unfragmented placeholders
    for (const [placeholder, value] of Object.entries(replacements)) {
      if (
        value &&
        typeof value === 'object' &&
        (Array.isArray(value) || value.list !== undefined)
      ) {
        continue
      }
      const escaped = this.#escapeXml(String(value))
      const placeholderEscaped = this.#escapeXml(placeholder)

      // Replace the XML-escaped form (e.g., {{name}} as {{name}})
      updated = updated.split(placeholderEscaped).join(escaped)
      // Replace the plain form (in case it's not escaped in the XML)
      updated = updated.split(placeholder).join(escaped)
    }

    return updated
  }

  /**
   * Processes all paragraphs (<a:p>) in the slide XML, normalizing text runs
   * within each paragraph to fix fragmented placeholders.
   *
   * @private
   * @param {string} slideXml
   * @param {Object.<string, string>} replacements
   * @returns {string}
   */
  #processParagraphs(slideXml, replacements) {
    // Find all <a:p>...</a:p> paragraphs
    let updated = slideXml
    let offset = 0

    const paragraphPattern = /<a:p>([\s\S]*?)<\/a:p>/g
    let match

    while ((match = paragraphPattern.exec(slideXml)) !== null) {
      const paragraphXml = match[0]
      const processedParagraph = this.#processParagraph(paragraphXml, replacements)

      if (processedParagraph !== paragraphXml) {
        // More reliable: replace from the beginning of the current search area
        updated =
          updated.substring(0, match.index + offset) +
          processedParagraph +
          updated.substring(match.index + offset + paragraphXml.length)

        offset += processedParagraph.length - paragraphXml.length
      }
    }

    return updated
  }

  /**
   * Processes a single paragraph, normalizing runs and replacing placeholders.
   *
   * @private
   * @param {string} paragraphXml - XML of a single <a:p> element.
   * @param {Object.<string, string>} replacements
   * @returns {string} Updated paragraph XML.
   */
  #processParagraph(paragraphXml, replacements) {
    // Extract all text runs from this paragraph
    const runs = this.#extractRuns(paragraphXml)

    if (runs.length === 0) return paragraphXml

    // Combine text from all runs
    const combinedText = runs.map(r => r.text).join('')

    // Check if any placeholder appears in the combined text
    let hasPlaceholder = false
    let matchedPlaceholder = null
    for (const placeholder of Object.keys(replacements)) {
      if (combinedText.includes(placeholder)) {
        hasPlaceholder = true
        matchedPlaceholder = placeholder
        break
      }
    }

    if (!hasPlaceholder) return paragraphXml

    const replacementValue = replacements[matchedPlaceholder]
    const isList =
      replacementValue &&
      (Array.isArray(replacementValue) ||
        (typeof replacementValue === 'object' && replacementValue.list !== undefined))

    if (isList) {
      const listConfig = Array.isArray(replacementValue)
        ? { list: replacementValue }
        : replacementValue
      const { ValidationEngine } = require('./ValidationEngine.js')
      const validation = ValidationEngine.validateList(listConfig)
      if (!validation.valid) {
        throw new Error(`List validation failed: ${validation.errors.join(', ')}`)
      }
      return this.generateListParagraphs(paragraphXml, runs[0], listConfig)
    }

    // Perform replacement on combined text
    let replacedText = combinedText
    for (const [placeholder, value] of Object.entries(replacements)) {
      replacedText = replacedText.split(placeholder).join(String(value))
    }

    // Rebuild the paragraph: merge all runs into a single run using first run's format
    return this.#mergeRunsWithText(paragraphXml, runs, replacedText)
  }

  /**
   * Extracts all text runs from a paragraph's XML.
   * Returns run XML and the extracted text content.
   *
   * @private
   * @param {string} paragraphXml
   * @returns {Array<{xml: string, text: string, start: number, end: number}>}
   */
  #extractRuns(paragraphXml) {
    const runs = []
    const runPattern = /(<a:r(?:\s[^>]*)?>)([\s\S]*?)(<\/a:r>)/g
    let match

    while ((match = runPattern.exec(paragraphXml)) !== null) {
      const runXml = match[0]
      // Extract text from <a:t>...</a:t> within this run
      const tMatch = /<a:t>([\s\S]*?)<\/a:t>/.exec(runXml)
      const text = tMatch ? this.#unescapeXml(tMatch[1]) : ''

      runs.push({
        xml: runXml,
        text,
        start: match.index,
        end: match.index + runXml.length,
      })
    }

    return runs
  }

  /**
   * Merges all runs in a paragraph into a single run using the first run's
   * formatting as the template, with the replaced text as content.
   *
   * @private
   * @param {string} paragraphXml
   * @param {Array} runs - Extracted run info.
   * @param {string} newText - New text to inject.
   * @returns {string} Updated paragraph XML.
   */
  #mergeRunsWithText(paragraphXml, runs, newText) {
    if (runs.length === 0) return paragraphXml

    // Use the first run as the format template
    const firstRun = runs[0]

    // Build the replacement run: first run's format + new text
    const mergedRunXml = this.#setRunText(firstRun.xml, this.#escapeXml(newText))

    // Build new paragraph:
    // Keep everything before first run, insert merged run, remove the rest,
    // keep everything after last run
    const firstRunStart = firstRun.start
    const lastRunEnd = runs[runs.length - 1].end

    return (
      paragraphXml.substring(0, firstRunStart) + mergedRunXml + paragraphXml.substring(lastRunEnd)
    )
  }

  /**
   * Replaces the text content of a text run XML.
   *
   * @private
   * @param {string} runXml - Run XML string.
   * @param {string} text - New text content (already XML-escaped).
   * @returns {string} Updated run XML.
   */
  #setRunText(runXml, text) {
    const tPattern = /(<a:t>)([\s\S]*?)(<\/a:t>)/
    if (tPattern.test(runXml)) {
      return runXml.replace(tPattern, `$1${text}$3`)
    }
    // If no <a:t>, add one before </a:r>
    return runXml.replace('</a:r>', `<a:t>${text}</a:t></a:r>`)
  }

  /**
   * Checks if a string contains any placeholder from the map.
   *
   * @param {string} text
   * @param {Object.<string, string>} replacements
   * @returns {boolean}
   */
  containsPlaceholders(text, replacements) {
    return Object.keys(replacements).some(p => text.includes(p))
  }

  /**
   * Extracts all unique placeholder keys from an XML string.
   *
   * @param {string} xml - Slide XML.
   * @param {RegExp} [pattern] - Placeholder pattern.
   * @returns {string[]} Array of placeholder keys found.
   *
   * @example
   * engine.extractPlaceholders(slideXml);
   * // → ['{{title}}', '{{date}}', '{{company}}']
   */
  extractPlaceholders(xml, pattern = DEFAULT_PLACEHOLDER_PATTERN) {
    const placeholders = new Set()
    const textPattern = /<a:t>([\s\S]*?)<\/a:t>/g
    let match

    // Extract text content first, then find placeholders
    const allText = []
    while ((match = textPattern.exec(xml)) !== null) {
      allText.push(match[1])
    }

    const combined = allText.join('')
    const plPattern = new RegExp(pattern.source, 'g')
    let plMatch
    while ((plMatch = plPattern.exec(combined)) !== null) {
      placeholders.add(plMatch[0])
    }

    return Array.from(placeholders)
  }

  /**
   * Generates a block of list paragraph XML elements from a template paragraph,
   * a baseline run, and list options.
   *
   * @param {string} paragraphXml
   * @param {Object} firstRun - Run XML info.
   * @param {Object} listConfig - List styling and items.
   * @returns {string} XML string of multiple paragraphs.
   */
  generateListParagraphs(paragraphXml, firstRun, listConfig) {
    const list = listConfig.list || []
    const ordered = !!listConfig.ordered
    const style = listConfig.style || {}

    const flattenList = (items, currentLvl = 0) => {
      let flat = []
      for (const item of items) {
        if (typeof item === 'string') {
          flat.push({ text: item, lvl: currentLvl })
        } else if (typeof item === 'object' && item !== null) {
          const text = item.text || ''
          flat.push({ text, lvl: currentLvl })
          if (Array.isArray(item.children)) {
            flat = flat.concat(flattenList(item.children, currentLvl + 1))
          }
        }
      }
      return flat
    }

    const flatItems = flattenList(list)

    const rPrMatch = /(<a:rPr>[\s\S]*?<\/a:rPr>)/.exec(firstRun.xml)
    let baseRPr = rPrMatch ? rPrMatch[1] : '<a:rPr/>'

    if (style.fontSize) {
      const szVal = Math.round(style.fontSize * 100)
      if (/sz="\d+"/.test(baseRPr)) {
        baseRPr = baseRPr.replace(/sz="\d+"/, `sz="${szVal}"`)
      } else {
        baseRPr = baseRPr.replace('<a:rPr', `<a:rPr sz="${szVal}"`)
      }
    }
    if (style.color) {
      const cleanColor = style.color.replace('#', '')
      const newFill = `<a:solidFill><a:srgbClr val="${cleanColor}"/></a:solidFill>`
      if (/<a:solidFill>[\s\S]*?<\/a:solidFill>/.test(baseRPr)) {
        baseRPr = baseRPr.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, newFill)
      } else {
        if (baseRPr.endsWith('/>')) {
          baseRPr = baseRPr.replace('/>', `>${newFill}</a:rPr>`)
        } else {
          baseRPr = baseRPr.replace('</a:rPr>', `${newFill}</a:rPr>`)
        }
      }
    }
    if (style.fontFamily) {
      const latinXml = `<a:latin typeface="${style.fontFamily}"/><a:cs typeface="${style.fontFamily}"/>`
      baseRPr = baseRPr.replace(/<a:latin\s+[^>]*\/>/g, '').replace(/<a:cs\s+[^>]*\/>/g, '')
      if (baseRPr.endsWith('/>')) {
        baseRPr = baseRPr.replace('/>', `>${latinXml}</a:rPr>`)
      } else {
        baseRPr = baseRPr.replace('</a:rPr>', `${latinXml}</a:rPr>`)
      }
    }

    let paragraphsXml = ''
    for (const item of flatItems) {
      const lvl = item.lvl
      const marL = style.marL !== undefined ? style.marL : 381000 + lvl * 457200
      const indent = style.indent !== undefined ? style.indent : -228600

      let pPr = `<a:pPr lvl="${lvl}" marL="${marL}" indent="${indent}">`
      if (ordered) {
        const numType = style.numberType || 'arabicPeriod'
        pPr += `<a:buAutoNum type="${numType}"/>`
      } else {
        const bulletChar = style.bulletChar || '•'
        pPr += `<a:buChar char="${bulletChar}"/>`
      }

      if (style.bulletColor) {
        const cleanBClr = style.bulletColor.replace('#', '')
        pPr += `<a:buClr><a:srgbClr val="${cleanBClr}"/></a:buClr>`
      }

      if (style.bulletSize) {
        pPr += `<a:buSzPct val="${Math.round(style.bulletSize * 1000)}"/>`
      }

      pPr += `</a:pPr>`

      const runXml = `<a:r>${baseRPr}<a:t>${this.#escapeXml(item.text)}</a:t></a:r>`
      paragraphsXml += `<a:p>${pPr}${runXml}</a:p>`
    }

    return paragraphsXml
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
      .replace(/'/g, '&apos;')
  }

  /**
   * Unescapes XML entities.
   * @private
   * @param {string} str
   * @returns {string}
   */
  #unescapeXml(str) {
    return str
      .replace(/&amp;/g, '&')
      .replace(/&lt;/g, '<')
      .replace(/&gt;/g, '>')
      .replace(/&quot;/g, '"')
      .replace(/&apos;/g, "'")
  }
}

module.exports = { TemplateEngine }
