const { createLogger } = require('../utils/logger.js')

const logger = createLogger('TemplateEngine')

const DEFAULT_PLACEHOLDER_PATTERN = /\{\{([^{}]+)\}\}/g

class TemplateEngine {
  #xmlParser

  constructor(xmlParser) {
    this.#xmlParser = xmlParser
  }

  replaceTextInXml(slideXml, replacements) {
    if (!replacements || Object.keys(replacements).length === 0) {
      return slideXml
    }

    logger.debug(`Replacing ${Object.keys(replacements).length} placeholder(s)`)

    // Step 1: Process AST to handle fragmented text runs
    const slideObj = this.#xmlParser.parse(slideXml)
    this.#processParagraphsObj(slideObj, replacements)
    let updated = this.#xmlParser.build(slideObj)

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

      if (updated.includes(placeholderEscaped)) {
        updated = updated.replaceAll(placeholderEscaped, escaped)
      }
      if (updated.includes(placeholder)) {
        updated = updated.replaceAll(placeholder, escaped)
      }
    }

    return updated
  }

  #processParagraphsObj(obj, replacements) {
    if (!obj || typeof obj !== 'object') return

    for (const key of Object.keys(obj)) {
      const value = obj[key]
      if (key === 'a:p') {
        const paragraphs = Array.isArray(value) ? value : [value]
        const newParagraphs = []
        for (const p of paragraphs) {
          const processed = this.#processParagraphAst(p, replacements)
          if (Array.isArray(processed)) {
            newParagraphs.push(...processed)
          } else {
            newParagraphs.push(processed)
          }
        }
        obj[key] = newParagraphs
      } else if (typeof value === 'object') {
        this.#processParagraphsObj(value, replacements)
      }
    }
  }

  #processParagraphAst(p, replacements) {
    if (!p['a:r']) return p
    const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]
    if (runs.length === 0) return p

    const combinedText = runs.map(r => (r['a:t'] !== undefined ? String(r['a:t']) : '')).join('')

    let hasPlaceholder = false
    let matchedPlaceholder = null
    for (const placeholder of Object.keys(replacements)) {
      if (combinedText.includes(placeholder)) {
        hasPlaceholder = true
        matchedPlaceholder = placeholder
        break
      }
    }

    if (!hasPlaceholder) return p

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
      return this.generateListParagraphsAst(p, runs[0], listConfig)
    }

    let replacedText = combinedText
    for (const [placeholder, value] of Object.entries(replacements)) {
      if (replacedText.includes(placeholder)) {
        replacedText = replacedText.replaceAll(placeholder, String(value))
      }
    }

    const firstRun = this.#xmlParser.deepClone(runs[0])
    firstRun['a:t'] = replacedText
    p['a:r'] = [firstRun]

    return p
  }

  generateListParagraphsAst(pObj, firstRunObj, listConfig) {
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

    const baseRPr = firstRunObj['a:rPr'] ? this.#xmlParser.deepClone(firstRunObj['a:rPr']) : {}

    if (style.fontSize) {
      baseRPr['@_sz'] = Math.round(style.fontSize * 100)
    }
    if (style.bold !== undefined) {
      baseRPr['@_b'] = style.bold ? '1' : '0'
    }
    if (style.italic !== undefined) {
      baseRPr['@_i'] = style.italic ? '1' : '0'
    }
    if (style.underline) {
      baseRPr['@_u'] = typeof style.underline === 'string' ? style.underline : 'sng'
    } else if (style.underline === false) {
      delete baseRPr['@_u']
    }
    if (style.strike) {
      baseRPr['@_strike'] = style.strike === 'double' ? 'dblStrike' : 'sngStrike'
    } else if (style.strike === false) {
      delete baseRPr['@_strike']
    }
    if (style.superscript) {
      baseRPr['@_baseline'] = '30000'
    } else if (style.subscript) {
      baseRPr['@_baseline'] = '-25000'
    } else if (style.baseline !== undefined) {
      baseRPr['@_baseline'] = String(Math.round(style.baseline))
    }
    if (style.characterSpacing !== undefined || style.tracking !== undefined) {
      const spc = style.characterSpacing !== undefined ? style.characterSpacing : style.tracking
      baseRPr['@_spc'] = String(Math.round(spc * 100))
    }
    if (style.color) {
      const cleanColor = style.color.replace('#', '')
      baseRPr['a:solidFill'] = { 'a:srgbClr': { '@_val': cleanColor } }
    }
    if (style.fontFamily) {
      baseRPr['a:latin'] = { '@_typeface': style.fontFamily }
      baseRPr['a:ea'] = { '@_typeface': style.fontFamily }
      baseRPr['a:cs'] = { '@_typeface': style.fontFamily }
    }

    const newParagraphs = []

    for (const item of flatItems) {
      const lvl = item.lvl
      const marL = style.marL !== undefined ? style.marL : 381000 + lvl * 457200
      const indent = style.indent !== undefined ? style.indent : -228600

      const pPr = {
        '@_lvl': lvl,
        '@_marL': marL,
        '@_indent': indent,
      }

      if (ordered) {
        pPr['a:buAutoNum'] = { '@_type': style.numberType || 'arabicPeriod' }
      } else {
        pPr['a:buChar'] = { '@_char': style.bulletChar || '•' }
      }

      if (style.bulletColor) {
        pPr['a:buClr'] = { 'a:srgbClr': { '@_val': style.bulletColor.replace('#', '') } }
      }
      if (style.bulletSize) {
        pPr['a:buSzPct'] = { '@_val': Math.round(style.bulletSize * 1000) }
      }

      const run = {
        'a:rPr': this.#xmlParser.deepClone(baseRPr),
        'a:t': item.text,
      }

      // Preserve properties from original paragraph if necessary
      const newP = {
        'a:pPr': pPr,
        'a:r': [run],
      }

      newParagraphs.push(newP)
    }

    return newParagraphs
  }

  // To preserve backwards compatibility for tests calling generateListParagraphs with XML strings
  generateListParagraphs(paragraphXml, firstRun, listConfig) {
    const pObj = this.#xmlParser.parse(paragraphXml)
    // extract firstRunObj
    const p = pObj['a:p'] ? pObj['a:p'] : pObj
    const runs = p['a:r'] ? (Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']]) : []
    const firstRunObj = runs.length > 0 ? runs[0] : {}

    const newParagraphs = this.generateListParagraphsAst(p, firstRunObj, listConfig)

    let res = ''
    for (const newP of newParagraphs) {
      res += this.#xmlParser.build({ 'a:p': newP })
    }
    return res
  }

  containsPlaceholders(text, replacements) {
    return Object.keys(replacements).some(p => text.includes(p))
  }

  extractPlaceholders(xml, pattern = DEFAULT_PLACEHOLDER_PATTERN) {
    const placeholders = new Set()
    const textPattern = /<a:t>([\s\S]*?)<\/a:t>/g
    let match

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

  #escapeXml(str) {
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;')
  }

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
