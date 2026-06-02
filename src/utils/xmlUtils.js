/**
 * @fileoverview XML validation, repair, recovery, and security diagnostics utilities.
 *
 * Provides tools to check if generated XML is well-formed, protect against
 * XML Entity Attacks (XXE, DTD abuse, Billion Laughs), and recover diagnostics.
 */

const { XMLValidator } = require('fast-xml-parser')
const { XMLParser } = require('../parsers/XMLParser.js')
const { PPTXError } = require('./errors.js')

const parser = new XMLParser()

/**
 * Helper to compute line and column numbers from a string index.
 */
function getLineAndCol(str, index) {
  let line = 1
  let col = 1
  for (let i = 0; i < index; i++) {
    if (str[i] === '\n') {
      line++
      col = 1
    } else if (str[i] !== '\r') {
      col++
    }
  }
  return { line, col }
}

/**
 * Validates that an XML string is secure and well-formed.
 * Checks for DTDs, recursive/custom entities, external references (XXE), and malformed tags.
 *
 * @param {string} xmlString - Raw XML content.
 * @returns {{ valid: boolean, error: string|null, line: number|null, column: number|null, recommendation: string|null }}
 */
function validateXml(xmlString) {
  if (typeof xmlString !== 'string') {
    return {
      valid: false,
      error: 'Invalid XML input: expected string.',
      line: 1,
      column: 1,
      recommendation: 'Ensure XML input is passed as a string.',
    }
  }

  // 1. Check for external references (XXE)
  if (/SYSTEM\b/i.test(xmlString) || /PUBLIC\b/i.test(xmlString)) {
    const match = xmlString.match(/(SYSTEM|PUBLIC)\b/i)
    const index = match ? match.index : 0
    const { line, col } = getLineAndCol(xmlString, index)
    return {
      valid: false,
      error: 'External reference SYSTEM/PUBLIC detected',
      line,
      column: col,
      recommendation: 'Remove external system/public identifiers to prevent XXE attacks.',
    }
  }

  // 2. Check for entity declarations (prevent custom/recursive entities)
  if (/<!ENTITY/i.test(xmlString)) {
    const index = xmlString.search(/<!ENTITY/i)
    const { line, col } = getLineAndCol(xmlString, index)
    return {
      valid: false,
      error: 'Custom entity declaration detected',
      line,
      column: col,
      recommendation: 'Do not declare custom entities to protect against XML entity injection.',
    }
  }

  // 3. Check for DTD / DOCTYPE declarations (DTD abuse, recursive entities, Billion Laughs)
  if (/<!DOCTYPE/i.test(xmlString)) {
    const index = xmlString.search(/<!DOCTYPE/i)
    const { line, col } = getLineAndCol(xmlString, index)
    return {
      valid: false,
      error: 'DTD/DOCTYPE declaration detected: entity expansion limit exceeded / DTD abuse',
      line,
      column: col,
      recommendation:
        'Remove DOCTYPE declarations or DTD abuse to prevent entity expansion attacks.',
    }
  }

  // 4. Check for oversized entity references to prevent DoS (exceeding 50,000 entity references)
  const entityCount = (xmlString.match(/&[a-zA-Z0-9#x]+;/g) || []).length
  if (entityCount > 50000) {
    return {
      valid: false,
      error: `Entity expansion limit exceeded: ${entityCount} references (max 50000)`,
      line: 1,
      column: 1,
      recommendation: 'Reduce the density of standard entity references.',
    }
  }

  // 5. Well-formedness check using XMLValidator
  const validation = XMLValidator.validate(xmlString)
  if (validation !== true) {
    return {
      valid: false,
      error: validation.err.msg || 'Malformed XML',
      line: validation.err.line || 1,
      column: validation.err.col || 1,
      recommendation:
        'Fix XML syntax errors (unclosed tags, invalid characters, mismatched brackets).',
    }
  }

  return {
    valid: true,
    error: null,
    line: null,
    column: null,
    recommendation: null,
  }
}

/**
 * Validates that an XML string is well-formed.
 * Backwards compatibility wrapper for original validateXML.
 *
 * @param {string} xmlString - XML to validate.
 * @returns {{ valid: boolean, error: string|null }} Validation result.
 */
function validateXML(xmlString) {
  const result = validateXml(xmlString)
  return {
    valid: result.valid,
    error: result.error,
  }
}

/**
 * Safely parses XML with validation, recovery diagnostics, and fallback reporting.
 *
 * @param {string} xmlString - Raw XML content.
 * @param {string} filename - Filename for error reporting context.
 * @param {XMLParser} [xmlParserInstance] - Optional parser instance.
 * @returns {Object} Parsed JS object.
 * @throws {PPTXError} If parsing fails or security limits are violated.
 */
function safeParseXml(xmlString, filename = 'unknown.xml', xmlParserInstance = null) {
  const validation = validateXml(xmlString)
  if (!validation.valid) {
    const errorDetails = {
      file: filename,
      line: validation.line || 1,
      column: validation.column || 1,
      error: validation.error,
      recommendation: validation.recommendation || 'Malformed entity reference detected',
    }
    const err = new PPTXError(`XML parse validation error in ${filename}: ${validation.error}`)
    err.diagnostic = errorDetails
    throw err
  }

  try {
    const p = xmlParserInstance || parser
    return p.parse(xmlString, filename)
  } catch (err) {
    let line = 1
    let col = 1
    const lineMatch = err.message.match(/line:?\s*(\d+)/i) || err.message.match(/:(\d+):\d+$/)
    const colMatch = err.message.match(/col(umn)?:?\s*(\d+)/i) || err.message.match(/:\d+:(\d+)$/)
    if (lineMatch) line = parseInt(lineMatch[1], 10)
    if (colMatch) col = parseInt(colMatch[2] || colMatch[1], 10)

    const errorDetails = {
      file: filename,
      line,
      column: col,
      error: err.message,
      recommendation: 'Ensure all XML tags are closed properly and entity syntax is valid.',
    }
    const newErr = new PPTXError(`XML parse error in ${filename}: ${err.message}`)
    newErr.diagnostic = errorDetails
    throw newErr
  }
}

/**
 * Attempts to repair common XML corruption issues in PPTX files.
 *
 * @param {string} xmlString - Potentially broken XML.
 * @returns {{ xml: string, repaired: boolean, changes: string[] }}
 */
function repairXML(xmlString) {
  const changes = []
  let xml = xmlString

  // Fix 1: Remove invalid XML control characters
  const before = xml
  xml = xml.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
  if (xml !== before) changes.push('Removed invalid control characters')

  // Fix 2: Fix unescaped ampersands in text content (not in entities)
  const fixedAmp = xml.replace(/&(?!amp;|lt;|gt;|quot;|apos;|#\d+;|#x[0-9a-fA-F]+;)/g, '&amp;')
  if (fixedAmp !== xml) {
    xml = fixedAmp
    changes.push('Escaped unescaped ampersands')
  }

  // Fix 3: Replace null bytes
  if (xml.includes('\x00')) {
    xml = xml.replace(/\x00/g, '')
    changes.push('Removed null bytes')
  }

  // Fix 4: Ensure XML declaration is present
  if (!xml.trimStart().startsWith('<?xml')) {
    xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml
    changes.push('Added missing XML declaration')
  }

  return {
    xml,
    repaired: changes.length > 0,
    changes,
  }
}

/**
 * Checks if an XML string contains a specific element.
 */
function xmlContainsElement(xmlString, elementName) {
  return xmlString.includes(`<${elementName}`) || xmlString.includes(`<${elementName}>`)
}

/**
 * Counts occurrences of an element in XML.
 */
function countElements(xmlString, elementName) {
  const pattern = new RegExp(`<${elementName}[\\s>/]`, 'g')
  return (xmlString.match(pattern) || []).length
}

/**
 * Extracts all attribute values for a given attribute name.
 */
function extractAttributeValues(xmlString, attrName) {
  const pattern = new RegExp(`${attrName.replace(':', '\\:')}="([^"]*)"`, 'g')
  const values = []
  let match
  while ((match = pattern.exec(xmlString)) !== null) {
    values.push(match[1])
  }
  return values
}

/**
 * Scans an XML string for entity references.
 *
 * @param {string} xmlString - XML string to scan.
 * @returns {{ standard: number, custom: number, numeric: number, hex: number, total: number, entities: string[] }}
 */
function scanForEntities(xmlString) {
  const result = {
    standard: 0,
    custom: 0,
    numeric: 0,
    hex: 0,
    total: 0,
    entities: [],
  }
  if (typeof xmlString !== 'string') return result

  const entityRegex = /&[a-zA-Z0-9#x_:-]+;/g
  const matches = xmlString.match(entityRegex) || []
  result.total = matches.length

  const standardSet = new Set(['&amp;', '&lt;', '&gt;', '&quot;', '&apos;'])

  matches.forEach(match => {
    result.entities.push(match)
    if (standardSet.has(match)) {
      result.standard++
    } else if (match.startsWith('&#x')) {
      result.hex++
    } else if (match.startsWith('&#')) {
      result.numeric++
    } else {
      result.custom++
    }
  })

  return result
}

/**
 * Analyzes XML properties.
 *
 * @param {string} xmlString - XML content.
 * @returns {{ sizeBytes: number, lineCount: number, elementCount: number, attributeCount: number, entityStats: Object }}
 */
function analyzeXmlFile(xmlString) {
  if (typeof xmlString !== 'string') {
    return { sizeBytes: 0, lineCount: 0, elementCount: 0, attributeCount: 0, entityStats: {} }
  }

  const sizeBytes = Buffer.byteLength(xmlString, 'utf8')
  const lineCount = xmlString.split('\n').length
  const elementCount = (xmlString.match(/<[a-zA-Z0-9_:-]+/g) || []).length
  const attributeCount = (xmlString.match(/\s[a-zA-Z0-9_:-]+=/g) || []).length
  const entityStats = scanForEntities(xmlString)

  return {
    sizeBytes,
    lineCount,
    elementCount,
    attributeCount,
    entityStats,
  }
}

/**
 * Reports complexity indicators of the XML document.
 *
 * @param {string} xmlString - XML content.
 * @returns {{ maxDepth: number, nodeCount: number, ratioTextToMarkup: number }}
 */
function reportXmlComplexity(xmlString) {
  if (typeof xmlString !== 'string') {
    return { maxDepth: 0, nodeCount: 0, ratioTextToMarkup: 0 }
  }

  let currentDepth = 0
  let maxDepth = 0
  let nodeCount = 0

  const tagRegex = /<\/?([a-zA-Z0-9_:-]+)(\s[^>]*)*>/g
  let match
  while ((match = tagRegex.exec(xmlString)) !== null) {
    const rawTag = match[0]
    nodeCount++
    if (rawTag.startsWith('</')) {
      currentDepth--
    } else if (rawTag.endsWith('/>')) {
      if (currentDepth + 1 > maxDepth) {
        maxDepth = currentDepth + 1
      }
    } else {
      currentDepth++
      if (currentDepth > maxDepth) {
        maxDepth = currentDepth
      }
    }
  }

  const textOnly = xmlString.replace(/<[^>]+>/g, '')
  const textLength = textOnly.length
  const xmlLength = xmlString.length
  const ratioTextToMarkup = xmlLength > 0 ? parseFloat((textLength / xmlLength).toFixed(4)) : 0

  return {
    maxDepth,
    nodeCount,
    ratioTextToMarkup,
  }
}

module.exports = {
  validateXml,
  validateXML,
  safeParseXml,
  repairXML,
  xmlContainsElement,
  countElements,
  extractAttributeValues,
  scanForEntities,
  analyzeXmlFile,
  reportXmlComplexity,
}
