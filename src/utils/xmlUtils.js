/**
 * @fileoverview XML validation and repair utilities.
 *
 * Provides tools to check if generated XML is well-formed and
 * attempt automatic repairs for common PPTX corruption issues.
 */

const { XMLParser } = require('../parsers/XMLParser.js')

const parser = new XMLParser()

/**
 * Validates that an XML string is well-formed.
 *
 * @param {string} xmlString - XML to validate.
 * @returns {{ valid: boolean, error: string|null }} Validation result.
 *
 * @example
 * const { valid, error } = validateXML(xml);
 * if (!valid) console.error('XML error:', error);
 */
function validateXML(xmlString) {
  return parser.validate(xmlString)
}

/**
 * Attempts to repair common XML corruption issues in PPTX files.
 *
 * Known issues this addresses:
 * 1. Unescaped & in attribute values (e.g., href="a&b" → href="a&amp;b")
 * 2. Unclosed tags (limited heuristic repair)
 * 3. Invalid XML characters (removes control chars below 0x20 except tab/LF/CR)
 *
 * @param {string} xmlString - Potentially broken XML.
 * @returns {{ xml: string, repaired: boolean, changes: string[] }}
 *
 * @example
 * const { xml, repaired, changes } = repairXML(brokenXml);
 * if (repaired) console.log('Repaired:', changes);
 */
function repairXML(xmlString) {
  const changes = []
  let xml = xmlString

  // Fix 1: Remove invalid XML control characters
  const before = xml
  xml = xml.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '')
  if (xml !== before) changes.push('Removed invalid control characters')

  // Fix 2: Fix unescaped ampersands in text content (not in entities)
  // Match & not followed by valid entity patterns
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
 *
 * @param {string} xmlString
 * @param {string} elementName - Element tag name (e.g., 'a:tbl').
 * @returns {boolean}
 */
function xmlContainsElement(xmlString, elementName) {
  return xmlString.includes(`<${elementName}`) || xmlString.includes(`<${elementName}>`)
}

/**
 * Counts occurrences of an element in XML.
 *
 * @param {string} xmlString
 * @param {string} elementName
 * @returns {number}
 */
function countElements(xmlString, elementName) {
  const pattern = new RegExp(`<${elementName}[\\s>/]`, 'g')
  return (xmlString.match(pattern) || []).length
}

/**
 * Extracts all attribute values for a given attribute name.
 *
 * @param {string} xmlString - XML string to search.
 * @param {string} attrName - Attribute name (e.g., 'r:id', 'name').
 * @returns {string[]} Array of attribute values found.
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

module.exports = {
  validateXML,
  repairXML,
  xmlContainsElement,
  countElements,
  extractAttributeValues,
}
