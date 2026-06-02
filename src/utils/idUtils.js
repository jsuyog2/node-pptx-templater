/**
 * @fileoverview Unique ID generation utilities.
 * Used for generating shape IDs, slide IDs, etc. in OpenXML.
 */

const { randomBytes } = require('crypto')

/**
 * Generates a unique integer ID for use as a shape or slide ID.
 * OpenXML shape IDs are positive integers unique within a slide.
 *
 * @param {number[]} [existingIds] - Array of existing IDs to avoid.
 * @returns {number} Unique positive integer.
 */
function generateUniqueId(existingIds = []) {
  const maxId = existingIds.length > 0 ? Math.max(...existingIds) : 0
  return maxId + 1
}

/**
 * Generates a UUID v4 string.
 * Used for chart GUID references and extension URIs in OpenXML.
 *
 * @returns {string} UUID v4 string (e.g., '{A1B2C3D4-E5F6-...}')
 */
function generateGuid() {
  const bytes = randomBytes(16)
  bytes[6] = (bytes[6] & 0x0f) | 0x40 // Version 4
  bytes[8] = (bytes[8] & 0x3f) | 0x80 // Variant RFC4122

  const hex = bytes.toString('hex')
  const guid = [
    hex.slice(0, 8),
    hex.slice(8, 12),
    hex.slice(12, 16),
    hex.slice(16, 20),
    hex.slice(20, 32),
  ]
    .join('-')
    .toUpperCase()

  return `{${guid}}`
}

/**
 * Generates a slide ID (numeric, starts at 256 by convention in OpenXML).
 * PowerPoint seems to start slide IDs at 256.
 *
 * @param {string[]} [existingSlideIds] - Existing slide ID strings.
 * @returns {string} New slide ID string.
 */
function generateSlideId(existingSlideIds = []) {
  const existingNums = existingSlideIds.map(id => parseInt(id, 10)).filter(n => !isNaN(n))
  const maxId = existingNums.length > 0 ? Math.max(...existingNums) : 255
  return String(maxId + 1)
}

module.exports = {
  generateUniqueId,
  generateGuid,
  generateSlideId,
}
