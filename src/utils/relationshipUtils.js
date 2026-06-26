/**
 * @fileoverview Relationship ID utilities.
 *
 * OpenXML relationship IDs follow the format rId1, rId2, rId3, ...
 * They must be unique within each .rels file.
 *
 * These utilities generate collision-free IDs when adding new relationships.
 */

/**
 * Generates the next available relationship ID given an array of existing IDs.
 * Always uses the format "rId{N}" where N is the next integer after the max.
 *
 * @param {string[]} existingIds - Array of existing rId strings (e.g., ['rId1', 'rId2']).
 * @returns {string} New relationship ID (e.g., 'rId3').
 *
 * @example
 * generateRelationshipId(['rId1', 'rId3'])  // → 'rId4'
 * generateRelationshipId([])                 // → 'rId1'
 */
function generateRelationshipId(existingIds) {
  if (!existingIds || existingIds.length === 0) return 'rId1'

  const maxNum = existingIds.reduce((max, id) => {
    const match = /^rId(\d+)$/.exec(id)
    if (!match) return max
    return Math.max(max, parseInt(match[1], 10))
  }, 0)

  return `rId${maxNum + 1}`
}

/**
 * Parses a relationship ID string and returns its numeric value.
 *
 * @param {string} rId - Relationship ID (e.g., 'rId5').
 * @returns {number} Numeric value (e.g., 5), or -1 if not a valid rId.
 *
 * @example
 * parseRelationshipId('rId5')  // → 5
 * parseRelationshipId('foo')   // → -1
 */
function parseRelationshipId(rId) {
  const match = /^rId(\d+)$/.exec(rId)
  return match ? parseInt(match[1], 10) : -1
}

/**
 * Checks if a string is a valid relationship ID.
 *
 * @param {string} str
 * @returns {boolean}
 */
function isValidRelationshipId(str) {
  return /^rId\d+$/.test(str)
}

/**
 * Remaps old relationship IDs to new ones within an XML string.
 * Used when cloning slides to avoid rId conflicts.
 *
 * @param {string} xml - XML content containing rId references.
 * @param {Map<string, string>} idMap - Map of old rId → new rId.
 * @returns {string} Updated XML with remapped rIds.
 *
 * @example
 * remapRelationshipIds(xml, new Map([['rId1', 'rId5'], ['rId2', 'rId6']]));
 */
function remapRelationshipIds(xml, idMap) {
  if (!idMap || idMap.size === 0) return xml

  // Match any relationship ID attribute reference prefixed with r: (e.g. r:id, r:embed, r:link, r:dm, r:lo, r:qs, r:cs)
  return xml.replace(
    /\b(r:[a-zA-Z0-9_]+)=(["'])(rId\d+)\2/g,
    (match, attr, quote, id) => {
      if (idMap.has(id)) {
        return `${attr}=${quote}${idMap.get(id)}${quote}`
      }
      return match
    }
  )
}

module.exports = {
  generateRelationshipId,
  parseRelationshipId,
  isValidRelationshipId,
  remapRelationshipIds,
}
