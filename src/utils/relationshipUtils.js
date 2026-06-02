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
  let updated = xml

  // Sort by length descending to avoid partial replacements (e.g., rId1 replacing part of rId10)
  const sortedEntries = Array.from(idMap.entries()).sort(([a], [b]) => b.length - a.length)

  for (const [oldId, newId] of sortedEntries) {
    // Replace rId references in attribute values: r:id="rId1", r:embed="rId1"
    const pattern = new RegExp(`(r:[a-zA-Z]+=")${oldId}(")|rId="${oldId}(")`, 'g')
    updated = updated.replace(pattern, (match, pre, post) => {
      if (pre) return `${pre}${newId}${post}`
      return match.replace(oldId, newId)
    })

    // Simple global replace as fallback
    updated = updated.split(`"${oldId}"`).join(`"${newId}"`)
  }

  return updated
}

module.exports = {
  generateRelationshipId,
  parseRelationshipId,
  isValidRelationshipId,
  remapRelationshipIds,
}
