const { REL_TYPES } = require('../RelationshipManager.js');

class ChartRelationshipManager {
  /**
   * Validates and fixes chart relationships.
   *
   * @param {RelationshipManager} relationshipManager
   * @param {ZipManager} zipManager
   * @param {string} chartZipPath
   * @returns {Object} validation issues
   */
  static validateChartRelationships(relationshipManager, zipManager, chartZipPath) {
    const issues = { errors: [], warnings: [] };
    const rels = relationshipManager.getRelationships(chartZipPath);

    let hasWorkbook = false;
    for (const rel of rels) {
      if (rel.type === REL_TYPES.PACKAGE) {
        hasWorkbook = true;
        const xlsxPath = relationshipManager.resolveTarget(chartZipPath, rel.target);
        if (!zipManager.hasFile(xlsxPath)) {
          issues.errors.push(`Embedded workbook missing: ${xlsxPath}`);
        }
      }
    }

    if (!hasWorkbook) {
      issues.warnings.push(`Chart ${chartZipPath} has no embedded workbook relationship. Live editing in PowerPoint may fail.`);
    }

    return issues;
  }
}

module.exports = { ChartRelationshipManager };
