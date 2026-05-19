/**
 * node-pptx-templater
 *
 * A low-level PowerPoint OpenXML templating engine for Node.js.
 * Generates and edits PPTX files directly through XML manipulation
 * without relying on PowerPoint generation libraries.
 *
 * @module node-pptx-templater
 * @author node-pptx-templater contributors
 * @license MIT
 *
 * Architecture Overview:
 * ┌─────────────────────────────────────────────────────────────┐
 * │                   PPTXTemplater                        │
 * │  ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌──────────┐  │
 * │  │ZipManager│  │XMLParser │  │SlideManager│ │ChartMgr  │  │
 * │  └──────────┘  └──────────┘  └──────────┘  └──────────┘  │
 * │  ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌──────────┐  │
 * │  │TableMgr  │  │HyperlinkMgr│ │MediaMgr  │  │RelMgr    │  │
 * │  └──────────┘  └──────────┘  └──────────┘  └──────────┘  │
 * │                   ┌──────────────┐                         │
 * │                   │OutputWriter  │                         │
 * │                   └──────────────┘                         │
 * └─────────────────────────────────────────────────────────────┘
 */

const { PPTXTemplater } = require('./core/PPTXTemplater.js');
const { ZipManager } = require('./managers/ZipManager.js');
const { XMLParser } = require('./parsers/XMLParser.js');
const { SlideManager } = require('./managers/SlideManager.js');
const { ChartManager } = require('./managers/ChartManager.js');
const { TableManager } = require('./managers/TableManager.js');
const { HyperlinkManager } = require('./managers/HyperlinkManager.js');
const { MediaManager } = require('./managers/MediaManager.js');
const { RelationshipManager } = require('./managers/RelationshipManager.js');
const { OutputWriter } = require('./core/OutputWriter.js');
const { TemplateEngine } = require('./core/TemplateEngine.js');

// Utility exports
const { generateRelationshipId, parseRelationshipId } = require('./utils/relationshipUtils.js');
const { validateXML, repairXML } = require('./utils/xmlUtils.js');
const { createLogger } = require('./utils/logger.js');
const { PPTXError, SlideNotFoundError, ChartNotFoundError, TableNotFoundError } = require('./utils/errors.js');

module.exports = {
  PPTXTemplater,
  ZipManager,
  XMLParser,
  SlideManager,
  ChartManager,
  TableManager,
  HyperlinkManager,
  MediaManager,
  RelationshipManager,
  OutputWriter,
  TemplateEngine,
  generateRelationshipId,
  parseRelationshipId,
  validateXML,
  repairXML,
  createLogger,
  PPTXError,
  SlideNotFoundError,
  ChartNotFoundError,
  TableNotFoundError
};
