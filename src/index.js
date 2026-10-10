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
 * │  ┌──────────┐  ┌──────────┐  ┌──────────┐  ┌──────────┐  │
 * │  │ShapeMgr  │  │ImageMgr  │  │TextMgr   │  │ZOrderMgr │  │
 * │  └──────────┘  └──────────┘  └──────────┘  └──────────┘  │
 * │                   ┌──────────────┐                         │
 * │                   │OutputWriter  │                         │
 * │                   └──────────────┘                         │
 * └─────────────────────────────────────────────────────────────┘
 */

const { PPTXTemplater } = require('./core/PPTXTemplater.js')
const { ZipManager } = require('./managers/ZipManager.js')
const { XMLParser, Z_ORDER_SYMBOL } = require('./parsers/XMLParser.js')
const { SlideManager } = require('./managers/SlideManager.js')
const { ChartManager } = require('./elements/ChartManager.js')
const { TableManager } = require('./elements/TableManager.js')
const { ShapeManager } = require('./elements/ShapeManager.js')
const { ImageManager } = require('./elements/ImageManager.js')
const { TextManager } = require('./elements/TextManager.js')
const { HyperlinkManager } = require('./managers/HyperlinkManager.js')
const { MediaManager } = require('./managers/MediaManager.js')
const { RelationshipManager } = require('./managers/RelationshipManager.js')
const { ZOrderManager } = require('./spatial/ZOrderManager.js')
const { OutputWriter } = require('./core/OutputWriter.js')
const { TemplateEngine } = require('./core/TemplateEngine.js')
const { ValidationEngine } = require('./core/ValidationEngine.js')
const { PackageLoader, isPkZip } = require('./core/PackageLoader.js')
const { VirtualFileSystem } = require('./core/VirtualFileSystem.js')
const spatialUnits = require('./spatial/units.js')

// Utility exports
const { generateRelationshipId, parseRelationshipId } = require('./utils/relationshipUtils.js')
const {
  validateXml,
  validateXML,
  safeParseXml,
  repairXML,
  scanForEntities,
  analyzeXmlFile,
  reportXmlComplexity,
} = require('./utils/xmlUtils.js')
const { createLogger, setGlobalLogLevel, resetLogLevel } = require('./utils/logger.js')
const {
  PPTXError,
  SlideNotFoundError,
  ChartNotFoundError,
  TableNotFoundError,
} = require('./utils/errors.js')

module.exports = {
  PPTXTemplater,
  PPTXTemplate: PPTXTemplater,
  ZipManager,
  XMLParser,
  Z_ORDER_SYMBOL,
  SlideManager,
  ChartManager,
  TableManager,
  ShapeManager,
  ImageManager,
  TextManager,
  HyperlinkManager,
  MediaManager,
  RelationshipManager,
  ZOrderManager,
  OutputWriter,
  TemplateEngine,
  ValidationEngine,
  PackageLoader,
  isPkZip,
  VirtualFileSystem,
  ...spatialUnits,
  generateRelationshipId,
  parseRelationshipId,
  validateXml,
  validateXML,
  safeParseXml,
  repairXML,
  scanForEntities,
  analyzeXmlFile,
  reportXmlComplexity,
  createLogger,
  setGlobalLogLevel,
  resetLogLevel,
  PPTXError,
  SlideNotFoundError,
  ChartNotFoundError,
  TableNotFoundError,
}
