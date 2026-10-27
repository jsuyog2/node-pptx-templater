/**
 * node-pptx-templater (v3.0.0-alpha.0)
 *
 * Enterprise-grade, spatial-aware OpenXML PowerPoint templating and layout engine.
 *
 * @module node-pptx-templater
 * @license MIT
 */

export { PPTXTemplater, PPTXTemplater as PPTXTemplate } from './core/PPTXTemplater.js'
export { PackageLoader, isPkZip } from './core/PackageLoader'
export { VirtualFileSystem } from './core/VirtualFileSystem'
export { OutputWriter } from './core/OutputWriter.js'
export { TemplateEngine } from './core/TemplateEngine.js'
export { ValidationEngine } from './core/ValidationEngine.js'

export { ChartManager } from './elements/ChartManager.js'
export { TableManager } from './elements/TableManager.js'
export { ShapeManager } from './elements/ShapeManager.js'
export { ImageManager } from './elements/ImageManager.js'
export { TextManager } from './elements/TextManager.js'

export { SlideManager } from './managers/SlideManager.js'
export { ZipManager } from './managers/ZipManager.js'
export { RelationshipManager } from './managers/RelationshipManager.js'
export { ContentTypesManager } from './managers/ContentTypesManager.js'
export { HyperlinkManager } from './elements/HyperlinkManager.js'
export { ElementManager } from './elements/ElementManager.js'
export { MediaManager } from './managers/MediaManager.js'

export { ZOrderManager } from './spatial/ZOrderManager.js'
export {
  EMU_PER_PT,
  EMU_PER_IN,
  EMU_PER_PX,
  ptToEmu,
  inToEmu,
  pxToEmu,
  emuToPt,
  emuToIn,
  emuToPx,
  toEmu,
  checkCollision,
  isContained,
  getBounds,
  setBounds,
  move,
  resize,
} from './spatial/units'
export type { BoundingBox } from './spatial/units'

export { XMLParser, Z_ORDER_SYMBOL } from './parsers/XMLParser.js'

export {
  generateRelationshipId,
  parseRelationshipId,
} from './utils/relationshipUtils.js'
export {
  validateXml,
  validateXML,
  safeParseXml,
  repairXML,
  scanForEntities,
  analyzeXmlFile,
  reportXmlComplexity,
} from './utils/xmlUtils.js'
export { createLogger, setGlobalLogLevel, resetLogLevel } from './utils/logger.js'
export {
  PPTXError,
  SlideNotFoundError,
  ChartNotFoundError,
  TableNotFoundError,
} from './utils/errors.js'
