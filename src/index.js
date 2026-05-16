/**
 * pptx-templater
 *
 * A low-level PowerPoint OpenXML templating engine for Node.js.
 * Generates and edits PPTX files directly through XML manipulation
 * without relying on PowerPoint generation libraries.
 *
 * @module pptx-templater
 * @author pptx-templater contributors
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

export { PPTXTemplater } from './core/PPTXTemplater.js';
export { ZipManager } from './managers/ZipManager.js';
export { XMLParser } from './parsers/XMLParser.js';
export { SlideManager } from './managers/SlideManager.js';
export { ChartManager } from './managers/ChartManager.js';
export { TableManager } from './managers/TableManager.js';
export { HyperlinkManager } from './managers/HyperlinkManager.js';
export { MediaManager } from './managers/MediaManager.js';
export { RelationshipManager } from './managers/RelationshipManager.js';
export { OutputWriter } from './core/OutputWriter.js';
export { TemplateEngine } from './core/TemplateEngine.js';

// Utility exports
export { generateRelationshipId, parseRelationshipId } from './utils/relationshipUtils.js';
export { validateXML, repairXML } from './utils/xmlUtils.js';
export { createLogger } from './utils/logger.js';
export { PPTXError, SlideNotFoundError, ChartNotFoundError, TableNotFoundError } from './utils/errors.js';
