# node-pptx-templater (v3.0.0-alpha.0) Codebase Map

## 1. Directory Tree
`	ext
src
├── cli
│   ├── commands
│   │   ├── build.js
│   │   ├── debug.js
│   │   ├── extract.js
│   │   ├── inspect.js
│   │   └── validate.js
│   └── index.js
├── core
│   ├── OutputWriter.js
│   ├── PackageLoader.js
│   ├── PackageLoader.ts
│   ├── PPTXTemplater.js
│   ├── TemplateEngine.js
│   ├── ValidationEngine.js
│   ├── VirtualFileSystem.js
│   └── VirtualFileSystem.ts
├── index.d.ts
├── index.js
├── index.mjs
├── managers
│   ├── ChartManager.js
│   ├── charts
│   │   ├── ChartCacheGenerator.js
│   │   ├── ChartParser.js
│   │   ├── ChartRelationshipManager.js
│   │   └── ChartWorkbookUpdater.js
│   ├── ContentTypesManager.js
│   ├── HyperlinkManager.js
│   ├── ImageManager.js
│   ├── MediaManager.js
│   ├── RelationshipManager.js
│   ├── ShapeManager.js
│   ├── SlideManager.js
│   ├── TableManager.js
│   ├── TextManager.js
│   ├── ZipManager.js
│   └── ZOrderManager.js
├── parsers
│   └── XMLParser.js
├── templates
│   ├── blankPptx.js
│   └── slideTemplate.js
└── utils
    ├── contentTypesHelper.js
    ├── errors.js
    ├── idUtils.js
    ├── imageMetadata.js
    ├── logger.js
    ├── relationshipUtils.js
    ├── units.js
    └── xmlUtils.js
`

## 2. File-by-File Registry

### src/cli/commands/build.js
**Primary Responsibility:** Internal module.

**Internal Dependencies:** ../../index.js
**External Dependencies:** chalk, ora, fs/promises, path

### src/cli/commands/debug.js
**Primary Responsibility:** Internal module.

**Internal Dependencies:** ../../index.js
**External Dependencies:** chalk, ora, path

### src/cli/commands/extract.js
**Primary Responsibility:** Internal module.

**Internal Dependencies:** ../../index.js
**External Dependencies:** chalk, path, fs/promises, jszip

### src/cli/commands/inspect.js
**Primary Responsibility:** Internal module.

**Internal Dependencies:** ../../index.js
**External Dependencies:** chalk, ora, path

### src/cli/commands/validate.js
**Primary Responsibility:** Internal module.

**Internal Dependencies:** ../../index.js
**External Dependencies:** chalk, ora, path

### src/cli/index.js
**Primary Responsibility:** Internal module.

**Internal Dependencies:** ./commands/build.js, ./commands/validate.js, ./commands/inspect.js, ./commands/extract.js, ./commands/debug.js
**External Dependencies:** commander, chalk, path

### src/core/OutputWriter.js
**Primary Responsibility:** Serializes the modified PPTX to various output formats.

**Key Classes / Exports:**
- class OutputWriter
- class OutputWriter

**Internal Dependencies:** ../parsers/XMLParser.js, ../utils/logger.js, ../utils/errors.js
**External Dependencies:** fs-extra, path

### src/core/PackageLoader.js
**Primary Responsibility:** Loads an OpenXML PPTX package from multiple source formats into a VirtualFileSystem.

**Key Classes / Exports:**
- class PackageLoader
- class PackageLoader

**Internal Dependencies:** ./VirtualFileSystem.js, ../utils/errors.js, ../utils/logger.js
**External Dependencies:** path, fs/promises, jszip

### src/core/PackageLoader.ts
**Primary Responsibility:** Internal module.

**Key Classes / Exports:**
- class PackageLoader

**External Dependencies:** jszip

### src/core/PPTXTemplater.js
**Primary Responsibility:** Main engine class for PPTX template manipulation.

**Key Classes / Exports:**
- class PPTXTemplater
- class for
- class PPTXTemplater

**Internal Dependencies:** ../managers/ZipManager.js, ../parsers/XMLParser.js, ../managers/ContentTypesManager.js, ../managers/SlideManager.js, ../managers/ChartManager.js, ../managers/TableManager.js, ../managers/HyperlinkManager.js, ../managers/MediaManager.js, ../managers/RelationshipManager.js, ../managers/ShapeManager.js, ../managers/ImageManager.js, ../managers/TextManager.js, ../managers/ZOrderManager.js, ./ValidationEngine.js, ./OutputWriter.js, ./TemplateEngine.js, ../utils/logger.js, ../utils/errors.js, ../managers/charts/ChartRelationshipManager.js, ../utils/units.js
**External Dependencies:** perf_hooks, crypto, fs-extra, path

### src/core/TemplateEngine.js
**Primary Responsibility:** Internal module.

**Key Classes / Exports:**
- class TemplateEngine

**Internal Dependencies:** ../utils/logger.js, ./ValidationEngine.js

### src/core/ValidationEngine.js
**Primary Responsibility:** Runs audits and verification checks on PPTX files.

**Key Classes / Exports:**
- class ValidationEngine
- class ValidationEngine


### src/core/VirtualFileSystem.js
**Primary Responsibility:** In-memory store for all OpenXML package assets.

**Key Classes / Exports:**
- class VirtualFileSystem
- class VirtualFileSystem

**Internal Dependencies:** ../utils/errors.js, ../utils/logger.js
**External Dependencies:** util

### src/core/VirtualFileSystem.ts
**Primary Responsibility:** Internal module.

**Key Classes / Exports:**
- class VirtualFileSystem

**External Dependencies:** util

### src/index.d.ts
**Primary Responsibility:** Internal module.

**Key Classes / Exports:**
- class PPTXTemplater
- class ZipManager
- class XMLParser
- class SlideManager
- class ChartManager
- class TableManager
- class ShapeManager
- class ImageManager
- class TextManager
- class HyperlinkManager
- class MediaManager
- class RelationshipManager
- class OutputWriter
- class TemplateEngine
- class ValidationEngine
- class PPTXError
- class SlideNotFoundError
- class ChartNotFoundError
- class TableNotFoundError


### src/index.js
**Primary Responsibility:** Internal module.

**Internal Dependencies:** ./core/PPTXTemplater.js, ./managers/ZipManager.js, ./parsers/XMLParser.js, ./managers/SlideManager.js, ./managers/ChartManager.js, ./managers/TableManager.js, ./managers/ShapeManager.js, ./managers/ImageManager.js, ./managers/TextManager.js, ./managers/HyperlinkManager.js, ./managers/MediaManager.js, ./managers/RelationshipManager.js, ./core/OutputWriter.js, ./core/TemplateEngine.js, ./core/ValidationEngine.js, ./utils/relationshipUtils.js, ./utils/xmlUtils.js, ./utils/logger.js, ./utils/errors.js

### src/index.mjs
**Primary Responsibility:** Internal module.


### src/managers/ChartManager.js
**Primary Responsibility:** Handles chart data updates by directly manipulating chart XML files.

**Key Classes / Exports:**
- class ChartManager
- class ChartManager

**Internal Dependencies:** ../utils/logger.js, ../utils/errors.js, ./RelationshipManager.js, ./charts/ChartWorkbookUpdater.js, ./charts/ChartCacheGenerator.js, ../core/ValidationEngine, ../utils/idUtils.js
**External Dependencies:** jszip

### src/managers/charts/ChartCacheGenerator.js
**Primary Responsibility:** Internal module.

**Key Classes / Exports:**
- class ChartCacheGenerator

**Internal Dependencies:** ./ChartWorkbookUpdater.js

### src/managers/charts/ChartParser.js
**Primary Responsibility:** Internal module.

**Key Classes / Exports:**
- class ChartParser


### src/managers/charts/ChartRelationshipManager.js
**Primary Responsibility:** Internal module.

**Key Classes / Exports:**
- class ChartRelationshipManager

**Internal Dependencies:** ../RelationshipManager.js

### src/managers/charts/ChartWorkbookUpdater.js
**Primary Responsibility:** Internal module.

**Key Classes / Exports:**
- class ChartWorkbookUpdater

**Internal Dependencies:** ../../utils/logger.js
**External Dependencies:** jszip

### src/managers/ContentTypesManager.js
**Primary Responsibility:** Internal module.

**Key Classes / Exports:**
- class ContentTypesManager

**Internal Dependencies:** ../utils/logger.js, ../utils/errors.js

### src/managers/HyperlinkManager.js
**Primary Responsibility:** Manages hyperlink creation and modification in PPTX slides.

**Key Classes / Exports:**
- class HyperlinkManager
- class HyperlinkManager

**Internal Dependencies:** ../utils/logger.js, ../utils/errors.js, ./RelationshipManager.js

### src/managers/ImageManager.js
**Primary Responsibility:** Manages image elements and relationships on individual slides.

**Key Classes / Exports:**
- class ImageManager
- class ImageManager

**Internal Dependencies:** ../utils/logger.js, ../utils/errors.js, ./RelationshipManager.js, ../utils/imageMetadata.js

### src/managers/MediaManager.js
**Primary Responsibility:** Manages media embedding, deduplication, and retrieval in PPTX files.

**Key Classes / Exports:**
- class MediaManager
- class MediaManager

**Internal Dependencies:** ../utils/logger.js, ../utils/errors.js
**External Dependencies:** crypto, fs-extra

### src/managers/RelationshipManager.js
**Primary Responsibility:** Parses, manages, and writes OpenXML relationship (.rels) files.

**Key Classes / Exports:**
- class RelationshipManager
- class RelationshipManager

**Internal Dependencies:** ../utils/logger.js, ../utils/relationshipUtils.js

### src/managers/ShapeManager.js
**Primary Responsibility:** Manages shape elements inside PPTX slides.

**Key Classes / Exports:**
- class ShapeManager
- class ShapeManager

**Internal Dependencies:** ../utils/logger.js, ../utils/errors.js, ../parsers/XMLParser.js

### src/managers/SlideManager.js
**Primary Responsibility:** Manages slide loading, ordering, modification, and creation.

**Key Classes / Exports:**
- class SlideManager
- class SlideManager

**Internal Dependencies:** ../utils/logger.js, ../utils/errors.js, ./RelationshipManager.js, ../templates/slideTemplate.js, ../utils/relationshipUtils.js, ../utils/idUtils.js, ../core/PPTXTemplater.js

### src/managers/TableManager.js
**Primary Responsibility:** Handles table operations in PPTX templates.

**Key Classes / Exports:**
- class TableManager
- class TableManager

**Internal Dependencies:** ../utils/logger.js, ../utils/errors.js

### src/managers/TextManager.js
**Primary Responsibility:** Manages text elements, replacements, and search inside slide XML.

**Key Classes / Exports:**
- class TextManager
- class TextManager

**Internal Dependencies:** ../utils/logger.js, ../utils/errors.js, ../core/ValidationEngine.js

### src/managers/ZipManager.js
**Primary Responsibility:** Manages the PPTX ZIP archive — reading, modifying, and re-packaging it.

**Key Classes / Exports:**
- class ZipManager
- class ZipManager

**Internal Dependencies:** ../utils/logger.js, ../utils/errors.js, ../templates/blankPptx.js, ../parsers/XMLParser.js
**External Dependencies:** jszip, fs-extra, path, util

### src/managers/ZOrderManager.js
**Primary Responsibility:** Manages stacking layers and z-index of shapes on slides.

**Key Classes / Exports:**
- class ZOrderManager
- class ZOrderManager

**Internal Dependencies:** ../utils/errors.js, ../parsers/XMLParser.js

### src/parsers/XMLParser.js
**Primary Responsibility:** Provides XML parsing and serialization with OpenXML-aware configuration.

**Key Classes / Exports:**
- class XMLParser
- class XMLParser

**Internal Dependencies:** ../utils/errors.js
**External Dependencies:** fast-xml-parser

### src/templates/blankPptx.js
**Primary Responsibility:** Internal module.


### src/templates/slideTemplate.js
**Primary Responsibility:** Internal module.


### src/utils/contentTypesHelper.js
**Primary Responsibility:** Internal module.

**Key Classes / Exports:**
- class ContentTypesHelper

**Internal Dependencies:** ./logger.js

### src/utils/errors.js
**Primary Responsibility:** Base error class for all node-pptx-templater errors.

**Key Classes / Exports:**
- class PPTXError
- class for
- class PPTXError
- class SlideNotFoundError
- class SlideNotFoundError
- class ChartNotFoundError
- class ChartNotFoundError
- class TableNotFoundError
- class TableNotFoundError
- class XMLParseError
- class XMLParseError
- class InvalidTemplateError
- class InvalidTemplateError
- class MediaEmbedError
- class MediaEmbedError


### src/utils/idUtils.js
**Primary Responsibility:** Internal module.

**External Dependencies:** crypto

### src/utils/imageMetadata.js
**Primary Responsibility:** Internal module.

**External Dependencies:** fs, util

### src/utils/logger.js
**Primary Responsibility:** Internal module.

**Internal Dependencies:** ./logger
**External Dependencies:** node-pptx-templater

### src/utils/relationshipUtils.js
**Primary Responsibility:** Internal module.


### src/utils/units.js
**Primary Responsibility:** Internal module.


### src/utils/xmlUtils.js
**Primary Responsibility:** Internal module.

**Internal Dependencies:** ../parsers/XMLParser.js, ./errors.js
**External Dependencies:** fast-xml-parser

## 3. Feature-to-File Matrix

**Ingestion (PPTX, ZIP, XML directory loading):**
- src/core/PackageLoader.ts
- src/core/VirtualFileSystem.ts

**In-Memory Virtual File System (VFS):**
- src/core/VirtualFileSystem.ts

**OpenXML Parser & Serializer (AST handling, namespaces):**
- src/core/XMLParser.js
- src/core/TemplateEngine.js
- src/utils/xmlSecurity.js

**Presentation & Slide Management (CRUD, slide duplication, rId tracking):**
- src/managers/SlideManager.js
- src/managers/RelationshipManager.js
- src/core/PPTXTemplater.js

**Spatial Engine & Unit Utilities (EMUs, bounding boxes, z-order):**
- src/utils/units.js
- src/managers/ShapeManager.js
- src/managers/ZOrderManager.js

**Element Handlers (Text, Shape, Image, Chart + embedded .xlsx sync):**
- src/managers/ImageManager.js
- src/managers/ChartManager.js
- src/managers/ChartWorkbookUpdater.js
- src/managers/ShapeManager.js
- src/core/TemplateEngine.js

**Table Engine & Auto-Pagination (splitTableOnOverflow):**
- src/managers/TableManager.js
- src/core/PPTXTemplater.js

**Export Adapters (File, Stream, Buffer, Base64, Express):**
- src/core/OutputWriter.js
- src/core/PPTXTemplater.js

**CLI & Entry Points:**
- src/cli/index.js
- src/index.js

## 4. Pending Improvements or Known Gaps

### src/cli/index.js
- **Legacy Sync Detected:** fs.readFileSync

### src/core/PackageLoader.js
- **Legacy Sync Detected:** readFileSync, statSync

### src/core/PackageLoader.ts
- **Legacy Sync Detected:** readFileSync, statSync

### src/core/PPTXTemplater.js
- **Legacy Sync Detected:** fs.readFileSync, existsSync

