# Project Architecture Summary: `node-pptx-templater`

> **Version**: 3.0.0-alpha.0  
> **Repository Type**: Pure JavaScript Low-Level OpenXML Presentation Manipulation Engine  
> **Standards Compliance**: ISO/IEC 29500 (ECMA-376) Parts 1 & 2 (OPC & PresentationML)  
> **Runtime Target**: Node.js (`>=18.x`, `20.x`, `22.x+`), AWS Lambda, Cloudflare Workers, Vercel Edge

---

## 1. High-Level System Architecture

`node-pptx-templater` is architectured as an in-memory virtual OpenXML package manipulator. Rather than depending on heavyweight office automation binaries or high-level abstraction wrappers, it interacts directly with the OpenXML AST (Abstract Syntax Tree) and the underlying ZIP container.

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                                INGESTION & PARSING                                     │
│  ┌─────────────────────────┐  ┌──────────────────────────┐  ┌────────────────────────┐ │
│  │ PackageLoader.js        │  │ FlatOpcConverter.js      │  │ LegacyConverter.js     │ │
│  │ (Buffer/Stream/URL/Dir) │  │ (<pkg:package> Single-XML│  │ (.ppt / .odp detection)│ │
│  └───────────┬─────────────┘  └────────────┬─────────────┘  └───────────┬────────────┘ │
└──────────────┼─────────────────────────────┼────────────────────────────┼──────────────┘
               ▼                             ▼                            ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                         IN-MEMORY VIRTUAL FILE SYSTEM (VFS)                            │
│  VirtualFileSystem.js: Map<string, Uint8Array> — Zero synchronous disk blocking        │
│  Files: [Content_Types].xml, _rels/.rels, ppt/presentation.xml, ppt/slides/slide*.xml  │
└───────────────────────────────────────────┬────────────────────────────────────────────┘
                                            ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                         AST PARSING & NORMALIZATION LAYER                              │
│  XMLParser.js: fast-xml-parser engine with dual-mode parsing                          │
│  • Normal Parser: Fast structural navigation with @_ attributes & #text                │
│  • Ordered Parser (preserveOrder: true): Extracts precise <p:spTree> Z-order symbol    │
└───────────────────────────────────────────┬────────────────────────────────────────────┘
                                            ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                           SPECIALIZED SUBSYSTEM MANAGERS                               │
│  ┌────────────────────────┐ ┌────────────────────────┐ ┌─────────────────────────────┐ │
│  │ SlideManager.js        │ │ TableManager.js        │ │ ChartManager.js             │ │
│  │ (Clone, reorder, sldId)│ │ (Row CRUD, merge, auto)│ │ (Dual-sync: cache + XLSX)   │ │
│  ├────────────────────────┤ ├────────────────────────┤ ├─────────────────────────────┤ │
│  │ ShapeManager.js        │ │ ImageManager.js        │ │ TextManager.js              │ │
│  │ (Geometry, fill, stroke│ │ (MIME, aspect ratio)   │ │ (Typography, run styling)   │ │
│  ├────────────────────────┤ ├────────────────────────┤ ├─────────────────────────────┤ │
│  │ ZOrderManager.js       │ │ HyperlinkManager.js    │ │ MediaManager.js             │ │
│  │ (Layer stacking, group)│ │ (External & internal)  │ │ (Audio, video, embedded)    │ │
│  ├────────────────────────┤ ├────────────────────────┤ ├─────────────────────────────┤ │
│  │ RelationshipManager.js │ │ ContentTypesManager.js │ │ TemplateEngine.js           │ │
│  │ (_rels target chains)  │ │ (Overrides & Defaults) │ │ (Run healing, list builder) │ │
│  └────────────────────────┘ └────────────────────────┘ └─────────────────────────────┘ │
└───────────────────────────────────────────┬────────────────────────────────────────────┘
                                            ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                        SERIALIZATION & STRICT 18-METHOD I/O                            │
│  OutputWriter.js & FrameworkAdapters.js: Multi-target streaming, buffer, and exports   │
│  Targeting: File, Buffer, Stream, ArrayBuffer, Base64, Response, Express, Fastify      │
└────────────────────────────────────────────────────────────────────────────────────────┘
```

---

## 2. Core Engine Components

### 2.1 The Virtual File System (`src/core/VirtualFileSystem.js`)
- **Storage Model**: All package assets (XML documents, binary PNG/JPEG images, Excel workbooks, `.rels` manifests) are stored in an internal `Map<string, Uint8Array>`.
- **Normalization**: Enforces uniform forward-slash relative paths matching standard ZIP entry names (e.g., `ppt/slides/slide1.xml`, `[Content_Types].xml`).
- **Zero Disk Blocking**: Decoupled from physical disk storage. Read and write operations occur strictly in memory via `readBytes()`, `writeBytes()`, `readText()`, and `writeText()`.

### 2.2 Package Ingestion Engine (`src/core/PackageLoader.js`)
- **Polymorphic Ingestion**: Autodetects and loads presentations from 6 distinct source categories:
  1. Local file paths (`.pptx`, `.pptm`, `.potx`, `.potm`, `.ppsx`, `.ppsm`, `.ppam`).
  2. Binary `Buffer` or `Uint8Array` (e.g., AWS S3 `getObject` body, multipart uploads).
  3. Node.js `Readable` streams and Web Standard `ReadableStream`.
  4. PKZip containers via magic bytes (`0x50 0x4B 0x03 0x04`).
  5. Uncompressed OpenXML directory trees containing `[Content_Types].xml`, `_rels/`, and `ppt/`.
  6. Remote HTTP/HTTPS URLs via streaming fetch.
  7. Cloud object configurations (AWS S3, Google Cloud Storage, Azure Blob).
  8. Microsoft Flat OPC single-file XML packages (`<pkg:package>`).

### 2.3 XML Parser & Serializer (`src/parsers/XMLParser.js`)
- **Fast XML Processing**: Powered by `fast-xml-parser` configured for lossless round-tripping.
- **Attributes & Namespaces**: Attributes are prefixed with `@_` (e.g., `@_id`, `@_name`). XML namespaces (`a:`, `p:`, `c:`, `r:`) are strictly preserved.
- **Dual-Mode AST Parsing**:
  - Standard Parser: Parses documents into indexed JavaScript objects for rapid node traversal.
  - Ordered Parser (`preserveOrder: true`): Used on `<p:spTree>` containers to accurately record drawing element visual stacking sequences without disturbing DOM structure.
- **Z-Order Symbol Binding**: Attaches `[Z_ORDER_SYMBOL]` to `<p:spTree>` and `<p:grpSp>` nodes to guarantee element rendering order is preserved across serialization.
- **Non-Visual Properties Protection**: Guarantees `<p:nvGrpSpPr>` and `<p:grpSpPr>` header elements remain at the top of the container during rebuilding.

### 2.4 Template Engine (`src/core/TemplateEngine.js`)
- **AST-Level Text Run Healing**: Solves the classic OpenXML problem where Microsoft PowerPoint fragments placeholder tokens across multiple `<a:r>` and `<a:t>` nodes (e.g., `{{` in run 1, `FIRST_` in run 2, `NAME}}` in run 3).
- **Format Preservation**: Merges run text strings, performs substitutions, and preserves the original formatting (`<a:rPr>`) from the first run.
- **Native List Engine**: Generates multi-level bulleted or numbered list paragraphs directly in the AST with custom indentation and font formatting.

### 2.5 Validation & Security Engine (`src/core/ValidationEngine.js`, `src/utils/xmlUtils.js`)
- **XXE Guard**: Rejects documents containing `<!DOCTYPE>` declarations with `SYSTEM` or `PUBLIC` external entity identifiers.
- **Billion Laughs / Entity Bomb Shield**: Detects custom entity declarations and rejects recursive expansion attacks.
- **Package Integrity Audit**: Verifies presentation manifests, slide ID lists, relationship targets, and Content Types overrides.

---

## 3. Subsystem Managers & Responsibilities

| Subsystem Manager | File Location | Key Responsibilities |
| :--- | :--- | :--- |
| **`SlideManager`** | `src/managers/SlideManager.js` | Slide indexing (1-based), deep cloning, deletion, reordering, layout mapping, cross-deck slide import, section synchronization. |
| **`TableManager`** | `src/elements/TableManager.js` | Dynamic row/column addition, cell data mutation, alternating zebra striping, custom borders, cell merging (`gridSpan`, `rowSpan`), `<a16:rowId>` allocation. |
| **`ChartManager`** | `src/elements/ChartManager.js` | Bar, Column, Line, Pie, Doughnut, Area, Scatter chart updates; series data & category mutations; data label controls. |
| **`ChartWorkbookUpdater`** | `src/elements/charts/ChartWorkbookUpdater.js` | Synchronizes embedded Excel workbooks (`sheet1.xml`, `sharedStrings.xml`) and clears conflicting table definitions. |
| **`ShapeManager`** | `src/elements/ShapeManager.js` | Adding vector shapes, setting fills (solid, gradient), strokes, geometry, text inside shapes, aligning and distributing shapes. |
| **`CellShapeManager`** | `src/elements/TableManager.js` | Embedding vector shapes, badges, and status pills directly into table cells. |
| **`ImageManager`** | `src/elements/ImageManager.js` | PNG, JPEG, SVG, GIF, WebP insertion; in-place image swapping; aspect ratio locking; Base64 data URI handling. |
| **`MediaManager`** | `src/managers/MediaManager.js` | Managing binary media files, MIME type mappings, and audio/video relationship links. |
| **`TextManager`** | `src/elements/TextManager.js` | In-place text replacements, font sizing, font families, color fills, character styling (bold, italic, underline, strikethrough). |
| **`ZOrderManager`** | `src/spatial/ZOrderManager.js` | Element layer management: `bringToFront`, `sendToBack`, `stepForward`, `stepBackward`, `setZIndex`, `moveObjectBefore`, `moveObjectAfter`. |
| **`RelationshipManager`**| `src/managers/RelationshipManager.js` | Managing `.rels` files across presentation root, slides, slide layouts, and slide masters; generating unique `rId` keys. |
| **`ContentTypesManager`**| `src/managers/ContentTypesManager.js` | Managing `[Content_Types].xml`; registering MIME type `<Default>` extensions and part `<Override>` elements. |
| **`CommentManager`** | `src/managers/CommentManager.js` | Managing collaborative review comments, authors manifest (`ppt/commentAuthors.xml`), comment parts (`ppt/comments/commentN.xml`), and slide relationship links. |
| **`ZipManager`** | `src/managers/ZipManager.js` | In-memory compression and decompression of OpenXML packages using JSZip. |

---

## 4. Strict 18-Method I/O API Whitelist

The entire I/O surface of `PPTXTemplater` is strictly defined by the following 18 methods:

```text
STATIC LIFECYCLE & FACTORY METHODS (7)
├── PPTXTemplater.load(source, options)
├── PPTXTemplater.createBlank(options)
├── PPTXTemplater.extractPptx(pptxPath, outputDir, options)
├── PPTXTemplater.buildPptx(sourceDir, outputPath, options)
├── PPTXTemplater.preload(source, options)
├── PPTXTemplater.fromCache(source)
└── PPTXTemplater.clearCache()

EXPORT INSTANCE METHODS (8)
├── ppt.saveToFile(destinationPath, options)
├── ppt.toBuffer(options)
├── ppt.toStream(options)
├── ppt.toBase64(options)
├── ppt.toArrayBuffer(options)
├── ppt.toResponse(options)
├── ppt.toExpress(res, filename, options)
└── ppt.toFastify(reply, filename, options)

EXTRACTION & CONVERSION INSTANCE METHODS (3)
├── ppt.extractToZip(destinationPath, options)
├── ppt.extractToXml(destinationDir, options)
└── ppt.toFlatXml(destinationPathOrOptions, options)
```

---

## 4.1 Strict 23-Method Slide Management API Whitelist

The Slide Management domain of `PPTXTemplater` is strictly locked to **exactly 23 approved methods**:

```text
SLIDE SELECTION & SCOPING (2)
├── ppt.useSlide(...slideRefs)
└── ppt.useAllSlides()

SLIDE CREATION & DUPLICATION (2)
├── ppt.addSlide(options, position)
└── ppt.duplicateSlide(sourceIndex, atPosition)

SLIDE LIFECYCLE & REORDERING (3)
├── ppt.removeSlide(position)
├── ppt.moveSlide(fromIndex, toIndex)
└── ppt.reorderSlides(order)

SLIDE INSPECTION & METADATA (5)
├── ppt.getSlides()
├── ppt.getSlide(index)
├── ppt.tagSlide(slideIndex, tag)
├── ppt.getSlideDimensions()
└── ppt.getSlideLayout(slideIndex)

CROSS-DECK TRANSFER (2)
├── ppt.importSlideFrom(sourcePpt, sourceIndex, atPosition)
└── ppt.exportSlides(...slideIndices)

SECTION MANAGEMENT (<p14:sectionLst>) (5)
├── ppt.useSection(sectionRef)
├── ppt.addSection(name, startSlideIndex)
├── ppt.updateSection(sectionIdentifier, options)
├── ppt.removeSection(sectionIdentifier, removeSlides)
└── ppt.getSection(sectionRef)

SLIDE VISIBILITY & CANVAS RESET (3)
├── ppt.hideSlide(...slideRefs)
├── ppt.unhideSlide(...slideRefs)
└── ppt.resetSlide(slideRef, options)

COLLABORATIVE COMMENTS (2)
├── ppt.addComment(text, options)
└── ppt.removeComment(commentIdentifier, options)
```

All non-standard aliases (such as `cloneSlide`, `deleteSlide`, `insertSlide`, `importSlides`, `setSlideDimensions`, `setSlideSize`, `setAspectRatio`, `getSlideSize`) are strictly prohibited and purged from the library to preserve clean API boundaries and zero ambiguity.

### 4.2 Coordinate Subsystem Overhaul & Positional Import (v3.3.0 Scope)
1. **Global Percentage-Based Coordinates**:
   - Universal translation of percentages (`'0%'`–`'100%'` or numeric values with `unit: '%'`) into exact English Metric Units (EMUs) mapped against presentation canvas dimensions (16:9 widescreen default `12192000 x 6858000`).
   - Standard 4-corner mapping: Top-Left `(0, 0)`, Top-Right `(100%, 0)`, Bottom-Left `(0, 100%)`, Bottom-Right `(100%, 100%)`.
2. **Positional Slide Import (`importSlideFrom`)**:
   - Insertion index targeting via `atPosition` (1-based index) with boundary clamping (`<= 0` -> 1, `> count` -> end).
3. **PowerPoint Section Protocol (`addSection`)**:
   - Generates unique GUIDs and `<p14:section>` with `<p14:sldIdLst>` within presentation extension lists (`<p:extLst>`), properly registering `xmlns:p14`.
4. **Rich Metadata & Geometry Extraction (`getSlide` / `getSlides`)**:
   - Comprehensive extraction of geometry (EMUs and px), element types, text runs, placeholders, notes, transitions, and animations with zero circular references.
5. **Non-Blocking In-Memory Lifecycles**:
   - Validation is fully decoupled from core mutation flows (`createBlank`, `load`, `addShape`, `importSlideFrom`, exports) and accessible on-demand via `ppt.validatePresentation()`.

---

## 5. Command-Line Interface (CLI)

The CLI (`src/cli/index.js`) is installed globally as `node-pptx-templater` and provides 5 commands:

1. **`build <template> <output>`**:
   Builds a PPTX from a template with data injected from a JSON file.
   - `-d, --data <file>`: JSON file with template data.
   - `-s, --slide <numbers>`: Comma-separated slide numbers (e.g. `1,3,5`).
2. **`validate <file>`**:
   Validates the structural integrity and OpenXML compliance of a presentation.
   - `--strict`: Treats warnings as fatal errors.
3. **`inspect <file>`**:
   Inspects the internal parts, slide elements, charts, tables, media, and relationship tree.
   - `--slides`, `--charts`, `--tables`, `--media`, `--rels`, `--all`.
4. **`extract <file>`**:
   Extracts slide XML, chart XML, relationship files, or entire uncompressed folder trees.
   - `-s, --slide <num>`, `-o, --out <path>`, `--all`, `--folder <path>`, `--zip <path>`, `--format-xml`.
5. **`debug <file>`**:
   Analyzes potentially corrupted presentations and reports diagnostics.
   - `--fix`: Attempts automatic XML repairs.
   - `-o, --out <path>`: Destination path for repaired presentation.

---

## 6. Documentation Portal Architecture (`docs-site/`)

The documentation portal is built using **Next.js 14+ (App Router)** and deployed to **Cloudflare Pages**:
- **Static Export**: Configured with `output: 'export'`, `images: { unoptimized: true }`, and `trailingSlash: true` in `docs-site/next.config.mjs`.
- **Wrangler Configuration**: Deploys via `docs-site/wrangler.toml` targeting `pages_build_output_dir = "out"`.
- **Decoupled JSON Schemas**: All documentation content is driven by structured data files in `docs-site/src/data/`:
  - `apiReference.json`: 103 documented methods across 10 functional categories with signatures, OpenXML impact, edge cases, parameters, returns, and runnable code samples.
  - `features.json`: Granular feature catalog and capability matrix.
  - `docsNavigation.json`: Hierarchical sidebar navigation tree.
  - `compatibility.json`: Version support matrices (Node, Cloudflare Workers, AWS Lambda, browsers).
  - `siteMeta.json`: SEO titles, metadata, OpenGraph tags, and repository links.
- **Design System**: Glassmorphism UI using Tailwind CSS, supporting dark/light mode (`ThemeContext.jsx`), GSAP smooth page transitions (`gsapTransitions.js`), sticky header, and fixed non-scrolling documentation sidebar.
