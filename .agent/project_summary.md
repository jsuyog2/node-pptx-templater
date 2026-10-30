# Project Architecture Summary: `node-pptx-templater`

> **Version**: 3.0.0-beta.1  
> **Runtime**: Pure JavaScript (Node.js `>=18.x`, AWS Lambda, Cloudflare Workers, Vercel Edge, Deno, Bun)  
> **Standard**: ISO/IEC 29500 (ECMA-376) Parts 1 & 2 (OPC & PresentationML)  
> **CI/CD Pipeline**: Operational (Decoupled Push Gate, Release Publishing, Manual Docs Hosting)  
> **API Catalog**: See `docs-site/src/data/apiReference.json` for full schemas and usage examples.

---

## 1. Architectural Blueprint & Subsystem Managers

`node-pptx-templater` is an in-memory OpenXML package manipulator operating on a Virtual File System (`Map<string, Uint8Array>`) with zero native binaries or office dependencies.

```
PackageLoader / Converters ──► In-Memory VFS (VirtualFileSystem)
                                    └──► XMLParser (Lossless AST & Z-Order)
                                          ├── Core: TemplateEngine, ValidationEngine
                                          ├── Visual Elements: TableManager, ChartManager, ShapeManager,
                                          │                     TextManager, ImageManager, ElementManager,
                                          │                     HyperlinkManager, ZOrderManager
                                          ├── Package Managers: SlideManager, CommentManager, MediaManager,
                                          │                     ZipManager, ContentTypesManager, RelationshipManager
                                          └── Output: OutputWriter (18-method I/O boundary)
```

### Core Subsystems:
- **`VirtualFileSystem`** (`src/core/VirtualFileSystem.js`): Normalized in-memory store (`Map<string, Uint8Array>`); zero intermediate disk writes.
- **`PackageLoader`** (`src/core/PackageLoader.js`): Ingests local files, buffers, streams, URLs, directories, and Flat OPC XML.
- **`XMLParser`** (`src/parsers/XMLParser.js`): Dual-mode AST parser with `[Z_ORDER_SYMBOL]` preserving shape stacking.
- **`TemplateEngine`** (`src/core/TemplateEngine.js`): AST-level run fragment healing across split `<a:r>`/`<a:t>` tags and native list generation.
- **`ValidationEngine`** (`src/core/ValidationEngine.js`): Hardened against XXE, Billion Laughs, and Zip-Slip path traversals.
- **`SlideManager`** (`src/managers/SlideManager.js`): 1-based slide operations, deep cloning, layout binding, sections, and comments.
- **`TableManager`** (`src/elements/TableManager.js`): DrawingML tables, rectangular cell merge/split, unique 32-bit `<a16:rowId>`, auto-pagination.
- **`ChartManager`** (`src/elements/ChartManager.js`): 17 chart types with DrawingML cache and embedded Excel workbook Dual-Sync.
- **`ShapeManager`** (`src/elements/ShapeManager.js`): Vector geometry, fills, borders, text, alignment, and coordinate math.
- **`TextManager`** (`src/elements/TextManager.js`): Variable substitutions, full text box replacement (`replaceTextByTag`), DrawingML typography.
- **`ZOrderManager`** (`src/spatial/ZOrderManager.js`): Visual stacking operations while preserving `<p:nvGrpSpPr>` and `<p:grpSpPr>` invariants.
- **`OutputWriter`** (`src/writer/OutputWriter.js`): Serialization pipeline with Express/Fastify framework streaming adapters.
- **`ZipManager`** (`src/managers/ZipManager.js`): In-memory ZIP compression/decompression via JSZip.

---

## 2. Standardized 8 Categories & Public API Boundary

All APIs, TypeScript interfaces (`index.d.ts`), and JSDocs strictly adhere to 8 canonical categories (130 methods total):

### A. Approved 18-Method I/O Whitelist (`Import/Export`)
- **Static Ingestion & Lifecycle (7)**: `PPTXTemplater.load`, `PPTXTemplater.createBlank`, `PPTXTemplater.extractPptx`, `PPTXTemplater.buildPptx`, `PPTXTemplater.preload`, `PPTXTemplater.fromCache`, `PPTXTemplater.clearCache` (plus utility `PPTXTemplater.setLogLevel`).
- **Export Instance Methods (8)**: `saveToFile`, `toBuffer`, `toStream`, `toBase64`, `toArrayBuffer`, `toResponse`, `toExpress`, `toFastify`.
- **Extraction & Conversion (3)**: `extractToZip`, `extractToXml`, `toFlatXml`.
- *Mandate*: Deprecated aliases (`toFile`, `save`, `exportZip`, `exportXml`, `writeToBuffer`) are strictly prohibited.

### B. Approved 24-Method Slide Management Whitelist (`Slides`)
- **Slide CRUD & Selection (8)**: `useSlide`, `useAllSlides`, `addSlide`, `duplicateSlide`, `removeSlide`, `getSlides`, `getSlide`, `resetSlide`.
- **Cross-Deck & Ordering (4)**: `importSlideFrom`, `exportSlides`, `moveSlide`, `reorderSlides`.
- **Layout & Metadata (4)**: `tagSlide`, `getSlideDimensions`, `getSlideLayout`, `hideSlide`, `unhideSlide`.
- **Sections & Comments (8)**: `useSection`, `addSection`, `updateSection`, `removeSection`, `getSection`, `addComment`, `removeComment`.

### C. Other Canonical Categories (Detailed in `docs-site/src/data/apiReference.json`)
- **`Text & Placeholders`** (10): Text replacement, tag styling, list generation.
- **`Tables`** (30): Table lifecycle, rows/cols CRUD, cell styling, rectangular merge/split, auto-pagination.
- **`Charts`** (11): Chart CRUD, data sync, series manipulation, data labels.
- **`Shapes`** (11): Shape geometry, grouping, borders, fills.
- **`Images & Media`** (11): Images, audio, video embedding, playback configuration.
- **`Elements`** (15): Unified element querying (4), Z-ordering (5), and Hyperlinks (6).

---

## 3. Documentation Portal (`docs-site/`)

The documentation portal is built with Next.js 14+ App Router (`output: 'export'`, `trailingSlash: true`) and deployed to Cloudflare Pages (`https://node-pptx-templater.pages.dev`). Content and API references are fully decoupled into JSON schemas under `docs-site/src/data/`, with 100% API coverage enforced via `npm run docs:validate`.

---

## 4. CI/CD Topology & Staged Release Architecture

- **Push Gating (`.github/workflows/ci.yml`)**: On push/PR to `main`, `develop`, `v3.0.0`, executes sequential verification: (1) `ci` (lint, format:check, test), (2) `verify-package-managers` (npm, pnpm, yarn with `nodeLinker: node-modules`, bun), (3) `build` (compile and dry-run pack). Hard stop with zero publishing.
- **Release Publishing (`.github/workflows/publish.yml`)**: Triggered strictly on GitHub Releases or manual dispatch. Ingests pre-built tarball, runs idempotent checks (`npm view`), and publishes to npm (`node-pptx-templater`) and GitHub Packages (`@<owner>/node-pptx-templater`) using `--ignore-scripts`.
- **Documentation Deployment (`.github/workflows/docs.yml`)**: Manual-only `workflow_dispatch` deploying static exports to Cloudflare Pages (`npm run docs:deploy`).
