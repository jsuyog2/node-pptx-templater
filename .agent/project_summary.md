# Project Architecture Summary: `node-pptx-templater`

> **Version**: 3.0.0-beta.0  
> **Runtime**: Pure JavaScript (Node.js `>=18.x`, AWS Lambda, Cloudflare Workers, Vercel Edge, Deno, Bun)  
> **Standard**: ISO/IEC 29500 (ECMA-376) Parts 1 & 2 (OPC & PresentationML)  
> **API Catalog**: See `docs-site/src/data/apiReference.json` for full parameter definitions & examples

---

## 1. Architectural Blueprint & Subsystem Managers

`node-pptx-templater` is an in-memory OpenXML package manipulator operating on a Virtual File System (`Map<string, Uint8Array>`) with zero native binaries or office dependencies.

```
Ingestion (PackageLoader, FlatOpcConverter, LegacyConverter)
  └─► In-Memory VFS (VirtualFileSystem)
        └─► AST Parsing & Serializing (XMLParser: standard & ordered [Z_ORDER_SYMBOL])
              ├─► Core: TemplateEngine (run healing, lists), ValidationEngine (XXE, integrity)
              ├─► Slide Subsystems: SlideManager, CommentManager
              ├─► Element Subsystems: ElementManager, HyperlinkManager, ZOrderManager, TextManager, TableManager, ChartManager, ShapeManager, ImageManager, MediaManager
              ├─► Package Manifests: RelationshipManager, ContentTypesManager, ZipManager
              └─► Output Pipeline: OutputWriter & FrameworkAdapters (18-method I/O boundary)
```

### Core Subsystem Managers
- **`VirtualFileSystem`** (`src/core/VirtualFileSystem.js`): Normalized in-memory file store (`Map<string, Uint8Array>`).
- **`PackageLoader`** (`src/core/PackageLoader.js`): Ingests local files, buffers, streams, URLs, directories, cloud objects, and Flat OPC XML.
- **`XMLParser`** (`src/parsers/XMLParser.js`): Lossless `fast-xml-parser` with dual-mode AST parsing and Z-order preservation (`[Z_ORDER_SYMBOL]`).
- **`StyleBuilder`** (`src/parsers/StyleBuilder.js`): DrawingML typography and shape style builder.
- **`TemplateEngine`** (`src/core/TemplateEngine.js`): AST-level text run healing across fragmented `<a:r>`/`<a:t>` tags and native list generation.
- **`ValidationEngine`** (`src/core/ValidationEngine.js`): XXE defense, entity bomb prevention, and structural package verification.
- **`SlideManager`** (`src/managers/SlideManager.js`): 1-based slide operations, deep cloning, reordering, layout binding, sections, and comments.
- **`TableManager`** (`src/elements/TableManager.js`): DrawingML table lifecycle, row/column CRUD, rectangular cell merges/splits, `<a16:rowId>` allocation, and polymorphic cell badges.
- **`ChartManager`** (`src/elements/ChartManager.js`): DrawingML (`<c:chartSpace>`) and extended (`<cx:chartSpace>`) chart types with Dual-Sync updating of embedded Excel workbooks.
- **`ShapeManager`** (`src/elements/ShapeManager.js`): Vector geometry, fills, borders, text, alignment, grouping, and coordinate math.
- **`TextManager`** (`src/elements/TextManager.js`): Run substitutions, full text box replacement (`replaceTextByTag`), and DrawingML styling.
- **`ElementManager`** (`src/elements/ElementManager.js`): Unified slide element inspection, querying, relationship-safe removal, and canvas resets.
- **`HyperlinkManager`** (`src/elements/HyperlinkManager.js`): External web URLs, internal slide jumps (`action="ppaction://hlinksldjump"`), and slideshow actions.
- **`ZOrderManager`** (`src/spatial/ZOrderManager.js`): Visual stacking operations while preserving `<p:nvGrpSpPr>` and `<p:grpSpPr>` header invariants.
- **`ImageManager`** & **`MediaManager`** (`src/elements/ImageManager.js`, `src/managers/MediaManager.js`): Image, audio, and video insertion, MIME detection, aspect ratio, and timing tracks.
- **`RelationshipManager`** & **`ContentTypesManager`** (`src/managers/`): `.rels` target graphs and `[Content_Types].xml` default/override mappings.
- **`ZipManager`** (`src/managers/ZipManager.js`): In-memory ZIP packaging and extraction via JSZip.
- **`OutputWriter`** (`src/writer/OutputWriter.js`): Multi-target output serialization with Express/Fastify adapters.

---

## 2. Exact 8 Standardized Categories & 130 Public Methods

The public API surface of `PPTXTemplater` is strictly organized across 8 categories:

1. **`Import/Export`** (18 Methods):
   - *Static (7)*: `load`, `createBlank`, `extractPptx`, `buildPptx`, `preload`, `fromCache`, `clearCache` (plus `setLogLevel`).
   - *Export (8)*: `saveToFile`, `toBuffer`, `toStream`, `toBase64`, `toArrayBuffer`, `toResponse`, `toExpress`, `toFastify`.
   - *Extraction (3)*: `extractToZip`, `extractToXml`, `toFlatXml`.
2. **`Slides`** (24 Methods):
   - `useSlide`, `useAllSlides`, `addSlide`, `duplicateSlide`, `removeSlide`, `getSlides`, `getSlide`, `importSlideFrom`, `exportSlides`, `moveSlide`, `reorderSlides`, `tagSlide`, `getSlideDimensions`, `getSlideLayout`, `useSection`, `addSection`, `updateSection`, `removeSection`, `getSection`, `hideSlide`, `unhideSlide`, `resetSlide`, `addComment`, `removeComment`.
3. **`Text & Placeholders`** (10 Methods):
   - `replaceText`, `addText`, `replaceTextByTag`, `updateText`, `addLists`, `getList`, `findText`, `addStyle`, `removeStyle`, `updateStyle`.
4. **`Tables`** (30 Methods):
   - `addTable`, `updateTable`, `updateTableData`, `removeTable`, `getTables`, `getTable`, `cloneTable`, `addTableRow`, `addTableColumn`, `removeTableRow`, `removeTableColumn`, `updateRowData`, `updateColumnData`, `updateTableCell`, `clearRowData`, `clearColumnData`, `clearCellData`, `getTableData`, `getColumnData`, `getRowData`, `getColumn`, `getRow`, `getCell`, `getCellData`, `mergeCells`, `splitCell`, `styleTable`, `styleColumn`, `styleRow`, `styleCell`.
5. **`Charts`** (11 Methods):
   - `addChart`, `updateChart`, `updateChartData`, `removeChart`, `clearCharts`, `getCharts`, `getChart`, `cloneChart`, `getDataLabels`, `getChartLabelPositions`.
6. **`Shapes`** (11 Methods):
   - `addShape`, `updateShape`, `removeShape`, `clearShapes`, `getShapes`, `getShape`, `cloneShape`, `getGroupShapes`, `groupShapes`, `ungroupShapes`.
7. **`Images & Media`** (11 Methods):
   - `addImage`, `updateImage`, `removeImage`, `addVideo`, `updateVideo`, `removeVideo`, `addAudio`, `updateAudio`, `removeAudio`, `getMedias`, `getMedia`, `cloneMedia`.
8. **`Elements`** (15 Methods):
   - *Element Queries (4)*: `getElements`, `getElement`, `removeElement`, `removeAllElements`.
   - *Z-Ordering (5)*: `setZOrder`, `bringToFront`, `sendToBack`, `bringForward`, `sendBackward`.
   - *Hyperlinks (6)*: `addHyperlink`, `updateHyperlink`, `removeHyperlink`, `getHyperlinks`, `getHyperlink`, `setHyperlink`.
   - *Spatial Collision (2)*: `checkCollision`, `isContained`.

---

## 3. Documentation Portal (`docs-site/`)

The documentation portal is redesigned into an authentic Microsoft PowerPoint desktop software environment built with Next.js 14+ App Router (`output: 'export'`, `images: { unoptimized: true }`, `trailingSlash: true`) and deployed to Cloudflare Pages (`https://node-pptx-templater.pages.dev`).

### Architecture & UI System:
- **Microsoft PowerPoint Desktop Chrome**:
  - **Top Title Bar (36px)**: Quick Access icons (Save, Undo, Redo, Slide Show), file title `"node-pptx-templater.pptx — PowerPoint"`, AutoSave indicator, Theme Toggle, and Window Controls (`—`, `□`, `✕`).
  - **Subheader Ribbon (68px)**: Persistent `#D24726` "File" tab, primary navigation tabs, and quick action ribbon tools (npm copy, GitHub, npm).
  - **Fixed Left Slide Thumbnail Strip (280px)**: Vertical 16:9 thumbnail previews showing method/section titles with slide numbers and active `#D24726` selection outline.
  - **Center Slide Stage (16:9 Aspect Ratio)**: Centered presentation slide cards with syntax-highlighted code blocks, parameter tables, and 1-click Copy buttons.
  - **Bottom Status Bar (32px)**: Slide counter (`"Slide X of Y"`), pure JS OpenXML engine indicator, view mode selectors, and zoom controls.
- **Strict 4-Route Hierarchy**:
  - `/` — Main Home Landing Page with HeroSection, ScrollTrigger reveals, interactive recipe playground, and comparison matrix.
  - `/docs` — Documentation Overview presentation deck.
  - `/docs/installation` — Dedicated Installation & Setup presentation deck.
  - `/docs/capabilities` — Capabilities, Architecture & Runtime Compatibility Matrix.
  - `/docs/api-reference` — Complete API Reference presentation deck covering all 8 canonical categories.
- **GSAP Animation Engine**:
  - `docs-site/src/animations/gsapSlideTransitions.js`: ScrollTrigger staggered reveals on the hero section and smooth vertical push / slide-up slide transitions on wheel/keyboard navigation synchronizing with the thumbnail sidebar.
- **Comprehensive SEO & AI Ingestion**:
  - `docs-site/public/llms.txt`: Machine-optimized LLM crawler documentation for Perplexity, Google AI Overviews, Claude, and ChatGPT Search.
  - Schema.org JSON-LD structured data (`SoftwareApplication`, `TechArticle`, `FAQPage`).
  - Standard `sitemap.xml` and `robots.txt` generated automatically via Next.js route handlers.
- **v4.2.0 UI/UX & Responsive Enhancements**:
  - **Internal Card Scroll Guard**: Slide deck detects boundary conditions (`scrollTop <= 0` and `scrollTop + clientHeight >= scrollHeight - 2`) so internal content scrolls naturally before deck snapping triggers.
  - **Thumbnail Collision Fix**: Sidebar header has opaque backdrop fill (`bg-[#E6DEC8] dark:bg-[rgba(38,38,38,0.98)]`) and thumbnails sit in an isolated flex scroll sub-container with `pt-3` and `overscroll-contain`.
  - **Horizontal Folder Deck**: Landing page converts sections into a horizontal presentation folder flow powered by GSAP ScrollTrigger and touch swipes.
  - **Mobile Slide Filmstrip**: Responsive breakpoint switching (`max-lg:hidden` on left sidebar, bottom horizontal 16:9 filmstrip with auto-centering on `< 1024px`).
  - **Global Omnisearch**: Universal search dialog (`Cmd+K` / `Ctrl+K` and header trigger button) indexing all 135+ methods, categories, and slides with keyboard navigation and instant deep linking.
