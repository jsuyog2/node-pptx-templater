# `node-pptx-templater` — Project Architecture & System Summary

## 1. Project Overview & Mission

`node-pptx-templater` is an enterprise-grade, low-level PowerPoint (PPTX) OpenXML template and manipulation engine for Node.js. It operates with **zero external native binary dependencies** (no LibreOffice, no Java, no headless browser, no C++ addons) and performs non-blocking, in-memory OpenXML transformations.

The library solves the critical flaws of traditional string-based PowerPoint patchers:
- **Zero Repair Warnings:** Synchronizes DrawingML chart definitions alongside embedded Excel workbook binaries (`ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx`) so PowerPoint never displays "PowerPoint found a problem with content" repair dialogs.
- **Run Normalization (Anti-Fragmentation):** Reassembles fragmented text placeholder tokens (e.g. `{{` in one XML run, variable name in another) across adjacent `<a:r>` and `<a:t>` nodes without breaking font, color, or style declarations.
- **Strict Coordinate Precision:** Standardizes all spatial coordinates and dimensions in integer English Metric Units (EMUs).
- **Flagship Table Pagination:** Recursively slices overflowing tables across duplicated slides against vertical margins while repeating headers and preserving pinned summary/total rows.

---

## 2. Technology Stack & Runtimes

| Component | Technology | Rationale / Details |
| :--- | :--- | :--- |
| **Runtime Target** | Node.js `>= 20.12.0` (also compatible with Node 22 and 24) | Pure JavaScript execution suitable for AWS Lambda, Google Cloud Functions, Cloud Run, and Docker. |
| **Module Formats** | Dual ESM (`dist/index.mjs`) & CJS (`dist/index.js`) | Bundled via `tsup` with native `.d.ts` TypeScript declaration files. |
| **XML Parsing & Serialization** | `fast-xml-parser` `^5.9.3` | High-throughput AST parsing, attribute preservation (`@_`), and CDATA/entity escaping. |
| **ZIP Package Container** | `jszip` `^3.10.1` | In-memory archive extraction and non-blocking compression. |
| **Test Framework** | `vitest` `^4.1.9` | 306 unit, integration, and snapshot tests (47 test files) with v8 code coverage. |
| **Linter & Formatter** | `eslint` `^8.57.0`, `prettier` `^3.2.5` | Strict style consistency across `src/`, `tests/`, and `showcase/`. |
| **CLI Framework** | `commander` `^12.0.0`, `ora`, `chalk` | Developer CLI (`node-pptx-templater`) for inspection, extraction, and validation. |

---

## 3. Subsystem & Manager Architecture

The library is organized into specialized managers orchestrated through a single fluent interface:

```
                          ┌─────────────────────────────┐
                          │        PPTXTemplater        │
                          │   (Fluent Orchestrator)     │
                          └──────────────┬──────────────┘
                                         │
    ┌──────────────────┬─────────────────┼─────────────────┬──────────────────┐
    ▼                  ▼                 ▼                 ▼                  ▼
┌──────────────┐ ┌──────────────┐ ┌──────────────┐ ┌──────────────┐ ┌────────────────┐
│PackageLoader │ │SlideManager  │ │TableManager  │ │ChartManager  │ │ShapeManager    │
│& VirtualFS   │ │& RelManager  │ │& Pagination  │ │& Excel Sync  │ │& ZOrderManager │
└──────────────┘ └──────────────┘ └──────────────┘ └──────────────┘ └────────────────┘
    │                  │                 │                 │                  │
    ▼                  ▼                 ▼                 ▼                  ▼
┌──────────────┐ ┌──────────────┐ ┌──────────────┐ ┌──────────────┐ ┌────────────────┐
│ZipManager    │ │ContentTypes  │ │Cell Shapes   │ │ChartCache    │ │Spatial Units   │
│& XMLParser   │ │& Hyperlinks  │ │& Row ID Gen  │ │& Workbook    │ │(EMU Engine)    │
└──────────────┘ └──────────────┘ └──────────────┘ └──────────────┘ └────────────────┘
    │                  │                                               │
    ▼                  ▼                                               ▼
┌──────────────┐ ┌──────────────┐                       ┌─────────────────────────────┐
│LegacyFormat- │ │FlatOpc-      │                       │        OutputWriter         │
│Converter     │ │Converter     │                       │ & FrameworkAdapters         │
└──────────────┘ └──────────────┘                       └─────────────────────────────┘
                                                                       │
                                                                       ▼
                                                        ┌─────────────────────────────┐
                                                        │File, Buffer, Stream, Array, │
                                                        │Base64, Express, FlatXml, Web│
                                                        └─────────────────────────────┘
```

### Core Architecture Roles

1. **`PPTXTemplater` (`src/core/PPTXTemplater.js`):** The primary user-facing orchestrator. Exposes fluent methods for slide scoping (`useSlide`, `useAllSlides`), text substitution, element creation, and multi-target serialization.
2. **`PackageLoader` (`src/core/PackageLoader.js`):** Polymorphic ingestion layer that auto-detects inputs: local file paths (.pptx, .pptm, .potx, .potm, .ppsx, .ppsm, .ppam, .ppt, .odp), Microsoft Flat OPC XML packages (`.xml`, `<pkg:package>`), in-memory `Buffer`, `Uint8Array`, streams, remote web URLs, cloud object stores (S3/GCS/Azure), and raw unpacked OpenXML directory trees into an in-memory `VirtualFileSystem`.
3. **`FlatOpcConverter` (`src/converters/FlatOpcConverter.js`):** Parses and serializes Microsoft Flat OPC XML packages (`<pkg:package>`), extracting `<pkg:xmlData>` into VFS XML strings and base64 decoding `<pkg:binaryData>` into binary buffers for media and workbooks, with zero ZIP decompression overhead.
4. **`LegacyFormatConverter` (`src/converters/LegacyFormatConverter.js`):** Binary signature detector (`isLegacyPpt`, `isOdp`, `isPkZip`) that parses legacy OLE Compound binary PPT streams and OpenDocument ODP XML, normalizing text and slides into standard OpenXML VFS trees.
5. **`VirtualFileSystem` (`src/core/VirtualFileSystem.js`):** In-memory map (`Map<string, Uint8Array | string>`) holding raw presentation parts, preventing disk thrashing and concurrency conflicts.
6. **`XMLParser` (`src/parsers/XMLParser.js`):** Configured fast-xml-parser instance with custom symbol-based Z-order preservation (`Z_ORDER_SYMBOL`) and security sanitization.
7. **`SlideManager` (`src/managers/SlideManager.js`):** Slide graph life-cycle: enumeration, deep-duplication with relationship remapping, deletion with manifest cleanup, and slide sizing/dimensions.
8. **`TableManager` (`src/elements/TableManager.js`):** Table mutation, row/column CRUD, multi-directional cell merges (`rowspan`/`gridSpan`), zebra striping, cell borders, cell shapes, and recursive auto-pagination (`splitTableOnOverflow`).
9. **`ChartManager` & `ChartWorkbookUpdater` (`src/elements/ChartManager.js`, `src/managers/charts/`):** Mutates `<c:chartSpace>` series and synchronizes matching `.xlsx` workbook binary parts (`ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx`).
10. **`ShapeManager` & `ZOrderManager` (`src/elements/ShapeManager.js`, `src/spatial/ZOrderManager.js`):** Vector shape generation, absolute geometry calculations in EMUs, and `<p:spTree>` visual layer reordering (`bringToFront`, `sendToBack`).
11. **`ImageManager` & `MediaManager` (`src/elements/ImageManager.js`, `src/managers/MediaManager.js`):** Image insertion and replacement across PNG, JPEG, SVG, WebP, Base64 data URIs, and TIFF with aspect-ratio containment.
12. **`RelationshipManager` & `ContentTypesManager` (`src/managers/RelationshipManager.js`, `src/managers/ContentTypesManager.js`):** Manages `_rels/*.rels` relationship trees, sequential `rId` generation, and `[Content_Types].xml` extension/override registrations.
13. **`TemplateEngine` (`src/core/TemplateEngine.js`):** AST-level run normalization and token extraction/replacement across paragraph nodes.
14. **`BlankPresentation` (`src/templates/blankPresentation.js`):** Minimal valid OpenXML AST generator synthesizing `[Content_Types].xml`, `_rels/.rels`, `presentation.xml`, `slideMaster1.xml`, `slideLayout1.xml`, `theme1.xml`, and blank `slide1.xml` for template-free creation with configurable aspect ratios.
15. **`OutputWriter` (`src/writer/OutputWriter.js`):** Multi-target serialization engine supporting disk files (`saveToFile`), Node.js Buffers (`toBuffer`), streams (`toStream`), ArrayBuffers (`toArrayBuffer`), Base64 data URIs (`toBase64`), Flat OPC XML (`toFlatXml`), Express.js (`toExpress`), Fastify (`toFastify`), Web API Response (`toResponse`), and raw extraction (`extractToZip`, `extractToXml`).
16. **`FrameworkAdapters` (`src/writer/FrameworkAdapters.js`):** HTTP streaming adapters for Express (`sendExpress`), Fastify (`sendFastify`), Koa (`sendKoa`), and Web standard fetch Response (`createWebResponse`).

---

## 4. Key Primary APIs

### Lifecycle & Loading (Strict Whitelist: 7 Static, 8 Export, 3 Extraction)
```javascript
// Load standard OpenXML presentation or template (.pptx, .pptm, .potx, .potm, .ppsx)
const ppt = await PPTXTemplater.load('./template.pptx', { logLevel: 'warn' });

// Load from in-memory Buffer / Uint8Array (AWS S3, Google Cloud Storage, Azure Blob)
const pptFromBuffer = await PPTXTemplater.load(buffer);

// Load from remote HTTP/HTTPS URL
const pptFromUrl = await PPTXTemplater.load('https://example.com/templates/deck.pptx');

// Ingest cloud storage stream or buffer directly
const pptFromCloud = await PPTXTemplater.load({ provider: 's3', stream: s3Stream });

// Ingest Microsoft Flat OPC XML package (single-file XML, zero ZIP overhead)
const pptFromFlatXml = await PPTXTemplater.load('./sample.xml');

// Ingest legacy binary PPT or OpenDocument ODP with automatic normalization
const pptLegacy = await PPTXTemplater.load('./archive.ppt');

// Ingest unpacked OpenXML directory directly into memory
const pptFromDir = await PPTXTemplater.load('./extracted_folder');

// Create brand-new, template-free blank presentation with custom aspect ratio
const blankDeck = await PPTXTemplater.createBlank({ aspectRatio: '16:9' });

// High-concurrency template caching (<1ms deep clone)
await PPTXTemplater.preload('./template.pptx');
const cachedPpt = await PPTXTemplater.fromCache('./template.pptx');
PPTXTemplater.clearCache();
```

### Slide Context & Duplication
```javascript
ppt.useSlide(1);               // Target 1-based slide 1
ppt.useAllSlides();            // Target all slides in deck
await ppt.duplicateSlide(1, 2);// Clone slide 1 and insert at position 2
ppt.reorderSlides([2, 3, 1]);  // Permute slide order
await ppt.importSlideFrom(sourceDeck, 1); // Cross-deck slide import
```

### Text & Placeholders
```javascript
ppt.replaceTextByTag('{{name}}', 'Jane Doe', { caseSensitive: true });
ppt.replaceMultiple({ '{{company}}': 'Acme Inc.', '{{quarter}}': 'Q3' });
const matches = ppt.findText('{{');
ppt.appendTextRun('Shape1', ' [Appended Text]', { style: { bold: true, color: '#059669' } });
```

### Table Operations & Auto-Pagination
```javascript
ppt.updateTable('Table1', [['Col 1', 'Col 2'], ['Val 1', 'Val 2']]);
ppt.addTableRow('Table1', ['Region', ['Q1', 'Q2'], '$10,000'], { mergeStrategy: 'rowspan' });
ppt.stripeRows('Table1', { oddColor: 'F8FAFC', evenColor: 'EDF2F7', startRow: 1 });
await ppt.splitTableOnOverflow('Table1', { bottomMargin: '0.5in', repeatHeaders: 1, summaryRows: 1 });
const rows = await ppt.getTableRows('Table1', { includeMetadata: true });
```

### Native Chart Updates
```javascript
await ppt.updateChart('RevenueChart', {
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'Direct Sales', values: [150, 185, 220, 290] },
    { name: 'Channel', values: [85, 110, 140, 195] }
  ]
});
ppt.updateDataLabels('RevenueChart', { showVal: true, showCatName: false });
```

### Shapes & Z-Order
```javascript
await ppt.addShape('roundedRectangle', {
  name: 'KpiCard', x: 50, y: 100, width: 400, height: 200, fill: '#1E293B'
});
ppt.bringToFront('KpiCard');
ppt.alignShapes(['ShapeA', 'ShapeB'], 'middle');
ppt.distributeShapes(['ShapeA', 'ShapeB', 'ShapeC'], 'horizontally');
```

### Media & Images
```javascript
await ppt.addImage('./assets/logo.png', { x: 100, y: 100, width: 150, height: 150 });
await ppt.replaceImage('demo-logo', './assets/new-logo.png');
```

### Serialization & Multi-Target Export
```javascript
// Local file write
await ppt.saveToFile('./output/report.pptx');

// In-memory Buffer for Lambda / Cloud Run / S3
const buffer = await ppt.toBuffer();

// Node.js stream for piping
const stream = await ppt.toStream();

// Web-standard ArrayBuffer for Edge runtimes (Cloudflare Workers, Vercel Edge)
const arrayBuffer = await ppt.toArrayBuffer();

// Standard Base64 string or Data URI
const dataUri = await ppt.toBase64({ dataUri: true });

// Export to Microsoft Flat OPC XML string
const flatXml = await ppt.toFlatXml();

// Direct Express.js HTTP download response
await ppt.toExpress(res, 'report.pptx');

// Direct Fastify HTTP reply download
await ppt.toFastify(reply, 'report.pptx');

// Standard Web API Response for Next.js Route Handlers, Remix, SvelteKit
return await ppt.toResponse({ filename: 'report.pptx' });

// Extract raw compressed ZIP archive
await ppt.extractToZip('./backups/deck.zip');

// Unpack presentation into folder with pretty-printed XML for Git inspection
await ppt.extractToXml('./unpacked-deck', { formatXml: true });
```

---

## 5. Performance & Profiling Methods

- **Sub-Second Scaling:** Benchmarked to process 200 concurrent shape/text replacements in `< 400ms`.
- **Cached Memory Loading:** Pre-parsed cache loading is `> 10x` faster than disk/ZIP decompression.
- **Concurrent Safety:** Supports 50+ concurrent slide generation instances from the same cached template without cross-talk.
- **Benchmarking Suite:** Run benchmarks via `npm run benchmark` (`node benchmarks/run.js`).
- **Comprehensive Showcase:** Run all 12 feature scenarios with `npm run examples` (`node showcase/examples.js`).

---

## 6. Next.js Documentation Portal (`docs-site`)

- **Architecture & Deployment:** Transitioned from legacy static HTML to Next.js (App Router) configured for pure Static HTML Export (`docs-site/`, `next.config.mjs` with `output: 'export'`, `images: { unoptimized: true }`, `trailingSlash: true`). Deployed natively to Cloudflare Pages (`wrangler.toml` pointing to `pages_build_output_dir = "out"`).
- **Design System & UI Stack:** Modern Tailwind CSS Glassmorphism theme (`glass-card`, `glass-header`, `glass-sidebar`, ambient mesh gradients). Native Light and Dark modes with smooth transitions, system preference detection, and `localStorage` persistence via `ThemeContext` and `ThemeToggle`.
- **Motion & Interaction Engine:** GSAP animation engine (`src/animations/gsapTransitions.js`) powering staggered hero typography reveals (`animateHero`), feature grid reveals (`animateFeaturesGrid`), and smooth API card category transitions (`animateApiCards`).
- **Fixed Always-on-Top Sidebar:** Dedicated `DocsSidebar` with fixed viewport positioning (`top-16 left-0 h-[calc(100vh-4rem)] w-72 z-40 overflow-y-auto`), non-scrolling layout locking, and active route tracking via `usePathname`.
- **Site Route Structure:**
  - **Header Navigation:** Home (`/`), Docs (`/docs`), Features (`/features`), and API Reference (`/docs/api`), with Theme Toggle and GitHub star link.
  - **Docs Section (`/docs/*` with `DocsLayout` & `DocsSidebar`):**
    - 📌 **Overview** (`/docs`): Library philosophy, visual PowerPoint template mutation vs legacy code canvas assembly.
    - 🚀 **Installation Guides** (`/docs/installation`): Package managers (npm, pnpm, yarn, bun), Node.js >= 18 requirements, ESM/CJS, and native TypeScript types.
    - ⚡ **Compatibility Matrix** (`/docs/compatibility`): Zero Office/Electron runtime guarantees, cold boot benchmarks (<15ms, <2MB), AWS Lambda and Cloudflare Workers streaming recipes.
    - 📚 **API Reference** (`/docs/api`): 103 public methods across all 10 domain categories with live search, category filters, parameter tables, and copyable snippets.
  - **Features Catalog (`/features`):** Full 96 itemized capabilities across 12 domain categories with interactive search, category filters, and flagship indicators.
- **Content Decoupling & Ingestion:**
  - Content source: Type-safe, modular JSON schemas in `docs-site/src/data/`: `features.json` (96 features), `apiReference.json` (103 methods), `docsNavigation.json`, `compatibility.json`, `guides.json`, and `siteMeta.json`.
  - Sitemaps & Robots automated via Next.js metadata handlers (`sitemap.js` and `robots.js`).


