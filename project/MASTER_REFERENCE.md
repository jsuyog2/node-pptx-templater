# Master Reference & Architectural State: `node-pptx-templater`

## 1. Executive Summary & Core Philosophy
`node-pptx-templater` is a low-level, high-throughput OpenXML & Flat OPC PowerPoint template engine and presentation generator built in pure JavaScript.
- **Runtime Compatibility**: Node.js (`>= 18.0.0`, 20.x, 22.x LTS) and serverless edge environments (AWS Lambda, Cloudflare Workers, Vercel Edge).
- **Zero Native / Office Dependencies**: Uses pure JS (`jszip`, `fast-xml-parser`, `mime-types`). Zero reliance on Microsoft Office, LibreOffice (`soffice`), Electron, Java, or native C/C++ bindings (`node-gyp`, `canvas`, `sharp`).
- **Visual Template Paradigm**: Designers build slides visually in PowerPoint, Google Slides, or Keynote using `{{placeholders}}` and named objects; developers populate text, tables, charts, shapes, and images programmatically via a fluent, chainable API.
- **Silent-by-Default Telemetry**: Internal operations route through a configurable logger (`silent` by default; supports `error`, `warn`, `info`, `debug`, `verbose`).

---

## 2. Core Architecture & Sub-Manager Ecosystem
The central orchestrator class `PPTXTemplater` (`src/core/PPTXTemplater.js`) coordinates an in-memory Virtual File System (VFS) and specialized domain managers:

1. **`PackageLoader` & `VirtualFileSystem` (VFS)**:
   - Normalizes internal package paths (`\` to `/`, strips leading `/`, blocks `../` Zip-Slip traversal).
   - Polymorphic ingestion supporting PKZip `.pptx` archives, uncompressed XML folders, Buffers, Streams, and single-file Microsoft Office Flat OPC XML (`<pkg:package>`) files.
2. **`ZipManager`**:
   - Handles in-memory decompression and archiving of Open Packaging Conventions (OPC) containers via `jszip`.
3. **`XMLParser`**:
   - Hybrid `fast-xml-parser` engine utilizing `preserveOrder: true` where DrawingML schemas mandate strict child element ordering (`<p:spTree>`, `<a:tcPr>`, `<a:pPr>`).
4. **`TemplateEngine` & `TextManager`**:
   - Performs AST-level run fragment healing across split `<a:r>` and `<a:t>` nodes (recombining broken tags like `{{` + `company` + `}}` into the leading run while preserving `<a:rPr>` typography), and manages lists and text styling.
5. **`SlideManager`**:
   - Handles slide targeting (`useSlide`, `useAllSlides`), cloning, insertion, deletion, reordering, standalone export, and cross-presentation slide importing with media/layout deduplication.
6. **`TableManager`**:
   - Manages cell updates, dynamic row insertion/cloning, 32-bit unsigned `<a16:rowId>` generation, rectangular cell merges (`gridSpan`, `rowSpan`, `hMerge`, `vMerge`), auto-fitting, overflow pagination, and JSON extraction (`getTableRows`).
7. **`ChartManager` & `ChartWorkbookUpdater`**:
   - Dual-syncs visual DrawingML caches (`ppt/charts/chartX.xml` `<c:strCache>`, `<c:numCache>`, `<c:ptCount>`) and the underlying embedded Excel workbook (`ppt/embeddings/*.xlsx` `xl/worksheets/sheet1.xml` & `xl/sharedStrings.xml`), plus custom data labels.
8. **`ShapeManager` & `CellShapeManager`**:
   - Manages EMU/pixel geometry, shape creation/cloning, alignments, and table cell overlay badges (`addCellShape`, `updateCellShape`) without triggering table layout reflow.
9. **`ZOrderManager`**:
   - Controls visual stacking order (`bringToFront`, `sendToBack`, `bringForward`, `sendBackward`, `setZIndex`, `swapObjects`) by reordering nodes in `<p:spTree>`.
10. **`MediaManager`, `ImageManager` & `HyperlinkManager`**:
    - SHA-1 media deduplication, aspect-ratio preserving image swaps, external URL links, and internal slide navigation actions.
11. **`RelationshipManager` & `ContentTypesManager`**:
    - Maintains `.rels` dependency graphs (`rId` uniqueness, `slideId >= 256`) and `[Content_Types].xml` MIME overrides.
12. **`OutputWriter`**:
    - Serializes modified VFS state to files, Buffers, Streams, Base64, ArrayBuffers, Web Responses, Express/Fastify replies, uncompressed XML folders, and Flat OPC XML.

---

## 3. Strict 18-Method I/O API Boundary
All deprecated import/export aliases have been removed. The library strictly enforces these **18 approved I/O methods** (demonstrated in `showcase/input_export.js`):

### Static Ingestion, Build & Cache API (7 Methods)
1. `PPTXTemplater.load(source, options)`: Loads `.pptx`, Flat OPC `.xml` (`<pkg:package>`), unpacked XML directory, Buffer, Uint8Array, Stream, or URL into the VFS.
2. `PPTXTemplater.createBlank(options)`: Generates a schema-compliant blank presentation in memory (16:9, 4:3, or custom dimensions) with preloaded VFS state and zero PowerPoint repair warnings.
3. `PPTXTemplater.extractPptx(pptxPath, outputDir, options)`: Extracts a `.pptx` archive into a version-control-friendly directory of raw XML files.
4. `PPTXTemplater.buildPptx(sourceDir, outputPath, options)`: Compiles an unpacked XML directory tree back into a `.pptx` binary.
5. `PPTXTemplater.preload(source, options)`: Parses and caches a template in static memory.
6. `PPTXTemplater.fromCache(source)`: Clones a preloaded template from memory in `< 1ms` for high-concurrency servers.
7. `PPTXTemplater.clearCache()`: Purges all cached templates from static memory.

### Instance Export Methods (8 Methods)
8. `ppt.saveToFile(destinationPath, options)`: Commits AST changes and writes the `.pptx` file to disk.
9. `ppt.toBuffer(options)`: Serializes the presentation into a Node.js `Buffer`.
10. `ppt.toStream(options)`: Emits a `ReadableStream`.
11. `ppt.toBase64(options)`: Outputs a Base64 string or Data URI.
12. `ppt.toArrayBuffer(options)`: Outputs a standard `ArrayBuffer` for Edge runtimes.
13. `ppt.toResponse(options)`: Returns a standard Web Fetch API `Response` object for Next.js Route Handlers and Cloudflare Workers.
14. `ppt.toExpress(res, filename, options)`: Pipes the presentation directly to an Express.js `res` object with OpenXML MIME and `Content-Disposition` headers.
15. `ppt.toFastify(reply, filename, options)`: Sends the presentation via a Fastify `reply` object with configured headers.

### Instance Extraction & Conversion Methods (3 Methods)
16. `ppt.extractToZip(destinationPath, options)`: Writes the raw compressed ZIP package to disk.
17. `ppt.extractToXml(destinationDir, options)`: Unpacks all in-memory presentation XML parts into a folder.
18. `ppt.toFlatXml(destinationPathOrOptions, options)`: Serializes the presentation into a single-file Microsoft Office Flat OPC XML document (`<pkg:package>`).

---

## 4. OpenXML Integrity, Security & Corruption Prevention
- **Flat OPC XML Ingestion & Serialization**: Detects `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">`, strips leading slashes from `pkg:name`, extracts `<pkg:xmlData>` as UTF-8 strings, and base64-decodes `<pkg:binaryData>` (images and `.xlsx` files) into `Uint8Array` buffers in the VFS.
- **`createBlank()` Repair Prevention**: Ensures `[Content_Types].xml` includes `<Default>` entries for `rels`/`xml` and `<Override>` entries for `presentation.xml`, `slideMaster1.xml`, `slideLayout1.xml`, `theme1.xml`, and `slide1.xml`; wires bidirectional `.rels` between slide, layout, and master; assigns `slideId="256"`; and preloads VFS caches upon initialization.
- **XML Security Guardrails**: Rejects `<!DOCTYPE>` and `<!ENTITY>` declarations to block XXE and Billion Laughs exponential entity expansion attacks; blocks `../` path traversal in ZIP and VFS entries.

---

## 5. Documentation Portal (`docs-site/`) & SEO Architecture
- **Framework**: Next.js (App Router) inside `docs-site/` configured with `output: 'export'` and `images: { unoptimized: true }` for static hosting on Cloudflare Pages (the legacy `/docs` folder is permanently removed).
- **Design System**: Tailwind CSS Glassmorphism UI (`backdrop-blur-xl`), Light & Dark Mode with `localStorage` persistence (`ThemeContext.jsx`), and GSAP animations (`gsapTransitions.js`).
- **Navigation Layout**:
  - **Sticky Glass Header**: `Home` (`/`), `Docs` (`/docs`), and `Features` (`/features`).
  - **Fixed Non-Scrolling Docs Sidebar (`DocsSidebar.jsx`)**: Anchored (`fixed top-16 left-0 h-[calc(100vh-4rem)]`) with `Overview` (`/docs`), `Installation Guides` (`/docs/installation`), `Compatibility` (`/docs/compatibility`), and `API Reference` (`/docs/api`).
- **Decoupled JSON Content (`docs-site/src/data/`)**:
  - `features.json` (synced from `project/Features.md`), `apiReference.json` (all 10 categories + 18 I/O methods), `docsNavigation.json`, `compatibility.json`, and `siteMeta.json`.
- **SEO & Social Previews**: Automated `sitemap.js` and `robots.js`, Schema.org `SoftwareApplication` JSON-LD, OpenGraph (`1200x630` banner for WhatsApp/LinkedIn/Discord), and Twitter Summary Large Image cards.

---

## 6. `.agent/` Autonomous Workspace Layout
The repository maintains its AI agent context inside `.agent/`:
- `.agent/AGENTS.md`: Master agent roster (`architecture-reviewer`, `openxml-debugger`, `docs-architect`, `qa-validator`) and core project mandates.
- `.agent/.agentignore` & `.agent/mcp_config.json`: Context exclusion rules and MCP tool definitions.
- `.agent/project_summary.md`: Architectural overview and subsystem map.
- `.agent/rules/`: `general.md`, `features.md`, and `io-whitelist.md`.
- `.agent/skills/`: Playbooks for `openxml-editing.md`, `chart-sync.md`, `table-merging.md`, `flat-opc-xml.md`, and `docs-site-maintenance.md`.
- `.agent/memory/architectural-decisions.md`: ADR history.
- `.agent/context/domain-dictionary.md`: OpenXML/DrawingML/OPC terminology.
- `.agent/workflows/`: SOPs for `add-new-feature.md`, `fix-xml-corruption.md`, and `build-docs.md`.
