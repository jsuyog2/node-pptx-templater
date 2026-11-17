# Technical Reference & Knowledge Base: `node-pptx-templater`

> **Comprehensive System Documentation for Gemini Notebook**  
> **Package:** `node-pptx-templater` (v3.0.0-alpha.0)  
> **Source Repository:** [github.com/jsuyog2/node-pptx-templater](https://github.com/jsuyog2/node-pptx-templater)  
> **Runtime Target:** Node.js (>= 18.0.0, tested on 20.x, 22.x LTS), AWS Lambda, Cloudflare Workers, Vercel Edge  
> **Module Distribution:** Dual ECMAScript Module (`dist/index.mjs`) & CommonJS (`dist/index.cjs`), native `.d.ts` types

---

## Table of Contents

1. [Library Overview & Execution Model](#1-library-overview--execution-model)
   - [1.1 Purpose & Core Philosophy](#11-purpose--core-philosophy)
   - [1.2 Visual Template Design Paradigm](#12-visual-template-design-paradigm)
   - [1.3 Runtime Targets & Environment Compatibility](#13-runtime-targets--environment-compatibility)
   - [1.4 Dependency Boundaries & Pure JavaScript Mandate](#14-dependency-boundaries--pure-javascript-mandate)
   - [1.5 Open Packaging Conventions (OPC) Anatomy](#15-open-packaging-conventions-opc-anatomy)
2. [Core Architecture & Sub-Manager Subsystem](#2-core-architecture--sub-manager-subsystem)
   - [2.1 Orchestrator: PPTXTemplater](#21-orchestrator-pptxtemplater)
   - [2.2 Ingestion & VFS: PackageLoader & VirtualFileSystem](#22-ingestion--vfs-packageloader--virtualfilesystem)
   - [2.3 Package Compression: ZipManager](#23-package-compression-zipmanager)
   - [2.4 XML Parsing: XMLParser & Order Preservation](#24-xml-parsing-xmlparser--order-preservation)
   - [2.5 Slide Management: SlideManager & Relationship Graphs](#25-slide-management-slidemanager--relationship-graphs)
   - [2.6 Table Engine: TableManager](#26-table-engine-tablemanager)
   - [2.7 Native Chart Engine: ChartManager & Workbook Dual-Sync](#27-native-chart-engine-chartmanager--workbook-dual-sync)
   - [2.8 Spatial Engine: ShapeManager, CellShapeManager & Units](#28-spatial-engine-shapemanager-cellshapemanager--units)
   - [2.9 Layer Stacking: ZOrderManager](#29-layer-stacking-zordermanager)
   - [2.10 Typography & Text: TemplateEngine & TextManager](#210-typography--text-templateengine--textmanager)
   - [2.11 Media & Assets: MediaManager & HyperlinkManager](#211-media--assets-mediamanager--hyperlinkmanager)
   - [2.12 Serialization & Output: OutputWriter](#212-serialization--output-outputwriter)
3. [OpenXML Corruption Prevention & Security Guardrails](#3-openxml-corruption-prevention--security-guardrails)
   - [3.1 AST-Level Run Fragment Healing](#31-ast-level-run-fragment-healing)
   - [3.2 Native Chart & Excel Spreadsheet Dual-Sync](#32-native-chart--excel-spreadsheet-dual-sync)
   - [3.3 Table Matrix Consistency & 32-Bit Unsigned `<a16:rowId>`](#33-table-matrix-consistency--32-bit-unsigned-a16rowid)
   - [3.4 Rectangular Cell Merging (Origin Spans & Covered Cells)](#34-rectangular-cell-merging-origin-spans--covered-cells)
   - [3.5 XML Security (XXE, Billion Laughs, DTD Entity Guardrails)](#35-xml-security-xxe-billion-laughs-dtd-entity-guardrails)
4. [Complete API & Method Specification](#4-complete-api--method-specification)
   - [4.1 Lifecycle, Ingestion & In-Memory Preloading](#41-lifecycle-ingestion--in-memory-preloading)
   - [4.2 Slide Management & Coordinate Sizing](#42-slide-management--coordinate-sizing)
   - [4.3 Text Manipulation & Placeholder Templating](#43-text-manipulation--placeholder-templating)
   - [4.4 Table Operations, Merging & Auto-Pagination](#44-table-operations-merging--auto-pagination)
   - [4.5 Native Chart Updates, Data Labels & Positioning](#45-native-chart-updates-data-labels--positioning)
   - [4.6 Shapes, Spatial Geometry & Alignment](#46-shapes-spatial-geometry--alignment)
   - [4.7 Table Cell Shapes (Non-Reflow Overlay Badging)](#47-table-cell-shapes-non-reflow-overlay-badging)
   - [4.8 Media, Images & Aspect-Ratio Preservation](#48-media-images--aspect-ratio-preservation)
   - [4.9 Hyperlinks & Interactive Navigation](#49-hyperlinks--interactive-navigation)
   - [4.10 Layer Stacking & Visual Z-Order](#410-layer-stacking--visual-z-order)
   - [4.11 Inspection, Diagnostics, Schema Validation & Repair](#411-inspection-diagnostics-schema-validation--repair)
   - [4.12 Export Adapters & Multi-Target Serialization](#412-export-adapters--multi-target-serialization)
5. [CLI, Logging, Tooling & Performance Profiling](#5-cli-logging-tooling--performance-profiling)
   - [5.1 Command-Line Interface (CLI) Commands](#51-command-line-interface-cli-commands)
   - [5.2 Logging Subsystem & Runtime Telemetry](#52-logging-subsystem--runtime-telemetry)
   - [5.3 High-Throughput Preloading & Template Caching](#53-high-throughput-preloading--template-caching)
   - [5.4 Built-In Performance Profiling & Metrics](#54-built-in-performance-profiling--metrics)

---

# 1. Library Overview & Execution Model

## 1.1 Purpose & Core Philosophy
`node-pptx-templater` is a low-level, high-throughput OpenXML PowerPoint template manipulation and programmatic generation engine written exclusively in pure JavaScript.

The library was conceived to eliminate the two primary historical failure modes of programmatic presentation generation in the Node.js ecosystem:
1. **The "Code-as-Canvas" Trap**: Constructing slide geometry, tables, corporate branding, themes, typography, and vector palettes purely imperatively in code (e.g., specifying coordinates for every title, bullet point, margin, and border). This approach is brittle, inaccessible to non-developers, and turns simple brand adjustments into multi-day engineering refactors.
2. **The "Headless Office Runtime" Trap**: Relying on headless instances of LibreOffice (`soffice`), Microsoft Office Automation, unoconv, or Electron/Puppeteer. These systems suffer from high startup latency (often seconds per invocation), excessive memory consumption (hundreds of megabytes to gigabytes), concurrency contention, and vulnerability to catastrophic segmentation faults or process hangs inside containerized serverless runtimes.

`node-pptx-templater` resolves both challenges by adopting an **in-memory Open Packaging Conventions (OPC) Virtual File System (VFS)** that ingests pre-designed `.pptx` decks as templates, parses their underlying DrawingML, PresentationML, and ChartML abstract syntax trees (AST), applies surgical mutations, and serializes pristine packages in sub-millisecond execution cycles.

```
┌─────────────────────────────────────────────────────────────────────────────────────────┐
│                                   DEVELOPER / DESIGNER                                  │
│   Designs master layouts, brand typography, vector shapes, tables & charts visually     │
│   in Microsoft PowerPoint, Google Slides, or Apple Keynote with {{placeholders}}       │
└────────────────────────────────────────────┬────────────────────────────────────────────┘
                                             │ Saves template.pptx
                                             ▼
┌─────────────────────────────────────────────────────────────────────────────────────────┐
│                           node-pptx-templater (Pure JS Engine)                          │
│                                                                                         │
│  ┌───────────────────────┐   ┌───────────────────────────┐   ┌────────────────────────┐ │
│  │ PackageLoader & VFS   │──►│ Abstract Syntax Tree (AST)│──►│ OutputWriter           │ │
│  │ In-memory ZIP package │   │ Run Healing, Dual-Sync,   │   │ Stream, Buffer, Disk,  │ │
│  │ Zero disk footprint   │   │ <a16:rowId>, EMU math     │   │ Folder, Express Pipe   │ │
│  └───────────────────────┘   └───────────────────────────┘   └────────────────────────┘ │
└────────────────────────────────────────────┬────────────────────────────────────────────┘
                                             │ Emits pristine OpenXML binary
                                             ▼
┌─────────────────────────────────────────────────────────────────────────────────────────┐
│                                   END CONSUMERS                                         │
│   • Microsoft PowerPoint (Windows, macOS, Web, iOS, Android) [No "Repair" Dialogs]      │
│   • Google Slides & Apple Keynote [100% Geometry Fidelity]                             │
│   • LibreOffice Impress & Apache OpenOffice [Standard OPC Compliance]                   │
└─────────────────────────────────────────────────────────────────────────────────────────┘
```

## 1.2 Visual Template Design Paradigm
Rather than defining presentations programmatically from a blank canvas, developers design baseline presentations visually using their presentation tool of choice:
- **Placeholders**: Text frames and shapes can contain tag tokens such as `{{company_name}}`, `{{quarter}}`, or `{{kpi_value}}`.
- **Pre-Styled Tables**: Tables are created with desired column widths, corporate header fills, font styles, and padding. The engine populates rows, handles rectangular cell merges, and auto-paginates table data across slides when rows overflow.
- **Pre-Styled Native Charts**: Bar, Column, Line, Pie, and Area charts are inserted directly into the master deck with corporate color palettes, axis typography, and legends pre-configured. At runtime, the engine mutates data points and series values without modifying the visual styling.
- **Shape & Image Anchors**: Static or placeholder shapes act as bounding boxes. At runtime, images are swapped into place while strictly maintaining aspect ratios, or vector shapes are aligned dynamically.

## 1.3 Runtime Targets & Environment Compatibility
`node-pptx-templater` is built to run anywhere modern JavaScript executes:
- **Node.js**: Minimum supported version is `Node.js >= 18.0.0` LTS; officially tested and validated across Node.js 20.x and Node.js 22.x LTS.
- **Serverless & Edge Platforms**:
  - **AWS Lambda**: Sub-50ms cold starts, zero `/tmp` disk footprint required (fully in-memory operations).
  - **Cloudflare Workers & Vercel Edge**: Compatible with web-standard `ArrayBuffer`, `ReadableStream`, and `Uint8Array` abstractions.
  - **Google Cloud Functions & Azure Functions**: Efficient multi-tenant concurrent execution without process locks.
- **Module Distribution**:
  - **CommonJS**: `dist/index.cjs` (via `require('node-pptx-templater')`).
  - **ES Modules**: `dist/index.mjs` (via `import { PPTXTemplater } from 'node-pptx-templater'`).
  - **TypeScript Declarations**: Fully typed public contracts in `src/index.d.ts` and `dist/index.d.ts`.

## 1.4 Dependency Boundaries & Pure JavaScript Mandate
The package enforces strict dependency hygiene:
- **Zero Native Addons**: No compiled C/C++ bindings (e.g., `node-gyp`, `@napi-rs`, `canvas`, or `sharp`).
- **Zero External Office Binaries**: Never spawns external processes or relies on system-installed office suites.
- **Core Production Dependencies**:
  - `jszip`: In-memory de-archiving and archiving of ZIP containers.
  - `fast-xml-parser`: High-throughput bidirectional XML parsing and serialization with attribute and tag order preservation.
  - `mime-types`: Precise MIME type resolution for embedded raster and vector media.

## 1.5 Open Packaging Conventions (OPC) Anatomy
A `.pptx` presentation is a compressed ZIP archive conforming to ECMA-376 and ISO/IEC 29500 (Open Packaging Conventions). Understanding this physical structure is essential when working with the low-level engine:

```
[Content_Types].xml               <-- Global MIME type registry for every part
_rels/
  └── .rels                       <-- Package-level relationships (points to ppt/presentation.xml)
docProps/
  ├── app.xml                     <-- Application metadata (slides count, app name)
  └── core.xml                    <-- Dublin Core metadata (author, created, modified)
ppt/
  ├── presentation.xml            <-- Slide master registry, slide list (<p:sldIdLst>), dimensions
  ├── _rels/
  │   └── presentation.xml.rels   <-- Maps slide rIds, slideMasters, themes to part URIs
  ├── slides/
  │   ├── slide1.xml              <-- Slide 1 Shape Tree (<p:spTree>), text, tables, graphic frames
  │   ├── slide2.xml
  │   └── _rels/
  │       ├── slide1.xml.rels     <-- Maps slide 1 relationships (charts, images, hyperlinks, layout)
  │       └── slide2.xml.rels
  ├── slideLayouts/
  │   ├── slideLayout1.xml        <-- Slide layout geometry
  │   └── _rels/
  ├── slideMasters/
  │   ├── slideMaster1.xml        <-- Master slide definitions (global themes, header/footer anchors)
  │   └── _rels/
  ├── theme/
  │   └── theme1.xml              <-- Color palettes, font schemes (major/minor fonts)
  ├── charts/
  │   ├── chart1.xml              <-- DrawingML chart markup (<c:chartSpace>, series, categories)
  │   └── _rels/
  │       └── chart1.xml.rels     <-- Points to embedded Excel workbook
  ├── embeddings/
  │   └── Microsoft_Excel_Worksheet1.xlsx <-- Backing spreadsheet binary for chart1
  └── media/
      ├── image1.png              <-- Raw media assets (PNG, JPEG, SVG, WebP)
      └── image2.jpeg
```

### Critical Relationship Rules:
1. **Explicit Part Names**: Every part inside `ppt/slides/` must have an explicit `<Override>` entry in `[Content_Types].xml` specifying `ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"`.
2. **Relationship Chains**: A slide element cannot reference a file path directly; it references a relationship identifier (e.g. `r:embed="rId2"` or `r:id="rId5"`). That identifier MUST exist inside `ppt/slides/_rels/slideX.xml.rels` and resolve to the target part.
3. **Double-Linked Presentation Tree**: Slides must be declared both in `ppt/_rels/presentation.xml.rels` (with a unique `rId`) and in `ppt/presentation.xml` within `<p:sldIdLst>` (with a unique `id >= 256`).


---

# 2. Core Architecture & Sub-Manager Subsystem

The architectural layout of `node-pptx-templater` follows a modular orchestrator-subsystem pattern. The public facade (`PPTXTemplater`) delegates domain operations to dedicated managers, ensuring separation of concerns while maintaining a fluent, chainable developer interface.

```
                                  ┌───────────────────────────┐
                                  │       PPTXTemplater       │
                                  │       (Orchestrator)      │
                                  └─────────────┬─────────────┘
                                                │
         ┌───────────────────┬──────────────────┼──────────────────┬──────────────────┐
         ▼                   ▼                  ▼                  ▼                  ▼
  ┌──────────────┐    ┌──────────────┐   ┌──────────────┐   ┌──────────────┐   ┌──────────────┐
  │ZipManager &  │    │  XMLParser   │   │ SlideManager │   │ TableManager │   │ ChartManager │
  │PackageLoader │    │(Order-Aware) │   │ & Rels Graph │   │& Row Auto-Fit│   │& Excel Sync  │
  └──────────────┘    └──────────────┘   └──────────────┘   └──────────────┘   └──────────────┘
         │                   │                  │                  │                  │
         ▼                   ▼                  ▼                  ▼                  ▼
  ┌──────────────┐    ┌──────────────┐   ┌──────────────┐   ┌──────────────┐   ┌──────────────┐
  │ ShapeManager │    │ZOrderManager │   │TemplateEngine│   │ MediaManager │   │ OutputWriter │
  │ & Cell Shapes│    │(<p:spTree>)  │   │& TextManager │   │& Hyperlinks  │   │ (Multi-Pipe) │
  └──────────────┘    └──────────────┘   └──────────────┘   └──────────────┘   └──────────────┘
```

---

## 2.1 Orchestrator: PPTXTemplater
The `PPTXTemplater` class (`src/core/PPTXTemplater.js`) serves as the primary coordination hub.
- **Fluent API Design**: Every state-mutating method returns `this`, enabling chained operations:
  ```javascript
  const ppt = await PPTXTemplater.load('template.pptx');
  await ppt
    .useSlide(1)
    .replaceText({ '{{title}}': 'Executive Briefing' })
    .updateTable('Table1', rows)
    .saveToFile('output.pptx');
  ```
- **Context Management**: Tracks selected slide context via `#selectedSlides` (supporting single indices `1`, arrays `[1, 3]`, or `'all'`).
- **Lifecycle Guarantees**: Coordinates lazy parsing of slide XML files upon selection, tracks dirty parts, coordinates sub-manager commits, and ensures proper cleanup on output generation.

---

## 2.2 Ingestion & VFS: PackageLoader & VirtualFileSystem
The ingestion layer (`src/core/PackageLoader.js` and `src/core/VirtualFileSystem.js`) abstracts away the physical storage of presentation parts.
- **`VirtualFileSystem` (VFS)**:
  - An in-memory key-value store mapping normalized path strings (`ppt/slides/slide1.xml`) to raw binary `Uint8Array` chunks or UTF-8 strings.
  - Normalizes Windows backslashes (`\`) to OpenXML forward slashes (`/`), strips leading `./`, and disallows directory traversal attacks (`../`).
  - Supports atomic snapshotting and cloning for high-concurrency caching.
- **`PackageLoader` Polymorphism**:
  Detects the input type dynamically:
  1. **Filesystem Path (`string`)**: Resolves absolute path. Checks if target is a file or a folder. If a folder, ingests unpacked XML files directly.
  2. **In-Memory `Buffer` or `Uint8Array`**: Inspects magic numbers (`0x50 0x4B 0x03 0x04` for PKZip). Decompresses package contents directly into the VFS.
  3. **Node.js `ReadableStream` or Web Stream**: Buffers incoming byte stream in memory before inflating.
  4. **Blank Presentation**: When initialized via `PPTXTemplater.create()`, loads an embedded minimal valid presentation bundle (`src/templates/blankPptx.js`).

---

## 2.3 Package Compression: ZipManager
The `ZipManager` (`src/managers/ZipManager.js`) manages the underlying JSZip instance.
- **Direct Part Access**: Exposes asynchronous file read/write methods (`readText(path)`, `writeText(path, text)`, `readBinary(path)`, `writeBinary(path, buffer)`).
- **Compression Tuning**: Exposes configurable compression levels (`STORE` for maximum speed in local testing, `DEFLATE` with levels 1–9 for production balance between CPU and payload size).
- **Binary Signature Verification**: Includes `isPkZip()` to validate package integrity prior to decompression, throwing descriptive errors on corrupted inputs.

---

## 2.4 XML Parsing: XMLParser & Order Preservation
OpenXML markup is strictly order-dependent. DrawingML and PresentationML XML schemas (XSD) mandate exact element ordering inside tags (for instance, `<p:spPr>` must precede `<p:txBody>`, and `<a:off>` must precede `<a:ext>` within `<a:xfrm>`).

`XMLParser` (`src/parsers/XMLParser.js`) solves this by utilizing `fast-xml-parser` with precise configuration:
- **`preserveOrder: false` (Standard Object Traversal)**: Used for reading element trees where fast key lookup is critical.
- **`preserveOrder: true` (Ordered Tuple AST)**: Used when mutating elements whose schemas strictly enforce element sequence.
- **Attribute Preservation**: Configured with `attributeNamePrefix: '@_'` and `ignoreAttributes: false` so that OpenXML namespaces (`xmlns:p`, `xmlns:a`, `xmlns:r`) and attributes (`r:id`, `val`, `sz`) remain intact.
- **Value Handling**: `parseTagValue: false` ensures phone numbers, hex strings, and numeric row IDs are not coerced to JavaScript numbers, which would strip leading zeroes or introduce precision loss.
- **Whitespace Retention**: `trimValues: false` preserves critical text run spacing in `<a:t>` tags.

---

## 2.5 Slide Management: SlideManager & Relationship Graphs
`SlideManager` (`src/managers/SlideManager.js`) manages slide lifecycles and presentation manifests.
- **Slide Selection (`useSlide`)**: Resolves 1-based slide numbers to internal part paths (`ppt/slides/slideN.xml`).
- **Deep Slide Cloning (`duplicateSlide`)**:
  1. Copies the target slide XML AST.
  2. Creates a new slide relationship file (`ppt/slides/_rels/slideM.xml.rels`), deep-cloning all referenced layout, master, image, and chart relationships.
  3. Allocates a unique `rId` in `ppt/_rels/presentation.xml.rels`.
  4. Registers the new slide in `ppt/presentation.xml` within `<p:sldIdLst>` with a newly allocated unique 32-bit slide ID (`id >= 256`).
  5. Registers the new slide part in `[Content_Types].xml`.
- **Cross-Deck Import (`importSlideFrom`)**:
  Imports a slide from an external presentation, deduplicating shared themes and layouts, generating fresh relationship IDs, and importing required media binaries without collisions.

---

## 2.6 Table Engine: TableManager
The table engine (`src/managers/TableManager.js` and `src/elements/TableManager.js`) provides high-level and low-level table manipulation:
- **Structure**: Targets `<a:tbl>` within `<p:graphicFrame>` elements.
- **Row Hydration (`updateTable`, `addTableRow`)**: Injects dynamic data into cells, auto-generating `<a:tc>`, `<a:txBody>`, and `<a:tcPr>` markup.
- **32-Bit Unsigned `<a16:rowId>` Injection**: Modern PowerPoint versions (Office 2016+) track table rows using unique 32-bit integers. `TableManager` automatically assigns unique `<a16:rowId>` tags to prevent PowerPoint "Repair" dialogs.
- **Rectangular Cell Merging (`mergeCells`)**: Calculates coordinates, adds `gridSpan` and `rowSpan` attributes to origin cells, and places shadowed `<a:tc hMerge="1">` / `<a:tc vMerge="1">` placeholder cells.
- **Dynamic Auto-Pagination (`splitTableOnOverflow`)**:
  Calculates cumulative row heights based on cell padding and text length. When a table exceeds the slide boundary:
  1. Duplicates the parent slide.
  2. Slices the table at the overflow boundary.
  3. Re-injects designated table headers onto subsequent slides.
  4. Cascades remaining rows across newly generated slides recursively.

---

## 2.7 Native Chart Engine: ChartManager & Workbook Dual-Sync
`ChartManager` (`src/managers/ChartManager.js`) coordinates visual DrawingML chart definitions and embedded spreadsheet binaries:
- **Dual-Sync Protocol**:
  Updating DrawingML (`ppt/charts/chartX.xml`) alone causes Microsoft Office to report corruptions or overwrite data when a user clicks "Edit Data in Excel".
  `ChartManager` delegates to `ChartWorkbookUpdater.js` to synchronously patch `ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx`, keeping categories, series values, and formula ranges (`Sheet1!$B$2:$B$10`) in complete parity.
- **Data Label Customization (`updateDataLabels`)**: Controls positioning (`bestFit`, `ctr`, `inBase`, `outEnd`), styling, and explicit value overrides for individual data points.
- **Chart Geometry & Layout**: Allows moving, resizing, and toggling axes, legends, and gridlines.

---

## 2.8 Spatial Engine: ShapeManager, CellShapeManager & Units
`ShapeManager` (`src/managers/ShapeManager.js`) and `CellShapeManager` govern spatial geometry.
- **OpenXML Unit Standard (EMU)**:
  All OpenXML spatial coordinates and dimensions are stored in English Metric Units (EMU).
  - $1 \text{ pt} = 12,700 \text{ EMUs}$
  - $1 \text{ inch} = 914,400 \text{ EMUs}$
  - $1 \text{ pixel (at 96 DPI)} = 9,525 \text{ EMUs}$
  - $1 \text{ cm} = 360,000 \text{ EMUs}$
  The `units.js` utility automatically converts human-readable strings (`'100px'`, `'2.5in'`, `'14pt'`) into integer EMUs, applying `Math.round()` to prevent invalid floating-point schema violations.
- **Cell Overlay Shapes (`CellShapeManager`)**:
  Enables embedding status badges, pills, and indicator shapes over specific table cells (`addCellShape`). Calculates exact cell bounding boxes (table offset + column offsets + row offsets) to place shapes in absolute slide space without causing table layout reflow.

---

## 2.9 Layer Stacking: ZOrderManager
Visual stacking order in PowerPoint is determined by the document order of child elements inside `<p:spTree>`:
- Elements rendered earlier in `<p:spTree>` appear beneath elements rendered later.
- `ZOrderManager` (`src/spatial/ZOrderManager.js` and `src/managers/ZOrderManager.js`) translates operations (`bringToFront`, `sendToBack`, `bringForward`, `sendBackward`, `setZIndex`) into node rearrangements in `<p:spTree>`, keeping non-visual properties (`<p:nvGrpSpPr>`, `<p:grpSpPr>`) pinned to the top.

---

## 2.10 Typography & Text: TemplateEngine & TextManager
Text operations are split between the template tokenizer and styling engine:
- **`TemplateEngine` (`src/core/TemplateEngine.js`)**: Scans paragraphs (`<a:p>`), identifies fragmented placeholder tags (`{{tag}}`) split across multiple `<a:r>` and `<a:t>` runs, coalesces them into the leading run while preserving `<a:rPr>` styling, and substitutes template values.
- **`TextManager` (`src/managers/TextManager.js`)**: Manages typography overrides (font family, font size in hundredths of a point, bold, italic, underline, color hex codes) and handles bulleted/numbered lists.

---

## 2.11 Media & Assets: MediaManager & HyperlinkManager
- **`MediaManager` (`src/managers/MediaManager.js`)**: Handles image ingestion (PNG, JPEG, SVG, WebP, GIF), assigns SHA-1 content hashes to deduplicate identical images, registers parts in `[Content_Types].xml`, and generates relationship mappings in `slideX.xml.rels`.
- **`HyperlinkManager` (`src/managers/HyperlinkManager.js`)**: Creates external URL relationships (`http://`, `mailto:`) and internal slide jump relationships (`ppaction://hlinksldjump`) and binds them to text runs or shapes via `r:id`.

---

## 2.12 Serialization & Output: OutputWriter
`OutputWriter` (`src/core/OutputWriter.js`) serializes the in-memory VFS into desired consumer formats:
- **File System (`toFile` / `saveToFile`)**: Asynchronously writes the ZIP container to disk.
- **Node.js `Buffer` (`toBuffer`)**: Returns raw in-memory Buffer for immediate serverless dispatch.
- **Standard `ArrayBuffer` (`toArrayBuffer`)**: Standardized binary format for Cloudflare Workers / browser fetch.
- **Base64 String (`toBase64`)**: Encodes binary with optional data URI prefix (`data:application/vnd.openxmlformats...;base64,...`).
- **Readable Stream (`toStream`)**: Emits chunked stream for piping to AWS S3 / Google Cloud Storage.
- **Express.js Pipe (`toExpress`)**: Pipes stream to `res` with auto-configured `Content-Type` and `Content-Disposition`.
- **Unpacked Folder (`saveToFolder` / `saveXml`)**: Dumps the raw OpenXML XML/media files to an uncompressed directory for version control inspection.


---

# 3. OpenXML Corruption Prevention & Security Guardrails

Microsoft PowerPoint enforces strict schema and relationship validation when opening presentations. A single missing relationship, non-integer coordinate, unclosed tag, or duplicate row ID causes PowerPoint to display its infamous warning:  
**"PowerPoint found a problem with content in presentation.pptx. PowerPoint can attempt to repair the presentation."**

`node-pptx-templater` incorporates five architectural guardrails that systematically eliminate presentation repair corruptions and protect against security vulnerabilities.

---

## 3.1 AST-Level Run Fragment Healing

### The Fragmentation Problem
When a user types a placeholder like `{{quarterly_revenue}}` into PowerPoint, PowerPoint's internal layout and spellchecking engine frequently fragments the text across multiple text runs (`<a:r>`) and text nodes (`<a:t>`).

For example, typing `{{quarterly_revenue}}` might be serialized in OpenXML as:
```xml
<a:p>
  <a:r>
    <a:rPr sz="1800" b="1"><a:solidFill><a:srgbClr val="003366"/></a:solidFill></a:rPr>
    <a:t>{{quar</a:t>
  </a:r>
  <a:r>
    <a:rPr sz="1800" b="1"><a:solidFill><a:srgbClr val="003366"/></a:solidFill></a:rPr>
    <a:t>terly_</a:t>
  </a:r>
  <a:r>
    <a:rPr sz="1800" b="1"><a:solidFill><a:srgbClr val="003366"/></a:solidFill></a:rPr>
    <a:t>revenue}}</a:t>
  </a:r>
</a:p>
```
A naïve regex or string replacement searching for `{{quarterly_revenue}}` will fail because the string does not exist contiguously in any individual `<a:t>` node.

### The Run Healing Algorithm
`TemplateEngine.js` resolves this at the Abstract Syntax Tree (AST) level:
1. **Paragraph Scanning**: Iterates over all paragraphs (`<a:p>`) inside text bodies (`<p:txBody>`).
2. **Text Run Tokenization**: Assembles a contiguous text string of the paragraph while tracking character index boundaries mapping back to individual `<a:r>` nodes.
3. **Placeholder Boundary Matching**: Evaluates regular expressions (e.g. `/\{\{([^}]+)\}\}/g`) against the unified paragraph string.
4. **Run Coalescence & Formatting Preservation**:
   - Identifies the start run ($R_{\text{start}}$) and end run ($R_{\text{end}}$) containing the placeholder token.
   - Preserves all formatting properties (`<a:rPr>`) from $R_{\text{start}}$ (font, size, color, bold, italic).
   - Injects the replacement text into $R_{\text{start}}$'s `<a:t>` node.
   - Removes or clears the consumed text from intervening and trailing runs ($R_{\text{start}+1}$ through $R_{\text{end}}$).
5. **Multi-Line Splitting**: If the replacement string contains line breaks (`\n`), the engine converts them into separate `<a:p>` (paragraph) elements or `<a:br>` (break) runs, copying the original paragraph's spacing and bullet properties.

---

## 3.2 Native Chart & Excel Spreadsheet Dual-Sync

### The Dual-Cache Architecture
A PowerPoint native chart comprises two synchronized data representations:
1. **DrawingML Visual Cache (`ppt/charts/chartX.xml`)**: Used by PowerPoint to render the visual chart on screen. Contains `<c:plotArea>`, `<c:ser>`, `<c:cat>`, `<c:val>`, and formula strings (e.g., `<c:f>Sheet1!$B$2:$B$5</c:f>`).
2. **Embedded Excel Binary (`ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx`)**: An actual zipped OpenXML spreadsheet containing the backing data grid (`xl/worksheets/sheet1.xml`). Used whenever a user clicks "Edit Data in Excel" inside PowerPoint.

### The Conflict & Repair Warning
If a library mutates only the DrawingML XML:
- PowerPoint opens the presentation showing the updated visual graph.
- As soon as the user clicks **"Edit Data in Excel"**, PowerPoint reloads the chart from the embedded `.xlsx` workbook.
- Because the embedded sheet contains old numbers or mismatched rows, PowerPoint displays a collision dialog: *"The data in the spreadsheet does not match the chart data"*, or silently overwrites the user's modifications with the stale workbook data.

### The Dual-Sync Implementation
`ChartManager` and `ChartWorkbookUpdater` perform simultaneous updates:
1. **DrawingML Update**:
   - Replaces category labels in `<c:cat><c:strRef><c:strCache>`.
   - Replaces series data points in `<c:val><c:numRef><c:numCache>`.
   - Synchronizes `<c:ptCount @_val="N"/>` to match the exact point count.
   - Regenerates 0-indexed contiguous point indices (`<c:pt @_idx="0">`, `<c:pt @_idx="1">`).
   - Updates formula ranges (e.g., expanding `Sheet1!$B$2:$B$5` to `Sheet1!$B$2:$B$10` if 5 new points were added).
2. **Embedded XLSX Update**:
   - Decompresses the embedded `.xlsx` binary in memory.
   - Updates `xl/worksheets/sheet1.xml` cell values (`<c r="A2"><v>...</v></c>`).
   - Synchronizes `xl/sharedStrings.xml` if string labels are modified.
   - Recompresses the `.xlsx` binary and updates the VFS entry.

---

## 3.3 Table Matrix Consistency & 32-Bit Unsigned `<a16:rowId>`

### The 32-Bit Unsigned `<a16:rowId>` Invariant
Starting in Microsoft PowerPoint 2016, PowerPoint tracks table rows using an OpenXML DrawingML 2014 Extension element:
```xml
<a:tr h="450000">
  <a16:rowId val="2948201948" xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"/>
  <a:tc>...</a:tc>
</a:tr>
```
If dynamic rows are appended or cloned without `<a16:rowId>`, or if duplicate `<a16:rowId>` values are present, PowerPoint's schema parser rejects the table and prompts for a document repair.

`TableManager` generates unique 32-bit unsigned integers:
```javascript
function generateUniqueRowId(existingIds) {
  let id;
  do {
    // Generate an unsigned 32-bit integer (1 to 2^31 - 1)
    id = Math.floor(Math.random() * 2147483647) + 1;
  } while (existingIds.has(id));
  existingIds.add(id);
  return id;
}
```

### The Equal Cell Matrix Rule
An OpenXML table (`<a:tbl>`) defines its column count in `<a:tblGrid>`:
```xml
<a:tblGrid>
  <a:gridCol w="2500000"/>
  <a:gridCol w="2500000"/>
  <a:gridCol w="2500000"/>
</a:tblGrid>
```
**Strict Invariant**: Every `<a:tr>` child MUST contain exactly 3 `<a:tc>` elements. If a row contains 2 cells or 4 cells, the table structure is corrupted. Even when cells are merged horizontally or vertically, dummy placeholder cells MUST be preserved in the XML tree.

---

## 3.4 Rectangular Cell Merging (Origin Spans & Covered Cells)

PowerPoint does not support arbitrary non-rectangular table cell merges. All cell merges must form a valid 2D rectangular bounding box defined by `[startRow, startCol]` to `[endRow, endCol]`.

### The Origin vs. Shadowed Cell Pattern
To merge a $2 \times 2$ block of cells across Row 1–2 and Col 1–2:
1. **Origin Cell (Top-Left, Row 1, Col 1)**:
   - Decorated with `gridSpan="2"` (spanning 2 columns).
   - Decorated with `rowSpan="2"` (spanning 2 rows).
   - Contains the visible text and styling.
2. **Horizontally Shadowed Cell (Row 1, Col 2)**:
   - Exists as `<a:tc hMerge="1">` with an empty `<a:txBody>`.
3. **Vertically Shadowed Cell (Row 2, Col 1)**:
   - Exists as `<a:tc vMerge="1">` with an empty `<a:txBody>`.
4. **Diagonally Shadowed Cell (Row 2, Col 2)**:
   - Exists as `<a:tc hMerge="1" vMerge="1">` with an empty `<a:txBody>`.

```
┌────────────────────────┬────────────────────────┐
│   ORIGIN CELL (1, 1)   │   SHADOWED CELL (1, 2) │
│ gridSpan="2" rowSpan="2│       hMerge="1"       │
├────────────────────────┼────────────────────────┤
│   SHADOWED CELL (2, 1) │   SHADOWED CELL (2, 2) │
│       vMerge="1"       │   hMerge="1" vMerge="1"│
└────────────────────────┴────────────────────────┘
```
Failing to place shadowed cells or omitting them from `<a:tr>` violates the Equal Cell Matrix Rule and corrupts the presentation.

---

## 3.5 XML Security (XXE, Billion Laughs, DTD Entity Guardrails)

Presentations ingested by serverless functions often originate from untrusted public uploads. Malicious OpenXML packages can embed dangerous XML payloads.

`XMLParser.js` and `xmlUtils.js` implement defensive security guardrails:

### 1. XML External Entity (XXE) Prevention
- External entity declarations (`<!ENTITY xxe SYSTEM "file:///etc/passwd">` or `<!ENTITY xxe SYSTEM "http://malicious.com">`) are strictly rejected.
- `fast-xml-parser` is configured to never resolve external DTDs or URIs.
- Pre-parse scanning scans input XML strings and throws a `PPTXError('Security Error: XXE injection attempt detected')` if external entity identifiers (`SYSTEM` or `PUBLIC`) are present.

### 2. Billion Laughs & Quadratic Blowup Prevention
- Nested entity expansion attacks (where an entity expands exponentially into gigabytes of RAM) are blocked by prohibiting `<!ENTITY>` definitions inside DOCTYPE declarations.
- XML inputs with recursive or circular entity references are rejected prior to DOM construction.

### 3. ZIP Path Traversal (Slip) Guardrails
- Zip entries containing `../` or absolute path prefixes (`/etc/`, `C:\`) are rejected upon loading, ensuring files cannot escape the virtual sandbox.


---

# 4. Complete API & Method Specification

This section documents every public method provided by `node-pptx-templater` across all functional domains. Each entry details parameter types, return signatures, underlying OpenXML node mutations, edge-case constraints, and an executable JavaScript sample.

---

## 4.1 Lifecycle, Ingestion & In-Memory Preloading

### `static PPTXTemplater.load(source, options)`
- **Signature**: `static async load(source: string | Buffer | Uint8Array | ReadableStream, options?: { logLevel?: string, overwrite?: boolean }): Promise<PPTXTemplater>`
- **Return Type**: `Promise<PPTXTemplater>` (New instance initialized with loaded presentation).
- **OpenXML & ZIP Impact**: Unpacks ZIP archive into in-memory `VirtualFileSystem`. Parses `[Content_Types].xml`, `_rels/.rels`, and `ppt/presentation.xml`. Does not touch the filesystem if `source` is a `Buffer` or `Stream`.
- **Constraints & Edge Cases**:
  - Throws `PPTXError` if `source` is invalid or contains non-ZIP binary data without proper PKZip signatures.
  - Rejects XML containing malicious DTD/XXE declarations.
- **Example**:
  ```javascript
  import { PPTXTemplater } from 'node-pptx-templater';
  import fs from 'node:fs';

  // From file path
  const ppt = await PPTXTemplater.load('templates/deck.pptx');

  // From in-memory Buffer
  const buffer = fs.readFileSync('templates/deck.pptx');
  const pptFromBuf = await PPTXTemplater.load(buffer, { logLevel: 'warn' });
  ```

---

### `static PPTXTemplater.create()`
- **Signature**: `static async create(): Promise<PPTXTemplater>`
- **Return Type**: `Promise<PPTXTemplater>`
- **OpenXML & ZIP Impact**: Generates an in-memory scratch deck from embedded OpenXML defaults (`[Content_Types].xml`, blank slide master, theme, and 1 blank slide in 16:9 widescreen format).
- **Constraints & Edge Cases**: Generates an empty presentation without requiring an on-disk template file.
- **Example**:
  ```javascript
  const ppt = await PPTXTemplater.create();
  console.log(ppt.slideCount); // 1
  ```

---

### `static PPTXTemplater.fromPresentationXml(options)`
- **Signature**: `static async fromPresentationXml(options: { width?: number, height?: number }): Promise<PPTXTemplater>`
- **Return Type**: `Promise<PPTXTemplater>`
- **OpenXML & ZIP Impact**: Initializes an empty presentation package with custom slide dimensions written to `<p:sldSz cx="..." cy="..."/>` in `ppt/presentation.xml`.
- **Constraints & Edge Cases**: Dimensions are converted to EMUs.
- **Example**:
  ```javascript
  const ppt = await PPTXTemplater.fromPresentationXml({ width: 1920, height: 1080 });
  ```

---

### `static PPTXTemplater.preload(source)` / `static PPTXTemplater.cache(source)`
- **Signature**: `static async preload(source: string | Buffer): Promise<void>`
- **Return Type**: `Promise<void>`
- **OpenXML & ZIP Impact**: Inflates and caches the VFS file entries in a static global memory cache keyed by file path or Buffer SHA-256 hash.
- **Constraints & Edge Cases**: Ideal for serverless or web servers handling high request volume against identical templates.
- **Example**:
  ```javascript
  // On server startup:
  await PPTXTemplater.preload('templates/invoice-template.pptx');
  ```

---

### `static PPTXTemplater.fromCache(source)`
- **Signature**: `static async fromCache(source: string | Buffer): Promise<PPTXTemplater>`
- **Return Type**: `Promise<PPTXTemplater>`
- **OpenXML & ZIP Impact**: Clones cached VFS entries into a new `PPTXTemplater` instance in < 1ms without re-reading disk or re-inflating zip streams.
- **Constraints & Edge Cases**: Throws `PPTXError` if the specified source was not previously registered via `preload()`.
- **Example**:
  ```javascript
  // Inside request handler:
  const ppt = await PPTXTemplater.fromCache('templates/invoice-template.pptx');
  ppt.useSlide(1).replaceText({ '{{INVOICE_ID}}': 'INV-9021' });
  ```

---

### `static PPTXTemplater.clearCache()`
- **Signature**: `static clearCache(): void`
- **Return Type**: `void`
- **OpenXML & ZIP Impact**: Purges all static preloaded VFS packages from memory.
- **Example**:
  ```javascript
  PPTXTemplater.clearCache();
  ```

---

### `static PPTXTemplater.extractPptx(pptxPath, outputPath, options)`
- **Signature**: `static async extractPptx(pptxPath: string, outputPath: string, options?: { formatXml?: boolean }): Promise<void>`
- **Return Type**: `Promise<void>`
- **OpenXML & ZIP Impact**: Decompresses the `.pptx` container into a disk folder, optionally formatting all XML files for human readability.
- **Example**:
  ```javascript
  await PPTXTemplater.extractPptx('deck.pptx', './unpacked-deck', { formatXml: true });
  ```

---

### `static PPTXTemplater.buildPptx(folderPath, pptxPath)`
- **Signature**: `static async buildPptx(folderPath: string, pptxPath: string): Promise<void>`
- **Return Type**: `Promise<void>`
- **OpenXML & ZIP Impact**: Ingests an uncompressed folder tree and writes a valid compressed `.pptx` archive.
- **Example**:
  ```javascript
  await PPTXTemplater.buildPptx('./unpacked-deck', './output/rebuilt.pptx');
  ```

---

## 4.2 Slide Management & Coordinate Sizing

### `useSlide(...slideRefs)`
- **Signature**: `useSlide(...slideRefs: Array<number | number[] | 'all'>): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Sets internal `#selectedSlides` state. Reads `ppt/slides/slideN.xml` into memory if not already cached.
- **Constraints & Edge Cases**:
  - Slide numbering is **1-based** ($1 \le \text{index} \le \text{slideCount}$).
  - Passing `'all'` targets every slide in the presentation.
  - Throws `SlideNotFoundError` if any requested slide index exceeds `slideCount`.
- **Example**:
  ```javascript
  ppt.useSlide(1);            // Target slide 1
  ppt.useSlide(1, 3);         // Target slides 1 and 3
  ppt.useSlide('all');        // Target all slides
  ```

---

### `useAllSlides()`
- **Signature**: `useAllSlides(): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Convenience shortcut equivalent to `useSlide('all')`.
- **Example**:
  ```javascript
  ppt.useAllSlides().replaceText({ '{{CONFIDENTIAL}}': 'INTERNAL USE ONLY' });
  ```

---

### `addSlide(options)` / `insertSlide(slideIndex, options)`
- **Signature**: `addSlide(options?: { layout?: string | number, master?: number }): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**:
  - Creates `ppt/slides/slideM.xml` and `ppt/slides/_rels/slideM.xml.rels`.
  - Appends slide reference to `ppt/presentation.xml` inside `<p:sldIdLst>`.
  - Registers override in `[Content_Types].xml`.
- **Example**:
  ```javascript
  ppt.addSlide({ layout: 1 });
  ```

---

### `duplicateSlide(slideIndex, atPosition)` / `cloneSlide(sourceSlideNumber, atPosition)`
- **Signature**: `async duplicateSlide(slideIndex: number, atPosition?: number): Promise<this>`
- **Return Type**: `Promise<this>` (Chainable).
- **OpenXML & ZIP Impact**:
  - Clones source slide XML and relationship file.
  - Duplicates all embedded relationship pointers, avoiding ID collisions.
  - Adds a new entry in `<p:sldIdLst>` with a newly generated slide ID ($id \ge 256$).
  - Updates `[Content_Types].xml`.
- **Constraints & Edge Cases**: If `atPosition` is omitted, the slide is appended directly after `slideIndex`.
- **Example**:
  ```javascript
  // Duplicate slide 1 and place the clone at position 2:
  await ppt.duplicateSlide(1, 2);
  ```

---

### `removeSlide(slideNumber)` / `deleteSlide(slideIndex)`
- **Signature**: `removeSlide(slideNumber: number): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**:
  - Deletes `<p:sldId>` entry from `ppt/presentation.xml`.
  - Purges relationship from `ppt/_rels/presentation.xml.rels`.
  - Removes slide override from `[Content_Types].xml`.
  - Deletes `ppt/slides/slideN.xml` and `ppt/slides/_rels/slideN.xml.rels` from VFS.
- **Constraints & Edge Cases**: Throws `PPTXError` if attempting to delete the only slide in a presentation.
- **Example**:
  ```javascript
  ppt.removeSlide(3);
  ```

---

### `moveSlide(fromIndex, toIndex)` / `reorderSlides(order)`
- **Signature**: `moveSlide(fromIndex: number, toIndex: number): this`  
  `reorderSlides(order: number[]): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Reorders `<p:sldId>` entries within `<p:sldIdLst>` in `ppt/presentation.xml`.
- **Example**:
  ```javascript
  ppt.moveSlide(1, 3);             // Move slide 1 to position 3
  ppt.reorderSlides([3, 1, 2]);    // Complete reordering of 3 slides
  ```

---

### `tagSlide(slideNumber, tag)`
- **Signature**: `tagSlide(slideNumber: number, tag: string): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Inserts metadata tag into slide non-visual properties or comment streams.
- **Example**:
  ```javascript
  ppt.tagSlide(1, 'ExecutiveSummary');
  ```

---

### `get slideCount(): number`
- **Signature**: `get slideCount(): number`
- **Return Type**: `number`
- **OpenXML & ZIP Impact**: Reads length of `<p:sldId>` elements in `ppt/presentation.xml`.
- **Example**:
  ```javascript
  console.log(`Presentation contains ${ppt.slideCount} slides`);
  ```

---

### `getSlides()`
- **Signature**: `getSlides(): Array<{ index: number, partPath: string, rId: string }>`
- **Return Type**: `Array<Object>`
- **Example**:
  ```javascript
  const slides = ppt.getSlides();
  ```

---

### `getSlideDimensions()` / `getSlideSize()`
- **Signature**: `getSlideDimensions(): { width: number, height: number, widthEmu: number, heightEmu: number, type: string }`
- **Return Type**: `Object`
- **OpenXML & ZIP Impact**: Reads `<p:sldSz cx="..." cy="..." type="..."/>` in `ppt/presentation.xml`.
- **Example**:
  ```javascript
  const dims = ppt.getSlideDimensions();
  console.log(`Slide width in EMUs: ${dims.widthEmu}, height: ${dims.heightEmu}`);
  ```

---

### `setSlideDimensions(widthOrOptions, height)` / `setSlideSize(width, height)` / `setAspectRatio(ratio)`
- **Signature**: `setSlideDimensions(width: number | string, height?: number | string): this`  
  `setAspectRatio(ratio: '16:9' | '16:10' | '4:3'): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Mutates `<p:sldSz>` in `ppt/presentation.xml` with converted integer EMUs.
- **Example**:
  ```javascript
  ppt.setSlideDimensions('13.333in', '7.5in'); // Standard 16:9 widescreen
  ppt.setAspectRatio('16:9');
  ```

---

### `getSlideLayout(slideIndex)`
- **Signature**: `getSlideLayout(slideIndex?: number): { name: string, type: string, master: string, rId: string }`
- **Return Type**: `Object`
- **OpenXML & ZIP Impact**: Resolves layout relationship in `slideN.xml.rels` pointing to `slideLayouts/slideLayoutX.xml`.
- **Example**:
  ```javascript
  const layout = ppt.getSlideLayout(1);
  console.log(`Slide 1 uses layout: ${layout.name}`);
  ```

---

### `importSlideFrom(sourceEngine, slideRef)` / `importSlides(slideIndices)`
- **Signature**: `async importSlideFrom(sourceEngine: PPTXTemplater, slideRef: number): Promise<this>`
- **Return Type**: `Promise<this>` (Chainable).
- **OpenXML & ZIP Impact**:
  - Copies slide XML and all dependent layouts, shapes, media, and fonts from `sourceEngine`.
  - Re-maps relationship IDs to guarantee unique keys.
  - Appends imported slide to `<p:sldIdLst>`.
- **Example**:
  ```javascript
  const sourceDeck = await PPTXTemplater.load('corporate-branding.pptx');
  await ppt.importSlideFrom(sourceDeck, 1);
  ```

---

### `exportSlides(...slideNumbers)`
- **Signature**: `async exportSlides(...slideNumbers: number[]): Promise<PPTXTemplater>`
- **Return Type**: `Promise<PPTXTemplater>` (New isolated PPTX instance containing only exported slides).
- **Example**:
  ```javascript
  const standaloneDeck = await ppt.exportSlides(1, 2);
  await standaloneDeck.saveToFile('extracted-slides.pptx');
  ```


---

## 4.3 Text Manipulation & Placeholder Templating

### `replaceText(replacements)` / `replaceMultiple(replacements)`
- **Signature**: `replaceText(replacements: Record<string, string | number>): this`  
  `replaceMultiple(replacements: Record<string, string | number>): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**:
  - Scans paragraphs `<a:p>` across targeted slides.
  - Performs AST run healing across fragmented `<a:r>` and `<a:t>` tags.
  - Replaces token values while retaining `<a:rPr>` typography styles (font, size, color, bold, italic).
  - Encodes XML special characters (`&`, `<`, `>`, `"`, `'`).
  - Converts newline characters (`\n`) into proper paragraph `<a:p>` or break `<a:br>` structures.
- **Constraints & Edge Cases**:
  - Token keys should include delimiter braces if used in the template (e.g. `'{{company_name}}'`).
  - If a placeholder is not present in the slide, it is skipped without error.
- **Example**:
  ```javascript
  ppt.useSlide(1).replaceText({
    '{{TITLE}}': 'Annual Shareholder Report',
    '{{YEAR}}': 2026,
    '{{REVENUE}}': '$14.2B'
  });
  ```

---

### `replaceTextByTag(tag, replacement)`
- **Signature**: `replaceTextByTag(tag: string, replacement: string | number): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Convenience method for replacing an individual tag token.
- **Example**:
  ```javascript
  ppt.replaceTextByTag('{{STATUS}}', 'Approved');
  ```

---

### `updateText(targetOrOptions, newText, options)`
- **Signature**: `updateText(targetOrOptions: string | { shapeId?: string, shapeName?: string }, newText: string, options?: { fontSize?: number, fontFace?: string, color?: string, bold?: boolean, italic?: boolean }): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Mutates the text body `<p:txBody>` of a specific shape or text box, optionally applying typography overrides in `<a:rPr>`.
- **Example**:
  ```javascript
  ppt.useSlide(1).updateText('SubtitleBox', 'Prepared for Acme Corp', {
    fontSize: 18,
    color: '64748B',
    italic: true
  });
  ```

---

### `findText(pattern, options)`
- **Signature**: `findText(pattern: string | RegExp, options?: { caseSensitive?: boolean }): Array<{ slideIndex: number, shapeId: string, shapeName: string, matchedText: string }>`
- **Return Type**: `Array<Object>`
- **OpenXML & ZIP Impact**: Scans `<a:p>` runs across active slides without modifying XML.
- **Example**:
  ```javascript
  const matches = ppt.findText(/TBD|TODO/i);
  console.log(`Found ${matches.length} unresolved notes.`);
  ```

---

### `appendTextRun(textOrOptions, runOptions)` / `prependTextRun(textOrOptions, runOptions)`
- **Signature**: `appendTextRun(textOrOptions: string | { shapeId: string, text: string }, runOptions?: { fontSize?: number, color?: string, bold?: boolean }): this`  
  `prependTextRun(textOrOptions: string | { shapeId: string, text: string }, runOptions?: { fontSize?: number, color?: string, bold?: boolean }): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Inserts a new `<a:r>` element before or after existing runs within the target shape's last paragraph `<a:p>`.
- **Example**:
  ```javascript
  ppt.appendTextRun({ shapeId: 'DisclaimerBox', text: ' (Confidential)' }, {
    color: 'EF4444',
    bold: true
  });
  ```

---

### `extractPlaceholders()`
- **Signature**: `extractPlaceholders(): string[]`
- **Return Type**: `string[]` (Array of unique placeholder strings).
- **OpenXML & ZIP Impact**: Inspects all active slides and extracts every matching `{{...}}` pattern across healed text runs.
- **Example**:
  ```javascript
  const tags = ppt.useAllSlides().extractPlaceholders();
  console.log('Unfilled placeholders:', tags); // ['{{NAME}}', '{{DATE}}']
  ```

---

### `getList(listIdOrShapeName)` / `validateList(listIdOrShapeName)`
- **Signature**: `getList(listIdOrShapeName: string): Array<{ text: string, level: number, bullet: boolean }>`  
  `validateList(listIdOrShapeName: string): boolean`
- **Return Type**: `Array<Object>` / `boolean`
- **OpenXML & ZIP Impact**: Inspects `<a:pPr lvl="...">` and `<a:buChar>` / `<a:buAutoNum>` structures within a shape.
- **Example**:
  ```javascript
  const listItems = ppt.getList('KeyTakeaways');
  ```

---

## 4.4 Table Operations, Merging & Auto-Pagination

### `addTable(tableIdOrOptions, options)`
- **Signature**: `addTable(tableIdOrOptions: string | { id?: string, rows: number, cols: number, x: number | string, y: number | string, width: number | string, height: number | string }, options?: object): string`
- **Return Type**: `string` (The allocated table element ID).
- **OpenXML & ZIP Impact**: Inserts `<p:graphicFrame>` containing `<a:graphicData uri=".../drawingml/2006/table">` with `<a:tbl>`, `<a:tblGrid>`, and initial `<a:tr>` rows into `<p:spTree>`.
- **Example**:
  ```javascript
  ppt.useSlide(1).addTable({
    id: 'FinancialTable',
    rows: 4,
    cols: 3,
    x: '1in',
    y: '2in',
    width: '8in',
    height: '3in'
  });
  ```

---

### `updateTable(tableId, rows)`
- **Signature**: `updateTable(tableId: string, rows: any[][]): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**:
  - Targets `<a:tbl>` matching `tableId` (by name or non-visual ID).
  - Populates cells `<a:tc>` with text values, copying styling and dimensions from existing rows.
  - Automatically injects unique 32-bit unsigned `<a16:rowId>` elements to newly inserted rows.
- **Example**:
  ```javascript
  ppt.updateTable('FinancialTable', [
    ['Metric', '2025', '2026'],
    ['ARR', '$10M', '$18M'],
    ['Customers', '120', '210'],
    ['Net Margin', '22%', '28%']
  ]);
  ```

---

### `getTableRows(tableId, options)`
- **Signature**: `getTableRows(tableId: string, options?: { raw?: boolean }): Promise<any[]> | any[]`
- **Return Type**: `Array<Array<string | Object>>`
- **OpenXML & ZIP Impact**: Parses `<a:tbl>` and returns 2D grid matrix of cell text, accounting for merged regions.
- **Example**:
  ```javascript
  const matrix = await ppt.getTableRows('FinancialTable');
  console.log(matrix);
  ```

---

### `addTableRow(tableId, rowData, options)` / `insertTableRow(tableId, rowIndex, rowData, options)`
- **Signature**: `addTableRow(tableId: string, rowData: any[], options?: { height?: number | string }): this`  
  `insertTableRow(tableId: string, rowIndex: number, rowData: any[], options?: object): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Appends or inserts a new `<a:tr>` element into `<a:tbl>`.
  - Sets row height in EMUs (`@_h`).
  - Generates unique `<a16:rowId>`.
  - Populates `<a:tc>` elements matching `<a:gridCol>` count.
- **Example**:
  ```javascript
  ppt.addTableRow('FinancialTable', ['EBITDA', '$3.2M', '$5.4M']);
  ppt.insertTableRow('FinancialTable', 1, ['Gross Profit', '$8.5M', '$15.1M']);
  ```

---

### `cloneTableRow(tableId, sourceRowIndex, targetRowIndex)`
- **Signature**: `cloneTableRow(tableId: string, sourceRowIndex: number, targetRowIndex?: number): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Duplicates `<a:tr>` node, copying background fills and cell borders, while generating a fresh, unique `<a16:rowId>`.
- **Example**:
  ```javascript
  ppt.cloneTableRow('FinancialTable', 1);
  ```

---

### `removeTableRow(tableId, rowIndex)`
- **Signature**: `removeTableRow(tableId: string, rowIndex: number): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Removes the `<a:tr>` element at zero-based index `rowIndex`.
- **Example**:
  ```javascript
  ppt.removeTableRow('FinancialTable', 2);
  ```

---

### `addColumn(tableId, colIndexOrData, colDataOrOptions, options)` / `deleteColumn(tableId, colIndex)`
- **Signature**: `addColumn(tableId: string, colIndex: number, colData?: any[], options?: { width?: number | string }): this`  
  `deleteColumn(tableId: string, colIndex: number): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**:
  - `addColumn`: Inserts `<a:gridCol w="..."/>` in `<a:tblGrid>` and inserts corresponding `<a:tc>` cells in every `<a:tr>`.
  - `deleteColumn`: Removes `<a:gridCol>` and strips the corresponding cell from each `<a:tr>`.
- **Example**:
  ```javascript
  ppt.addColumn('FinancialTable', 3, ['Target', '$20M', '250', '30%'], { width: '2in' });
  ppt.deleteColumn('FinancialTable', 0);
  ```

---

### `updateCell(tableId, rowIndex, colIndex, value, options)`
- **Signature**: `updateCell(tableId: string, rowIndex: number, colIndex: number, value: string | number, options?: { fontSize?: number, bold?: boolean, color?: string, fill?: string, align?: 'left'|'center'|'right' }): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Mutates the text body `<a:txBody>` and cell property `<a:tcPr>` of the cell at `(rowIndex, colIndex)`.
- **Example**:
  ```javascript
  ppt.updateCell('FinancialTable', 1, 1, '$11.5M', {
    bold: true,
    color: '10B981', // Emerald green
    fill: 'F0FDF4'
  });
  ```

---

### `mergeCells(tableId, startRow, startCol, endRow, endCol)`
- **Signature**: `mergeCells(tableId: string, startRow: number, startCol: number, endRow: number, endCol: number): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**:
  - Decorates origin cell `(startRow, startCol)` with `gridSpan` and `rowSpan`.
  - Places shadowed placeholder cells with `hMerge="1"` and `vMerge="1"` across the rectangular region.
- **Constraints & Edge Cases**: Region must form a valid rectangle. Validates that coordinates are within table bounds.
- **Example**:
  ```javascript
  // Merge top header across 3 columns:
  ppt.mergeCells('FinancialTable', 0, 0, 0, 2);
  ```

---

### `unmergeCells(tableId, startRow, startCol, endRow, endCol)`
- **Signature**: `unmergeCells(tableId: string, startRow: number, startCol: number, endRow: number, endCol: number): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Clears `gridSpan` and `rowSpan` on origin cell, removes `hMerge` and `vMerge` attributes on shadowed cells.
- **Example**:
  ```javascript
  ppt.unmergeCells('FinancialTable', 0, 0, 0, 2);
  ```

---

### `isMergedCell(tableId, rowIndex, colIndex)` / `getMergeRegion(tableId, rowIndex, colIndex)` / `getMergedCells(tableId)`
- **Signature**: `isMergedCell(tableId: string, rowIndex: number, colIndex: number): boolean`  
  `getMergeRegion(tableId: string, rowIndex: number, colIndex: number): { startRow: number, startCol: number, endRow: number, endCol: number } | null`  
  `getMergedCells(tableId: string): Array<Object>`
- **Return Type**: `boolean` / `Object | null` / `Array<Object>`
- **Example**:
  ```javascript
  if (ppt.isMergedCell('FinancialTable', 0, 1)) {
    const region = ppt.getMergeRegion('FinancialTable', 0, 1);
    console.log(`Cell belongs to region:`, region);
  }
  ```

---

### `stripeRows(tableId, options)`
- **Signature**: `stripeRows(tableId: string, options?: { evenColor?: string, oddColor?: string, startRow?: number }): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Iterates over `<a:tr>` elements, setting `<a:tcPr><a:solidFill>` on cells with alternating hex colors.
- **Example**:
  ```javascript
  ppt.stripeRows('FinancialTable', {
    oddColor: 'FFFFFF',
    evenColor: 'F8FAFC',
    startRow: 1 // Keep header un-striped
  });
  ```

---

### `autoFitTable(tableId, options)` / `resizeTable(tableId, dimensions)`
- **Signature**: `autoFitTable(tableId: string, options?: { maxRowHeight?: number }): this`  
  `resizeTable(tableId: string, dimensions: { width?: number | string, height?: number | string }): this`
- **Return Type**: `this` (Chainable).
- **Example**:
  ```javascript
  ppt.autoFitTable('FinancialTable');
  ```

---

### `splitTableOnOverflow(tableId, options)`
- **Signature**: `async splitTableOnOverflow(tableId: string, options?: { bottomMargin?: number | string, repeatHeaders?: number, preserveSummaryRows?: number }): Promise<{ splitCount: number, slidesCreated: number }>`
- **Return Type**: `Promise<{ splitCount: number, slidesCreated: number }>`
- **OpenXML & ZIP Impact**:
  - Calculates cumulative table height and compares against slide boundaries ($Y_{\text{limit}} = \text{SlideHeight} - \text{bottomMargin}$).
  - When rows overflow, duplicates the slide (`duplicateSlide`), splits table rows at the break threshold, re-injects designated header rows onto new slides, pins summary rows to the last slide, and updates slide relationship manifests.
- **Constraints & Edge Cases**: Recursively creates as many slides as necessary until all records are accommodated.
- **Example**:
  ```javascript
  // Inject 100 rows into a table and auto-paginate across slides:
  const result = await ppt.splitTableOnOverflow('LargeDataTable', {
    bottomMargin: '1in',
    repeatHeaders: 1,
    preserveSummaryRows: 1
  });
  console.log(`Created ${result.slidesCreated} pagination slides.`);
  ```

---

### `getCellBounds(tableId, rowIndex, colIndex)` / `getCellPosition(tableId, rowIndex, colIndex)`
- **Signature**: `getCellBounds(tableId: string, rowIndex: number, colIndex: number): { x: number, y: number, width: number, height: number, xEmu: number, yEmu: number, widthEmu: number, heightEmu: number }`
- **Return Type**: `Object`
- **OpenXML & ZIP Impact**: Reads table `<a:xfrm>`, sums `<a:gridCol>` widths up to `colIndex`, and sums `<a:tr h="...">` heights up to `rowIndex` to return the exact slide-relative position of a cell.
- **Example**:
  ```javascript
  const bounds = ppt.getCellBounds('FinancialTable', 1, 2);
  console.log(`Cell (1, 2) is at X: ${bounds.x}px, Y: ${bounds.y}px`);
  ```


---

## 4.5 Native Chart Updates, Data Labels & Positioning

### `updateChart(chartId, data)` / `updateChartData(chartId, seriesData)`
- **Signature**: `async updateChart(chartId: string, data: { categories?: string[], series: Array<{ name: string, values: number[] }> }): Promise<this>`  
  `async updateChartData(chartId: string, seriesData: Array<{ name: string, values: number[] }>): Promise<this>`
- **Return Type**: `Promise<this>` (Chainable).
- **OpenXML & ZIP Impact**:
  - Locates `ppt/charts/chartX.xml` associated with the target chart shape.
  - Updates DrawingML `<c:ser>` elements, `<c:cat>` string cache (`<c:strCache>`), and `<c:val>` numeric cache (`<c:numCache>`).
  - Sets `<c:ptCount @_val="N"/>` and generates 0-indexed contiguous `<c:pt @_idx="i">` nodes.
  - Automatically resolves and updates the embedded Excel spreadsheet in `ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx` to guarantee dual-sync cache parity.
- **Constraints & Edge Cases**:
  - `chartId` can be the PowerPoint shape name (e.g. `'SalesChart'`) or non-visual object ID.
  - Throws `ChartNotFoundError` if chart does not exist on the selected slide.
- **Example**:
  ```javascript
  await ppt.useSlide(1).updateChart('RevenueChart', {
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [
      { name: 'Target', values: [100, 120, 140, 160] },
      { name: 'Actual', values: [105, 128, 138, 172] }
    ]
  });
  ```

---

### `updateChartCategories(chartId, categories)`
- **Signature**: `async updateChartCategories(chartId: string, categories: string[]): Promise<this>`
- **Return Type**: `Promise<this>` (Chainable).
- **OpenXML & ZIP Impact**: Mutates category labels in both `<c:chartSpace>` and the embedded `.xlsx` workbook.
- **Example**:
  ```javascript
  await ppt.updateChartCategories('RevenueChart', ['Jan', 'Feb', 'Mar']);
  ```

---

### `updateChartTitle(chartId, title)`
- **Signature**: `updateChartTitle(chartId: string, title: string): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Sets `<c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>` to `title`.
- **Example**:
  ```javascript
  ppt.updateChartTitle('RevenueChart', 'Fiscal 2026 Performance');
  ```

---

### `replaceChartSeries(chartId, seriesIndexOrName, newSeriesData)`
- **Signature**: `async replaceChartSeries(chartId: string, seriesIndexOrName: number | string, newSeriesData: { name?: string, values: number[] }): Promise<this>`
- **Return Type**: `Promise<this>` (Chainable).
- **OpenXML & ZIP Impact**: Replaces a single series without altering sibling series data.
- **Example**:
  ```javascript
  await ppt.replaceChartSeries('RevenueChart', 0, {
    name: 'Revised Target',
    values: [110, 125, 145, 165]
  });
  ```

---

### `getDataLabels(chartId, options)` / `updateDataLabels(chartId, labelsOrOptions)`
- **Signature**: `async getDataLabels(chartId: string, options?: object): Promise<any>`  
  `async updateDataLabels(chartId: string, labelsOrOptions: { showVal?: boolean, showCatName?: boolean, showLegendKey?: boolean, showPercent?: boolean, position?: 'bestFit'|'b'|'ctr'|'inBase'|'inEnd'|'l'|'r'|'t' }): Promise<this>`
- **Return Type**: `Promise<any>` / `Promise<this>` (Chainable).
- **OpenXML & ZIP Impact**: Mutates `<c:dLbls>` container inside `<c:plotArea>`:
  - `<c:showVal @_val="1|0"/>`
  - `<c:showCatName @_val="1|0"/>`
  - `<c:dLblPos @_val="outEnd|ctr|inBase"/>`
- **Example**:
  ```javascript
  await ppt.updateDataLabels('RevenueChart', {
    showVal: true,
    showCatName: false,
    position: 'outEnd'
  });
  ```

---

### `getChartLabelPositions(chartId)` / `getChartBarPositions(chartId)`
- **Signature**: `async getChartLabelPositions(chartId: string): Promise<Array<{ index: number, x: number, y: number }>>`  
  `async getChartBarPositions(chartId: string): Promise<Array<{ seriesIndex: number, pointIndex: number, x: number, y: number, width: number, height: number }>>`
- **Return Type**: `Promise<Array<Object>>`
- **OpenXML & ZIP Impact**: Computes pixel coordinates for chart bars and labels relative to the slide canvas.
- **Example**:
  ```javascript
  const labelPositions = await ppt.getChartLabelPositions('RevenueChart');
  ```

---

### `addTextNearChartLabel(chartId, labelIndex, text, options)`
- **Signature**: `async addTextNearChartLabel(chartId: string, labelIndex: number, text: string, options?: { offsetX?: number, offsetY?: number, color?: string, fontSize?: number }): Promise<this>`
- **Return Type**: `Promise<this>` (Chainable).
- **OpenXML & ZIP Impact**: Resolves chart label spatial position and inserts an absolute text box in `<p:spTree>` near the target data point.
- **Example**:
  ```javascript
  await ppt.addTextNearChartLabel('RevenueChart', 3, '⭐ Record High!', {
    offsetY: -20,
    color: 'D97706',
    fontSize: 12
  });
  ```

---

## 4.6 Shapes, Spatial Geometry & Alignment

### `addShape(typeOrOptions, options)`
- **Signature**: `addShape(typeOrOptions: string | { type: string, id?: string, name?: string, x: number | string, y: number | string, width: number | string, height: number | string, fill?: string, stroke?: string, strokeWidth?: number, text?: string }, options?: object): string`
- **Return Type**: `string` (The created shape ID).
- **OpenXML & ZIP Impact**:
  - Inserts `<p:sp>` inside `<p:spTree>` with geometry preset `<a:prstGeom prst="...">` (e.g. `rect`, `roundRect`, `ellipse`, `triangle`).
  - Sets spatial transform `<a:xfrm>` (`<a:off>`, `<a:ext>`) with converted integer EMUs.
  - Applies fill `<a:solidFill>` and line styling `<a:ln>`.
- **Example**:
  ```javascript
  ppt.useSlide(1).addShape({
    type: 'roundRect',
    id: 'KpiCard',
    name: 'CardBackground',
    x: '1in',
    y: '1.5in',
    width: '4in',
    height: '2.5in',
    fill: 'F1F5F9',
    stroke: 'CBD5E1',
    strokeWidth: 1.5,
    text: 'Total Users: 24,500'
  });
  ```

---

### `updateShape(shapeId, options)` / `updateShapePosition(shapeId, bounds)` / `updateShapeText(shapeId, text, options)`
- **Signature**: `updateShape(shapeId: string, options: { x?: number | string, y?: number | string, width?: number | string, height?: number | string, fill?: string, stroke?: string, text?: string }): this`  
  `updateShapePosition(shapeId: string, bounds: { x?: number | string, y?: number | string, width?: number | string, height?: number | string }): this`  
  `updateShapeText(shapeId: string, text: string, options?: object): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Mutates `<a:xfrm>`, `<a:solidFill>`, or `<p:txBody>` within the target `<p:sp>` node.
- **Example**:
  ```javascript
  ppt.updateShape('KpiCard', {
    fill: 'E0F2FE',
    text: 'Total Users: 28,100 (+15%)'
  });
  ```

---

### `cloneShape(shapeId, options)` / `deleteShape(shapeId)` / `removeShape(shapeId)`
- **Signature**: `cloneShape(shapeId: string, options?: { newId?: string, x?: number | string, y?: number | string }): string`  
  `deleteShape(shapeId: string): this`  
  `removeShape(shapeId: string): this`
- **Return Type**: `string` (Cloned shape ID) / `this` (Chainable).
- **OpenXML & ZIP Impact**:
  - `cloneShape`: Duplicates `<p:sp>` AST node, updates shape ID and name, and appends to `<p:spTree>`.
  - `deleteShape`: Removes `<p:sp>` node from `<p:spTree>`.
- **Example**:
  ```javascript
  const card2Id = ppt.cloneShape('KpiCard', { x: '5.5in' });
  ppt.deleteShape('OldPlaceholderShape');
  ```

---

### `getBounds(objectIdOrName)` / `setBounds(objectIdOrName, bounds)`
- **Signature**: `getBounds(objectIdOrName: string): { x: number, y: number, width: number, height: number, xEmu: number, yEmu: number, widthEmu: number, heightEmu: number }`  
  `setBounds(objectIdOrName: string, bounds: { x?: number | string, y?: number | string, width?: number | string, height?: number | string }): this`
- **Return Type**: `Object` / `this` (Chainable).
- **OpenXML & ZIP Impact**: Reads or writes `<a:off @_x="..." @_y="..."/>` and `<a:ext @_cx="..." @_cy="..."/>`.
- **Example**:
  ```javascript
  const bounds = ppt.getBounds('CardBackground');
  console.log(`Card is at X: ${bounds.x}px, width: ${bounds.width}px`);

  ppt.setBounds('CardBackground', { width: '4.5in' });
  ```

---

### `move(objectIdOrName, options)` / `resize(objectIdOrName, options)`
- **Signature**: `move(objectIdOrName: string, options: { x: number | string, y: number | string }): this`  
  `resize(objectIdOrName: string, options: { width: number | string, height: number | string }): this`
- **Return Type**: `this` (Chainable).
- **Example**:
  ```javascript
  ppt.move('CardBackground', { x: '1.2in', y: '1.8in' });
  ppt.resize('CardBackground', { width: '5in', height: '3in' });
  ```

---

### `alignShapes(shapeIds, alignment)`
- **Signature**: `alignShapes(shapeIds: string[], alignment: 'left' | 'right' | 'center' | 'top' | 'bottom' | 'middle'): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Computes alignment bounding box across all specified shapes and aligns their respective `<a:off>` coordinates.
- **Example**:
  ```javascript
  // Align 3 KPI cards along their vertical middle:
  ppt.alignShapes(['Card1', 'Card2', 'Card3'], 'middle');
  ```

---

### `distributeShapes(shapeIds, direction)`
- **Signature**: `distributeShapes(shapeIds: string[], direction: 'horizontal' | 'vertical'): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Computes uniform spacing between consecutive shapes and updates `<a:off>` offsets.
- **Example**:
  ```javascript
  ppt.distributeShapes(['Card1', 'Card2', 'Card3'], 'horizontal');
  ```

---

### `checkCollision(idA, idB)` / `isContained(innerId, outerId)`
- **Signature**: `checkCollision(idA: string, idB: string): boolean`  
  `isContained(innerId: string, outerId: string): boolean`
- **Return Type**: `boolean`
- **OpenXML & ZIP Impact**: Geometric collision queries comparing `<a:xfrm>` boxes in EMU space without mutating XML.
- **Example**:
  ```javascript
  if (ppt.checkCollision('BadgeShape', 'TitleText')) {
    ppt.move('BadgeShape', { x: '0.5in', y: '0in' });
  }
  ```

---

### `addTextAtPosition(text, options)`
- **Signature**: `addTextAtPosition(text: string, options: { x: number | string, y: number | string, width?: number | string, height?: number | string, fontSize?: number, color?: string, bold?: boolean }): string`
- **Return Type**: `string` (The created text box ID).
- **Example**:
  ```javascript
  ppt.addTextAtPosition('Quarterly Status Report', {
    x: '1in',
    y: '0.8in',
    fontSize: 24,
    bold: true,
    color: '0F172A'
  });
  ```

---

## 4.7 Table Cell Shapes (Non-Reflow Overlay Badging)

### `addCellShape(tableId, rowIndex, colIndex, options)`
- **Signature**: `addCellShape(tableId: string, rowIndex: number, colIndex: number, options: { type?: string, width?: number | string, height?: number | string, fill?: string, stroke?: string, text?: string, textColor?: string, alignment?: 'center' | 'left' | 'right' }): string`
- **Return Type**: `string` (The created shape ID).
- **OpenXML & ZIP Impact**:
  - Resolves cell bounding coordinates on the slide canvas using `getCellBounds()`.
  - Places an absolute vector shape in `<p:spTree>` directly centered or aligned over the cell.
  - Guarantees **non-reflow**: The shape visually tracks the cell without altering table cell padding or forcing row height expansion.
- **Example**:
  ```javascript
  // Embed a green "Active" pill badge inside Table 1, Row 2, Column 3:
  ppt.addCellShape('Table1', 2, 3, {
    type: 'roundRect',
    width: '1in',
    height: '0.3in',
    fill: 'DCFCE7', // Light green
    text: 'Active',
    textColor: '15803D',
    alignment: 'center'
  });
  ```

---

### `updateCellShape(tableId, rowIndex, colIndex, shapeIndex, options)` / `removeCellShape(tableId, rowIndex, colIndex, shapeIndex)`
- **Signature**: `updateCellShape(tableId: string, rowIndex: number, colIndex: number, shapeIndex: number, options: object): this`  
  `removeCellShape(tableId: string, rowIndex: number, colIndex: number, shapeIndex: number): this`
- **Return Type**: `this` (Chainable).
- **Example**:
  ```javascript
  ppt.updateCellShape('Table1', 2, 3, 0, {
    fill: 'FEE2E2',
    text: 'Expired',
    textColor: 'B91C1C'
  });
  ```

---

### `alignShapeToCell(shapeId, tableId, rowIndex, colIndex, alignment)`
- **Signature**: `alignShapeToCell(shapeId: string, tableId: string, rowIndex: number, colIndex: number, alignment: 'center' | 'left' | 'right' | 'top' | 'bottom'): this`
- **Return Type**: `this` (Chainable).
- **Example**:
  ```javascript
  ppt.alignShapeToCell('CustomIcon', 'Table1', 1, 0, 'center');
  ```


---

## 4.8 Media, Images & Aspect-Ratio Preservation

### `addImage(sourcePathOrBuffer, options)`
- **Signature**: `async addImage(sourcePathOrBuffer: string | Buffer | Uint8Array, options: { x: number | string, y: number | string, width?: number | string, height?: number | string, name?: string, lockAspectRatio?: boolean }): Promise<string>`
- **Return Type**: `Promise<string>` (The created picture ID).
- **OpenXML & ZIP Impact**:
  - Ingests image binary into `ppt/media/imageM.ext` in VFS.
  - Generates SHA-1 hash to deduplicate against identical existing media assets.
  - Adds `<Override>` to `[Content_Types].xml` (e.g. `image/png`, `image/jpeg`, `image/svg+xml`, `image/webp`).
  - Creates `<Relationship>` in `ppt/slides/_rels/slideN.xml.rels` with type `http://schemas.openxmlformats.org/officeDocument/2006/relationships/image`.
  - Inserts `<p:pic>` element in `<p:spTree>` containing `<p:blipFill><a:blip r:embed="rIdX"/></p:blipFill>` and bounding `<a:xfrm>`.
- **Constraints & Edge Cases**:
  - Automatically parses intrinsic image dimensions (PNG, JPEG, GIF, BMP, SVG, WebP) to preserve native aspect ratio if only `width` or `height` is supplied.
  - Supports Base64 data URIs (`data:image/png;base64,...`).
- **Example**:
  ```javascript
  const picId = await ppt.useSlide(1).addImage('./assets/company-logo.png', {
    x: '0.8in',
    y: '0.8in',
    width: '2.5in',
    lockAspectRatio: true
  });
  ```

---

### `replaceImage(imageIdOrName, sourcePathOrBuffer)`
- **Signature**: `async replaceImage(imageIdOrName: string, sourcePathOrBuffer: string | Buffer | Uint8Array): Promise<this>`
- **Return Type**: `Promise<this>` (Chainable).
- **OpenXML & ZIP Impact**:
  - Locates existing `<p:pic>` node matching `imageIdOrName`.
  - Extracts the relationship ID from `<a:blip @_r:embed="...">`.
  - Swaps the binary target in `ppt/media/` while retaining original slide coordinates, rotation, and Z-index position.
  - Re-computes bounding box using `object-fit: contain` logic to prevent image distortion.
- **Example**:
  ```javascript
  await ppt.replaceImage('ExecutivePhoto', './assets/new-headshot.jpg');
  ```

---

### `replaceShape(shapeIdOrName, sourcePathOrBuffer)`
- **Signature**: `async replaceShape(shapeIdOrName: string, sourcePathOrBuffer: string | Buffer): Promise<this>`
- **Return Type**: `Promise<this>` (Chainable).
- **OpenXML & ZIP Impact**: Replaces a vector `<p:sp>` placeholder shape with an image `<p:pic>` node inside the exact same bounding box geometry, allocating a new media relationship.
- **Example**:
  ```javascript
  await ppt.replaceShape('GreyPlaceholderBox', './assets/product-hero.webp');
  ```

---

### `removeImage(imageIdOrName)`
- **Signature**: `removeImage(imageIdOrName: string): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Removes `<p:pic>` node from `<p:spTree>` and purges orphaned relationship entry from `slideN.xml.rels`.
- **Example**:
  ```javascript
  ppt.removeImage('ObsoleteLogo');
  ```

---

### `getImages()`
- **Signature**: `getImages(): Array<{ id: string, name: string, rId: string, target: string, bounds: Object }>`
- **Return Type**: `Array<Object>`
- **Example**:
  ```javascript
  const images = ppt.getImages();
  ```

---

## 4.9 Hyperlinks & Interactive Navigation

### `addHyperlink(options)`
- **Signature**: `addHyperlink(options: { shapeId?: string, shapeName?: string, text?: string, url: string, tooltip?: string }): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**:
  - Creates `<Relationship>` in `ppt/slides/_rels/slideN.xml.rels` with type `http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink` and `TargetMode="External"`.
  - Attaches `<a:hlinkClick @_r:id="rIdX" @_tooltip="..."/>` to the target `<a:rPr>` or `<p:cNvPr>`.
- **Example**:
  ```javascript
  ppt.useSlide(1).addHyperlink({
    text: 'Acme Developer Portal',
    url: 'https://developer.acme.com',
    tooltip: 'Visit Developer Docs'
  });
  ```

---

### `addSlideLink(options)` / `addShapeNavigationLink(options)` / `addTextNavigationLink(options)`
- **Signature**: `addSlideLink(options: { shapeId?: string, shapeName?: string, text?: string, targetSlide: number | 'next' | 'previous' | 'first' | 'last' }): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**:
  - Resolves target slide relationship or action URI (`ppaction://hlinksldjump`).
  - Binds click navigation action `<a:hlinkClick @_action="ppaction://hlinksldjump" @_r:id="rIdX"/>`.
- **Example**:
  ```javascript
  // Add a button that jumps to Slide 5:
  ppt.addSlideLink({
    shapeId: 'JumpButton',
    targetSlide: 5
  });
  ```

---

### `getHyperlinks(slideIndex)`
- **Signature**: `getHyperlinks(slideIndex?: number): Array<{ rId: string, url: string, targetType: 'external' | 'internal', targetSlide?: number }>`
- **Return Type**: `Array<Object>`
- **Example**:
  ```javascript
  const links = ppt.getHyperlinks(1);
  ```

---

### `removeHyperlink(textOrOptions)`
- **Signature**: `removeHyperlink(textOrOptions: string | { shapeId?: string, text?: string }): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Strips `<a:hlinkClick>` tags from runs/shapes and removes unreferenced entries from `slideN.xml.rels`.
- **Example**:
  ```javascript
  ppt.removeHyperlink({ text: 'Acme Developer Portal' });
  ```

---

## 4.10 Layer Stacking & Visual Z-Order

Visual Z-order in PowerPoint is determined strictly by document order inside `<p:spTree>`.

### `bringToFront(objectIdOrName)`
- **Signature**: `bringToFront(objectIdOrName: string): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Moves the target element node to the very end of `<p:spTree>`, making it render on top of all other slide objects.
- **Example**:
  ```javascript
  ppt.useSlide(1).bringToFront('StatusBadge');
  ```

---

### `sendToBack(objectIdOrName)`
- **Signature**: `sendToBack(objectIdOrName: string): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Moves the target element node to the beginning of `<p:spTree>` (directly after non-visual group properties `<p:nvGrpSpPr>` and `<p:grpSpPr>`), placing it behind all other objects.
- **Example**:
  ```javascript
  ppt.sendToBack('CardBackground');
  ```

---

### `bringForward(objectIdOrName)` / `stepForward(objectIdOrName)`
- **Signature**: `bringForward(objectIdOrName: string): this`  
  `stepForward(objectIdOrName: string): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Swaps the target element node with its subsequent sibling node in `<p:spTree>`, advancing its visual layer by 1.
- **Example**:
  ```javascript
  ppt.stepForward('LogoImage');
  ```

---

### `sendBackward(objectIdOrName)` / `stepBackward(objectIdOrName)`
- **Signature**: `sendBackward(objectIdOrName: string): this`  
  `stepBackward(objectIdOrName: string): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Swaps the target element node with its preceding visual sibling in `<p:spTree>`, lowering its visual layer by 1.
- **Example**:
  ```javascript
  ppt.stepBackward('OverlayBox');
  ```

---

### `setZIndex(objectIdOrName, zIndex)`
- **Signature**: `setZIndex(objectIdOrName: string, zIndex: number): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Repositions the node to index `zIndex` within `<p:spTree>`.
- **Example**:
  ```javascript
  ppt.setZIndex('TitleText', 0); // Bottom layer
  ```

---

### `reorderObjects(slideIndex, objectIds)` / `applyZOrder(slideNumber, zOrderConfig)`
- **Signature**: `reorderObjects(slideIndex: number, objectIds: string[]): this`  
  `applyZOrder(slideNumber: number, zOrderConfig: Array<{ id: string, zIndex: number }>): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Re-sorts all child nodes in `<p:spTree>` according to the specified array sequence.
- **Example**:
  ```javascript
  ppt.reorderObjects(1, ['Background', 'DataGrid', 'StatusBadge', 'HeaderTitle']);
  ```

---

### `swapObjects(idA, idB)`
- **Signature**: `swapObjects(idA: string, idB: string): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Swaps the tree positions of elements `idA` and `idB`.
- **Example**:
  ```javascript
  ppt.swapObjects('Chart1', 'Chart2');
  ```

---

### `moveObjectBefore(targetId, referenceId)` / `moveObjectAfter(targetId, referenceId)`
- **Signature**: `moveObjectBefore(targetId: string, referenceId: string): this`  
  `moveObjectAfter(targetId: string, referenceId: string): this`
- **Return Type**: `this` (Chainable).
- **Example**:
  ```javascript
  ppt.moveObjectBefore('LegendBox', 'MainChart');
  ```

---

### `getObjectOrder(slideIndex)` / `getTopMostObject(slideIndex)` / `getBottomMostObject(slideIndex)`
- **Signature**: `getObjectOrder(slideIndex?: number): Array<{ id: string, name: string, type: string, zIndex: number }>`  
  `getTopMostObject(slideIndex?: number): Object`  
  `getBottomMostObject(slideIndex?: number): Object`
- **Return Type**: `Array<Object>` / `Object`
- **Example**:
  ```javascript
  const order = ppt.getObjectOrder(1);
  console.log('Top element is:', ppt.getTopMostObject(1).name);
  ```

---

### `normalizeZOrder(slideIndex)`
- **Signature**: `normalizeZOrder(slideIndex?: number): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Ensures required container metadata nodes (`<p:nvGrpSpPr>`, `<p:grpSpPr>`) remain at index 0 and 1, normalizing any corrupted tree sequences.
- **Example**:
  ```javascript
  ppt.normalizeZOrder(1);
  ```


---

## 4.11 Inspection, Diagnostics, Schema Validation & Repair

### `getElements(filter)` / `findElements(filterOrPredicate)`
- **Signature**: `getElements(filter?: { type?: string }): Array<{ id: string, name: string, type: 'text'|'shape'|'image'|'table'|'chart'|'group', bounds: Object }>`  
  `findElements(filterOrPredicate?: Function | Object): Array<Object>`
- **Return Type**: `Array<Object>`
- **OpenXML & ZIP Impact**: Traverses `<p:spTree>` on active slides and returns structured metadata.
- **Example**:
  ```javascript
  const elements = ppt.useSlide(1).getElements();
  const tables = elements.filter(e => e.type === 'table');
  ```

---

### `getTables()` / `getCharts()` / `getTextElements()` / `getShapes()`
- **Signature**: `getTables(): Array<Object>`  
  `getCharts(): Array<Object>`  
  `getTextElements(): Array<Object>`  
  `getShapes(): Array<Object>`
- **Return Type**: `Array<Object>`
- **Example**:
  ```javascript
  const charts = ppt.useSlide(1).getCharts();
  console.log(`Slide 1 contains ${charts.length} native charts.`);
  ```

---

### `getInfo()`
- **Signature**: `getInfo(): { slideCount: number, dimensions: Object, themes: string[], layouts: string[], chartsCount: number, tablesCount: number, imagesCount: number }`
- **Return Type**: `Object`
- **Example**:
  ```javascript
  const summary = ppt.getInfo();
  console.log(summary);
  ```

---

### `validate()` / `validatePresentation()` / `validatePresentationXml()` / `validateArchive()`
- **Signature**: `async validate(): Promise<{ valid: boolean, errors: Array<Object>, warnings: Array<Object> }>`  
  `async validatePresentation(): Promise<Object>`  
  `async validateArchive(): Promise<Object>`
- **Return Type**: `Promise<Object>`
- **OpenXML & ZIP Impact**:
  - Inspects `[Content_Types].xml`, `_rels/.rels`, and `ppt/_rels/presentation.xml.rels`.
  - Runs schema verification on all slides, ensuring every `rId` reference is satisfied.
  - Verifies table column counts and `<a16:rowId>` uniqueness.
  - Validates chart data point counts and Excel workbook consistency.
- **Example**:
  ```javascript
  const result = await ppt.validate();
  if (!result.valid) {
    console.error('Validation failures:', result.errors);
  }
  ```

---

### `validateSlide(slideIndex)` / `validateTable(tableId)` / `validateShape(shapeId)`
- **Signature**: `async validateSlide(slideIndex: number): Promise<Object>`  
  `async validateTable(tableId: string): Promise<Object>`  
  `async validateShape(shapeId: string): Promise<Object>`
- **Return Type**: `Promise<Object>`
- **Example**:
  ```javascript
  const tableCheck = await ppt.validateTable('FinancialTable');
  console.log('Table integrity:', tableCheck.valid);
  ```

---

### `validateCharts()` / `repairCharts()`
- **Signature**: `async validateCharts(): Promise<Object>`  
  `async repairCharts(): Promise<this>`
- **Return Type**: `Promise<Object>` / `Promise<this>` (Chainable).
- **OpenXML & ZIP Impact**: Validates and synchronizes all chart DrawingML definitions against embedded `.xlsx` workbooks, resolving discrepancies in formula ranges or point counts.
- **Example**:
  ```javascript
  await ppt.repairCharts();
  ```

---

### `repair()`
- **Signature**: `async repair(): Promise<this>`
- **Return Type**: `Promise<this>` (Chainable).
- **OpenXML & ZIP Impact**:
  - Automatically fixes missing relationship entries.
  - Generates missing `<a16:rowId>` table row IDs.
  - Adds missing shadowed `hMerge`/`vMerge` cells to unbalanced tables.
  - Rounds floating-point coordinates to integer EMUs.
- **Example**:
  ```javascript
  await ppt.repair();
  ```

---

### `inspectXML(xmlPath)` / `inspectSlide(slideIndex)` / `inspectChart(chartId)` / `inspectChartXML(chartFileName)`
- **Signature**: `async inspectXML(xmlPath: string): Promise<string>`  
  `async inspectSlide(slideIndex: number): Promise<string>`  
  `async inspectChart(chartId: string): Promise<string>`  
  `async inspectChartXML(chartFileName: string): Promise<string>`
- **Return Type**: `Promise<string>` (Pretty-printed XML string).
- **Example**:
  ```javascript
  const slide1Xml = await ppt.inspectSlide(1);
  console.log(slide1Xml.slice(0, 500));
  ```

---

### `debugRelationships()` / `debugChartRelationships()`
- **Signature**: `debugRelationships(): Array<{ source: string, rId: string, type: string, target: string, targetExists: boolean }>`  
  `debugChartRelationships(): Array<Object>`
- **Return Type**: `Array<Object>`
- **OpenXML & ZIP Impact**: Audits every `.rels` file in the package and flags dangling or broken targets.
- **Example**:
  ```javascript
  const relAudit = ppt.debugRelationships();
  console.table(relAudit.filter(r => !r.targetExists));
  ```

---

### `enableDebug()` / `enableDebugZip()`
- **Signature**: `enableDebug(): this`  
  `enableDebugZip(outputPath?: string): this`
- **Return Type**: `this` (Chainable).
- **OpenXML & ZIP Impact**: Activates verbose AST logging and instructs `OutputWriter` to preserve intermediate XML diagnostics on disk.
- **Example**:
  ```javascript
  ppt.enableDebug();
  ```

---

## 4.12 Export Adapters & Multi-Target Serialization

### `saveToFile(filePath, options)` / `toFile(filePath, options)` / `save(filePath, options)`
- **Signature**: `async saveToFile(filePath: string, options?: { compressionLevel?: number }): Promise<void>`  
  `async toFile(filePath: string, options?: object): Promise<void>`  
  `async save(filePath: string, options?: object): Promise<void>`
- **Return Type**: `Promise<void>`
- **OpenXML & ZIP Impact**: Commits all modified ASTs back into the VFS, compresses the package using JSZip (applying `compressionLevel` 1–9 or `STORE`), and writes the binary archive to local disk.
- **Example**:
  ```javascript
  await ppt.saveToFile('./output/Quarterly_Report.pptx', { compressionLevel: 6 });
  ```

---

### `toBuffer(options)`
- **Signature**: `async toBuffer(options?: { compressionLevel?: number }): Promise<Buffer>`
- **Return Type**: `Promise<Buffer>` (Node.js Buffer).
- **OpenXML & ZIP Impact**: Asynchronously compiles the VFS into a complete `.pptx` ZIP buffer entirely in memory. Zero disk I/O.
- **Constraints & Edge Cases**: Ideal for AWS Lambda, Google Cloud Functions, or microservices returning file streams.
- **Example**:
  ```javascript
  const pptxBuffer = await ppt.toBuffer();
  console.log(`Generated buffer of size: ${pptxBuffer.length} bytes`);
  ```

---

### `toArrayBuffer(options)`
- **Signature**: `async toArrayBuffer(options?: object): Promise<ArrayBuffer>`
- **Return Type**: `Promise<ArrayBuffer>`
- **OpenXML & ZIP Impact**: Returns a web-standard `ArrayBuffer` suitable for Cloudflare Workers, browser fetch, or Blob instantiation.
- **Example**:
  ```javascript
  const arrayBuffer = await ppt.toArrayBuffer();
  ```

---

### `toBase64(options)`
- **Signature**: `async toBase64(options?: { dataUri?: boolean }): Promise<string>`
- **Return Type**: `Promise<string>`
- **OpenXML & ZIP Impact**:
  - If `dataUri: false` (default): Returns raw Base64 string.
  - If `dataUri: true`: Prepend MIME prefix: `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,...`.
- **Example**:
  ```javascript
  const dataUri = await ppt.toBase64({ dataUri: true });
  // Ready for frontend download links: <a href={dataUri} download="report.pptx">
  ```

---

### `toExpress(res, filename, options)`
- **Signature**: `async toExpress(res: any, filename?: string, options?: object): Promise<void>`
- **Return Type**: `Promise<void>`
- **OpenXML & ZIP Impact**:
  - Automatically sets headers:
    - `Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation`
    - `Content-Disposition: attachment; filename="${filename || 'presentation.pptx'}"`
  - Generates the package buffer and streams it to the Express HTTP response object.
- **Example**:
  ```javascript
  import express from 'express';
  import { PPTXTemplater } from 'node-pptx-templater';

  const app = express();

  app.get('/download-report', async (req, res) => {
    const ppt = await PPTXTemplater.load('templates/report.pptx');
    ppt.useSlide(1).replaceText({ '{{USER}}': req.query.name || 'Valued Client' });
    await ppt.toExpress(res, 'Personalized_Report.pptx');
  });
  ```

---

### `toStream(options)` / `saveToStream(writableOrOptions, options)`
- **Signature**: `async toStream(options?: object): Promise<ReadableStream>`  
  `async saveToStream(writable: WritableStream | NodeJS.WritableStream, options?: object): Promise<void>`
- **Return Type**: `Promise<ReadableStream>` / `Promise<void>`
- **OpenXML & ZIP Impact**: Streams compressed package bytes in chunks, minimizing peak memory usage when exporting large presentations to object storage (AWS S3, Google Cloud Storage, Azure Blob).
- **Example**:
  ```javascript
  import { Upload } from '@aws-sdk/lib-storage';
  import { S3Client } from '@aws-sdk/client-s3';

  const stream = await ppt.toStream();
  const upload = new Upload({
    client: new S3Client(),
    params: {
      Bucket: 'my-corp-reports',
      Key: 'reports/q3.pptx',
      Body: stream,
      ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    }
  });
  await upload.done();
  ```

---

### `saveXml(folderPath)` / `saveToFolder(folderPath)`
- **Signature**: `async saveXml(folderPath: string): Promise<void>`  
  `async saveToFolder(folderPath: string): Promise<void>`
- **Return Type**: `Promise<void>`
- **OpenXML & ZIP Impact**: Extracts and writes all VFS files directly to `folderPath` as an uncompressed directory tree (`[Content_Types].xml`, `_rels/`, `ppt/slides/`, `ppt/charts/`, etc.).
- **Example**:
  ```javascript
  // Extract modified presentation to a folder for Git inspection:
  await ppt.saveToFolder('./debug-presentation-folder');
  ```


---

# 5. CLI, Logging, Tooling & Performance Profiling

`node-pptx-templater` includes a full suite of development tooling, diagnostic commands, a non-intrusive logging architecture, and built-in execution profilers designed for mission-critical enterprise environments.

---

## 5.1 Command-Line Interface (CLI) Commands

The library provides a standalone CLI binary registered in `package.json` under `bin: { "node-pptx-templater": "src/cli/index.js" }`.

### Global Installation & Direct Invocation
```bash
# Global installation
npm install -g node-pptx-templater

# Direct npx execution
npx node-pptx-templater --help
```

---

### 1. `build` Command
Builds a `.pptx` presentation from a template by injecting data from a JSON file.

- **Syntax**: `node-pptx-templater build <template> <output> [options]`
- **Options**:
  - `-d, --data <file>`: Path to JSON file containing template replacement values.
  - `-s, --slide <numbers>`: Comma-separated slide numbers to include in the output (e.g. `1,3,5`).
  - `--no-banner`: Suppresses the CLI startup banner.
- **Example**:
  ```bash
  # data.json: { "{{company}}": "Acme Corp", "{{year}}": "2026" }
  node-pptx-templater build template.pptx output.pptx -d data.json -s 1,2
  ```

---

### 2. `validate` Command
Validates the Open Packaging Conventions (OPC), relationship graphs, and DrawingML schema integrity of a PowerPoint file.

- **Syntax**: `node-pptx-templater validate <file> [options]`
- **Options**:
  - `--strict`: Exits with non-zero error code on warnings as well as errors.
- **Output**: Reports relationship graph health, table matrix checks, `<a16:rowId>` uniqueness, and chart workbook sync status.
- **Example**:
  ```bash
  node-pptx-templater validate generated_deck.pptx --strict
  ```

---

### 3. `inspect` Command
Inspects the internal OpenXML architecture of a `.pptx` presentation, printing structured summaries to stdout.

- **Syntax**: `node-pptx-templater inspect <file> [options]`
- **Options**:
  - `--slides`: Displays slide count, layout references, and slide dimensions.
  - `--charts`: Lists embedded charts, chart types, series names, and workbook bindings.
  - `--tables`: Summarizes table IDs, row/column dimensions, and merged cells.
  - `--media`: Lists embedded images, formats, dimensions, and SHA-1 hashes.
  - `--rels`: Prints full relationship dependency trees.
  - `--all`: Dumps full inspection report.
- **Example**:
  ```bash
  node-pptx-templater inspect presentation.pptx --charts --tables
  ```

---

### 4. `extract` Command
Extracts individual XML parts or media assets from a `.pptx` container for inspection or version control.

- **Syntax**: `node-pptx-templater extract <file> [options]`
- **Options**:
  - `-s, --slide <number>`: Slide number whose raw XML should be extracted.
  - `-o, --out <path>`: Destination output path (defaults to stdout if omitted).
  - `--chart <name>`: Extracts a specific chart's DrawingML markup.
  - `--rels`: Extracts relationship XML files.
- **Example**:
  ```bash
  node-pptx-templater extract deck.pptx -s 1 -o ./slide1_raw.xml
  ```

---

### 5. `debug` Command
Analyzes a corrupted or damaged presentation and optionally attempts automated repair.

- **Syntax**: `node-pptx-templater debug <file> [options]`
- **Options**:
  - `--fix`: Attempts automatic repairs (re-aligns tables, fixes missing `rIds`, re-syncs charts).
  - `-o, --out <path>`: Specifies destination file path for the repaired presentation.
- **Example**:
  ```bash
  node-pptx-templater debug broken.pptx --fix -o repaired.pptx
  ```

---

## 5.2 Logging Subsystem & Runtime Telemetry

### Silent Default Principle
In serverless functions, HTTP APIs, and command-line pipelines, extraneous `console.log` noise pollutes stdout and degrades performance.  
**Rule**: `node-pptx-templater` is **completely silent by default**. It produces zero console output unless explicitly configured.

### Log Levels
The logging engine (`src/utils/logger.js`) supports six granular log levels:
- `'verbose'`: Logs fine-grained AST tokenization, run healing steps, and ZIP offset allocations.
- `'debug'`: Logs element queries, coordinate conversions, and relationship mapping resolutions.
- `'info'`: Logs slide selections, table insertions, chart synchronizations, and save operations.
- `'warn'`: Logs missing optional placeholders, unused relationships, or non-fatal fallback behavior.
- `'error'`: Logs schema violations, missing mandatory parts, and security rejections.
- `'silent'`: Completely suppresses all log output (Default).

### Configuring Log Levels

#### 1. At Instance Initialization
```javascript
const ppt = await PPTXTemplater.load('template.pptx', { logLevel: 'debug' });
```

#### 2. Globally at Runtime
```javascript
import { PPTXTemplater, setGlobalLogLevel, resetLogLevel } from 'node-pptx-templater';

// Activate verbose logging across all instances:
PPTXTemplater.setLogLevel('verbose');

// Or via utility helper:
setGlobalLogLevel('info');

// Reset back to silent default:
resetLogLevel();
```

---

## 5.3 High-Throughput Preloading & Template Caching

In high-concurrency environments (e.g. generating hundreds of invoices or investor updates per second), re-reading template `.pptx` files from disk and decompressing ZIP containers for every request introduces unnecessary CPU and I/O overhead.

`node-pptx-templater` provides an in-memory preloading and cloning engine:

```
┌────────────────────────────────────────────────────────────────────────┐
│                        SERVER BOOTSTRAP PHASE                          │
│                                                                        │
│   await PPTXTemplater.preload('templates/report.pptx');                │
│   └── Ingests, unzips, and stores pristine VFS snapshot in RAM         │
└──────────────────────────────────┬─────────────────────────────────────┘
                                   │
                                   ▼
┌────────────────────────────────────────────────────────────────────────┐
│                        REQUEST EXECUTION PHASE                         │
│                                                                        │
│   Request 1 ──► PPTXTemplater.fromCache('templates/report.pptx')       │
│                 └── Instant in-memory deep-clone (< 1ms) ──► Mutate    │
│                                                                        │
│   Request 2 ──► PPTXTemplater.fromCache('templates/report.pptx')       │
│                 └── Instant in-memory deep-clone (< 1ms) ──► Mutate    │
└────────────────────────────────────────────────────────────────────────┘
```

### Usage Pattern
```javascript
import express from 'express';
import { PPTXTemplater } from 'node-pptx-templater';

const app = express();
const TEMPLATE_PATH = './templates/Executive_Template.pptx';

// 1. Preload during server startup:
await PPTXTemplater.preload(TEMPLATE_PATH);
console.log('Template preloaded into memory.');

// 2. Serve requests with near-instant instantiation:
app.post('/generate-deck', async (req, res) => {
  // Clones the pristine VFS snapshot in < 1ms:
  const ppt = await PPTXTemplater.fromCache(TEMPLATE_PATH);

  ppt.useSlide(1).replaceText({
    '{{NAME}}': req.body.userName,
    '{{DATE}}': new Date().toISOString().slice(0, 10)
  });

  await ppt.toExpress(res, 'Personalized_Deck.pptx');
});
```

---

## 5.4 Built-In Performance Profiling & Metrics

`PPTXTemplater` includes an integrated telemetry profiler utilizing Node.js `perf_hooks.performance`.

### Enabling Profiling
```javascript
const ppt = await PPTXTemplater.load('template.pptx');
ppt.enablePerformanceProfile();

// Perform operations...
ppt.useSlide(1);
ppt.replaceText({ '{{title}}': 'High-Throughput Benchmark' });
await ppt.updateChart('Chart1', chartData);
await ppt.splitTableOnOverflow('DataTable', { bottomMargin: '1in' });
const buffer = await ppt.toBuffer();

// Retrieve granular metrics:
const metrics = ppt.getPerformanceMetrics();
console.table(metrics.stages);
console.log(`Total Execution Time: ${metrics.totalDurationMs.toFixed(2)} ms`);
```

### Profiling Report Structure
```json
{
  "totalDurationMs": 14.82,
  "stages": [
    { "stage": "package_load", "durationMs": 3.12 },
    { "stage": "ast_parse", "durationMs": 1.45 },
    { "stage": "run_fragment_healing", "durationMs": 0.82 },
    { "stage": "text_replacement", "durationMs": 0.41 },
    { "stage": "chart_dual_sync", "durationMs": 4.19 },
    { "stage": "table_pagination", "durationMs": 2.11 },
    { "stage": "zip_serialization", "durationMs": 2.72 }
  ],
  "memory": {
    "heapUsedBeforeMb": 42.15,
    "heapUsedAfterMb": 45.30,
    "heapDeltaMb": 3.15
  }
}
```

---

## Summary of Reference Documents

| Resource | Description |
| :--- | :--- |
| [`GEMINI_NOTEBOOK_SOURCE.md`](file:///e:/Code/pptx-templater/GEMINI_NOTEBOOK_SOURCE.md) | **This document**: Complete, monolithic technical reference designed for Gemini Notebook ingestion. |
| [`.agent/project_summary.md`](file:///e:/Code/pptx-templater/.agent/project_summary.md) | High-level architectural overview and manager catalog for AI autonomous agents. |
| [`.agent/rules/features.md`](file:///e:/Code/pptx-templater/.agent/rules/features.md) | Domain-specific rules for OpenXML `<a16:rowId>`, chart dual-sync, and cell spans. |
| [`.agent/rules/general.md`](file:///e:/Code/pptx-templater/.agent/rules/general.md) | Engineering principles, pure-JS constraints, and anti-XXE security standards. |
| [`.agent/skills/chart-sync.md`](file:///e:/Code/pptx-templater/.agent/skills/chart-sync.md) | Deep guide on DrawingML `<c:chartSpace>` and embedded Excel synchronization. |
| [`.agent/skills/table-merging.md`](file:///e:/Code/pptx-templater/.agent/skills/table-merging.md) | Implementation details for `gridSpan`, `rowSpan`, and shadowed placeholder cells. |
| [`showcase/examples.js`](file:///e:/Code/pptx-templater/showcase/examples.js) | 12 end-to-end executable demonstrations covering every library feature. |
