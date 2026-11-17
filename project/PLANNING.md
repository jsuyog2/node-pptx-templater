# PPTX Engine — Architectural Specification & Product Plan

A high-performance Node.js library for programmatic PowerPoint (.pptx) template manipulation, spatial querying, dynamic table pagination, and multi-target export.

---

## 1. System Architecture


```

┌────────────────────────────────────────────────────────────────────────────────────────┐
│                                INGESTION LAYER                                         │
│  ┌─────────────────────────┐  ┌──────────────────────────┐  ┌────────────────────────┐ │
│  │ Direct PPTX (Buf/Path)  │  │ ZIP Package (Any Ext)    │  │ Raw XML Directory Tree │ │
│  └───────────┬─────────────┘  └────────────┬─────────────┘  └───────────┬────────────┘ │
└──────────────┼─────────────────────────────┼────────────────────────────┼──────────────┘
▼                             ▼                            ▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                              VIRTUAL OPENXML DOM (CORE)                                │
│  ┌──────────────────────────────────────────────────────────────────────────────────┐  │
│  │ Presentation Model: [Content_Types].xml, ppt/presentation.xml, Layout & Themes   │  │
│  └────────────────────────────────────────┬─────────────────────────────────────────┘  │
│                                           ▼                                            │
│  ┌──────────────────────────────────────────────────────────────────────────────────┐  │
│  │ Slide Graph: Slides, Relationships (_rels), Slide Masters, Layout References     │  │
│  └───────────────────┬───────────────────────────────┬──────────────────────────────┘  │
│                      ▼                               ▼                                 │
│  ┌─────────────────────────────────────┐  ┌─────────────────────────────────────────┐  │
│  │ Spatial & Element Tree              │  │ Table & Overflow Layout Engine          │  │
│  │ • X/Y/W/H/Rotation Normalizer (EMU) │  │ • Font & Cell Dimension Estimator       │  │
│  │ • Text, Shapes, Media & Native Chart│  │ • Pagination & Slide-Clone Dispatcher   │  │
│  └─────────────────────────────────────┘  └─────────────────────────────────────────┘  │
└───────────────────────────────────────────┬────────────────────────────────────────────┘
▼
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                                SERIALIZATION & EXPORT                                  │
│  ┌───────────────┬──────────────┬───────────────┬───────────────┬───────────────────┐  │
│  │  File System  │ Node Stream  │ Memory Buffer │ ArrayBuffer   │ Base64 / Data URI │  │
│  │  (toFile)     │ (toStream)   │ (toBuffer)    │ (toArrayBuf)  │ (toBase64)        │  │
│  └───────────────┴──────────────┴───────────────┴───────────────┴───────────────────┘  │
│  ┌──────────────────────────────────────────────────────────────────────────────────┐  │
│  │ Express.js HTTP Response Pipe (`toExpress(res, filename)`)                       │  │
│  └──────────────────────────────────────────────────────────────────────────────────┘  │
└────────────────────────────────────────────────────────────────────────────────────────┘

```

---

## 2. Runtime Targets & Compatibility

* **Node.js**: Minimum `Node.js 18.x` LTS, compatible through modern Node versions (`20.x`, `22.x+`).
* **Bundle Formats**: Dual-build **ESM** (`import`) and **CommonJS** (`require`) via `tsup`.
* **TypeScript**: Native TypeScript codebase emitting `.d.ts` declaration files with strict mode enabled.
* **Server Frameworks**: Seamless integration with Express.js, Fastify, NestJS, AWS Lambda, and standard Node HTTP servers.

---

## 3. Dependencies & Tech Stack

| Dependency | Category | Purpose |
| :--- | :--- | :--- |
| **`fflate`** or **`jszip`** | Runtime | Fast, low-memory ZIP archive decompression and packing. |
| **`fast-xml-parser`** | Runtime | High-throughput XML parsing and serialization with attribute preservation. |
| **`mime-types`** | Runtime | MIME resolution when embedding/replacing images and media. |
| **`typescript`** | DevDependency | Strict type definitions and type safety. |
| **`tsup`** | DevDependency | Fast esbuild-based dual bundler (ESM + CJS). |
| **`vitest`** | DevDependency | Unit and snapshot testing engine. |
| **`@types/express`** | Optional/Peer | Type definitions for direct Express response piping. |

---

## 4. Implementation & Execution Plan

### Phase 1: Ingestion & Zip/XML Normalizer
* Build `PackageLoader` to detect input source: buffer, file path, stream, or directory path.
* Normalize inputs into an in-memory virtual file system (`Map<string, string | Uint8Array>`).
* Support direct `.pptx`, `.zip` archives (reading binary ZIP signatures), and raw extracted XML directories.
* Configure XML parser to preserve OpenXML namespaces (`p:`, `a:`, `r:`, `c:`).

### Phase 2: Slide Graph & Spatial Coordinate Engine
* Parse `ppt/presentation.xml` and `ppt/_rels/presentation.xml.rels`.
* Implement OpenXML unit conversion helpers:
  * $1 \text{ pt} = 12700 \text{ EMUs}$
  * $1 \text{ in} = 914400 \text{ EMUs}$
  * $1 \text{ px (96 DPI)} = 9525 \text{ EMUs}$
* Build spatial query selectors (`slide.findElementByName()`, `slide.getElements()`).
* Implement Slide deep-cloning: clone XML, regenerate unique relationship IDs (`rId`), update `[Content_Types].xml`, and register new slide entries in presentation manifests.

### Phase 3: Shape, Image, Chart & Table Manipulators
* **Text & Shapes**: Target `<p:sp>`; modify `<a:t>`, fill styles, font sizes, colors, and line spacing.
* **Images**: Target `<p:pic>`; manage `ppt/media/*` binary replacements and relationship updates in `slideX.xml.rels`.
* **Charts**: Target `<c:chartSpace>`; mutate series data and categories, maintaining optional Excel sheet streams.
* **Tables**: Build and mutate `<a:tbl>`, rows `<a:tr>`, and cells `<a:tc>`; compute cell spans for merging (`gridSpan`, `rowSpan`).

### Phase 4: Dynamic Table Overflow & Pagination Engine
* Develop text measurement logic estimating cell heights from font size, line count, cell padding, and column width.
* Compute available slide vertical space ($Y_{\text{limit}} = \text{SlideHeight} - \text{BottomMargin}$).
* When cumulative row height exceeds $Y_{\text{limit}}$:
  1. Deep-clone the host slide.
  2. Split table rows at the break point.
  3. Render fitting rows on Slide $N$, duplicate headers onto Slide $N+1$, and repeat until all rows are placed.

### Phase 5: Export Adapters & Production Polish
* Implement target serializers: File, Stream, Buffer, ArrayBuffer, and Base64.
* Implement `toExpress(res, filename)` with automatic `Content-Disposition`, `Content-Type`, and stream piping.
* Run compatibility validation against Microsoft PowerPoint and LibreOffice.

---

## 5. Complete Feature List

### A. Ingestion & Input Flexibility
1. **Direct PPTX Import**: Load `.pptx` documents via file path, Buffer, or Stream.
2. **ZIP Container Import**: Ingest presentations saved or renamed with `.zip` extensions.
3. **Unpacked XML Directory Ingestion**: Ingest unzipped OpenXML directory structures directly.
4. **Blank Deck Generation**: Create presentations from scratch with customizable aspect ratios (16:9 widescreen or 4:3 standard).

### B. Slide Management & Inspection
5. **Slide Lifecycle (CRUD)**: Add blank slides, remove slides by index/ID, and reorder slide positions.
6. **Deep Slide Duplication**: Clone existing slides with all shapes, layouts, relationships, and styles preserved.
7. **Deck Metadata Retrieval**: Extract slide count, slide dimensions (in EMUs, points, or pixels), and active theme metadata.
8. **Slide Element Tree Querying**: Fetch all slide elements, filtering by type (`text`, `shape`, `image`, `table`, `chart`, or `group`).

### C. Spatial & Layout Engine
9. **Coordinate Reading**: Retrieve computed `x`, `y`, `width`, `height`, and `rotation` for any slide element.
10. **Spatial Transformation**: Reposition and resize any element with automatic unit conversions.
11. **Z-Index Ordering**: Reorder visual layers (bring to front, send to back, move forward, move backward).

### D. Text, Shapes & Media
12. **Text Manipulation**: In-place replacement, insertion, font size, color, font family, alignment, and formatting (bold, italic, underline).
13. **Shape Engine**: Add standard geometric shapes, customize fill colors/transparency, and adjust stroke weights.
14. **Text Inside Shapes**: Write and style text directly inside any shape with vertical and horizontal alignment.
15. **Image In-Place Replacement**: Replace template placeholder images while preserving bounding boxes and aspect ratios.
16. **Media Addition**: Insert new PNG, JPEG, SVG, or WebP images onto any slide.
17. **Hyperlinks**: Attach external URLs or internal slide-jump navigation links to text, shapes, or images.

### E. Table Engine & Pagination (Flagship)
18. **Table Structure Management**: Create tables and add, remove, or update rows, columns, and cells.
19. **Cell Merging & Styling**: Horizontal/vertical cell merge/unmerge, border controls, cell background fills, and internal padding.
20. **Nested Cell Elements**: Insert shapes, icons, and small indicators directly inside table cells.
21. **Table Metadata Extraction**: Inspect cell coordinates, column widths, row heights, and table bounds.
22. **Auto-Pagination on Overflow**: Automatically compute row heights, duplicate slides upon overflow, carry over table headers, and split data across consecutive slides.

### F. Native Charts
23. **Native Chart Data Mutation**: Update series values, categories, and titles for Bar, Column, Line, Pie, and Area charts.
24. **Chart Layout Control**: Reposition charts and toggle legends, labels, and axes.

### G. Output & Export Pipeline
25. **File System Export (`toFile`)**: Asynchronously save files to local disk paths.
26. **Node.js Stream Export (`toStream`)**: Stream generated files for piping to AWS S3, GCS, or disk streams.
27. **Buffer & ArrayBuffer Export (`toBuffer`, `toArrayBuffer`)**: Low-latency in-memory export for serverless and worker runtimes.
28. **Base64 Export (`toBase64`)**: Generate Base64 strings (or Data URIs) for JSON payloads or webhooks.
29. **Express.js Direct Dispatch (`toExpress`)**: Output directly to Express `res` with appropriate presentation MIME headers and attachment naming.