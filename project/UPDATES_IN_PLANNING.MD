# node-pptx-templater (v2.0) — Modernization & Refactoring Plan

A comprehensive refactoring and architectural blueprint to modernize `node-pptx-templater`, transitioning it from a string-replacement template patcher into a full-featured, spatial-aware PowerPoint manipulation engine.

---

## 1. Executive Summary & Codebase Audit

The current version of `node-pptx-templater` is primarily designed as a basic find-and-replace templating tool:
* Unpacks `.pptx` archives via zip utilities.
* Searches for regex matches (e.g., `{{placeholder}}`) across XML text nodes.
* Replaces cell values, swaps image blobs, and updates chart categories/series cache alongside embedded mini-Excel sheets.

### Current Architectural Bottlenecks
1. **Coupled File I/O:** Dependent on local disk file paths and raw synchronous buffers (`fs.readFileSync`/`writeFileSync`), preventing clean integration with serverless runtimes, microservices, and HTTP response pipelines.
2. **Missing Spatial Engine:** Lacks an AST-level spatial coordinate engine. Cannot calculate bounding boxes, detect collisions, or adjust layout coordinates ($X, Y, W, H$).
3. **No Overflow / Pagination Handling:** Tables that exceed slide bounds overlap headers/footers or clip off the edge.
4. **Fragile Relationship ID Tracking:** Adding, cloning, or deleting slides often breaks internal OpenXML relationships (`rId`), triggering Microsoft PowerPoint "Repair Mode" alerts.

---

## 2. Refactoring Roadmap: Delete, Enhance, and Add


```
            ┌──────────────────────────────────────────────┐
            │          node-pptx-templater (v2.0)          │
            └──────────────────────┬───────────────────────┘
                                   │
     ┌─────────────────────────────┼─────────────────────────────┐
     ▼                             ▼                             ▼



┌───────────────────┐        ┌───────────────────┐        ┌────────────────────┐
│   WHAT TO PURGE   │        │  WHAT TO ENHANCE  │        │   WHAT TO ADD      │
│ • Direct FS lock  │        │ • OpenXML Parser  │        │ • Multi-format In  │
│ • Raw string regex│        │ • Table Engine    │        │ • 5x Export Targets│
│ • Global mutators │        │ • Chart Workbook  │        │ • Spatial Engine   │
│ • Monolithic sync │        │ • Slide Cloner    │        │ • Table Pagination │
└───────────────────┘        └───────────────────┘        └────────────────────┘
```


---

### A. Functions to Delete or Deprecate

| Legacy Method / Pattern | Reason for Deprecation | Modern Replacement |
| :--- | :--- | :--- |
| **`sync` File I/O methods** | Blocks the Node.js event loop during large presentation generation. | Pure async pipeline using `Uint8Array`, Node `ReadableStream`, and web streams. |
| **Direct Regex Node Swapping (`/<p:txBody>.*<\/p:txBody>/`)** | Brittle; splits OpenXML runs (`<a:r>`) and breaks styling when tags span multiple XML runs. | DOM/AST run-reconstruction tokenizer that normalizes fragmented runs prior to replacement. |
| **Path-Only Ingestion (`loadFromFile`)** | Incompatible with cloud functions (AWS Lambda, Cloud Run) and memory-only workflows. | Universal `PackageLoader` supporting paths, buffers, streams, and extracted directories. |
| **Hardcoded `rId` Counters** | Causes ID collision corruption when slides are dynamically duplicated or deleted. | Centralized `RelationshipManager` with deterministic UUID/sequential `rId` hashing. |

---

### B. Functions to Refactor & Enhance

#### 1. Chart Engine & Embedded Spreadsheet Synchronizer
* **Issue:** Updating only `<c:chartSpace>` nodes leaves the embedded Excel file (`ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx`) out of sync, prompting repair warnings in Microsoft 365.
* **Enhancement:** Synchronize both `<c:chartSpace>` data series and the embedded `.xlsx` binary cache synchronously on update.

#### 2. Table Engine (`TableElement`)
* **Issue:** Restricted to basic cell text replacement.
* **Enhancement:**
  * Add multi-directional cell merging (`gridSpan` / `rowSpan`) and unmerging.
  * Support embedding nested elements (icons, mini-indicators, custom shapes) inside table cells (`<a:tc>`).
  * Support programmatic row/column insertion, deletion, and auto-resizing.

#### 3. Slide Cloner (`Slide.duplicate()`)
* **Issue:** Shallow copy of `slideX.xml` that corrupts when shared assets (images, master layouts) are referenced.
* **Enhancement:** Deep-clone pipeline that copies the target slide, isolates its `.rels` file, updates `[Content_Types].xml`, re-registers master relationships, and appends the entry to `ppt/presentation.xml`.

---

### C. Brand New Features to Add

#### 1. Universal Ingestion Module (`PackageLoader`)
Support multiple entry formats natively:
* **Direct PPTX:** Accepts `Buffer`, file path, or stream.
* **ZIP Archive:** Detects PKZip magic bytes (`0x50 0x4B 0x03 0x04`) regardless of file extension (`.zip`, `.pptx`, `.tmp`).
* **Raw OpenXML Directory:** Loads unzipped directory trees into the virtual file system (`Map<string, Uint8Array>`) without requiring an intermediate zip step.

#### 2. Multi-Target Export Engine (`ExportAdapter`)
Expose 5 distinct export interfaces:
```typescript
// 1. Local File System
await preso.toFile('./dist/output.pptx');

// 2. Node.js Readable Stream (for piping to S3/GCS)
const stream = preso.toStream();

// 3. Memory Buffers
const buffer = await preso.toBuffer();
const arrayBuf = await preso.toArrayBuffer();

// 4. Base64 / Data URI
const base64Str = await preso.toBase64({ dataUri: true });

// 5. Direct Express.js HTTP Pipe
await preso.toExpress(res, 'quarterly-report.pptx');

```

#### 3. Spatial Coordinate Engine & Unit Normalizer

Implement bi-directional conversions between points, pixels, inches, and OpenXML EMUs (English Metric Units):


$$\text{EMU} = \text{Points} \times 12700 = \text{Inches} \times 914400 = \text{Pixels (at 96 DPI)} \times 9525$$

* Expose programmatic bounding boxes: `element.getBounds()` returning `{ x, y, width, height, rotation }`.
* Expose transformations: `element.setBounds()`, `element.move()`, `element.resize()`, `element.bringToFront()`, and `element.sendToBack()`.

#### 4. Dynamic Table Pagination Engine (Flagship)

Automate multi-slide pagination when table rows exceed the available slide boundary:

```typescript
slide.splitTableOnOverflow({
  tableId: 'financial-summary-table',
  bottomMargin: '0.75in',
  repeatHeaderOnSplit: true,
  keepSummaryOnLastSlide: true
});

```

**Algorithm Execution:**

1. Calculate the slide's vertical boundary: $Y_{\text{limit}} = \text{SlideHeight} - \text{BottomMargin}$.
2. Traverse rows, calculating estimated height based on font size, line count, and cell padding.
3. If row $R_k$ exceeds $Y_{\text{limit}}$, deep-clone the slide.
4. Truncate the table on the current slide at row $R_{k-1}$.
5. Insert the cloned slide immediately after, inject the table headers, and populate rows $R_k \dots R_n$ recursively.

---

## 3. Step-by-Step Modernization Execution Plan

1. **Restructure Repository & Modern TS Build:**
* Migrate the build system to `tsup` targeting dual ESM and CommonJS outputs.
* Set up strict TypeScript checking, absolute path aliases, and configure `vitest` for unit and snapshot testing.


2. **Build Ingestion & Virtual File System (VFS):**
* Implement `PackageLoader` to read direct PPTX buffers, inspect ZIP headers, and traverse raw OpenXML directory trees into an in-memory virtual file system (`Map<string, Uint8Array>`).


3. **Implement Spatial & Element Tree Engine:**
* Build AST wrappers around `<p:sp>`, `<p:pic>`, `<p:graphicFrame>`, and `<c:chartSpace>`.
* Create the EMU unit conversion math module to support querying and mutating coordinates (`x`, `y`, `width`, `height`).


4. **Implement Table Merging and Dynamic Pagination:**
* Develop the table measurement and row-splitting algorithm.
* Validate that cloned slides copy media relationships without causing PowerPoint corruption.


5. **Build Export Adapters and Integration Tests:**
* Implement `.toFile()`, `.toStream()`, `.toBuffer()`, `.toArrayBuffer()`, `.toBase64()`, and `.toExpress()`.
* Run automated verification with LibreOffice and Microsoft PowerPoint CLI to confirm generated presentations pass strict validation.



```

<ElicitationsGroup message="To help you execute this plan:">
  <Elicitation label="Generate complete TypeScript API signatures" query="Write the complete TypeScript definitions, classes, and method signatures for this PPTX library's core API."/>
  <Elicitation label="Draft table overflow pagination implementation" query="Show the mathematical formula and TypeScript implementation for estimating table row heights and splitting overflowing slides."/>
  <Elicitation label="Provide project scaffolding and package.json" query="Provide the package.json, tsup configuration, and folder structure to initialize this npm library."/>
</ElicitationsGroup>