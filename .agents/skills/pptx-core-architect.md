---
name: pptx-core-architect
description: Skill for implementing advanced presentation features in node-pptx-templater v3.0.0-alpha.0, including multi-format ingestion, spatial AST coordinate conversions (EMU), table overflow pagination, and multi-target export adapters.
version: 3.0.0-alpha.0
---

# PPTX Core Architect Skill (v3.0.0-alpha.0)

## Mission & Purpose
Direct the architecture and implementation of core v3 engine capabilities: multi-format package ingestion, spatial coordinate normalization, table auto-pagination, and multi-target export serialization.

---

## Architectural Directives & Specifications

### 1. Universal Ingestion Engine (`PackageLoader`)
Support 3 distinct input sources:
- **Direct PPTX:** File path string, Node.js `Buffer`, or `ReadableStream`.
- **ZIP Packaging:** Detect PKZip binary signatures (`0x50 0x4B 0x03 0x04`) regardless of file extension (`.zip`, `.pptx`, `.tmp`).
- **Unpacked XML Directory Trees:** Ingest extracted directories containing `[Content_Types].xml` and `ppt/` directly into an in-memory Virtual File System (`Map<string, Uint8Array>`) without requiring an intermediate zip compression step.

### 2. Spatial Engine & Units Normalizer
- Standardize internal dimensions and positions in OpenXML English Metric Units (EMUs):
  - $1\text{ pt} = 12,700\text{ EMUs}$
  - $1\text{ in} = 914,400\text{ EMUs}$
  - $1\text{ px (96 DPI)} = 9,525\text{ EMUs}$
- Always round EMU coordinates using `Math.round()` to prevent floating-point schema violations in `<a:off>` and `<a:ext>`.
- Expose spatial operations across elements: `getBounds()`, `setBounds()`, `move()`, `resize()`, `bringToFront()`, and `sendToBack()`.

### 3. Dynamic Table Engine & Overflow Pagination
- **Structure:** Support row/column CRUD, custom edge borders, background fills, and multi-cell merging (`gridSpan` / `rowSpan`).
- **Height Estimation:** Compute expected row heights dynamically based on font metrics, cell padding, line count, and column width.
- **Overflow Detection:** Compare cumulative table height against available slide vertical space:
  $$Y_{\text{limit}} = \text{SlideHeight} - \text{BottomMargin}$$
- **Auto-Pagination Execution:**
  1. Slice table data at the last fitting row on the current slide.
  2. Deep-clone the host slide layout, persistent headers, footers, and background styling.
  3. Re-inject table headers onto the cloned slide.
  4. Continue rendering remaining rows recursively across subsequent slides.

### 4. Multi-Target Export Pipeline (`ExportAdapter`)
Expose 6 output formats:
- `.toFile(outputPath: string): Promise<void>`
- `.toStream(): Readable`
- `.toBuffer(): Promise<Buffer>`
- `.toArrayBuffer(): Promise<ArrayBuffer>`
- `.toBase64(options?: { dataUri?: boolean }): Promise<string>`
- `.toExpress(res: Response, filename: string): Promise<void>` (sets presentation MIME headers and attachment disposition)