---
name: pptx-engine-unified-modernizer
description: Master skill for node-pptx-templater v3.0.0-alpha.0. Combines legacy code cleanup and OpenXML relationship stabilization with the complete v3 architecture, spatial engine, dynamic table auto-pagination, and multi-target export adapters.
version: 3.0.0-alpha.0
---

# PPTX Engine Unified Modernizer Skill (v3.0.0-alpha.0)

## Mission & Purpose
Provide an end-to-end operational guide for transitioning `node-pptx-templater` into `v3.0.0-alpha.0`. This unified skill governs the removal of legacy anti-patterns, OpenXML schema stabilization, and the implementation of spatial layout and streaming export features.

---

## Operational Execution Lifecycle

### Phase 1: Legacy Code Purge & Architecture Stabilization
1. **Strip Synchronous I/O:** Remove all calls to `fs.readFileSync` and `fs.writeFileSync`. Route package mutations through an in-memory Virtual File System (`Map<string, Uint8Array>`).
2. **Eliminate Regex Parsing:** Replace raw string regex replacements with an AST run normalizer that merges fragmented `<a:r>` text runs inside `<a:p>` nodes before substituting template variables.
3. **Fix Relationship Identifiers:** Stop manual, hardcoded `rId` increments. Use an isolated `RelationshipManager` to generate unique identifiers and isolated `.rels` manifests.
4. **Synchronize Chart Workbooks:** Whenever `<c:chartSpace>` nodes are modified, update the embedded `.xlsx` workbook cache in `ppt/embeddings/` to prevent Microsoft PowerPoint repair prompts.

### Phase 2: Core Ingestion & Spatial Systems
1. **Universal PackageLoader:**
   - Ingest `.pptx` via path, Buffer, or Stream.
   - Ingest `.zip` archives by verifying PKZip header bytes (`0x50 0x4B 0x03 0x04`).
   - Ingest extracted OpenXML directory trees directly into memory without requiring prior zipping.
2. **Spatial Coordinate Math:**
   - Convert all coordinates and dimensions to OpenXML English Metric Units (EMUs):
     $$\text{EMU} = \text{Points} \times 12,700 = \text{Inches} \times 914,400 = \text{Pixels (96 DPI)} \times 9,525$$
   - Enforce integer values using `Math.round()`.
   - Provide element controls: `getBounds()`, `setBounds()`, `move()`, `resize()`, `bringToFront()`, and `sendToBack()`.

### Phase 3: Table Engine & Dynamic Pagination
1. **Table Manipulation:** Support cell merging (`gridSpan`/`rowSpan`), borders, cell fills, and inserting nested elements (icons, mini-shapes) inside table cells.
2. **Dynamic Height Calculation:** Calculate estimated row height using cell padding, font size, line count, and column width.
3. **Slide Overflow Pagination:**
   - Detect when cumulative table height exceeds $Y_{\text{limit}} = \text{SlideHeight} - \text{BottomMargin}$.
   - Split rows at the boundary, deep-clone the template slide layout, duplicate table headers, and populate remaining rows across subsequent slides.

### Phase 4: Multi-Target Serialization
Expose 6 standard export methods:
- `toFile(outputPath: string): Promise<void>`
- `toStream(): Readable`
- `toBuffer(): Promise<Buffer>`
- `toArrayBuffer(): Promise<ArrayBuffer>`
- `toBase64(options?: { dataUri?: boolean }): Promise<string>`
- `toExpress(res: Response, filename: string): Promise<void>`

---

## Technical Directives & Guardrails
- **Language & Runtime:** TypeScript (Strict Mode) on Node.js 18+ LTS.
- **Build System:** Bundled via `tsup` targeting dual ESM (`dist/index.mjs`) and CommonJS (`dist/index.js`) with generated `.d.ts` files.
- **Testing:** Unit tests, OpenXML schema validation, and snapshot tests using `vitest`.
- **OpenXML Namespaces:** Retain all XML namespaces (`p:`, `a:`, `r:`, `c:`, `w:`) and prefixes throughout parse and serialization cycles.