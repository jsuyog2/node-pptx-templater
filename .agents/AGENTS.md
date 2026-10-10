---
trigger: always_on
---

# Agent Directives & Architecture: node-pptx-templater (v3.0.0-alpha.0)

This repository houses `node-pptx-templater` v3.0.0-alpha, an enterprise-grade, spatial-aware OpenXML PowerPoint manipulation, templating, and layout engine for Node.js.

---

# WORKSPACE ROOT DIRECTIVE
- ALL operations must be strictly scoped to the repository root directory containing `package.json`.
- Do NOT search, list, or inspect parent directories, user home folders (`C:\Users\...`), or global system paths.
- All file paths must be resolved relative to `./` within this repository.

---

## 1. Core Operating Constraints

### Rule 1: Zero Synchronous / Blocking I/O
- **Forbidden:** `fs.readFileSync`, `fs.writeFileSync`, `fs.readdirSync`, `fs.statSync`.
- **Mandated:** Pure non-blocking I/O using `Uint8Array`, Node.js `stream.Readable`, and in-memory byte arrays.
- All internal transformations run against an in-memory Virtual File System (`Map<string, Uint8Array>`).

### Rule 2: Strict OpenXML AST Manipulation (No Regex Tag Swapping)
- **Forbidden:** String regex replacements across XML tags (e.g., `/<p:txBody>.*<\/p:txBody>/`). PowerPoint fragments placeholder tokens across distinct `<a:r>` runs and `<a:t>` text nodes.
- **Mandated:** Use an AST-level run normalizer that merges adjacent text runs inside paragraph nodes (`<a:p>`) before token matching or replacement.
- Preserve all OpenXML namespaces (`p:`, `a:`, `r:`, `c:`, `w:`) throughout parsing and serialization.

### Rule 3: Precise Spatial Units (EMUs)
- Normalize all spatial calculations to OpenXML English Metric Units (EMUs):
  - $1\text{ pt} = 12,700\text{ EMUs}$
  - $1\text{ in} = 914,400\text{ EMUs}$
  - $1\text{ px (96 DPI)} = 9,525\text{ EMUs}$
- Always use integer math (`Math.round()`) when writing `<a:off>` and `<a:ext>` tags.

### Rule 4: Workbook Cache Synchronization
- When updating series, categories, or values in `<c:chartSpace>`, update the embedded Excel binary cache inside `ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx` to avoid Microsoft Office repair warnings.

---

## 2. Ingestion & Export Pipelines

### PackageLoader Ingestion
Support 3 distinct input formats:
1. **Direct PPTX:** File path, Node.js `Buffer`, or `ReadableStream`.
2. **ZIP Containers:** Detect PKZip signatures (`0x50 0x4B 0x03 0x04`) regardless of extension (`.zip`, `.pptx`, `.tmp`).
3. **Unpacked XML Directories:** Recursively read an extracted folder containing `[Content_Types].xml` and `ppt/` directly into the VFS map without requiring an intermediate zip compression step.

### ExportAdapter Outputs
Expose 6 output formats:
- `toFile(outputPath: string): Promise<void>`
- `toStream(): Readable`
- `toBuffer(): Promise<Buffer>`
- `toArrayBuffer(): Promise<ArrayBuffer>`
- `toBase64(options?: { dataUri?: boolean }): Promise<string>`
- `toExpress(res: Response, filename: string): Promise<void>` (attaches presentation MIME type and disposition headers)

---

## 3. Flagship Table Auto-Pagination Engine
- Calculate row heights dynamically using font size, line count, cell padding, and column width.
- Identify slide vertical threshold:
  $$Y_{\text{limit}} = \text{SlideHeight} - \text{BottomMargin}$$
- When table rows exceed $Y_{\text{limit}}$:
  1. Slice rows at the last fitting position on the host slide.
  2. Deep-clone the host slide (duplicating layout, persistent shapes, and background).
  3. Generate isolated relationship files (`slideX.xml.rels`) and fresh `rId` keys.
  4. Repeat table headers on the cloned slide and populate remaining rows recursively.

---

## 4. Code Standards & Build Target
- Written in TypeScript (strict mode enabled).
- Bundled via `tsup` targeting dual ESM (`dist/index.mjs`) and CommonJS (`dist/index.js`).
- Unit and schema snapshot validation powered by `vitest`.