# Feature Implementation Rules

This document specifies mandatory rules and invariants for every feature domain in `node-pptx-templater`.

---

## 1. Text & Placeholder Templating

### 1.1 Run Healing Across Split Runs
- PowerPoint frequently fragments user-typed placeholder strings (e.g., `{{company_name}}`) across multiple `<a:r>` (text run) elements and `<a:t>` tags due to spelling checks, grammar tags, font switches, or editing pauses.
- **Rule**: Prior to token matching or replacement, the template engine must normalize and "heal" adjacent text runs within each paragraph (`<a:p>`).
- If a placeholder tag is split across runs:
  1. Detect the boundary of the tag (`{{` to `}}`).
  2. Merge the run text into the leading run while preserving the leading run's text properties (`<a:rPr>`).
  3. Clear or remove subsequent runs consumed by the placeholder tag.
- Text outside placeholder tags must remain unmodified.

### 1.2 Formatting Preservation
- When substituting text into `{{placeholder}}`, existing run styling (color `<a:solidFill>`, bold `<a:b>`, italic `<a:i>`, font size `sz`, font typeface `<a:latin>`) must be preserved unless explicit override styles are passed by the user.
- Multi-line replacements (containing `\n`) must be converted into separate `<a:p>` (paragraphs) or `<a:br>` (break runs) with matching paragraph properties (`<a:pPr>`).

---

## 2. Table Engine & Merging

### 2.1 Table Structure & OpenXML Elements
- A PowerPoint table resides in `<p:graphicFrame>` -> `<a:graphic>` -> `<a:graphicData>` -> `<a:tbl>`.
- The `<a:tbl>` element consists of:
  - `<a:tblPr>`: Table properties.
  - `<a:tblGrid>`: Column definitions with `<a:gridCol w="..."/>` widths in EMUs.
  - `<a:tr h="..."/>`: Rows with row height attribute and cell children.
  - `<a:tc>`: Table cells containing `<a:txBody>` (text body) and `<a:tcPr>` (cell properties).

### 2.2 Cell Merging Rules (Spans and Dummy Cells)
- **Horizontal Merges (`gridSpan`)**:
  - The primary (top-left) cell must define `gridSpan="N"`.
  - The remaining $N-1$ horizontally merged cells in that row must exist as `<a:tc hMerge="1">` placeholder cells.
- **Vertical Merges (`rowSpan`)**:
  - The top cell must define `rowSpan="M"`.
  - The lower $M-1$ cells directly below it in subsequent rows must exist as `<a:tc vMerge="1">` placeholder cells.
- **Total Columns Invariant**: Every `<a:tr>` must contain the exact same number of `<a:tc>` elements (including dummy `hMerge`/`vMerge` cells) matching the count of `<a:gridCol>` elements in `<a:tblGrid>`. Mismatches cause immediate presentation repair corruption.

### 2.3 Row Identification (`<a16:rowId>`)
- In PowerPoint 2016+, PowerPoint tracks table rows using 32-bit unsigned integers:
  `<a16:rowId val="123456789" xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"/>`
- **Rule**: When adding or duplicating rows dynamically, each newly generated `<a:tr>` MUST be assigned a unique 32-bit unsigned integer ID ($1 \le \text{rowId} \le 4,294,967,295$). Duplicate or non-numeric row IDs trigger PowerPoint's "Presentation Repair" dialog.

### 2.4 Auto-Pagination on Overflow
- When dynamic rows overflow the slide boundary:
  1. Calculate available vertical space ($Y_{\text{limit}} = \text{SlideHeight} - \text{MarginBottom}$).
  2. Estimate row heights based on font size, line count, and cell padding.
  3. Clone the host slide, split rows at the overflow threshold, copy table headers to the subsequent slide, and continue row insertion.

---

## 3. Native Chart Engine & Dual-Sync

### 3.1 Dual-Sync Architecture
- PowerPoint charts are hybrid structures consisting of two components:
  1. **DrawingML Chart Definition**: Stored in `ppt/charts/chartX.xml` (`<c:chartSpace>`), which PowerPoint uses to visually render the chart.
  2. **Embedded Excel Workbook**: Stored in `ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx`, which PowerPoint uses when the user clicks "Edit Data in Excel".
- **Rule**: When updating chart categories or series values, both `chartX.xml` and the embedded Excel sheet MUST be synchronized. Updating only the XML causes PowerPoint to show repair warnings or revert data upon user editing.

### 3.2 DrawingML `<c:chartSpace>` Structure
- Series data is represented within `<c:plotArea>` under the chart type element (e.g., `<c:barChart>`, `<c:lineChart>`, `<c:pieChart>`).
- Each `<c:ser>` contains:
  - `<c:tx>`: Series title/name (`<c:strRef>` or `<c:v>`).
  - `<c:cat>`: Category axis labels (`<c:strRef>` or `<c:numRef>`).
  - `<c:val>`: Values array (`<c:numRef>` with `<c:ptCount>` and `<c:pt idx="i"><c:v>val</c:v></c:pt>`).
- `<c:ptCount>` must strictly match the count of `<c:pt>` entries. Points must use 0-indexed contiguous `idx` attributes.

---

## 4. Shapes & Spatial Placement

### 4.1 EMU Coordinate Standards
- All coordinates and dimensions in OpenXML must be integer values represented in **English Metric Units (EMUs)**:
  - $1 \text{ pt} = 12,700 \text{ EMUs}$
  - $1 \text{ in} = 914,400 \text{ EMUs}$
  - $1 \text{ px (96 DPI)} = 9,525 \text{ EMUs}$
  - $1 \text{ cm} = 360,000 \text{ EMUs}$
- Floating-point EMU values are invalid in OpenXML schema and must always be rounded (`Math.round()`).

### 4.2 Element Positioning (`<a:xfrm>`)
- Elements must define their spatial box within `<p:spPr>` -> `<a:xfrm>`:
  - `<a:off x="X" y="Y"/>`: Offset from the slide top-left origin.
  - `<a:ext cx="WIDTH" cy="HEIGHT"/>`: Element extents (bounding box).
  - Optional `rot="ANGLE_60000"`: Rotation angle measured in $1/60,000\text{ths}$ of a degree.

### 4.3 Shapes in Table Cells
- Shapes embedded inside or visually mapped over table cells must be calculated using the cell's absolute slide coordinates (sum of table offset + column offsets + row offsets) to guarantee non-reflow during slide renders.

---

## 5. Z-Order & Layering

### 5.1 `<p:spTree>` Child Ordering
- In PowerPoint OpenXML, visual layering (Z-index) is determined strictly by the document order of child elements inside `<p:spTree>` (shape tree) in `ppt/slides/slideX.xml`.
- Elements rendered later in the `<p:spTree>` XML tree appear on top of earlier elements.
- **Rule**:
  - `bringToFront()` moves the target element node to the end of `<p:spTree>`.
  - `sendToBack()` moves the target element node to the beginning of `<p:spTree>` (directly after `<p:nvGrpSpPr>` and `<p:grpSpPr>`).
  - `moveForward()` / `moveBackward()` swaps the target element with its adjacent visual sibling.
- Do NOT alter element IDs (`id`) or non-visual property relationships when reordering `<p:spTree>` nodes.

---

## 6. Slide Lifecycle & Relationship Management

### 6.1 Slide Cloning
- When duplicating a slide:
  1. Create a unique part name (`ppt/slides/slideN.xml`).
  2. Duplicate the relationship file (`ppt/slides/_rels/slideN.xml.rels`).
  3. Register the new slide relationship in `ppt/_rels/presentation.xml.rels` with a newly generated, unique `rId` (e.g. `rIdN`).
  4. Register the new slide in `ppt/presentation.xml` inside `<p:sldIdLst>` with a unique `id` ($id \ge 256$).
  5. Register the part in `[Content_Types].xml` with content type `application/vnd.openxmlformats-officedocument.presentationml.slide+xml`.

### 6.2 Relationship Integrity
- Every media asset (`ppt/media/*`), chart (`ppt/charts/*`), or layout link referenced in a slide must have a corresponding `<Relationship>` entry in `ppt/slides/_rels/slideX.xml.rels`.
- Unused or orphaned relationship IDs must be purged to avoid OpenXML validation errors.

---

## 7. Documentation Portal & Next.js Static Export Rules

### 7.1 Pure Static HTML Export (`output: 'export'`)
- All documentation code residing in `docs-site/` must compile to static HTML for native zero-node hosting on Cloudflare Pages.
- `next.config.mjs` must strictly preserve:
  - `output: 'export'`
  - `images: { unoptimized: true }` (Required because static export does not include a Node image optimization server)
  - `trailingSlash: true` (Guarantees directory-style index.html generation for static web servers)
- Dynamic route segments (such as `/guides/[slug]`) must implement `generateStaticParams()` to pre-render every guide page statically at build time.

### 7.2 Strict Content Decoupling into JSON Data Schemas
- Never hardcode API method signatures, parameter tables, feature descriptions, or navigation links directly into React JSX components.
- All content must be managed in type-safe, decoupled JSON files in `docs-site/src/data/`:
  - `features.json`: Single source of truth for the 96 capabilities across 12 domain categories.
  - `apiReference.json`: Complete 111 public methods across all 10 engine domains.
  - `docsNavigation.json`: Navigation items, categories, and sidebar hierarchy.
  - `compatibility.json`: Runtimes matrix and serverless execution guarantees.
  - `siteMeta.json`: Global title, author, canonical domain, keywords, and OpenGraph metadata.
- Automated route handlers (`sitemap.js`, `robots.js`) must generate `sitemap.xml` and `robots.txt` automatically during `next build`.

---

## 8. Ingestion, Multi-Target Export & Extraction Rules

### 8.1 Unified Ingestion Pipeline & Signature Detection
- `PackageLoader` must detect container signatures prior to parsing:
  - PKZip standard magic bytes (`0x50 0x4B 0x03 0x04`) for `.pptx`, `.pptm`, `.potx`, `.potm`, `.ppsx`, `.ppsm`, `.ppam`, `.zip`.
  - Microsoft Flat OPC single-file XML packages (`<?xml` and `<pkg:package`).
  - OLE Compound File signature (`0xD0 0xCF 0x11 0xE0 0xA1 0xB1 0x1A 0xE1`) for legacy `.ppt` files.
  - OpenDocument signatures for `.odp` files (or PKZip containing `mimetype` with `application/vnd.oasis.opendocument.presentation`).
- Legacy binary PPT and ODP formats must be routed through `LegacyFormatConverter` to normalize extracted text atoms and slides into an OpenXML `VirtualFileSystem`.
- All ingestion routes polymorphic sources through `PPTXTemplater.load(source, options)`:
  - Local file paths, buffers, and streams.
  - Remote web URLs (`https://...`) are ingested via internal streaming fetch.
  - Cloud object storage configurations (`{ provider: 's3' | 'gcs' | 'azure', ... }`) ingest streams or buffers directly into the in-memory VFS.
  - Uncompressed directory trees mount directly without intermediate archiving.

### 8.2 Blank Presentation Generation Invariants
- `PPTXTemplater.createBlank(options)` must generate minimal valid OpenXML schemas purely in memory without relying on external file fixtures.
- Synthesized AST must include valid:
  - `[Content_Types].xml` with presentation, slide, slideMaster, slideLayout, theme, and relationship overrides.
  - `_rels/.rels` pointing to `ppt/presentation.xml`.
  - `ppt/presentation.xml` with `<p:sldSz>` (16:9, 4:3, 16:10, or custom EMU/pixel dimensions) and `<p:sldIdLst>`.
  - `ppt/slideMasters/slideMaster1.xml` and `ppt/slideLayouts/slideLayout1.xml` with layout relationships.
  - `ppt/theme/theme1.xml` with default font schemes and color palettes.
  - `ppt/slides/slide1.xml` containing a minimal valid `<p:sld>` ready for element additions.

### 8.3 Multi-Target Export & Framework Responses
- All serialization pipelines must route through `OutputWriter` and `FrameworkAdapters`.
- The public export interface consists strictly of:
  - `saveToFile(destinationPath, options)`: Commits presentation to local disk.
  - `toBuffer(options)`: Compiles presentation into a Node.js `Buffer`.
  - `toStream(options)`: Emits a readable stream (`ReadableStream`).
  - `toBase64(options)`: Returns standard Base64 string or Data URI.
  - `toArrayBuffer(options)`: Returns a web-standard `ArrayBuffer` / `Uint8Array` for Edge runtimes.
  - `toResponse(options)`: Emits a web-standard `Response` object for Next.js Route Handlers and Edge workers.
  - `toExpress(res, filename, options)`: Pipes presentation to Express.js HTTP response with OpenXML headers and handles client aborts.
  - `toFastify(reply, filename, options)`: Pipes presentation to Fastify reply with content type and disposition headers.
- Web framework adapters (`toExpress`, `toFastify`, `toResponse`) must automatically inject:
  - `Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation`
  - `Content-Disposition: attachment; filename="<filename>"` with RFC-5987 Unicode encoding
  - `Content-Length` where applicable.

### 8.4 In-Memory Template Caching Invariants
- Static template cache methods on `PPTXTemplater`:
  - `preload(source)`: Parses and caches template in static memory (`Map<string, PPTXTemplater>`).
  - `fromCache(source)`: Deep clones pre-parsed template in <1ms without disk I/O or ZIP decompression.
  - `clearCache()`: Purges cached templates from memory to prevent memory leaks in long-running processes.

### 8.5 XML & ZIP Extraction Invariants
- `extractToZip(destinationPath)`: Packages and dumps raw compressed ZIP archives to disk.
- `extractToXml(destinationDir, options)`: Unpacks all presentation parts into uncompressed folders. When `{ formatXml: true }` is enabled, all XML files must be formatted with 2-space indentation while preserving binary media files intact.

### 8.6 Flat OPC Single-File XML Ingestion & Export Rules
- Single-file XML packages starting with `<pkg:package>` must be parsed directly without ZIP overhead via `FlatOpcConverter`.
- Ingestion rules:
  - Strip leading slashes from `pkg:name` attributes (e.g., `/ppt/presentation.xml` -> `ppt/presentation.xml`) to mount into standard VFS paths.
  - Extract `<pkg:xmlData>` child nodes and serialize them as clean XML strings into the in-memory `VirtualFileSystem`.
  - Decode base64 strings from `<pkg:binaryData>` into binary `Uint8Array` / `Buffer` for embedded Excel workbooks (`ppt/embeddings/`) and images (`ppt/media/`).
  - Synthesize `[Content_Types].xml` with default extension mappings and part overrides so standard OpenXML managers operate transparently.
- Export rules:
  - `toFlatXml(destinationPathOrOptions)`: Serializes all VFS entries back into standard `<pkg:package>` XML with `<pkg:part pkg:name="..." pkg:contentType="...">` tags.
  - When passed a string, writes directly to local disk at `destinationPath` and returns the XML string.
  - When passed options or omitted, returns the XML string in memory.
  - XML parts emit `<pkg:xmlData>` with raw XML DOM/AST content; binary parts emit `<pkg:binaryData>` containing base64 strings.
- Standard operations (`useSlide`, `replaceText`, `updateChart`, `updateTable`, `saveToFile`, `toBuffer`) must execute seamlessly on decks loaded from Flat OPC XML.

