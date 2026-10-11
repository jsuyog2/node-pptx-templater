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
  - `features.json`: Single source of truth for the 86 capabilities, parsed from `project/Features.md` via `scripts/parse-features.js`.
  - `apiReference.json`: Complete 101 public methods across all 10 engine domains.
  - `docsNavigation.json`: Navigation items, categories, and sidebar hierarchy.
  - `compatibility.json`: Runtimes matrix and serverless execution guarantees.
  - `siteMeta.json`: Global title, author, canonical domain, keywords, and OpenGraph metadata.
- Automated route handlers (`sitemap.js`, `robots.js`) must generate `sitemap.xml` and `robots.txt` automatically during `next build`.

