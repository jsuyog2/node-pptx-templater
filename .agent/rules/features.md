# Rule: Technical Feature Invariants & OpenXML Protocol

> **Scope**: Implementation standards for core OpenXML features in `node-pptx-templater`.  
> **Applies to**: `TemplateEngine`, `ChartManager`, `TableManager`, `ZOrderManager`, `SlideManager`, `FlatOpcConverter`.

---

## 1. Text Run Healing Invariant

PowerPoint frequently splits contiguous text strings across multiple XML runs (`<a:r>`), especially when edits, spellchecks, or language tags occur. A template placeholder like `{{COMPANY_NAME}}` is often serialized as:
```xml
<a:p>
  <a:r><a:rPr sz="1400" b="1"/><a:t>{{</a:t></a:r>
  <a:r><a:rPr sz="1400" b="1"/><a:t>COMPANY_</a:t></a:r>
  <a:r><a:rPr sz="1400" b="1"/><a:t>NAME}}</a:t></a:r>
</a:p>
```

### Invariant Implementation:
1. **Paragraph-Level Run Aggregation**:
   - The engine must iterate over paragraph `<a:p>` elements and concatenate child `<a:r>/<a:t>` string values into a consolidated paragraph text buffer.
2. **Tag Matching & Substitution**:
   - Check if any requested placeholder key exists in the combined text buffer. If found, perform substitution across the buffer.
3. **Format Retention**:
   - The new text MUST NOT be dumped into raw unstyled text.
   - Clone the first run's `<a:rPr>` properties (font size, bold, italic, colors, Latin typeface) into a single new `<a:r>` node.
   - Replace the multiple fragmented runs with the unified run:
     ```javascript
     const firstRun = this.#xmlParser.deepClone(runs[0]);
     firstRun['a:t'] = replacedText;
     p['a:r'] = [firstRun];
     ```

---

## 2. Chart Dual-Sync Invariant

In OpenXML PresentationML, a native chart consists of two interdependent components:
1. **DrawingML Visual Presentation Cache**: `ppt/charts/chartX.xml` (`<c:chartSpace>`), which contains cached series names (`<c:strRef>/<c:strCache>`) and data points (`<c:numRef>/<c:numCache>`).
2. **Embedded Excel Workbook**: `ppt/embeddings/Microsoft_Excel_WorksheetX.xlsx`, which contains the underlying spreadsheet data (`xl/worksheets/sheet1.xml` and `xl/sharedStrings.xml`).

> [!WARNING]
> Updating only the `<c:chartSpace>` cache without updating the embedded `.xlsx` workbook will cause Microsoft PowerPoint to prompt the user to **"Repair Presentation"** as soon as they click "Edit Data in Excel".

### Invariant Implementation:
1. **Cache Synchronization (`ChartCacheGenerator.js`)**:
   - Mutate `<c:ser>` elements in `ppt/charts/chartX.xml`.
   - Rebuild `<c:cat>/<c:strRef>/<c:strCache>` with updated category count and point `<c:pt idx="i">`.
   - Rebuild `<c:val>/<c:numRef>/<c:numCache>` with updated series values and format codes (`General` or custom).
2. **Workbook Synchronization (`ChartWorkbookUpdater.js`)**:
   - Decompress the embedded `.xlsx` binary via JSZip.
   - Parse `xl/worksheets/sheet1.xml` and `xl/sharedStrings.xml`.
   - Update cell references: series titles in row 1 (`B1`, `C1`, etc.), categories in column A (`A2`, `A3`, etc.), and numeric values in the grid (`B2`, `C2`, etc.).
   - Clean up existing table definitions (`xl/tables/`) and sheet relationships to avoid schema validation conflicts.
   - Recompress with DEFLATE (level 6) and write back to the VFS.

---

## 3. Table Integrity & Row ID Invariant

PowerPoint tables (`<a:tbl>`) require strict conformance to row identification and cell merge rules:

### Invariant 1: Unique `<a16:rowId>` 32-Bit Unsigned Integers
- In modern PowerPoint (DrawingML 2014 extension lists), each table row `<a:tr>` contains an extension block:
  ```xml
  <a:extLst>
    <a:ext uri="{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}">
      <a16:rowId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="384918234"/>
    </a:ext>
  </a:extLst>
  ```
- **Rule**: When adding or cloning rows (`addTableRow()`, row duplication, or table auto-pagination), NEVER copy the existing `a16:rowId` value.
- **Action**: Generate a new, unique random 32-bit unsigned integer (`0` to `4294967295`) for every newly inserted or cloned row. Duplicate `val` attributes trigger PowerPoint repair dialogs.

### Invariant 2: Cell Merging Protocol (`gridSpan`, `rowSpan`, `hMerge`, `vMerge`)
- When merging a rectangular block of cells from `(startRow, startCol)` to `(endRow, endCol)`:
  - **Origin Cell** `(startRow, startCol)`:
    - Set `@_gridSpan` = column span count (if `> 1`).
    - Set `@_rowSpan` = row span count (if `> 1`).
  - **Horizontally Shadowed Cells** (`startRow`, `startCol + 1` ... `endCol`):
    - Set `@_hMerge="1"`.
    - Strip `@_gridSpan`.
  - **Vertically Shadowed Cells** (`startRow + 1` ... `endRow`, `col`):
    - Set `@_vMerge="1"`.
    - Strip `@_rowSpan`.
  - Delete duplicate content from shadowed cells to prevent corrupted visual overlapping.

---

## 4. Z-Order & `<p:spTree>` Child Ordering Invariant

Within any slide `<p:sld>`, the main visual drawing container is `<p:spTree>` (or `<p:grpSp>` for groups).

### Strict Element Order Requirement:
ECMA-376 Part 1 §19.3.1.51 strictly requires container property headers to appear **BEFORE** any visual drawing shapes:
1. `<p:nvGrpSpPr>` (Non-visual group shape properties) — MUST BE FIRST.
2. `<p:grpSpPr>` (Group shape transform & visual properties) — MUST BE SECOND.
3. Drawing shapes (`<p:sp>`, `<p:pic>`, `<p:graphicFrame>`, `<p:grpSp>`, `<p:cxnSp>`) in their exact Z-order sequence.

### Invariant Implementation:
- In `XMLParser.js` and `ZOrderManager.js`:
  - Visual shapes are reordered using the `[Z_ORDER_SYMBOL]` internal array.
  - When rebuilding the XML string via `serializeContainer()`, header elements (`<p:nvGrpSpPr>` and `<p:grpSpPr>`) are ALWAYS serialized at the very top.
  - Never allow a shape element (`<p:sp>`, `<p:pic>`) to precede `<p:nvGrpSpPr>`.

---

## 5. Blank Presentation & Flat OPC Invariant

### 5.1 Blank Presentation Deck Generation (`PPTXTemplater.createBlank()`)
To ensure presentations created from scratch open cleanly in PowerPoint, the generated VFS must contain:
1. `[Content_Types].xml`: Must include default MIME types and explicit `<Override>` entries for `/ppt/presentation.xml`, `/ppt/slideMasters/slideMaster1.xml`, `/ppt/slideLayouts/slideLayout1.xml`, `/ppt/theme/theme1.xml`, and `/ppt/slides/slide1.xml`.
2. Slide ID Minimum: In `ppt/presentation.xml`, the `<p:sldId>` attribute `@_id` MUST be `>= 256` (ECMA-376 standard; values `< 256` are reserved by PowerPoint).
3. Slide Master ID: `<p:sldMasterId id="2147483648" r:id="rId1"/>` (starts at $2^{31}$).
4. Bidirectional `.rels` Chains:
   - `_rels/.rels` -> `ppt/presentation.xml`
   - `ppt/_rels/presentation.xml.rels` -> `slideMasters/slideMaster1.xml`, `slides/slide1.xml`, `theme/theme1.xml`
   - `ppt/slides/_rels/slide1.xml.rels` -> `../slideLayouts/slideLayout1.xml`

### 5.2 Flat OPC XML (`<pkg:package>`) Handling
- ISO/IEC 29500-2 Part 2 Annex A Flat OPC packages encapsulate an entire PPTX archive into a single XML document.
- **XML Parts**: Serialized inside `<pkg:part pkg:name="/..." pkg:contentType="...">` wrapping `<pkg:xmlData>`. The root XML declaration (`<?xml ...?>`) must be stripped inside `<pkg:xmlData>`.
- **Binary Parts**: Images and `.xlsx` workbooks must be Base64-encoded inside `<pkg:binaryData>`.
- **Content Types Synthesis**: When converting Flat OPC to VFS, synthesize `[Content_Types].xml` dynamically if omitted in the flat package.

---

## 6. Slide Lifecycle, Section Management & Collaboration Invariants

### 6.1 Strict 23-Method Slide Management API Boundary
- The public slide management surface consists of **strictly 23 methods**:
  `useSlide`, `useAllSlides`, `addSlide`, `duplicateSlide`, `removeSlide`, `getSlides`, `getSlide`, `importSlideFrom`, `exportSlides`, `moveSlide`, `reorderSlides`, `tagSlide`, `getSlideDimensions`, `getSlideLayout`, `useSection`, `addSection`, `updateSection`, `removeSection`, `getSection`, `hideSlide`, `unhideSlide`, `resetSlide`, `addComment`, `removeComment`.
- Never introduce, restore, or allow deprecated aliases (e.g., `cloneSlide`, `deleteSlide`, `insertSlide`, `importSlides`, `setSlideDimensions`, `setSlideSize`, `setAspectRatio`, `getSlideSize`).
- All slide indexing is **1-based** (`1` to `slideCount`).

### 6.2 PowerPoint Sections Protocol (`<p14:sectionLst>`)
- Sections are structured inside `<p:presentation>` -> `<p:extLst>` -> `<p:ext uri="{521415D9-36F7-43E2-AB2F-EE99092638D7}">` -> `<p14:sectionLst>` using the `http://schemas.microsoft.com/office/powerpoint/2010/main` namespace.
- Each section `<p14:section>` must have a valid GUID `id` (e.g., `{UUID}`), a `name`, and a child `<p14:sldIdLst>` listing member slide IDs (`<p14:sldId id="..."/>`).
- When removing a section with `removeSlides: true`, member slides must be cleanly unlinked and purged from `ppt/presentation.xml`, `[Content_Types].xml`, and physical slide parts.

### 6.3 Slide Visibility Protocol (`show="0|1"`)
- Toggling slide visibility must directly set or strip the `show="0"` attribute on the root `<p:sld>` element in `ppt/slides/slideN.xml`.
- Never delete DrawingML shapes, slide layout bindings, or relationship entries when hiding a slide.

### 6.4 Canvas Reset Invariant (`resetSlide`)
- When clearing a slide's visual canvas, user-created drawing shapes (`<p:sp>`, `<p:pic>`, `<p:graphicFrame>`, `<p:grpSp>`) are purged from `<p:spTree>`.
- Non-visual container properties (`<p:nvGrpSpPr>` and `<p:grpSpPr>`) **MUST ALWAYS** be preserved at the top of `<p:spTree>`.
- If `options.preservePlaceholders` is `true` (default), shapes containing layout placeholders (`<p:ph>` inside `<p:nvSpPr>`) must be retained intact.

### 6.5 Collaborative Comments Architecture
- Author entries are managed in `ppt/commentAuthors.xml` (`application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml`).
- Slide comments are stored in dedicated parts `ppt/comments/commentN.xml` (`application/vnd.openxmlformats-officedocument.presentationml.comments+xml`) and linked to parent slides via relationship type `http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments`.
- Targeted removal supports zero-based comment index, author ID, or author name matching.

### 6.6 Non-Blocking Lifecycle Validation Protocol
- Internal lifecycle methods (`createBlank`, `load`, `addShape`, `importSlideFrom`, `saveToFile`, `toBuffer`, `toStream`) must never automatically invoke blocking presentation validation.
- `ppt.validatePresentation()` and `ppt.validateArchive()` serve as on-demand, non-blocking diagnostic utilities unless explicitly opted in via `options.strict` or `options.validate`.

### 6.7 Global Percentage Coordinate Invariant (0–100%)
- All visual elements (`addShape`, `addTable`, `addChart`, `addImage`, `addTextAtPosition`, `setBounds`) support percentage strings (`'0%'`, `'50%'`, `'100%'`) and percentage numbers (`unit: '%'`).
- Coordinates resolve dynamically against the slide canvas dimensions defined in `<p:sldSz cx="..." cy="..."/>` in `ppt/presentation.xml` (default 16:9 widescreen: `12192000 x 6858000` EMUs).
- Coordinate mapping rules:
  - `(0, 0)`: Top-Left corner ($x = 0$, $y = 0$).
  - `(100%, 0)`: Top-Right corner ($x = 100\%$, $y = 0$).
  - `(0, 100%)`: Bottom-Left corner ($x = 0$, $y = 100\%$).
  - `(100%, 100%)`: Bottom-Right corner ($x = 100\%$, $y = 100\%$).


