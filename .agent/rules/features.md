# Rule: Technical Feature Invariants & OpenXML Protocol

> **Scope**: Implementation standards for PresentationML & DrawingML features in `node-pptx-templater`.

---

## 1. Run Fragment Healing
- **Mechanism**: Recombines placeholders split across multiple `<a:r>` runs (e.g. `{{`, `TAG`, `}}`) caused by PowerPoint spellchecks or edits.
- **Protocol**: Concatenate text across all `<a:r>/<a:t>` runs in paragraph `<a:p>`, execute replacements on joined string, and re-emit a consolidated `<a:r>` run preserving the first run's typography (`<a:rPr>`).

## 2. Chart Dual-Sync
- **Dual-Store Mandate**: Concurrently update DrawingML visual cache `ppt/charts/chartX.xml` (`<c:chartSpace>` / `<cx:chartSpace>`) and embedded Excel workbook `ppt/embeddings/*.xlsx` across all 17 supported chart types.
- **Cache Update**: Rewrite `<c:ser>`, `<c:cat>/<c:strRef>/<c:strCache>`, and `<c:val>/<c:numRef>/<c:numCache>`.
- **Workbook Update**: Rewrite `sheet1.xml` and `sharedStrings.xml`. Purge conflicting `xl/tables/` to eliminate Excel dimension mismatch repair prompts.
- **clearCharts**: Clears chart frames from target slide(s) and purges unreferenced chart/workbook parts while preserving non-chart shapes and tables.

## 3. Table Matrix & Row ID Invariants
- **Unique `<a16:rowId>`**: Every `<a:tr>` row extension (`<a:ext uri="{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}">/<a16:rowId>`) must contain a unique 32-bit unsigned integer (`0` to `4294967295`). Never copy row IDs when cloning or creating rows.
- **Cell Merging Protocol**:
  - Origin Cell `(startRow, startCol)`: Set `@_gridSpan` (if columns $> 1$), `@_rowSpan` (if rows $> 1$); strip `hMerge`/`vMerge`.
  - Horizontally Shadowed Cells: Set `@_hMerge="1"`; delete `@_gridSpan` and text content.
  - Vertically Shadowed Cells: Set `@_vMerge="1"`; delete `@_rowSpan` and text content.
- **Cell Splitting**: Removes `rowSpan`, `gridSpan`, `hMerge`, and `vMerge` attributes across the region, restoring independent cell boundaries.

## 4. Section Grouping Protocol (`<p14:sectionLst>`)
- **Schema Location**: Located in `<p:presentation>` -> `<p:extLst>` -> `<p:ext uri="{521415E9-21F7-4876-8800-410E7D737B1B}">` -> `<p14:sectionLst>` (`http://schemas.microsoft.com/office/powerpoint/2010/main`).
- **Structure**: Each `<p14:section>` requires a GUID `id`, string `name`, and child `<p14:sldIdLst>` referencing numerical slide IDs (`<p14:sldId id="256"/>`).

## 5. Normalized Percentage Coordinate System
- **0–100% Mapping**: Coordinates accepted as percentages (`'0%'`–`'100%'`) resolved against canvas dimensions in `<p:sldSz>` (default 16:9: `12192000 x 6858000` EMUs).
- **Exact Integer EMUs**: Conversions must use exact integer EMUs (`1 pt = 12700`, `1 px = 9525`, `1 in = 914400`, `1 cm = 360000`). Never emit floating-point EMUs.

## 6. Elements, Z-Order & Header Invariants
- **Container Header Order**: In `<p:spTree>` and `<p:grpSp>`, non-visual headers must strictly precede visual children:
  1. `<p:nvGrpSpPr>`
  2. `<p:grpSpPr>`
  3. Visual shapes (`<p:sp>`, `<p:pic>`, `<p:graphicFrame>`, `<p:cxnSp>`) in Z-order sequence.
- **Relationship Cleanup**: Purging elements (`removeElement`, `removeHyperlink`) unlinks referenced `rId` definitions in `_rels/slideN.xml.rels` while strictly preserving slide layout bindings.

## 7. Full Text Box Replacement (`replaceTextByTag`)
- Targets shapes by ID (`<p:cNvPr id="...">`), name, description, or semantic tag. Clears existing runs in `<p:txBody>`, partitions `\n` into distinct `<a:p>` paragraphs, and applies root `<a:rPr>` typography from the primary run.
