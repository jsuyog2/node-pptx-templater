# Rule: Technical Feature Invariants & OpenXML Protocol

> **Scope**: Implementation standards for PresentationML & DrawingML features in `node-pptx-templater`.

---

## 1. Run Fragment Healing
- **PowerPoint Problem**: Editors split placeholders (e.g., `{{`, `TAG`, `}}`) across multiple `<a:r>` runs.
- **Protocol**: Concatenate `<a:r>/<a:t>` nodes per paragraph `<a:p>`, match and replace placeholders across the joined string, clone the first run's formatting `<a:rPr>`, and re-emit a consolidated `<a:r>` run.

## 2. Charts Engine & Excel Dual-Sync
- **Strict 8-Method Whitelist**:
  1. `addChart(slideId, chartType, data, options)`
  2. `updateChart(chartId, options)`
  3. `updateChartData(chartId, data)`
  4. `removeChart(chartId)`
  5. `clearCharts([slideRefOrChartId])`
  6. `getCharts([slideId])`
  7. `getChart(chartId)`
  8. `cloneChart(chartId, clonedChartId, clonedOptions)`
- **17 Supported Chart Types**: Column (`column`, `col`), Line (`line`), Pie (`pie`), Bar (`bar`), Area (`area`), XY Scatter (`scatter`, `xy`, `xyScatter`), Map (`map`), Stock (`stock`), Surface (`surface`), Radar (`radar`), Treemap (`treemap`), Sunburst (`sunburst`), Histogram (`histogram`), Box & Whisker (`boxWhisker`, `boxAndWhisker`, `box_and_whisker`), Waterfall (`waterfall`), Funnel (`funnel`), Combo (`combo`).
- **Dual-Store Architecture**: Synchronize DrawingML visual cache `ppt/charts/chartX.xml` (`<c:chartSpace>`) or extended parts (`<cx:chartSpace>`) and embedded Excel workbook `ppt/embeddings/Microsoft_Excel_WorksheetX.xlsx` concurrently.
- **Cache Update**: Rewrite `<c:ser>`, `<c:cat>/<c:strRef>/<c:strCache>`, and `<c:val>/<c:numRef>/<c:numCache>`.
- **Workbook Update**: Rewrite `xl/worksheets/sheet1.xml` and `xl/sharedStrings.xml`. Purge conflicting `xl/tables/` definitions to eliminate Excel dimension mismatch repair dialogs.
- **clearCharts Operation**: Evacuates all chart graphic frames from targeted slide(s), unlinks relationship nodes in `ppt/slides/_rels/slideN.xml.rels`, and purges orphaned parts and embedded workbooks while strictly preserving layout shapes, tables, and text frames.

## 3. Strict 30-Method Tables Subsystem & Invariants
- **Strict 30-Method Whitelist**:
  1. `addTable(slideId, rows, columns, data, options)`
  2. `updateTable(tableId, options)`
  3. `updateTableData(tableId, data)`
  4. `removeTable(tableId)`
  5. `getTables(slideId)`
  6. `getTable(tableId)`
  7. `cloneTable(tableId, clonedTableId, clonedOptions)`
  8. `addTableRow(tableId, rowIndex, data)`
  9. `addTableColumn(tableId, columnIndex, data)`
  10. `removeTableRow(tableId, rowIndex)`
  11. `removeTableColumn(tableId, columnIndex)`
  12. `updateRowData(tableId, rowIndex, data)`
  13. `updateColumnData(tableId, columnIndex, data)`
  14. `updateTableCell(tableId, rowIndex, columnIndex, options)`
  15. `clearRowData(tableId, rowIndex)`
  16. `clearColumnData(tableId, columnIndex)`
  17. `clearCellData(tableId, rowIndex, columnIndex)`
  18. `getTableData(tableId)`
  19. `getColumnData(tableId, columnIndex)`
  20. `getRowData(tableId, rowIndex)`
  21. `getColumn(tableId, columnIndex)`
  22. `getRow(tableId, rowIndex)`
  23. `getCell(tableId, rowIndex, columnIndex)`
  24. `getCellData(tableId, rowIndex, columnIndex)`
  25. `mergeCells(tableId, startRow, startColumn, endRow, endColumn)`
  26. `splitCell(tableId, rowIndex, columnIndex)`
  27. `styleTable(tableId, options)`
  28. `styleColumn(tableId, columnIndex, options)`
  29. `styleRow(tableId, rowIndex, options)`
  30. `styleCell(tableId, rowIndex, columnIndex, options)`
- **Purged Non-Whitelisted Routines**: Methods outside this exact 30-method whitelist (such as `getTableRows`, `autoFitTable`, `splitTableOnOverflow`, `cloneMergedRegion`, `addCellShape`, `updateCellShape`, `removeCellShape`, `getCellShape`, `getCellBounds`, `getCellPosition`, `stripeRows`) are strictly prohibited and purged from the codebase.
- **Unique `<a16:rowId>`**: Each `<a:tr>` row extension `<a:extLst>/<a:ext uri="{D42A27DB-BD31-4B8C-A354-60D93AC583D0}">/<a16:rowId>` must contain a unique 32-bit unsigned integer (`1` to `2147483647`). Never copy row IDs when cloning or adding rows.
- **Cell Merging Protocol**:
  - Origin Cell `(startRow, startCol)`: Set `@_gridSpan` (if columns $> 1$), `@_rowSpan` (if rows $> 1$); strip `hMerge`/`vMerge`.
  - Horizontally Shadowed Cells: Set `@_hMerge="1"`; delete `@_gridSpan` and text content.
  - Vertically Shadowed Cells: Set `@_vMerge="1"`; delete `@_rowSpan` and text content.
- **Cell Splitting Protocol**: Remove `rowSpan`, `gridSpan`, `hMerge`, and `vMerge` attributes across the region, restoring independent cell boundaries. Works whether called on primary anchor or secondary continuation cells.
- **Fluent Method Chaining**: All table mutation methods must return `this` (or `Promise<PPTXTemplater>` when async).
- **Polymorphic Table Cell Ingestion & Overlay Badges**:
  - `addTable`, `updateTable`, `updateTableData`, `updateTableCell`, `updateRowData`, `addTableRow`, and `addTableColumn` accept polymorphic inputs:
    1. **Primitives**: `string`, `number`, `boolean` seamlessly serialized to DrawingML text runs `<a:t>`.
    2. **Styled Values**: `{ value, style: { fontSize, fontFace, color, bold, italic, alignment, verticalAlignment, fill, ... } }`.
    3. **Cell Shape Overlays**: Positioned shape overlays (e.g. `roundRect`, `circle`, `pill`) aligned with cell bounding coordinates without modifying `<a:tbl>` table structure or causing table reflow:
       - `textPosition: 'inside'`: Renders inside `<p:txBody>` of the overlay shape; underlying `<a:tc>` text remains empty.
       - `textPosition: 'outside'`: Shape renders without text; cell DrawingML `<a:txBody>` retains text and margin padding (`@_marL` / `@_marR`) is auto-applied to prevent collision.
  - Life-cycle integration automatically handles cell badge repositioning, cleaning on cell updates/clears, and cloning with coordinate offsets.

## 4. Z-Order & `<p:spTree>` Container Order
- **ECMA-376 Header Order**: Inside `<p:spTree>` and `<p:grpSp>`, non-visual headers must appear before drawing children:
  1. `<p:nvGrpSpPr>` (Non-visual group properties)
  2. `<p:grpSpPr>` (Group shape transform & visual properties)
  3. Visual shapes (`<p:sp>`, `<p:pic>`, `<p:graphicFrame>`, `<p:grpSp>`, `<p:cxnSp>`) in Z-order sequence.
- **Symbol Preservation**: `XMLParser` tracks element stacking with `[Z_ORDER_SYMBOL]`. Serialization must prepend container headers.

## 5. Section Grouping Protocol (`<p14:sectionLst>`)
- **Schema Location**: Located in `<p:presentation>` -> `<p:extLst>` -> `<p:ext uri="{521415E9-21F7-4876-8800-410E7D737B1B}">` (or `{521415D9-36F7-43E2-AB2F-B90AF26B5E84}`) -> `<p14:sectionLst>` using namespace `http://schemas.microsoft.com/office/powerpoint/2010/main`.
- **Section Structure**: Each `<p14:section>` requires a GUID `id`, `name`, and child `<p14:sldIdLst>` referencing contiguous slide IDs.

## 6. Global Percentage Coordinate System
- **0–100% Normalized Coordinates**: All visual elements support percentage strings (`'0%'`–`'100%'`) resolved against canvas dimensions in `<p:sldSz>` (default 16:9: `12192000 x 6858000` EMUs).
- **Mapping**: `(0, 0)` Top-Left, `(100%, 0)` Top-Right, `(0, 100%)` Bottom-Left, `(100%, 100%)` Bottom-Right.
- **Units**: Conversions use integer EMUs (`1 pt = 12700`, `1 px = 9525`, `1 in = 914400`, `1 cm = 360000`). Never emit floating-point EMUs.

## 7. Full Text Box Replacement (`replaceTextByTag`)
- **Resolution**: Targets shapes by ID (`<p:cNvPr id="...">`), name (`<p:cNvPr name="...">`), description, or semantic tag.
- **Execution**: Clears all existing runs inside `<p:txBody>`, partitions `\n` into distinct `<a:p>` paragraphs, and applies root `<a:pPr>` and `<a:rPr>` typography from the primary run. Non-matches log a warning and return gracefully.

## 8. Blank Decks & Flat OPC Packaging
- **Blank Presentations**: Slide IDs in `ppt/presentation.xml` must be $\ge 256$. Slide Master ID starts at $2^{31}$. Register all part overrides in `[Content_Types].xml` and construct bidirectional `.rels` chains.
- **Flat OPC XML (`<pkg:package>`)**: Encapsulates OPC archives in a single XML file. XML parts use `<pkg:xmlData>` (root XML declaration stripped), binaries use Base64 `<pkg:binaryData>`.

## 9. Strict 10-Method Shapes & Geometry Whitelist Boundary
- **Approved 10 Methods**:
  1. `addShape(slideId, type, options)`
  2. `updateShape(targetId, options)`
  3. `removeShape(shapeId)`
  4. `clearShapes()`
  5. `getShapes()`
  6. `getShape(shapeId)`
  7. `cloneShape(shapeId, clonedShapeId, clonedOptions)`
  8. `getGroupShapes(groupId)`
  9. `groupShapes(shapeIds)`
  10. `ungroupShapes(groupId)`
- **Mandates**:
  - Exclusively these 10 methods are exposed on `PPTXTemplater` and `ShapeManager`.
  - Mutating operations must return fluent `this`.
  - `clearShapes()` evacuates user shapes while strictly preserving layout placeholders.
  - Grouping calculates composite `<a:xfrm>` (`<a:chOff>`, `<a:chExt>`) under `<p:grpSp>`, and ungrouping dissolves the group with exact canvas coordinate re-projection.

## 10. Strict 12-Method Images & Media Whitelist Boundary
- **Approved 12 Methods** (canonical camelCase with optional `slideId` when using `useSlide()`, with PascalCase aliases supported):
  1. `addImage([slideId], source, options)`
  2. `updateImage(imageId, options)`
  3. `removeImage(imageId)`
  4. `addVideo([slideId], source, options)`
  5. `updateVideo(videoId, options)`
  6. `removeVideo(videoId)`
  7. `addAudio([slideId], source, options)`
  8. `updateAudio(audioId, options)`
  9. `removeAudio(audioId)`
  10. `getMedias([slideId])`
  11. `getMedia(mediaId)`
  12. `cloneMedia(mediaId, clonedMediaId, clonedOptions)`
- **OpenXML Media Architecture**:
  - **Context Resolution**: Methods bind to active slide context from `ppt.useSlide(slideNumber)` when `slideId` is omitted.
  - **Universal Ingestion**: Supports local paths, remote URLs, Base64 Data URIs, raw Base64 strings, Buffers, and streams.
  - **Storage**: Media streams are ingested into `ppt/media/` with automatic SHA-1 hash deduplication.
  - **Content Types**: Extensions mapped in `[Content_Types].xml` (`image/png`, `video/mp4`, `audio/mpeg`, etc.).
  - **Relationships**: Tracked via `RelationshipManager` with namespaces for images (`.../relationships/image`), audio/video (`.../relationships/video`, `.../relationships/audio`, `.../relationships/media`).
  - **Shape Structures**:
    - Images: `<p:pic>` with `<p:blipFill>` and DrawingML coordinates (`<a:xfrm>`). Supports percentage & EMU coordinates, cropping (`<a:srcRect>`), rotation, borders, shadows, transparency (`<a:alphaModFix>`).
    - Video: `<p:pic>` with `<a:videoFile>` referencing video rel, poster frame blip, and `<p:timing>` sequence bindings for play/pause/stop and loop triggers.
    - Audio: `<p:pic>` with `<a:audioFile>` (or sound node) in `<p:spTree>`, optional `<p:timing>` audio actions, loop tags, and toggle icon visibility via `<p:cNvPr hidden="...">`.
  - **Fluent API**: All mutation methods return `this` (or `Promise<PPTXTemplater>`).
  - **Non-destructive Removal & Timing Pruning**: Deleting shapes clears timing sequence nodes in `<p:timing>` and verifies relationship targets are not shared before removing `.rels` entries.

