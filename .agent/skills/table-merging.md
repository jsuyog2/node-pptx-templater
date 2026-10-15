# Skill: Table Cell Merging, Row ID Integrity & Auto-Pagination

> **Domain**: PresentationML Table Processing (`<a:tbl>`), Extension Lists (`<a16:rowId>`), and Overflow Pagination  
> **Target Subsystems**: `TableManager.js`, `SlideManager.js`, `spatial/units.js`

---

## 1. Table Architecture in OpenXML

A table in PresentationML is contained within a `<p:graphicFrame>` whose `<a:graphicData>` has URI `http://schemas.openxmlformats.org/drawingml/2006/table`:
```xml
<p:graphicFrame>
  <p:nvGraphicFramePr>
    <p:cNvPr id="5" name="Table 4"/>
  </p:nvGraphicFramePr>
  <p:xfrm>
    <a:off x="914400" y="1828800"/>
    <a:ext cx="10363200" cy="4572000"/>
  </p:xfrm>
  <a:graphic>
    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
      <a:tbl>
        <a:tblPr/>
        <a:tblGrid>
          <a:gridCol w="2590800"/>
          <a:gridCol w="2590800"/>
        </a:tblGrid>
        <a:tr h="370840">
          <a:tc><a:txBody><a:p><a:r><a:t>Cell 1</a:t></a:r></a:p></a:txBody></a:tc>
          <a:tc><a:txBody><a:p><a:r><a:t>Cell 2</a:t></a:r></a:p></a:txBody></a:tc>
          <a:extLst>
            <a:ext uri="{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}">
              <a16:rowId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" val="12938472"/>
            </a:ext>
          </a:extLst>
        </a:tr>
      </a:tbl>
    </a:graphicData>
  </a:graphic>
</p:graphicFrame>
```

---

## 2. Cell Merging Playbook

### Coordinate System:
- Tables use 0-based indexing: `(startRow, startCol)` to `(endRow, endCol)`.
- Example: Merging header row across columns 0 to 3: `ppt.mergeCells('TableId', 0, 0, 0, 3)`.

### Step-by-Step Procedure:
1. **Locate Target Table**:
   Retrieve the `<a:tbl>` AST node for the specified table identifier.
2. **Compute Spans**:
   ```javascript
   const gridSpan = endCol - startCol + 1;
   const rowSpan = endRow - startRow + 1;
   ```
3. **Configure Origin Cell**:
   At `(startRow, startCol)`:
   - If `gridSpan > 1`, set `@_gridSpan = String(gridSpan)`.
   - If `rowSpan > 1`, set `@_rowSpan = String(rowSpan)`.
   - Strip any `@_hMerge` or `@_vMerge`.
4. **Configure Horizontally Shadowed Cells**:
   For cells `(r, c)` where `r = startRow` and `startCol < c <= endCol`:
   - Set `@_hMerge = "1"`.
   - Delete `@_gridSpan` and empty the text frame `<a:txBody>` (or leave a single empty `<a:p>`).
5. **Configure Vertically Shadowed Cells**:
   For cells `(r, c)` where `startRow < r <= endRow` and `startCol <= c <= endCol`:
   - Set `@_vMerge = "1"`.
   - If also horizontally shadowed, set `@_hMerge = "1"`.
   - Delete `@_rowSpan` and clear cell text.

---

## 3. Row ID Integrity Rule (`<a16:rowId>`)

Whenever rows are added or cloned:
```javascript
#updateRowId(rowObj) {
  const randomVal = String(Math.floor(Math.random() * 4294967295));
  if (!rowObj['a:extLst']) {
    rowObj['a:extLst'] = {
      'a:ext': {
        '@_uri': '{0D108BD9-81ED-4DB2-BD59-A6C34878D82A}',
        'a16:rowId': {
          '@_xmlns:a16': 'http://schemas.microsoft.com/office/drawing/2014/main',
          '@_val': randomVal
        }
      }
    };
  } else {
    // Traverse existing extLst and replace the val attribute with randomVal
  }
}
```
**Never omit or duplicate `a16:rowId`.**

---

## 4. Table Auto-Pagination Engine (`splitTableOnOverflow`)

When large datasets are populated into a table, rows can overflow beyond the bottom edge of the slide canvas. The flagship `splitTableOnOverflow()` method dynamically paginates the table:

```
Slide 1: Initial Slide                  Slide 2: Cloned Slide
┌─────────────────────────────┐        ┌─────────────────────────────┐
│ Header Title                │        │ Header Title                │
│ ┌─────────────────────────┐ │        │ ┌─────────────────────────┐ │
│ │ Header Row (Repeated)   │ │        │ │ Header Row (Repeated)   │ │
│ ├─────────────────────────┤ │        │ ├─────────────────────────┤ │
│ │ Rows 1 to 15            │ │  ───►  │ │ Rows 16 to 30           │ │
│ └─────────────────────────┘ │        │ ├─────────────────────────┤ │
│ (Overflows Bottom Margin)   │        │ │ Summary Row (Pinned)    │ │
│                             │        │ └─────────────────────────┘ │
└─────────────────────────────┘        └─────────────────────────────┘
```

### Protocol:
1. **Calculate Available Space**:
   - $Y_{\text{table}} = \text{Table bounding box top offset}$
   - $H_{\text{slide}} = \text{Slide total height in EMUs}$
   - $M_{\text{bottom}} = \text{Bottom margin in EMUs (default: 0.5in)}$
   - $Y_{\text{limit}} = H_{\text{slide}} - M_{\text{bottom}}$
2. **Estimate Row Heights**:
   - Inspect font size, typeface metrics, column width, and line wrapping for every cell in the row.
   - Cumulative table height = $\sum \text{RowHeight}_i$.
3. **Partition Rows**:
   - Find the cutoff row index $k$ where the table reaches $Y_{\text{limit}}$.
   - Slice rows $1 \dots k$ on the current slide.
4. **Deep-Clone Host Slide**:
   - Call `slideManager.duplicateSlide()`, generating unique relationship IDs and `[Content_Types].xml` overrides.
5. **Inject Repeated Headers & Pinned Summaries**:
   - If `repeatHeaders > 0`, prepend the top header rows to the table on the newly cloned slide.
   - If `summaryRows > 0`, pin the final total/summary rows to the last generated slide.
6. **Recurse**:
   - If the remaining rows still exceed $Y_{\text{limit}}$, recursively continue paginating onto subsequent slides until all rows fit cleanly.
