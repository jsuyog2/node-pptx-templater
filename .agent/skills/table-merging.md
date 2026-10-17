# Skill: Table Cell Merging, Row ID Integrity & Auto-Pagination

> **Domain**: PresentationML Table Processing (`<a:tbl>`), Extension Lists (`<a16:rowId>`), and Overflow Pagination  
> **Target Subsystems**: `TableManager.js`, `SlideManager.js`, `spatial/units.js`

---

## 1. Cell Merging Playbook

Tables use 0-based indexing: `(startRow, startCol)` to `(endRow, endCol)`.

1. **Calculate Spans**:
   `gridSpan = endCol - startCol + 1;` `rowSpan = endRow - startRow + 1;`
2. **Origin Cell `(startRow, startCol)`**:
   Set `@_gridSpan` (if `gridSpan > 1`), `@_rowSpan` (if `rowSpan > 1`). Remove `@_hMerge` and `@_vMerge`.
3. **Horizontally Shadowed Cells** (`r = startRow`, `startCol < c <= endCol`):
   Set `@_hMerge = "1"`, remove `@_gridSpan`, clear `<a:txBody>` text.
4. **Vertically Shadowed Cells** (`startRow < r <= endRow`, `startCol <= c <= endCol`):
   Set `@_vMerge = "1"` (and `@_hMerge = "1"` if also horizontally shadowed), remove `@_rowSpan`, clear text.

---

## 2. Row ID Integrity (`<a16:rowId>`)

Modern PowerPoint (Office 2016+) tracks table rows using `<a16:rowId>` inside `<a:extLst>`:
- **Mandate**: When creating, cloning, or duplicating table rows, ALWAYS allocate a fresh random 32-bit unsigned integer (`val = String(Math.floor(Math.random() * 4294967295))`).
- **Prohibition**: Never duplicate `rowId` values across rows; duplicates trigger instant repair warnings.

---

## 3. Table Auto-Pagination Engine (`splitTableOnOverflow`)

When table records exceed the slide boundary:
1. **Compute Available Height**: $Y_{\text{limit}} = H_{\text{slide}} - M_{\text{bottom}}$ (in EMUs).
2. **Estimate Row Heights**: Calculate physical row heights based on font size, cell margins, and wrapped text lines.
3. **Partition & Duplicate**: Slice rows at cutoff index $k$. Deep-clone the host slide via `slideManager.duplicateSlide()`.
4. **Inject Headers & Summaries**: Prepend repeated headers (`repeatHeaders > 0`) on cloned slides; pin summary rows (`summaryRows > 0`) to the final slide.
5. **Recurse**: Repeat pagination until all table rows fit cleanly.
