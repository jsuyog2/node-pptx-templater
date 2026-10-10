# Skill: OpenXML Table Merging & Row Generation

This skill details how to manipulate DrawingML tables (`<a:tbl>`), apply horizontal and vertical cell merges, and generate rows with valid `<a16:rowId>` identifiers.

---

## 1. Structure of an OpenXML Table

An OpenXML table in PresentationML is defined under `<a:tbl>`:

```xml
<a:tbl>
  <a:tblPr rtl="0" firstRow="1">
    <a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId>
  </a:tblPr>
  <a:tblGrid>
    <a:gridCol w="2500000"/>
    <a:gridCol w="2500000"/>
    <a:gridCol w="2500000"/>
  </a:tblGrid>
  <a:tr h="450000">
    <a16:rowId val="1001" xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main"/>
    <a:tc>...</a:tc>
    <a:tc>...</a:tc>
    <a:tc>...</a:tc>
  </a:tr>
</a:tbl>
```

---

## 2. Horizontal Merging (`gridSpan` & `hMerge`)

To merge cells across columns in the same row:

1. **Origin Cell (Top-Left)**:
   Add the attribute `gridSpan="COL_COUNT"` to the origin `<a:tc>`.
2. **Covered Cells (Right)**:
   Add dummy `<a:tc>` elements for the remaining `COL_COUNT - 1` columns with the attribute `hMerge="1"`.

```xml
<!-- A 3-column row where cell 1 spans across columns 1 and 2 -->
<a:tr h="400000">
  <!-- Origin cell spanning 2 columns -->
  <a:tc gridSpan="2">
    <a:txBody>
      <a:bodyPr/>
      <a:p><a:r><a:t>Spanning Columns 1 & 2</a:t></a:r></a:p>
    </a:txBody>
    <a:tcPr/>
  </a:tc>
  <!-- Dummy covered cell -->
  <a:tc hMerge="1">
    <a:txBody><a:bodyPr/><a:p/></a:txBody>
    <a:tcPr/>
  </a:tc>
  <!-- Normal third cell -->
  <a:tc>
    <a:txBody><a:bodyPr/><a:p><a:r><a:t>Column 3</a:t></a:r></a:p></a:txBody>
    <a:tcPr/>
  </a:tc>
</a:tr>
```

---

## 3. Vertical Merging (`rowSpan` & `vMerge`)

To merge cells vertically down rows:

1. **Origin Cell (Top)**:
   Add the attribute `rowSpan="ROW_COUNT"` to the origin `<a:tc>`.
2. **Covered Cells (Below)**:
   In the subsequent `ROW_COUNT - 1` rows, place dummy `<a:tc>` elements in that column with the attribute `vMerge="1"`.

```xml
<!-- Row 1: Cell 1 spans 2 rows down -->
<a:tr h="400000">
  <a:tc rowSpan="2">
    <a:txBody>
      <a:bodyPr/>
      <a:p><a:r><a:t>Spanning Rows 1 & 2</a:t></a:r></a:p>
    </a:txBody>
    <a:tcPr/>
  </a:tc>
  <a:tc>...</a:tc>
</a:tr>

<!-- Row 2: Cell 1 is vertically covered -->
<a:tr h="400000">
  <a:tc vMerge="1">
    <a:txBody><a:bodyPr/><a:p/></a:txBody>
    <a:tcPr/>
  </a:tc>
  <a:tc>...</a:tc>
</a:tr>
```

---

## 4. Crucial Invariant: The Equal Cell Matrix Rule

**RULE**: Every `<a:tr>` MUST have the exact same number of `<a:tc>` children as the total number of `<a:gridCol>` definitions in `<a:tblGrid>`.

- Even if a cell is merged horizontally or vertically, its placeholder `<a:tc>` MUST exist in the XML tree.
- Omitting covered cells will cause PowerPoint to fail schema validation and corrupt the presentation.

---

## 5. Generating Unique `<a16:rowId>` Values

Modern versions of PowerPoint (2016+) assign a unique 32-bit unsigned integer to each row.

```javascript
/**
 * Generates a unique 32-bit unsigned row ID for an OpenXML table row.
 * @param {Set<number>} existingIds Set of currently used row IDs in the presentation.
 * @returns {number}
 */
export function generateUniqueRowId(existingIds = new Set()) {
  let id;
  do {
    // Generate an unsigned 32-bit integer (1 to 2^31 - 1)
    id = Math.floor(Math.random() * 2147483647) + 1;
  } while (existingIds.has(id));
  existingIds.add(id);
  return id;
}

// Applying to newly cloned or created row AST:
newRow['a16:rowId'] = {
  '@_val': String(generateUniqueRowId(existingRowIds)),
  '@_xmlns:a16': 'http://schemas.microsoft.com/office/drawing/2014/main'
};
```

**Never**:
- Leave `<a16:rowId>` as duplicate when cloning a row.
- Assign non-numeric or negative strings.
- Exceed $4,294,967,295$ (the 32-bit unsigned limit).
