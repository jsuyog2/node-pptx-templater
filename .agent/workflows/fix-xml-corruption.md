# Workflow: Fixing OpenXML Corruption

This workflow provides a systematic procedure for debugging, isolating, and resolving OpenXML corruption and PowerPoint repair warnings.

---

## Step 1: Reproduce & Isolate

1. **Pinpoint the Trigger**:
   - Identify the exact code path, data structure, or template causing PowerPoint to trigger the "PowerPoint found a problem with content... Do you want to repair?" warning.
2. **Minimal Reproduction**:
   - Write a minimal failing test case in `tests/debug_repro.test.js` or run the specific example in `showcase/examples.js`.
   - Export the corrupted presentation to disk: `await pptx.toFile('corrupted.pptx')`.

---

## Step 2: Unpack and Format Package Contents

1. **Extract Archive**:
   Extract both the clean input template and the corrupted output into separate directories:
   ```bash
   mkdir clean_dump && unzip template.pptx -d clean_dump/
   mkdir corrupt_dump && unzip corrupted.pptx -d corrupt_dump/
   ```
2. **Pretty-Print XML**:
   Format XML files for diffing:
   ```bash
   # Format all XML files
   npx prettier --write "corrupt_dump/**/*.xml"
   npx prettier --write "clean_dump/**/*.xml"
   ```

---

## Step 3: Inspect Diff & Common Culprits

Execute a git diff or file diff across the directories:
```bash
git diff --no-index clean_dump/ corrupt_dump/
```

Check the following known failure patterns:

1. **Table Structure**:
   - Does `<a:tbl>` have an `<a:tblGrid>` with $N$ `<a:gridCol>` elements?
   - Does every `<a:tr>` contain exactly $N$ `<a:tc>` elements?
   - Are merged cells accompanied by dummy cells (`<a:tc hMerge="1">` or `<a:tc vMerge="1">`)?
   - Does every `<a:tr>` contain a unique 32-bit unsigned `<a16:rowId>`?
2. **Relationships (`.rels`)**:
   - Does every `r:id="rIdX"` in `ppt/slides/slideN.xml` exist in `ppt/slides/_rels/slideN.xml.rels`?
   - Do all target file paths in `.rels` exist in the ZIP container?
   - Is `[Content_Types].xml` updated with an `<Override>` entry for any newly added part?
3. **Coordinates & Measurements**:
   - Are any values in `<a:off @_x="..." @_y="..."/>` or `<a:ext @_cx="..." @_cy="..."/>` floating-point numbers? (They must be integer EMUs).
4. **Native Charts**:
   - Does `<c:ptCount @_val="N"/>` equal the number of `<c:pt>` entries?
   - Are `idx` attributes 0-indexed and continuous?
   - Is the embedded Excel spreadsheet in `ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx` synchronized with the DrawingML series?

---

## Step 4: Implement Fix & Add Regression Test

1. **Apply the Fix**:
   - Modify the relevant element manipulator or manager in `src/`.
2. **Write a Vitest Regression Test**:
   - Add a test in `tests/` that recreates the exact scenario and asserts:
     - The output XML contains the required tags, attributes, and dummy cells.
     - `validateOpenXMLPackage(pptxBuffer)` reports `isValid === true`.
3. **Run Validation**:
   ```bash
   npm test
   npm run validate
   ```

---

## Step 5: Verification in PowerPoint / LibreOffice

1. Open the resulting `.pptx` in Microsoft PowerPoint or LibreOffice Impress.
2. Verify:
   - File opens immediately without any repair prompts or modal alerts.
   - Merged table cells render with correct visual bounds.
   - Charts render properly, and clicking "Edit Data in Excel" opens the spreadsheet with accurate values.
   - Text formatting, colors, and line breaks match expectations.
