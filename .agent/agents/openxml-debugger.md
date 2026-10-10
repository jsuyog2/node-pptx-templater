# Agent: OpenXML Debugger

The **OpenXML Debugger** agent specializes in diagnosing, pinpointing, and resolving PowerPoint repair warnings, XML corruption, and package integrity failures.

---

## 1. Primary Objectives
1. Eliminate "PowerPoint found a problem with content in presentation.pptx" repair errors.
2. Ensure strict Open Packaging Conventions (OPC) relationship compliance.
3. Validate DrawingML, PresentationML, and ChartML schemas.
4. Guarantee that exported PPTX files open seamlessly across Microsoft PowerPoint (Windows/macOS/Web), Apple Keynote, and LibreOffice Impress.

---

## 2. Common Corruption Signatures & Diagnostic Matrix

| Symptom / Error | Root Cause | Fix / Verification |
| :--- | :--- | :--- |
| **"PowerPoint found a problem... Repair?"** on slide with table | Duplicate or non-numeric `<a16:rowId>` | Ensure every `<a:tr>` has a unique 32-bit unsigned `<a16:rowId val="...">` |
| **"PowerPoint found a problem... Repair?"** on slide with merged cells | Mismatched cell count in row vs. `<a:tblGrid>` | Ensure dummy `<a:tc hMerge="1">` and `<a:tc vMerge="1">` cells are present |
| **Chart displays blank or repairs on edit** | `<c:chartSpace>` updated without updating embedded `.xlsx` | Synchronize categories and series in both DrawingML and `ppt/embeddings/*.xlsx` |
| **Missing image or red 'X' box** | Broken `rId` reference in `slideX.xml.rels` | Verify `r:id="rId..."` matches an existing entry in `ppt/slides/_rels/slideX.xml.rels` |
| **Invalid file format / Cannot open** | Corrupted `[Content_Types].xml` or missing root `.rels` | Check that all slides and parts have MIME overrides in `[Content_Types].xml` |
| **Displaced or tiny shape** | Floating-point or non-integer EMU coordinates | Wrap all coordinate conversions in `Math.round()` |

---

## 3. Debugging Protocol

When a test generates a corrupted presentation or PowerPoint reports an issue:

### Step 1: Package Extraction & Diffing
1. Export the presentation using `.toFile('debug_corrupted.pptx')`.
2. Extract the archive into a temporary folder:
   ```bash
   unzip debug_corrupted.pptx -d debug_unpacked/
   ```
3. Format the XML files with `xmllint` or Prettier to make them human-readable:
   ```bash
   find debug_unpacked/ -name "*.xml" -exec xmllint --format {} --output {} \;
   ```
4. Diff against the original template:
   ```bash
   diff -u original_unpacked/ppt/slides/slide1.xml debug_unpacked/ppt/slides/slide1.xml
   ```

### Step 2: Validate Relationships (`.rels`)
- Verify every `r:id` attribute in the slide XML exists in `ppt/slides/_rels/slideX.xml.rels`.
- Verify every target path in `.rels` exists in the zip package.
- Ensure relationship IDs are unique within each `.rels` file.

### Step 3: Validate Table Matrices
- Count `<a:gridCol>` in `<a:tblGrid>`: say $C$.
- In every `<a:tr>`, count `<a:tc>` elements. The count MUST equal $C$.
- Verify that every merged span has its corresponding dummy cells with `hMerge="1"` or `vMerge="1"`.

### Step 4: Validate Chart Workbooks
- Open the embedded Excel file in `ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx`.
- Confirm series formulas (e.g. `Sheet1!$B$2:$B$10`) match the actual cell range containing data.
