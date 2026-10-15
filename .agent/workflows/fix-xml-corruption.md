# Workflow: Diagnosing & Fixing PowerPoint XML Corruption

> **Purpose**: Systematic forensic workflow for identifying, diagnosing, and resolving Microsoft PowerPoint "Needs Repair" or "PowerPoint found a problem with content" alerts.

---

## 1. Common Causes of PowerPoint Repair Alerts

Microsoft PowerPoint strictly enforces ISO/IEC 29500 (ECMA-376) schema validation. The top causes of corruption alerts in generated presentations are:

1. **Duplicate or Copied `<a16:rowId>` Values**:
   Cloned table rows sharing identical `rowId` values trigger instant table corruption alerts in PowerPoint 2016+.
2. **Missing `[Content_Types].xml` Overrides**:
   Adding a new slide (`slide2.xml`) without an `<Override PartName="/ppt/slides/slide2.xml" ContentType="..."/>` entry.
3. **Broken or Orphaned Relationship Chains**:
   A slide containing `r:id="rId5"` or `r:embed="rId2"` when that identifier does not exist in `ppt/slides/_rels/slideX.xml.rels`.
4. **Displaced `<p:spTree>` Header Elements**:
   Visual shapes (`<p:sp>`, `<p:pic>`) appearing before `<p:nvGrpSpPr>` or `<p:grpSpPr>` in `<p:spTree>` or `<p:grpSp>`.
5. **Chart Desynchronization**:
   Updating series values in `<c:chartSpace>` without updating the embedded `.xlsx` workbook, or leaving stale `xl/tables/` definitions.
6. **Invalid Slide ID Ranges**:
   Slide IDs (`<p:sldId id="...">`) in `ppt/presentation.xml` set to values `< 256`.

---

## 2. Step-by-Step Diagnostic Protocol

```
┌────────────────────────────────────────────────────────┐
│ Step 1: Extract Presentation with Formatting           │
│ npx node-pptx-templater extract corrupted.pptx --all   │
│     --folder ./unpacked_debug --format-xml             │
└──────────────────────────┬─────────────────────────────┘
                           ▼
┌────────────────────────────────────────────────────────┐
│ Step 2: Run Automated Integrity Validator             │
│ npx vitest run tests/helpers/packageIntegrity.js       │
└──────────────────────────┬─────────────────────────────┘
                           ▼
┌────────────────────────────────────────────────────────┐
│ Step 3: Inspect Relationship Graphs & Content Types    │
│ • Compare presentation.xml sldIdLst vs .rels           │
│ • Compare slideX.xml r:id references vs slideX.rels    │
│ • Check [Content_Types].xml for all part overrides     │
└──────────────────────────┬─────────────────────────────┘
                           ▼
┌────────────────────────────────────────────────────────┐
│ Step 4: Inspect Element Ordering & Extensions          │
│ • Inspect <p:spTree> header elements                   │
│ • Check <a16:rowId> values across <a:tr> rows          │
│ • Check chart cache vs xl/worksheets/sheet1.xml        │
└──────────────────────────┬─────────────────────────────┘
                           ▼
┌────────────────────────────────────────────────────────┐
│ Step 5: Apply Fix, Rebuild & Verify                    │
│ npx node-pptx-templater build ./unpacked_debug         │
│     ./repaired.pptx                                    │
│ npx node-pptx-templater validate ./repaired.pptx       │
└────────────────────────────────────────────────────────┘
```

---

## 3. Targeted Remediation Playbooks

### Issue A: Missing `[Content_Types].xml` Override
- **Symptom**: PowerPoint opens but slides after slide 1 are blank or missing.
- **Fix**: Open `[Content_Types].xml` and ensure every part in `ppt/slides/` has an override:
  ```xml
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  ```
- **Code Fix**: Verify that `SlideManager.js` calls `contentTypesManager.addOverride()` during `duplicateSlide()`.

### Issue B: Duplicate `<a16:rowId>` in Cloned Rows
- **Symptom**: "PowerPoint found a problem with content in presentation.pptx. PowerPoint can attempt to repair the presentation." Clicking repair removes table rows.
- **Fix**: Open `ppt/slides/slideX.xml` and search for `a16:rowId`. Ensure every row has a unique random 32-bit unsigned integer value.
- **Code Fix**: In `src/elements/TableManager.js`, ensure `#updateRowId(rowObj)` is invoked on every newly inserted or cloned row.

### Issue C: Displaced `<p:spTree>` Header Nodes
- **Symptom**: Schema validation fails on `<p:spTree>`.
- **Fix**: In `ppt/slides/slideX.xml`, ensure `<p:nvGrpSpPr>` is the first child of `<p:spTree>`, followed immediately by `<p:grpSpPr>`.
- **Code Fix**: In `src/parsers/XMLParser.js`, check `serializeContainer()`. Ensure `headerXml` (`<p:nvGrpSpPr>` and `<p:grpSpPr>`) is prepended before any visual shapes.

### Issue D: Stale Embedded Excel Table Dimensions
- **Symptom**: PowerPoint opens cleanly, but clicking "Edit Data in Excel" triggers an Excel repair dialog.
- **Fix**: In `ppt/embeddings/*.xlsx`, remove `xl/tables/` and sheet relationship references. Re-sync `sheet1.xml` with `ChartWorkbookUpdater`.
