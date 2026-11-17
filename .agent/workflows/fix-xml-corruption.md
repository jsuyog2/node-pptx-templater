# Workflow: Diagnosing & Fixing PowerPoint XML Corruption

> **Purpose**: Forensic SOP for diagnosing and fixing Microsoft PowerPoint "Needs Repair" or schema validation errors.

---

## 1. Top Corruption Causes

1. **Duplicate `<a16:rowId>`**: Cloned rows sharing identical `rowId` values trigger instant repair alerts in Office 2016+.
2. **Missing `[Content_Types].xml` Overrides**: Unregistered slide or part paths in package manifest.
3. **Dangling Relationships**: Missing target files or invalid `rId` keys in `_rels/` files.
4. **Displaced `<p:spTree>` Headers**: Visual shapes placed before `<p:nvGrpSpPr>` or `<p:grpSpPr>`.
5. **Chart Desynchronization**: DrawingML series updated without updating embedded Excel workbook, or stale `xl/tables/` definitions.
6. **Invalid Slide ID**: Slide IDs in `ppt/presentation.xml` with `@_id < 256`.

---

## 2. Diagnostic Commands

```bash
# 1. Unpack presentation into formatted XML tree
npx node-pptx-templater extract bad.pptx --all --folder ./debug_unpacked --format-xml

# 2. Run automated package integrity audit
npx vitest run tests/helpers/packageIntegrity.js

# 3. Apply fixes and recompile
npx node-pptx-templater build ./debug_unpacked ./fixed.pptx

# 4. Validate repaired presentation
npx node-pptx-templater validate ./fixed.pptx --strict
```

---

## 3. Targeted Fixes

- **Missing Content Type**: Add `<Override PartName="/ppt/slides/slideN.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>` to `[Content_Types].xml`.
- **Duplicate Row ID**: Generate unique 32-bit uint32 (`0` to `4294967295`) for `<a16:rowId val="...">` on every row.
- **Displaced Headers**: In `XMLParser.serializeContainer()`, ensure `<p:nvGrpSpPr>` and `<p:grpSpPr>` are serialized ahead of child visual shapes.
- **Excel Table Mismatch**: In `ppt/embeddings/*.xlsx`, delete `xl/tables/` definitions and remove table references from worksheet relationships.
