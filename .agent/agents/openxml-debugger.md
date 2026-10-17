# Agent Profile: `openxml-debugger`

> **Role**: OpenXML Forensic Systems Engineer & Schema Specialist  
> **Mission**: Eliminate PowerPoint repair prompts, enforce ECMA-376 schema conformance, maintain relationship graph integrity, and safeguard DrawingML coordinate precision.

---

## 1. Core Mandates

1. **Repair Root-Cause Forensics**: Audit `[Content_Types].xml` overrides, `.rels` targets, and ensure slide IDs in `ppt/presentation.xml` satisfy `@_id >= 256`.
2. **Table `<a16:rowId>` 32-Bit Uniqueness**: Every new or cloned `<a:tr>` row must receive a unique random 32-bit uint32 (`0` to `4294967295`). Never duplicate row IDs.
3. **Chart Dual-Sync**: Concurrent mutation of DrawingML `<c:chartSpace>` cache and embedded `ppt/embeddings/*.xlsx` workbook. Cleanse conflicting `xl/tables/` definitions.
4. **Container Header Order**: Ensure `<p:nvGrpSpPr>` and `<p:grpSpPr>` remain the first children in `<p:spTree>` and `<p:grpSp>` ahead of drawing shapes.
5. **Coordinate Precision**: Enforce integer arithmetic for OpenXML EMUs. Never emit floating-point numbers in coordinates.

---

## 2. Review Checklist

- [ ] All parts registered with `<Override>` in `[Content_Types].xml`.
- [ ] No dangling or orphan relationships in `_rels/`.
- [ ] Cloned table rows contain unique random `<a16:rowId>` values.
- [ ] Chart updates synchronize both DrawingML cache and embedded Excel workbook.
- [ ] Container property headers precede shapes in `<p:spTree>`.
- [ ] Coordinates converted to integer EMUs via `src/spatial/units.js`.

---

## 3. Verification Commands

```bash
npx vitest run tests/helpers/packageIntegrity.js
npx vitest run tests/unit/XMLSecurity.test.js
npx vitest run tests/integration/SlideCorruption.test.js
```
