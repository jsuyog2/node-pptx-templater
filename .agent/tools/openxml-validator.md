# Tool: OpenXML Package Validator

The **OpenXML Package Validator** is an automated validation tool used by agents, CI pipelines, and developers to verify Open Packaging Conventions (OPC) and OpenXML schema compliance of generated presentations.

---

## 1. Capabilities & Validation Scope

1. **ZIP Container & OPC Package Verification**:
   - Confirms valid ZIP structure and lack of directory traversal entries (`../`).
   - Ensures `[Content_Types].xml` exists and contains correct MIME overrides for all slides, layouts, themes, and media.
   - Ensures `_rels/.rels` and `ppt/_rels/presentation.xml.rels` exist and point to valid part targets.

2. **Relationship Graph Consistency**:
   - Verifies that every `r:id` referenced inside a slide (`<p:sld>`) exists in the slide's corresponding `_rels/slideX.xml.rels`.
   - Checks that all relationship target files exist within the archive.
   - Ensures no dangling or duplicate relationship IDs exist.

3. **Slide & DrawingML Integrity**:
   - Verifies that every shape in `<p:spTree>` has valid spatial offsets (`<a:off>`) and extents (`<a:ext>`).
   - Ensures coordinates are rounded integer EMUs (rejects floating points).
   - Validates that text runs (`<a:r>`) contain valid `<a:t>` tags.

4. **Table Matrix Integrity**:
   - Compares `<a:gridCol>` count in `<a:tblGrid>` against cell count in each `<a:tr>`.
   - Checks that cell spans (`gridSpan`, `rowSpan`) have corresponding dummy cells (`hMerge="1"`, `vMerge="1"`).
   - Verifies that each `<a:tr>` has a unique 32-bit unsigned `<a16:rowId>`.

5. **Chart & Embedded Workbook Sync**:
   - Verifies that `<c:ptCount>` matches the length of `<c:pt>` data points.
   - Verifies that `idx` values are contiguous starting from 0.
   - If an embedded `.xlsx` workbook exists, checks that series data matches the embedded sheet.

---

## 2. Invocations

### 2.1 Programmatic Validation (Vitest / Integration Tests)
```javascript
import { validateOpenXMLPackage } from '../src/utils/openXmlValidator.js';
import fs from 'node:fs';

const pptxBuffer = fs.readFileSync('output.pptx');
const result = await validateOpenXMLPackage(pptxBuffer);

if (!result.isValid) {
  console.error('Validation Errors:', result.errors);
  throw new Error(`OpenXML Validation failed with ${result.errors.length} errors`);
}
```

### 2.2 CLI / npm Script
```bash
# Validate generated showcase examples
npm run validate

# Run targeted validation test suite
npx vitest run tests/validation.test.js
```

---

## 3. Diagnostic Report Format

When failures occur, the validator outputs structured diagnostic reports:

```json
{
  "isValid": false,
  "summary": "2 errors found in ppt/slides/slide2.xml",
  "errors": [
    {
      "code": "TABLE_CELL_COUNT_MISMATCH",
      "part": "ppt/slides/slide2.xml",
      "element": "a:tr[3]",
      "message": "Row 3 has 4 cells, but tblGrid defines 5 columns. Missing hMerge/vMerge dummy cell."
    },
    {
      "code": "NON_INTEGER_EMU",
      "part": "ppt/slides/slide2.xml",
      "element": "a:off",
      "message": "Attribute @_x contains floating point value '12700.5'. Must be integer."
    }
  ],
  "warnings": [
    {
      "code": "UNUSED_RELATIONSHIP",
      "part": "ppt/slides/_rels/slide2.xml.rels",
      "message": "Relationship rId4 (ppt/media/image2.png) is not referenced by any shape."
    }
  ]
}
```
