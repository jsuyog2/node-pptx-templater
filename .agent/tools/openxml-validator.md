# Tool: OpenXML Package & XML Schema Validator

> **Target**: `src/core/ValidationEngine.js`, `tests/helpers/packageIntegrity.js`, CLI `validate`/`debug`  
> **Purpose**: Automated structural and schema auditing for OpenXML presentation packages.

---

## 1. Programmatic API (`ValidationEngine.js`)

```javascript
const ppt = await PPTXTemplater.load('presentation.pptx');
const result = await ppt.validatePresentation(); // { valid: boolean, errors: string[], warnings: string[] }
```

Static validators:
- `ValidationEngine.validatePresentation(vfs, xmlParser)`: Package-wide ECMA-376 compliance.
- `ValidationEngine.validateTable(tableObj)`: Grid columns, row heights, and `<a16:rowId>` uniqueness.
- `ValidationEngine.validateList(listConfig)`: List paragraph structure, styles, and nesting depths.

---

## 2. Test Harness Helper (`packageIntegrity.js`)

Used in Vitest suites to audit output presentation buffers:
```javascript
import { validatePackageIntegrity } from '../helpers/packageIntegrity.js';
const errors = await validatePackageIntegrity(await ppt.toBuffer());
expect(errors).toHaveLength(0);
```
Validates:
1. `[Content_Types].xml`, `ppt/presentation.xml`, and `_rels/.rels` presence.
2. Slide relationship IDs vs `p:sldIdLst` count and valid part targets.
3. Zero orphan slides and complete `<Override>` coverage in `[Content_Types].xml`.
4. Valid `.rels` target resolutions for all `r:embed`/`r:id` references.

---

## 3. CLI Validation Commands

```bash
# Structural validation
npx node-pptx-templater validate output/presentation.pptx [--strict]

# Deep debug and auto-repair
npx node-pptx-templater debug broken.pptx --fix -o repaired.pptx
```
