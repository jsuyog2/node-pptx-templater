# Tool: OpenXML Package & XML Schema Validator

> **Tool Target**: `src/core/ValidationEngine.js`, `tests/helpers/packageIntegrity.js`, CLI `validate`/`debug`  
> **Purpose**: Automated structural and schema auditing for OpenXML presentation packages.

---

## 1. Overview

The OpenXML Validator tool suite provides multi-layered verification of PowerPoint files, checking for:
1. Core package manifest integrity (`[Content_Types].xml`, `_rels/.rels`, `ppt/presentation.xml`).
2. Relationship reference completeness (ensuring no dangling `rId` references).
3. Slide ID uniqueness and value ranges (`slideId >= 256`).
4. Table structure integrity (row ID uniqueness, merge coordinate validity).
5. XML security compliance (detecting XXE and entity bombs).

---

## 2. API Validation Methods (`ValidationEngine.js`)

The engine provides built-in validation methods callable directly on any `PPTXTemplater` instance:

```javascript
const ppt = await PPTXTemplater.load('presentation.pptx');

// Run comprehensive presentation audit
const result = await ppt.validatePresentation();
console.log(result.valid);  // true | false
console.log(result.errors); // Array of error strings
console.log(result.warnings); // Array of non-fatal warnings
```

### Static Engine Validators:
- **`ValidationEngine.validatePresentation(vfs, xmlParser)`**:
  Validates entire VFS package against ECMA-376 rules.
- **`ValidationEngine.validateTable(tableObj)`**:
  Validates grid columns, row heights, and `<a16:rowId>` uniqueness across table rows.
- **`ValidationEngine.validateList(listConfig)`**:
  Validates list paragraph structure, levels, styles, and nesting depths.

---

## 3. Test Harness Validator (`tests/helpers/packageIntegrity.js`)

Used in Vitest integration tests to audit output buffers before assertions:

```javascript
import { validatePackageIntegrity } from '../helpers/packageIntegrity.js';

const buffer = await ppt.toBuffer();
const errors = await validatePackageIntegrity(buffer);

if (errors.length > 0) {
  console.error('Integrity failures:', errors);
}
expect(errors).toHaveLength(0);
```

### What `validatePackageIntegrity` Checks:
1. Presence of `[Content_Types].xml`, `ppt/presentation.xml`, and `_rels/.rels`.
2. `p:sldIdLst` count in `presentation.xml` exactly matches the slide relationship count in `ppt/_rels/presentation.xml.rels`.
3. Every slide ID points to a part that actually exists in `ppt/slides/`.
4. Zero orphan slides (every `slideX.xml` part is registered in presentation manifests).
5. Every slide has an explicit `<Override>` entry in `[Content_Types].xml`.
6. Every slide has a valid `.rels` file, and all `r:embed`, `r:id`, and `r:link` attributes in `slideX.xml` resolve to targets defined in `slideX.xml.rels`.

---

## 4. CLI Validation & Debug Commands

```bash
# 1. Structural validation via CLI
npx node-pptx-templater validate output/presentation.pptx

# 2. Strict validation (fails on warnings)
npx node-pptx-templater validate output/presentation.pptx --strict

# 3. Deep debug and auto-repair
npx node-pptx-templater debug broken.pptx --fix -o repaired.pptx
```
