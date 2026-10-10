---
name: pptx-legacy-cleaner
description: Skill for auditing, refactoring, and purging legacy patterns in node-pptx-templater. Removes synchronous blocking I/O, strips brittle regex text replacements, fixes OpenXML relationship ID collisions, and syncs embedded chart workbooks.
version: 3.0.0-alpha.0
---

# PPTX Legacy Cleaner Skill (v3.0.0-alpha.0)

## Mission & Purpose
Audit existing files in `node-pptx-templater`, purge deprecated and unstable implementations, and stabilize OpenXML packaging before developing greenfield features.

---

## Deprecation & Refactoring Directives

### 1. Zero Synchronous / Blocking File I/O
- **Scan & Purge:** Delete all occurrences of `fs.readFileSync`, `fs.writeFileSync`, `fs.readdirSync`, and `fs.statSync`.
- **Async & In-Memory Pipeline:** Replace disk-bound operations with non-blocking byte arrays (`Uint8Array`, `Buffer`) and Node.js `stream.Readable`.
- **Decouple Public APIs:** Remove constraints requiring disk file paths. Ingestion must accept in-memory buffers or streams directly.

### 2. Purge Regex-Based XML Tag Replacement
- **Identify Brittle Patterns:** Eliminate regular expressions targeting XML bodies (such as `/<p:txBody>.*<\/p:txBody>/` or `/<a:t>.*<\/a:t>/`).
- **OpenXML Run Splitting Issue:** PowerPoint commonly splits template tags across multiple runs (e.g., `<a:t>{{first_</a:t>` and `<a:t>name}}</a:t>`). Direct string replacement corrupts or misses these tokens.
- **AST Run Reconstruction:** Reconstruct and normalize text runs inside paragraph nodes (`<a:p>`) into unified tokens at the DOM/AST level before evaluating expressions and substituting values.

### 3. Stabilize OpenXML Relationship Management (`_rels`)
- **Prevent ID Collisions:** Remove hardcoded or naive global counters for relationship IDs (`rId`).
- **Isolated Manifests:** When duplicating slides, generate dedicated relationship maps (`ppt/slides/_rels/slide{M}.xml.rels`) and register references safely in `[Content_Types].xml` and `ppt/presentation.xml`.

### 4. Synchronize Embedded Excel Cache on Chart Updates
- **PowerPoint Repair Warning Bug:** Editing `<c:chartSpace>` data series without updating the corresponding embedded spreadsheet (`ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx`) causes Microsoft Office to display a file repair prompt.
- **Mandated Fix:** Synchronize both the chart XML nodes and the embedded binary workbook cache simultaneously.