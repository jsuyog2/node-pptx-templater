# Memory: Architectural Decision Records (ADRs)

> **Repository**: `node-pptx-templater` | **Status**: Approved & Enforced

---

### ADR-001: Pure JavaScript Runtime & Zero Office Dependencies
- **Context**: Office binaries (LibreOffice, COM, `node-gyp`) fail on serverless platforms (AWS Lambda, Cloudflare Workers, Edge).
- **Decision**: Zero native dependencies. Operate purely on JS data structures, `fast-xml-parser`, and JSZip.
- **Consequence**: Sub-100ms startup, instant edge compatibility, no compilation overhead.

### ADR-002: In-Memory Virtual File System (VFS)
- **Context**: Disk I/O causes concurrency bottlenecks and fails in read-only serverless runtimes.
- **Decision**: In-memory `Map<string, Uint8Array>` (`VirtualFileSystem.js`) handles all package reads and writes.
- **Consequence**: Zero synchronous disk blocking; physical disk I/O deferred exclusively to `saveToFile()`.

### ADR-003: Dual-Mode XML Parsing & Z-Order Preservation
- **Context**: Standard XML object parsing loses element ordering, which PresentationML uses to define visual Z-order in `<p:spTree>`.
- **Decision**: Dual-mode `XMLParser`: object mode for properties; ordered mode with `[Z_ORDER_SYMBOL]` for shape sequence. Container headers (`<p:nvGrpSpPr>`, `<p:grpSpPr>`) always serialized first.
- **Consequence**: Shape stacking hierarchy preserved without corrupting non-visual headers.

### ADR-004: AST-Level Text Run Healing
- **Context**: PowerPoint splits placeholders (e.g. `{{`, `NAME`, `}}`) across multiple `<a:r>`/`<a:t>` nodes.
- **Decision**: In `TemplateEngine.js`, aggregate runs per `<a:p>`, substitute text across the joined buffer, and re-emit a single `<a:r>` inheriting the first run's `<a:rPr>`.
- **Consequence**: 100% reliable placeholder substitutions with zero formatting loss.

### ADR-005: Dual-Sync for Native Charts & Embedded Workbooks
- **Context**: Updating only DrawingML `<c:chartSpace>` without updating embedded Excel workbooks causes repair alerts on edit.
- **Decision**: Synchronize DrawingML cache and embedded `.xlsx` (`sheet1.xml`, `sharedStrings.xml`) concurrently in `ChartManager.js`. Cleanse stale `xl/tables/`.
- **Consequence**: Zero repair warnings in PowerPoint; immediate editing in Excel supported.

### ADR-006: Unique 32-Bit Unsigned `<a16:rowId>` Allocation
- **Context**: Office 2016+ flags presentations as corrupt if cloned table rows share identical `<a16:rowId>` values.
- **Decision**: Generate unique random 32-bit uint32 (`0` to `4294967295`) for every new or cloned row.
- **Consequence**: Completely prevents table repair prompts on dynamic tables.

### ADR-007: Strict 18-Method I/O Whitelist & Framework Adapters
- **Context**: Inconsistent I/O aliases created maintenance overhead and API confusion.
- **Decision**: Lock public I/O surface to exactly 18 methods (7 Static, 8 Export, 3 Extraction); provide Express and Fastify adapters.
- **Consequence**: Unambiguous API boundary, clean TypeScript types, strict linting.

### ADR-008: Dynamic Row-Height Estimation & Table Auto-Pagination
- **Context**: Dynamic tables overflow slide bounds, causing off-screen text.
- **Decision**: `splitTableOnOverflow()` computes row heights, slices rows exceeding $Y_{\text{limit}}$, deep-clones host slides, and repeats header rows recursively.
- **Consequence**: Automated, publication-ready multi-slide table pagination.

### ADR-009: Decoupled Data-Driven Static Documentation
- **Context**: Hardcoded doc components lead to stale references and broken links.
- **Decision**: Next.js App Router static export (`output: 'export'`) on Cloudflare Pages driven by decoupled JSON schemas (`docs-site/src/data/`). 100% DocCoverage enforced via Vitest.
- **Consequence**: Single source of truth, blazing-fast static documentation, CI doc integrity checks.

### ADR-010: Native Flat OPC XML (`<pkg:package>`) Support
- **Context**: Binary `.pptx` archives cannot be diffed in Git or stored easily in text database columns.
- **Decision**: Implement `FlatOpcConverter.js` for single-file XML packages (`<pkg:package>`).
- **Consequence**: Git-diffable presentations, direct database text storage, zero-ZIP processing.

### ADR-011: Strict 23-Method Slide Management Whitelist
- **Context**: Slide methods had accumulated deprecated aliases and lacked structured section, visibility, and comment management.
- **Decision**: Enforce strict 23-method whitelist (1-based indexing). Implement section grouping via `<p14:sectionLst>`, visibility via `show="0"`, canvas reset (`resetSlide`), and `CommentManager`.
- **Consequence**: Predictable slide manipulation with zero repair alerts.
