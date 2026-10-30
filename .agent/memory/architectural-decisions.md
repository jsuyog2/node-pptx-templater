# Memory: Architectural Decision Records (ADRs)

> **Repository**: `node-pptx-templater` | **Status**: Approved & Enforced

---

### ADR-001: Pure JavaScript Runtime & Zero Office Dependencies
- **Context**: External Office binaries (LibreOffice, COM, `node-gyp`) fail on serverless/edge runtimes (Lambda, Workers, Edge).
- **Decision**: Build engine exclusively on pure JS data structures, `fast-xml-parser`, and JSZip with zero native dependencies.
- **Consequence**: Sub-50ms cold starts, universal portability across Node `>=18.x`, Deno, Bun, and Cloudflare Workers.

### ADR-002: In-Memory Virtual File System (VFS)
- **Context**: Physical disk writes cause concurrency bottlenecks and fail in read-only serverless environments.
- **Decision**: In-memory `Map<string, Uint8Array>` (`VirtualFileSystem.js`) manages all package files.
- **Consequence**: Zero temporary files written to disk; physical I/O deferred exclusively to explicit `saveToFile()` calls.

### ADR-003: Dual-Mode XML Parsing & Z-Order Preservation
- **Context**: Standard object XML parsing strips element order, corrupting PresentationML visual stacking in `<p:spTree>`.
- **Decision**: Dual-mode `XMLParser`: object mode for properties; ordered mode with `[Z_ORDER_SYMBOL]` for shape sequences.
- **Consequence**: Visual stacking accurately preserved while container property headers remain intact.

### ADR-004: AST-Level Text Run Healing
- **Context**: PowerPoint editors fragment placeholder strings (`{{`, `TAG`, `}}`) across multiple `<a:r>`/`<a:t>` runs.
- **Decision**: Aggregate runs per paragraph, perform substitutions across joined text, and emit single `<a:r>` inheriting primary styles.
- **Consequence**: 100% reliable placeholder substitutions with zero formatting degradation.

### ADR-005: Dual-Sync for Native Charts & Embedded Workbooks
- **Context**: Mutating DrawingML `<c:chartSpace>` without updating embedded Excel workbooks triggers PowerPoint repair alerts.
- **Decision**: Synchronize DrawingML cache and embedded `.xlsx` (`sheet1.xml`, `sharedStrings.xml`) concurrently in `ChartManager.js`.
- **Consequence**: Zero repair prompts; full Excel spreadsheet data editing retained across 17 native chart types.

### ADR-006: Unique 32-Bit Unsigned `<a16:rowId>` Allocation
- **Context**: Duplicating table rows with identical `<a16:rowId>` flags presentations as corrupt in Office 2016+.
- **Decision**: Allocate unique random 32-bit uint32 (`0` to `4294967295`) for every new or cloned row.
- **Consequence**: Guaranteed schema validity on dynamic tables.

### ADR-007: Strict 18-Method I/O Whitelist
- **Context**: Proliferation of deprecated aliases (`toFile`, `save`, `exportZip`) created maintenance bloat.
- **Decision**: Lock public I/O surface to exactly 18 methods (7 Static, 8 Export, 3 Extraction); reject all aliases.
- **Consequence**: Clean, unambiguous API contract and deterministic TypeScript typings.

### ADR-008: Dynamic Row-Height Estimation & Table Auto-Pagination
- **Context**: Unbounded table data overflows single slides, clipping content.
- **Decision**: Estimate EMU heights, slice rows exceeding limits, deep-clone host slides, and repeat headers via `splitTableOnOverflow`.
- **Consequence**: Automated, production-ready multi-slide table splitting.

### ADR-009: Decoupled Data-Driven Static Documentation
- **Context**: Manual documentation sites suffer from stale signatures and broken links.
- **Decision**: Next.js App Router static export on Cloudflare Pages driven by decoupled JSON schemas (`docs-site/src/data/`).
- **Consequence**: Single source of truth with 100% API coverage verified via automated CI tests.

### ADR-010: Native Flat OPC XML (`<pkg:package>`) Support
- **Context**: Binary `.pptx` archives cannot be diffed in Git or stored in plain-text databases.
- **Decision**: Implement `FlatOpcConverter.js` supporting single-file XML packages (`<pkg:package>`).
- **Consequence**: Direct Git diffing, text database storage, and zero-ZIP template workflows.

### ADR-011: Strict 24-Method Slide Management Whitelist
- **Context**: Slide operations lacked structured boundaries for sections, visibility, and layout binding.
- **Decision**: Enforce strict 24-method whitelist (1-based indexing) with section grouping and visibility controls.
- **Consequence**: Robust slide lifecycle management without repair alerts.

### ADR-012: Unified Elements & Hyperlinks Subsystem
- **Context**: Hyperlinks and visual layer manipulation were previously split across managers, fragmenting relationship routines.
- **Decision**: Consolidate hyperlinks (6 methods) under `src/elements/` within the canonical `Elements` category (15 methods total).
- **Consequence**: Centralized relationship lifecycle, streamlined fluent chaining, zero orphaned `.rels` references.

### ADR-013: Canonical 8 Standardized Categories Architecture
- **Context**: Discrepancies between documentation categories and internal code caused developer friction.
- **Decision**: Enforce strictly 8 approved categories across schemas, types, and docs; purge standalone diagnostics category.
- **Consequence**: Uniform categorization across all 130 public methods.
