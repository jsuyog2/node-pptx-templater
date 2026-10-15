# Memory: Architectural Decision Records (ADRs)

> **Repository**: `node-pptx-templater`  
> **Status**: Approved & Enforced Across Workspace

---

## Index of Architectural Decision Records

- [ADR-001: Pure JavaScript Runtime & Zero Office Dependencies](#adr-001-pure-javascript-runtime--zero-office-dependencies)
- [ADR-002: In-Memory Virtual File System (VFS) as Primary Package Abstraction](#adr-002-in-memory-virtual-file-system-vfs-as-primary-package-abstraction)
- [ADR-003: Dual-Mode XML Parsing (`fast-xml-parser` + Ordered Parser for Z-Order)](#adr-003-dual-mode-xml-parsing-fast-xml-parser--ordered-parser-for-z-order)
- [ADR-004: AST-Level Text Run Healing for Template Placeholders](#adr-004-ast-level-text-run-healing-for-template-placeholders)
- [ADR-005: Dual-Sync Strategy for Native Charts and Embedded Workbooks](#adr-005-dual-sync-strategy-for-native-charts-and-embedded-workbooks)
- [ADR-006: Unique 32-Bit Unsigned `<a16:rowId>` Allocation for Cloned Rows](#adr-006-unique-32-bit-unsigned-a16rowid-allocation-for-cloned-rows)
- [ADR-007: Strict 18-Method I/O Whitelist and Framework Adapters](#adr-007-strict-18-method-io-whitelist-and-framework-adapters)
- [ADR-008: Dynamic Row-Height Estimation and Table Auto-Pagination](#adr-008-dynamic-row-height-estimation-and-table-auto-pagination)
- [ADR-009: Decoupled Data-Driven Next.js Static Documentation Portal](#adr-009-decoupled-data-driven-nextjs-static-documentation-portal)
- [ADR-010: Flat OPC XML (`<pkg:package>`) First-Class Native Support](#adr-010-flat-opc-xml-pkgpackage-first-class-native-support)

---

### ADR-001: Pure JavaScript Runtime & Zero Office Dependencies
- **Status**: Approved
- **Context**: Many document generation libraries rely on COM interop, LibreOffice/OpenOffice headless daemons, Python unoconv processes, or native C/C++ bindings (`node-gyp`). These native dependencies fail in serverless environments (AWS Lambda, Cloudflare Workers, Vercel Edge) due to execution time limits, filesystem restrictions, and missing system binaries.
- **Decision**: Build the entire engine using pure JavaScript / TypeScript. No native binaries, no external office packages. The library operates directly on OpenXML ZIP packages and XML DOM trees.
- **Consequences**: Enables instantaneous cold-starts, sub-100ms execution times, zero native compilation issues on npm install, and 100% portability across modern serverless runtimes.

---

### ADR-002: In-Memory Virtual File System (VFS) as Primary Package Abstraction
- **Status**: Approved
- **Context**: Operating on physical disk files creates I/O bottlenecks in high-concurrency web servers (Express, Fastify) and is prohibited in read-only serverless runtimes.
- **Decision**: Introduce `VirtualFileSystem` (`src/core/VirtualFileSystem.js`), an in-memory `Map<string, Uint8Array>` representing all parts of the OpenXML package. All sub-managers read and write exclusively to the VFS.
- **Consequences**: Disk I/O is deferred entirely to the optional `saveToFile()` step. In-memory workflows (`toBuffer()`, `toStream()`, `fromCache()`) execute with zero disk interaction, maximizing throughput.

---

### ADR-003: Dual-Mode XML Parsing (`fast-xml-parser` + Ordered Parser for Z-Order)
- **Status**: Approved
- **Context**: Standard object-based XML parsers discard element order when multiple elements of different tags coexist in a parent container. However, OpenXML PresentationML visual stacking order (Z-index) in `<p:spTree>` depends strictly on the document sequence of sibling elements (`<p:sp>`, `<p:pic>`, `<p:graphicFrame>`, `<p:grpSp>`).
- **Decision**: Use `fast-xml-parser` in dual-mode:
  1. Main Parser: Fast indexed JS object for semantic access and property updates (`@_` attribute prefix, `#text`).
  2. Ordered Parser (`preserveOrder: true`): Used specifically on `<p:spTree>` to extract element ordering and attach `[Z_ORDER_SYMBOL]` to container nodes.
  3. Header Ordering: During serialization, explicitly place `<p:nvGrpSpPr>` and `<p:grpSpPr>` ahead of all visual shapes in `<p:spTree>` and `<p:grpSp>`.
- **Consequences**: Full fidelity of shape stacking order is preserved across round-trips without corrupting container non-visual properties.

---

### ADR-004: AST-Level Text Run Healing for Template Placeholders
- **Status**: Approved
- **Context**: Word processors and presentation editors fragment text into multiple `<a:r>` runs whenever characters are edited or spellchecked. Simple regex search-and-replace fails when a placeholder like `{{NAME}}` is split across two or three `<a:r>/<a:t>` nodes.
- **Decision**: In `TemplateEngine.js`, implement AST-level run aggregation. Concatenate text across all runs in an `<a:p>` paragraph, detect placeholders, execute replacements, and re-emit a consolidated `<a:r>` run that inherits the styling properties (`<a:rPr>`) of the first run.
- **Consequences**: 100% reliable placeholder replacement regardless of how PowerPoint fragmented the XML, while preserving bold, italic, font family, color, and size.

---

### ADR-005: Dual-Sync Strategy for Native Charts and Embedded Workbooks
- **Status**: Approved
- **Context**: PowerPoint charts store data in both the DrawingML `<c:chartSpace>` cache (`ppt/charts/chartX.xml`) and an embedded Excel workbook (`ppt/embeddings/Microsoft_Excel_WorksheetX.xlsx`). If only the DrawingML cache is updated, clicking "Edit Data in Excel" reveals outdated data and triggers PowerPoint repair warnings.
- **Decision**: Implement dual-synchronization in `ChartManager.js`. Every `updateChart()` call mutates both the DrawingML category/value caches and the embedded XLSX `sheet1.xml` and `sharedStrings.xml`. Conflicting `xl/tables/` definitions are automatically cleansed.
- **Consequences**: Charts open cleanly in PowerPoint, display correct figures immediately, and allow seamless editing in Microsoft Excel without repair alerts.

---

### ADR-006: Unique 32-Bit Unsigned `<a16:rowId>` Allocation for Cloned Rows
- **Status**: Approved
- **Context**: Modern PowerPoint (Office 2016+) tracks table rows using an extension element `<a16:rowId>` inside `<a:extLst>`. When rows are duplicated or appended during templating, copying the existing `rowId` creates duplicate identifiers, causing PowerPoint to declare the table corrupted.
- **Decision**: Enforce generation of unique random 32-bit unsigned integers (`0` to `4294967295`) for the `val` attribute of `<a16:rowId>` whenever rows are added, cloned, or paginated.
- **Consequences**: Completely eliminates PowerPoint table repair errors when dynamically hydrating multi-row tables.

---

### ADR-007: Strict 18-Method I/O Whitelist and Framework Adapters
- **Status**: Approved
- **Context**: Over iterative versions, numerous inconsistent import/export methods and aliases (`toFile`, `save`, `exportZip`, `exportXml`, `writeToBuffer`) accumulated, causing confusion and maintenance burden.
- **Decision**: Lock the public I/O API to exactly 18 whitelisted methods (7 Static, 8 Export, 3 Extraction & Conversion). Prohibit all deprecated aliases. Provide first-class framework adapters for Express (`toExpress`) and Fastify (`toFastify`).
- **Consequences**: Clear, unambiguous developer API contract, predictable TypeScript types, and seamless enterprise HTTP streaming.

---

### ADR-008: Dynamic Row-Height Estimation and Table Auto-Pagination
- **Status**: Approved
- **Context**: Injecting dynamic database records into template tables often overflows the bottom slide boundary, resulting in overlapping visual clutter or off-screen text.
- **Decision**: Develop `splitTableOnOverflow()` in `TableManager.js`. Using text metrics (font size, cell padding, column widths), calculate physical row heights. When cumulative height exceeds $Y_{\text{limit}}$, slice rows, deep-clone the host slide, repeat header rows, and recursively distribute records.
- **Consequences**: Flagship enterprise reporting capability: automated, publication-ready multi-slide table pagination with zero manual layout calculation.

---

### ADR-009: Decoupled Data-Driven Next.js Static Documentation Portal
- **Status**: Approved
- **Context**: Hardcoding documentation into UI components leads to out-of-sync documentation, broken links, and difficult maintenance.
- **Decision**: Build `docs-site/` using Next.js App Router with static export (`output: 'export'`) deployed to Cloudflare Pages. Drive all documentation from decoupled JSON schemas in `docs-site/src/data/` (`apiReference.json`, `features.json`, etc.). Enforce 100% DocCoverage via automated Vitest tests.
- **Consequences**: Blazing fast static site on global CDN, unified single-source-of-truth for API documentation, and automated CI validation of doc accuracy.

---

### ADR-010: Flat OPC XML (`<pkg:package>`) First-Class Native Support
- **Status**: Approved
- **Context**: Standard PPTX files are binary ZIP archives, preventing line-by-line git version control and requiring decompression on serverless runtimes.
- **Decision**: Implement `FlatOpcConverter.js` to natively ingest and export single-file Microsoft Flat OPC XML packages (`<pkg:package>`). Handle XML parts inside `<pkg:xmlData>` and binary media/workbooks inside `<pkg:binaryData>` Base64 containers.
- **Consequences**: Enables git-tracked, diffable slide decks, direct database text column storage, and zero-ZIP serverless execution.
