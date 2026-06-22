# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.1.3] - 2026-06-22

### Fixed

- **Layout Preservation in `addCellShape()`**: Fixed a critical layout bug where querying cell bounds or adding cell shapes dynamically (via `addCellShape()` or `updateCellShape()`) mutated table row heights inside the XML. The sizing/positioning logic now runs in a layout-only mode, keeping the table XML completely untouched and respecting original template custom row heights.

## [1.1.2] - 2026-06-22

### Fixed

- **Slide Link Relationships Target Paths**: Fixed PowerPoint "Repair Mode" errors when adding slide-to-slide hyperlinks (via `addSlideLink` or `addImageLink`). The relative target path resolves correctly to the sibling filename `slideX.xml` instead of using the redundant parent prefix `../slides/slideX.xml` (which exited and re-entered the same directory, violating PowerPoint's relative path resolution rules).

### Added

- **Redundant Traversal Validation**: Added automated check in `ValidationEngine` to identify and error on redundant relative path traversals inside relationship files (e.g. referencing `../slides/slideX.xml` from a source slide file already located in `ppt/slides/`).

## [1.1.0] - 2026-06-12

### Added

- **Runtime Log Level Control**: `PPTXTemplater.load(path, { logLevel: 'debug' })` — configure logging at load time without environment variables. Also added `PPTXTemplater.setLogLevel(level)` static method and `ppt.enableDebug()` instance shortcut. Supported levels: `verbose`, `debug`, `info`, `warn` (default), `error`, `silent`. The new `setGlobalLogLevel()` and `resetLogLevel()` functions are now exported from the public API.

- **`verbose` Log Level**: New log level below `debug` for maximum diagnostic output. Use `PPTXTemplater.setLogLevel('verbose')` to enable.

- **`getTableRows(tableId, options)`**: Extract table data as structured JSON. Supports three modes: default (array of objects using header row as keys), `{ raw: true }` (array of string arrays), and `{ includeMetadata: true }` (full metadata including row/column count and merged cell info).

- **Nested `addTableRow()` with `mergeStrategy`**: Add rows with nested arrays to create rowspan-merged cells. Options: `'rowspan'` (OpenXML vertical spans), `'auto'` (merge identical adjacent values), `'none'` (expand to flat rows).

- **8 new example files** in `examples/`: `image-operations.js`, `shape-operations.js`, `z-order.js`, `slide-import.js`, `text-search.js`, `table-extraction.js`, `nested-table-rows.js`, `xml-folder-workflow.js`.

- **Documentation Validation Tooling** (`scripts/validate-docs.js`): Checks that every public method has JSDoc. Exits with code 1 on violations. Run via `npm run docs:validate`.

- **Feature Inventory Generator** (`scripts/generate-feature-inventory.js`): Parses `PPTXTemplater.js` and outputs `docs/feature-inventory.json` and `docs/feature-inventory.md` with all methods grouped by category. Run via `npm run docs:inventory`.

- **Professional logo** (`assets/logo.png`): Library logo suitable for GitHub, NPM, and documentation.

- **New npm scripts**: `docs:validate`, `docs:inventory`, `example:images`, `example:shapes`, `example:zorder`, `example:slide-import`, `example:text-search`, `example:table-extraction`, `example:nested-rows`, `example:xml-folder`.

### Fixed

- **Horizontal merge preservation**: Fixed PowerPoint "repair mode" errors when using `addTableRow()` with rows containing `gridSpan`/`hMerge` attributes. The cloning logic now correctly preserves horizontal merge attributes and only clears vertical merge state.

- **`console.log` violations** in `ChartManager.js` (3 calls in `updateChartAsync()`), `OutputWriter.js` (5 calls in `printDebugZip()`), and `PPTXTemplater.js` (8 calls across `debugRelationships()`, `inspectSlide()`, `inspectXML()`, `inspectChart()`, `debugChartRelationships()`). All replaced with structured `logger.debug()` / `logger.info()` calls that respect the configured log level. **The library is now completely silent by default** — no terminal output at all unless explicitly enabled.

### Changed

- **Logger system overhauled**: Added `verbose` level, runtime `setGlobalLogLevel()` function, `resetLogLevel()` function, and module-level `runtimeLevel` override. All logger instances now share a mutable runtime level for live updates.

- **README**: Rewritten to be concise, modern, and SEO-friendly. Detailed API docs now live at https://jsuyog2.github.io/node-pptx-templater/.

- **Version**: Bumped from `1.0.21` → `1.1.0`.

---

## [1.0.6] - 2026-06-02


### Added

- **XML Validation & Diagnostics Engine**: Introduced a suite of tools in `src/utils/xmlUtils.js` for XML safety and diagnostics:
  - `validateXml(xmlString)` — Validates that an XML string is secure and well-formed, checking for DTDs, custom/recursive entities, and XXE.
  - `safeParseXml(xmlString, file)` — Unified wrapper that runs validation and captures detailed diagnostic error logs (file, line, col, error details, and recommendations) on failure.
  - `scanForEntities(xmlString)` — Scans and classifies all XML entity references (standard, custom, numeric, and hex).
  - `analyzeXmlFile(xmlString)` — Computes core file sizing and stats (bytes, lines, elements, attributes, entities).
  - `reportXmlComplexity(xmlString)` — Inspects structural metrics (maximum tag nesting depth, node count, text-to-markup ratio).
- **Public API Exports**: The new tools are exported from the main library entry point `src/index.js`.

### Fixed

- **XML Entity Expansion Limit Resolution**: Permanently resolved the `Entity expansion limit exceeded` parser crashes on large template files. Deactivated internal entity expansion in `fast-xml-parser` and replaced it with a fast, secure, non-recursive unescaper handling the 5 standard XML/HTML entities and numeric references (decimal and hex code points) natively.
- **Vulnerability Protections**: Integrated strict security checks directly into the validator to block DTD abuse, XML bombs (Billion Laughs), and XXE attacks safely before the parser processes them.

### Tests

- Added 13 new unit and integration tests in `tests/unit/XMLSecurity.test.js` validating security protections, large-scale entity processing, diagnostics error recovery, and complexity analysis.
- Total test count increased from 108 → 121 (all passing).

## [1.0.5] - 2026-06-02

### Added

- **Z-Order (Layer Management) System**: Full stacking control for all slide drawing objects — shapes, images, charts, tables, groups, connectors, and SmartArt. Directly manipulates the OpenXML `<p:spTree>` element order, matching PowerPoint's native Bring Forward / Send Backward behavior exactly. New APIs:
  - `getObjectOrder(slideIndex)` — Returns ordered metadata (id, type, zIndex) for every element on a slide, bottom-to-top.
  - `bringForward(options)` — Moves an object one layer up in the stack.
  - `sendBackward(options)` — Moves an object one layer down.
  - `bringToFront(options)` — Moves an object to the very top of the stack.
  - `sendToBack(options)` — Moves an object to the very bottom of the stack.
  - `setZIndex(options)` — Places an object at an exact 1-based stacking position.
  - `moveObjectBefore(options)` — Positions an object immediately below a named target.
  - `moveObjectAfter(options)` — Positions an object immediately above a named target.
  - `reorderObjects(options)` — Full bulk reorder of the slide stack from a given array.
  - `applyZOrder(slideIndex, configs)` — Applies multiple stacking rules sequentially in one call.
  - `swapObjects(slideIndex, id1, id2)` — Exchanges two objects' positions.
  - `sortObjects(slideIndex, compareFn)` — Sorts the stack using a custom comparator.
  - `getTopMostObject(slideIndex)` / `getBottomMostObject(slideIndex)` — Inspection helpers.
  - `normalizeZOrder(slideIndex)` — Re-derives and resets internal Z-order state from the current XML.
- **Z_ORDER_SYMBOL Export**: The `Z_ORDER_SYMBOL` is now exported from `src/index.js` for advanced integrations.
- **ZOrderManager**: New dedicated manager class (`src/managers/ZOrderManager.js`) encapsulating all layer logic.

### Fixed

- **`PPTXTemplater.create()` synchronous readiness**: Added `preloadAll()` call to `#initializeBlank()`. Previously, the blank PPTX template's pre-existing slides were registered but their XML was not cached, causing all synchronous operations (including ZOrderManager) to throw `"Slide N XML not pre-loaded"`.

### Changed

- **`XMLParser` hybrid parsing**: Added a secondary `preserveOrder: true` fast-xml-parser pass that runs during `parse()` whenever a slide `<p:spTree>` is detected. Extracts DOM element order and attaches it via `Z_ORDER_SYMBOL` to each container. The `build()` method uses a new `serializeContainer()` recursive function to serialize containers in Z_ORDER_SYMBOL order, injecting the result back into the output XML.
- **`ValidationEngine`**: `validate()` now audits the shape tree for duplicate shape IDs, reporting them as errors.

### Tests

- Added 12 new integration tests in `tests/integration/ZOrder.test.js` covering all Z-order operations.
- Total test count increased from 96 → 108 (all passing).

## [1.0.3] - 2026-06-02

### Added

- **Dynamic Formatting in updateTable**: Added support for inline cell styling (color fill `fill`, text alignment `align`, and `fontSize`) directly on cell objects passed to `updateTable`.
- **Comprehensive Tailwind Site**: Overhauled doc builder script to generate a premium Tailwind CSS documentation portal with clientside search, clipboard copying, sitemap.xml, robots.txt, and Schema.org metadata.

### Fixed

- **XML Element Ordering**: Enforced strict schema-valid element sequence (`a:pPr` -> runs -> `a:endParaRPr`) in slide table cell paragraphs. This resolves the bug where split cells inheriting from template merged cells had their text runs ignored by PowerPoint's XML compiler.
- **Template Style Inheritance**: Fixed a bug in `updateTable` where cloned rows always copied the first data row (`trs[1]`). The engine now correctly inherits formatting, alignment, and fill styles from matching indices in the template (`trs[i]`) when available.

## [1.0.2] - 2026-06-02

### Added

- **Table Cell Merging & Unmerging Engine**: Fully implemented horizontal cell spans (`gridSpan`, `hMerge`), vertical cell spans (`rowSpan`, `vMerge`), and rectangular block merges.
- **PowerPoint Repair Protection**: Implemented unique 32-bit unsigned `rowId` generation inside `<a16:rowId>` XML tags for all cloned and inserted rows, eliminating PowerPoint's "Repair Mode" error prompts.
- **Merge Integrations**: Integrated template-driven merges (`merge` configs array and cell-level `colSpan`/`rowSpan`) inside the main `updateTable` orchestrator.
- **Integration Test Suite**: Added a comprehensive merge test script under `tests/integration/PPTXMerge.test.js`.

## [1.0.1] - 2026-05-19

### Changed

- **CommonJS Target Conversion**: Converted the source code modules compilation and packaging layout from pure ES Modules (ESM) to CommonJS (CJS) to ensure compatibility with standard Node.js deployment, packaging, and edge runtime environments.

## [1.0.0] - 2026-05-17

### Added

- `PPTXTemplater` — main orchestrator class with fluent chainable API
- `ZipManager` — PPTX ZIP archive loading, reading, writing, and re-packaging
- `XMLParser` — high-performance XML parsing/building via `fast-xml-parser`
- `RelationshipManager` — OpenXML `.rels` file parsing and management
- `SlideManager` — slide discovery, ordering, addition, cloning, and removal
- `ChartManager` — direct chart XML data updates (bar, line, pie, area, scatter)
- `TableManager` — table row replacement preserving all formatting
- `HyperlinkManager` — external URL and slide-to-slide hyperlink injection
- `MediaManager` — image embedding with SHA-1 deduplication
- `TemplateEngine` — `{{placeholder}}` replacement with fragmented run normalization
- `OutputWriter` — file, buffer, and stream output
- CLI: `build`, `validate`, `inspect`, `extract`, `debug` commands
- Full JSDoc documentation throughout codebase
- Unit tests for all core components (Vitest)
- Integration tests with fixture-based testing
- Performance benchmarks
- GitHub Actions: CI, release, docs workflows
- ESLint + Prettier configuration
- MIT License

### Architecture

- Zero PPTX generation library dependencies
- Only uses: `jszip`, `fast-xml-parser`, `fs-extra`, `commander`, `chalk`, `ora`
- Async/await throughout
- Private class fields (`#field`) for encapsulation
- Modular architecture following SOLID principles

[1.0.4]: https://github.com/jsuyog2/node-pptx-templater/compare/v1.0.3...v1.0.4
[1.0.3]: https://github.com/jsuyog2/node-pptx-templater/compare/v1.0.2...v1.0.3
[1.0.2]: https://github.com/jsuyog2/node-pptx-templater/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/jsuyog2/node-pptx-templater/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/jsuyog2/node-pptx-templater/releases/tag/v1.0.0
