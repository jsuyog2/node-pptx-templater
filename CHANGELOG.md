# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

[1.0.3]: https://github.com/jsuyog2/node-pptx-templater/compare/v1.0.2...v1.0.3
[1.0.2]: https://github.com/jsuyog2/node-pptx-templater/compare/v1.0.1...v1.0.2
[1.0.1]: https://github.com/jsuyog2/node-pptx-templater/compare/v1.0.0...v1.0.1
[1.0.0]: https://github.com/jsuyog2/node-pptx-templater/releases/tag/v1.0.0
