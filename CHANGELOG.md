# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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
- `MediaManager` — image embedding with SHA-1 content deduplication
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
- Pure JavaScript ES Modules (no TypeScript)
- Zero PPTX generation library dependencies
- Only uses: `jszip`, `fast-xml-parser`, `fs-extra`, `commander`, `chalk`, `ora`
- Async/await throughout
- Private class fields (`#field`) for encapsulation
- Modular architecture following SOLID principles

[1.0.0]: https://github.com/your-org/pptx-templater/releases/tag/v1.0.0
