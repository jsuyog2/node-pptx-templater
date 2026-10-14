# Architectural Decision Records (ADR)

This file documents core architectural decisions, their historical context, alternatives evaluated, and consequences for `node-pptx-templater`.

---

## ADR-001: Pure JavaScript In-Memory Virtual File System (VFS)

- **Status**: Accepted
- **Context**: Programs manipulating `.pptx` often unpack zip archives to temporary disk folders, run operations, and repackage them. In cloud and serverless environments (AWS Lambda, Google Cloud Functions, Cloudflare Workers), disk access is slow, constrained, or read-only.
- **Decision**: Ingest presentations directly into an in-memory `Map<string, string | Uint8Array>`. Manipulate XML and media assets entirely in RAM.
- **Consequences**: Zero temporary file leakage, fast execution (sub-10ms for typical decks), zero disk I/O bottlenecks, and seamless compatibility with serverless runtimes.

---

## ADR-002: Dual CJS and ESM Distribution via `tsup`

- **Status**: Accepted
- **Context**: The Node.js ecosystem is divided between legacy CommonJS (`require`) and modern ECMAScript Modules (`import`).
- **Decision**: Use `tsup` to emit dual builds: `dist/index.cjs` (CommonJS) and `dist/index.mjs` (ESM), alongside generated `.d.ts` TypeScript definitions. Enforce explicit `exports` maps in `package.json`.
- **Consequences**: Seamless compatibility for all consumers without needing transpilation steps or polyfills.

---

## ADR-003: Zero Office / LibreOffice Runtime Dependencies

- **Status**: Accepted
- **Context**: Many document manipulation libraries depend on headless Microsoft Office, LibreOffice (`soffice`), or native C++ binaries (`unoconv`, `node-gyp`). These require heavy Docker containers, fail on serverless platforms, and suffer from high startup latency and memory overhead.
- **Decision**: Mandate 100% pure JavaScript execution. All presentation, layout, and chart manipulation is performed strictly by parsing and serializing Open Packaging Conventions (OPC) and OpenXML / DrawingML markup.
- **Consequences**: Lightweight deployments (package size < 5 MB), fast cold starts, zero native compilation failures, and cross-platform consistency across Windows, Linux, and macOS.

---

## ADR-004: Dual-Sync DrawingML Charts & Embedded Excel Workbooks

- **Status**: Accepted
- **Context**: PowerPoint charts store visual rendering caches in `ppt/charts/chartX.xml` (`<c:chartSpace>`) and editable tabular data in `ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx`. If only the XML cache is updated, clicking "Edit Data in Excel" in PowerPoint either causes repair errors or overwrites the chart with stale numbers.
- **Decision**: Implement a dual-sync updater that patches both the DrawingML series definitions and the embedded Excel worksheet binary simultaneously.
- **Consequences**: Guaranteed visual consistency and zero PowerPoint repair prompts when users inspect or edit data.

---

## ADR-005: 32-Bit Unsigned `<a16:rowId>` Table Row Identification

- **Status**: Accepted
- **Context**: Starting with PowerPoint 2016, Office tracks table rows using `<a16:rowId val="...">`. Duplicating or inserting rows without unique 32-bit unsigned integer IDs triggers a "PowerPoint found a problem with content" repair dialog.
- **Decision**: Enforce automatic generation of unique 32-bit unsigned integers ($1 \le \text{rowId} \le 4,294,967,295$) whenever rows are cloned or inserted.
- **Consequences**: Perfect compatibility with modern PowerPoint 2016, 2019, 2021, and Office 365.

---

## ADR-006: AST-Level Run Normalization & Token Healing for Text Placeholders

- **Status**: Accepted
- **Context**: PowerPoint's rich text editor frequently breaks placeholder tokens (e.g. `{{user_name}}`) across multiple `<a:r>` runs due to grammar markers, spelling dictionaries, or formatting variations. Regex-based string replacement on raw XML fails because the tag text is split.
- **Decision**: Implement AST-level run healing within `TemplateEngine.js`. Detect split placeholders across adjacent `<a:r>` nodes within each `<a:p>`, coalesce the text into the leading run while preserving its `<a:rPr>` styling, and clear the trailing consumed runs.
- **Consequences**: 100% reliable placeholder replacement without requiring users to retype placeholders in plain text editors.

---

## ADR-007: Silent Default Logging with Pluggable Logger

- **Status**: Accepted
- **Context**: Command-line utilities, web servers, and automated pipelines require clean stdout/stderr streams. Hardcoded `console.log` calls clutter consumer output and degrade performance.
- **Decision**: Ensure library execution is completely silent by default. Provide an optional `logger` option in presentation configuration for diagnostic telemetry.
- **Consequences**: Production-friendly integration into HTTP servers, Lambdas, and microservices.

---

## ADR-008: Strict XML Security (XXE & Expansion Prevention)

- **Status**: Accepted
- **Context**: Presentations may be loaded from untrusted user uploads. Malicious XML can contain XML External Entity (XXE) attacks or Billion Laughs entity expansions.
- **Decision**: Configure XML parsers to strictly disable DTD execution, external entity resolution, and entity expansions. Reject suspicious inputs before AST processing.
- **Consequences**: Robust protection against data exfiltration, local file inclusion, and denial-of-service attacks.

---

## ADR-009: Legacy `/docs` Removal & Transition to Next.js Static Export (`docs-site`)

- **Status**: Accepted
- **Context**: The legacy documentation in `/docs` consisted of static HTML/JS/CSS files generated by outdated builder scripts (`build-docs.js`). It was difficult to maintain, lacked modern component reusability, and lacked dynamic routing.
- **Decision**: Completely purged the legacy `/docs` directory. Built a modern documentation portal in `docs-site/` using Next.js (App Router) configured with `output: 'export'`, `images: { unoptimized: true }`, and `trailingSlash: true` for zero-server deployment on Cloudflare Pages (`wrangler.toml`).
- **Consequences**: 100% static HTML generation into `docs-site/out/`, sub-second page loads, automated sitemap and robots generation, and seamless integration with modern CI workflows.

---

## ADR-010: Glassmorphism Design System, Light/Dark Modes & GSAP Micro-Interactions

- **Status**: Accepted
- **Context**: Modern technical documentation requires high visual polish, accessible contrast, responsive readability, and fluid micro-interactions across devices.
- **Decision**: Implemented a modern Glassmorphism design system in Tailwind CSS (`.glass-card`, `.glass-header`, `.glass-sidebar`, ambient background mesh gradients). Added persistent Light and Dark themes via `ThemeContext` (detecting system preference and persisting to `localStorage`). Integrated GSAP animation utilities (`src/animations/gsapTransitions.js`) for staggered hero typography reveals, feature card entrance animations, and smooth API card filtering.
- **Consequences**: High aesthetic quality, accessibility in both bright and dark environments, smooth interactive feedback, and fixed non-scrolling sidebar navigation.

---

## ADR-011: Strict Documentation Content Decoupling into Modular JSON Schemas

- **Status**: Accepted
- **Context**: Hardcoding API signatures, feature inventories, or navigation links directly in JSX components introduces merge conflicts, stale text, and code-documentation drift.
- **Decision**: Decoupled all documentation content into modular JSON schemas under `docs-site/src/data/`:
  - `features.json`: Single source of truth for all 86 capabilities, parsed directly from `project/Features.md` via `scripts/parse-features.js`.
  - `apiReference.json`: Complete catalog of 101 public methods across all 10 engine domains.
  - `docsNavigation.json`: Navigation items, categories, and sidebar hierarchy.
  - `compatibility.json`: Runtime matrix and serverless execution guarantees.
  - `siteMeta.json`: Global SEO metadata, keywords, and OpenGraph configs.
- **Consequences**: Content maintainers update text and data files without touching UI code, automated scripts guarantee schema integrity, and unit tests verify 100% documentation coverage.

---

## ADR-012: Retirement of Legacy I/O Methods & Adoption of Unified PackageLoader / OutputWriter Pipelines

- **Status**: Accepted
- **Context**: Older iterations contained fragmented I/O helpers (`save`, `toFile`, `saveXml`, `fromPresentationXml`) with inconsistent option handling, ad-hoc stream logic, and limited format/runtime support. Serverless and Edge runtimes demand unified polymorphic ingestion and web-standard responses.
- **Decision**:
  - Retired legacy I/O methods (`save`, `toFile`, `saveXml`, `fromPresentationXml`), adding deprecation warnings that forward to the modern pipeline.
  - Adopted unified `PackageLoader` with polymorphic ingestion: file paths, Buffers, Uint8Array, streams, unpacked directories, Web URLs (`PPTXTemplater.fromUrl`), cloud object stores (`PPTXTemplater.fromCloud`), and legacy formats (`.ppt`, `.odp`) via `LegacyFormatConverter`.
  - Implemented multi-target serialization in `OutputWriter` and `FrameworkAdapters`: `saveToFile`, `toBuffer`, `toStream`, `toArrayBuffer`, `toBase64`, `toExpress`, `toFastify`, and standard Web API `toResponse`.
  - Added template-free blank deck generation (`PPTXTemplater.createBlank(options)` in `src/templates/blankPresentation.js`) and direct XML/ZIP extraction (`extractToZip`, `extractToXml`).
- **Consequences**: Clean separation of concerns, zero legacy I/O fragmentation, full Edge/Next.js/Remix compatibility, and seamless backwards compatibility with clear deprecation guidance.

---

## ADR-013: Native Microsoft Flat OPC XML (`<pkg:package>`) Package Ingestion & Serialization

- **Status**: Accepted
- **Context**: Microsoft PowerPoint supports Flat OPC XML format (single uncompressed XML documents containing `<pkg:package>` and `<pkg:part>` elements, often saved as `.xml` or `sample.xml`). Users and automated workflows require processing these presentations directly without compressing them into ZIP archives, avoiding ZIP overhead and enabling pure XML-based processing.
- **Decision**:
  - Implemented `FlatOpcConverter` (`src/converters/FlatOpcConverter.js`) with signature detection (`isFlatOpc()`).
  - Integrated into `PackageLoader.load()` and added static `PPTXTemplater.fromFlatOpc(source, options)`:
    - Strips leading slashes on `pkg:name` attributes to map to standard OpenXML VFS paths.
    - Serializes child nodes of `<pkg:xmlData>` to clean XML strings directly into the in-memory VFS.
    - Decodes base64 strings from `<pkg:binaryData>` into binary `Uint8Array` / `Buffer` for embedded `.xlsx` workbooks (`ppt/embeddings/`) and image assets (`ppt/media/`).
    - Synthesizes `[Content_Types].xml` with defaults and part overrides so standard OpenXML managers operate transparently.
  - Added multi-target serialization in `OutputWriter` and `PPTXTemplater`: `toFlatXml(options)` and `toFlatOpc(options)` to compile VFS back into valid `<pkg:package>` documents.
- **Consequences**: Presentations can be loaded from and exported to Flat OPC XML without ZIP compression/decompression overhead, while supporting all fluent mutation methods (`useSlide`, `replaceText`, `updateChart`, `updateTable`, `saveToFile`).

---

## ADR-014: Strict 18-Method I/O API Whitelist and Purge of Extraneous/Deprecated Aliases

- **Status**: Accepted
- **Context**: Across iterations, redundant aliases and utility methods (`fromUrl`, `fromCloud`, `fromFlatOpc`, `importPresentation`, `cache`, `saveToFolder`, `toZipBuffer`, `toFlatOpc`, `save`, `toFile`, `saveXml`) accumulated. Having multiple entry points for the same underlying capabilities increased API surface complexity, bloated documentation, and made TypeScript interfaces harder to navigate.
- **Decision**:
  - Enforce a strict 18-method I/O API whitelist:
    - **Static API (7 methods)**: `PPTXTemplater.load(source, options)`, `PPTXTemplater.createBlank(options)`, `PPTXTemplater.extractPptx(pptxPath, outputDir, options)`, `PPTXTemplater.buildPptx(sourceDir, outputPath, options)`, `PPTXTemplater.preload(source, options)`, `PPTXTemplater.fromCache(source)`, `PPTXTemplater.clearCache()`.
    - **Export Instance Methods (8 methods)**: `ppt.saveToFile(destinationPath, options)`, `ppt.toBuffer(options)`, `ppt.toStream(options)`, `ppt.toBase64(options)`, `ppt.toArrayBuffer(options)`, `ppt.toResponse(options)`, `ppt.toExpress(res, filename, options)`, `ppt.toFastify(reply, filename, options)`.
    - **Extraction & Conversion Instance Methods (3 methods)**: `ppt.extractToZip(destinationPath, options)`, `ppt.extractToXml(destinationDir, options)`, `ppt.toFlatXml(destinationPathOrOptions, options)`.
  - Unify all polymorphic inputs into `PPTXTemplater.load()` (which auto-detects local files, remote URLs, in-memory buffers, readable streams, cloud configs, and Flat OPC XML).
  - Unify Flat OPC export into `ppt.toFlatXml(destinationPathOrOptions, options)`, writing directly to disk if passed a file path string or returning the XML string in memory.
  - Purge all deprecated aliases and non-whitelisted methods across `src/`, `types/`, test suites, and documentation.
- **Consequences**: Minimal, unambiguous developer API surface, clean documentation catalog (103 public methods), unified polymorphic ingestion mental model, and zero deprecated code baggage.


