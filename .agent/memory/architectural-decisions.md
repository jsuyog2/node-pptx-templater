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
