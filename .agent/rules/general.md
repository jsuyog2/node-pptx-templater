# General Engineering Rules

These rules govern the architecture, design principles, runtime constraints, and coding standards for `node-pptx-templater`.

---

## 1. Core Engineering Principles

### 1.1 Zero Office / Electron Runtime Dependencies
- The library **must execute purely in JavaScript** across Node.js (>= 20.12.0) and modern serverless/edge environments.
- **NEVER** introduce dependencies requiring Microsoft Office, LibreOffice, unoconv, Puppeteer, Electron, or native C/C++ compiled binaries (e.g., `node-gyp`).
- All manipulation of presentation containers must operate directly on Open Packaging Conventions (OPC) zip archives and OpenXML / DrawingML markup.

### 1.2 In-Memory Virtual File System (VFS)
- Presentations must be ingested into and manipulated within an in-memory Virtual File System (`Map<string, string | Uint8Array>`).
- Disk I/O must only be performed when explicitly requested by file-path APIs (`saveToFile()`, `PPTXTemplater.load()`, `extractToXml()`, `extractToZip()`, `toFlatXml()`).
- File paths within the package must use forward slashes (`/`), be normalized (no leading `./`), and match OPC conventions (`ppt/slides/slide1.xml`, `ppt/_rels/presentation.xml.rels`, `[Content_Types].xml`).

### 1.3 Dual CJS and ESM Compatibility
- The library ships dual CommonJS (`dist/index.cjs`) and ECMAScript Module (`dist/index.mjs`) builds with full TypeScript declarations (`dist/index.d.ts`).
- Source code in `src/` must remain compatible with modern Node.js and the `tsup` bundling pipeline.
- Package exports in `package.json` must maintain strict export maps supporting `import`, `require`, `types`, and default conditions.

### 1.4 Immutability & Safe Mutation Patterns
- Templating operations should clone template assets when repeating or duplicating elements to avoid side-effects on the original presentation structure.
- When mutating XML AST nodes, deep copies of sub-trees (such as rows or slide elements) must be created before altering identifiers, relationship targets, or content runs.
- Original template layouts, slide masters, themes, and slide relationships must remain intact unless explicitly targeted for alteration.

---

## 2. Security Standards

### 2.1 Strict XML Security (XXE & Billion Laughs)
- All XML parsing through `fast-xml-parser` or `xmlUtils` must strictly disable:
  - DOCTYPE declarations (`allowBooleanAttributes: true`, but never execute entities).
  - External entity resolution (prevent XML External Entity / XXE attacks).
  - DTD entity expansion (prevent Billion Laughs / quadratic blowup attacks).
- Inputs containing suspicious entity definitions must be rejected before AST processing.

### 2.2 Path Traversal Prevention
- Any ZIP entry path extracted or written must be checked against directory traversal (`../` or absolute root prefixes).
- Relative targets in `.rels` files must be resolved strictly within the package virtual root.

---

## 3. Error Handling and Logging

### 3.1 Silent Default Logging
- The library must not pollute `console.log` or `console.error` during standard operation.
- Diagnostic logs and telemetry must be gated behind optional loggers or debug flags provided in the presentation options.

### 3.2 Actionable Error Messages
- Any OpenXML parsing failure, missing element, or relationship mismatch must throw a descriptive Error with actionable context:
  - Exact slide index or part name (`ppt/slides/slide2.xml`).
  - Target placeholder, shape name, or table coordinate.
  - Explanation of the schema violation or missing reference.

---

## 4. Code Quality & Verification Gates

### 4.1 Linting and Formatting
- All code must pass `npm run lint` (ESLint 9 flat config) and formatting checks with zero errors or warnings.
- Avoid loose `any` types in TypeScript declaration files.

### 4.2 Comprehensive Test Coverage
- Unit tests run using `vitest` (`npm test`).
- Every new feature, bugfix, or XML manipulator must include automated tests verifying:
  1. Correct AST generation.
  2. Relationship and content type registration.
  3. Presentation packaging integrity (can be opened and serialized without corruption).
- The test suite must remain 100% green at all times.
