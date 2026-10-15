# Rule: General Architecture & Security Invariants

> **Scope**: Repository-wide standard operating rules for `node-pptx-templater`.  
> **Applies to**: Core engine, sub-managers, CLI, tests, and documentation.

---

## 1. Pure JavaScript Runtime Mandate

1. **Zero Native Binaries**:
   - The engine MUST NOT depend on native C/C++ addons, `node-gyp`, compiled bindings, Python scripts, or external Office suites (Microsoft Office, LibreOffice, unoconv).
   - Any external dependency added to `package.json` must be pure JavaScript.
2. **Runtime Portability**:
   - The code must be compatible with Node.js LTS versions (`>=18.0.0`, `20.x`, `22.x+`), AWS Lambda, Cloudflare Workers, Vercel Edge Functions, and Deno.
   - Browser polyfill awareness: Core modules must avoid direct reliance on Node-specific globals (`fs`, `path`) where browser/edge execution is expected. When Node-specific modules are needed for CLI or file operations, wrap them or place them in appropriate writer/loader boundaries.
3. **Dual CommonJS and ESM Packaging**:
   - `src/index.js` provides the CommonJS entry point.
   - `src/index.mjs` provides the ES Module entry point.
   - Both must export identical interfaces.
   - Build step (`npm run build`) uses `tsup` targeting both formats with `src/index.d.ts` types.

---

## 2. Centralized Logging & Telemetry Rule

1. **No Direct `console.*` Calls**:
   - Never write `console.log()`, `console.warn()`, or `console.error()` inside `src/` (except within CLI command display handlers in `src/cli/`).
2. **Logger Factory Usage**:
   - Always instantiate module loggers via `createLogger`:
     ```javascript
     const { createLogger } = require('../utils/logger.js');
     const logger = createLogger('SubsystemName');
     ```
3. **Logging Levels**:
   - Supported log levels: `'verbose'`, `'debug'`, `'info'`, `'warn'`, `'error'`, `'silent'`.
   - Production default is `'warn'` or `'silent'`.
   - Never log sensitive binary payloads or huge XML strings at `'info'` level. Use `'debug'` or `'verbose'`.
4. **Runtime Control**:
   - Global level can be modified at runtime via `PPTXTemplater.setLogLevel('debug')` or environment variable `PPTX_LOG_LEVEL=silent`.

---

## 3. Documentation Coverage & Type Invariants

1. **Mandatory JSDoc for Public Methods**:
   - Every public method in `PPTXTemplater.js` and exported utility classes MUST have a comprehensive JSDoc block.
   - Missing JSDoc will immediately fail the automated test:
     ```bash
     npm run docs:validate  # Runs tests/unit/DocCoverage.test.js
     ```
   - Each JSDoc block must include:
     - `@param` with explicit types and descriptions.
     - `@returns` with return type (`PPTXTemplater` for fluent chaining).
     - `@example` with realistic usage snippet.
     - Description of OpenXML impact.
2. **TypeScript Declarations Sync**:
   - When modifying method signatures, adding methods, or deprecating parameters in `src/core/PPTXTemplater.js`, the corresponding signature in `src/index.d.ts` MUST be updated simultaneously.
   - Types must enforce strict null-checks and provide accurate autocomplete types.

---

## 4. XML Security & Parsing Guards

1. **XXE (XML External Entity) Defense**:
   - All untrusted XML documents ingested via `load()`, `FlatOpcConverter`, or `XMLParser` must be validated against `validateXml()` or `safeParseXml()`.
   - Any XML containing `<!ENTITY>` declarations referencing external URIs (`SYSTEM` or `PUBLIC`) must be rejected with a `PPTXError`.
2. **XML Bomb & Billion Laughs Mitigation**:
   - Disallow nested or quadratic entity expansions in incoming template XML.
   - Reject arbitrary custom DTD / DOCTYPE blocks in presentation parts.
3. **Zip-Slip Path Traversal Protection**:
   - When unpacking or extracting ZIP entries (`extractPptx()`, `ZipManager`, `extractToXml()`), resolve the destination directory and ensure that each resolved file path starts with the destination root:
     ```javascript
     const resolvedPath = path.resolve(destDir, entryName);
     if (!resolvedPath.startsWith(path.resolve(destDir) + path.sep)) {
       throw new PPTXError(`Zip-Slip security violation: illegal entry "${entryName}"`);
     }
     ```

---

## 5. Error Handling & Custom Exceptions

1. **Use Domain Errors**:
   - All errors thrown by the engine must inherit from `PPTXError` (from `src/utils/errors.js`).
   - Use specific subclasses where appropriate:
     - `SlideNotFoundError`: When a slide index or ID does not exist.
     - `ChartNotFoundError`: When a requested chart name or ID is absent.
     - `TableNotFoundError`: When a requested table name or ID is absent.
2. **Original Error Preservation**:
   - Always pass the underlying cause error to the `PPTXError` constructor:
     ```javascript
     throw new PPTXError(`Failed to parse slide: ${err.message}`, err);
     ```
