# Rule: General Architecture & Security Invariants

> **Scope**: Repository-wide standard operating rules for `node-pptx-templater` (v3.0.0-beta.0).

---

## 1. Pure JavaScript Runtime Mandate
- **Zero Native Binaries**: No C/C++ addons, `node-gyp`, Python, LibreOffice, OpenOffice, or COM automation. Pure JS dependencies only.
- **Isomorphic Portability**: Run seamlessly across Node.js (`>=18.x`, `20.x`, `22.x+`), AWS Lambda, Cloudflare Workers, Vercel Edge, Deno, and Bun.
- **Dual Packaging**: CommonJS (`src/index.js`) and ESM (`src/index.mjs`) must maintain identical exported interfaces and TypeScript definitions (`src/index.d.ts`).

## 2. Centralized Silent Logging
- **No Console Logging**: Direct `console.log/warn/error` in `src/` is strictly prohibited.
- **Logger Factory**: All telemetry must route through `createLogger('ModuleName')` from `src/utils/logger.js`.
- **Runtime Defaults**: Production default level is `'warn'` or `'silent'`. Verbose logs enabled via `PPTX_LOG_LEVEL=debug` or `PPTXTemplater.setLogLevel('debug')`.

## 3. Fluent Method Chaining
- All presentation and slide mutation methods on `PPTXTemplater` must return `this` (or `Promise<PPTXTemplater>` when async).

## 4. Hardened XML & Package Security
- **XXE Prevention**: Ingested XML is validated against `safeParseXml` / `validateXml`. Documents with `<!DOCTYPE>` or external `SYSTEM`/`PUBLIC` entities throw `PPTXError`.
- **Entity Bomb Protection**: Rejects nested/quadratic entity expansion and recursive custom DTD structures.
- **Zip-Slip Resistance**: Extraction routines (`ZipManager`, `extractPptx`, `extractToXml`) resolve target paths against destination roots; path traversal (`../`) throws security violations.

## 5. Strict 8 Standardized Categories
- All documentation schemas (`docs-site/src/data/`), TypeScript interfaces (`src/index.d.ts`), and JSDoc tags must strictly use these 8 canonical category names:
  1. `Import/Export` (18 methods)
  2. `Slides` (24 methods)
  3. `Text & Placeholders` (10 methods)
  4. `Tables` (30 methods)
  5. `Charts` (11 methods)
  6. `Shapes` (11 methods)
  7. `Images & Media` (11 methods)
  8. `Elements` (15 methods)
- Total whitelisted public methods: 130 methods. Standalone "Diagnostics & Validation" category is removed.

## 6. Documentation & Type Invariants
- **Mandatory JSDoc**: All public methods in `PPTXTemplater.js` require full JSDoc (`@param`, `@returns`, `@example`). Enforced via `npm run docs:validate`.
- **TypeScript Sync**: Signatures in `src/index.d.ts` must match runtime signatures in `src/core/PPTXTemplater.js`.
- **Domain Errors**: All exceptions must inherit from `PPTXError` (`src/utils/errors.js`) and preserve root causes.
