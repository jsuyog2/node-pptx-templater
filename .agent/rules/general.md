# Rule: General Architecture & Security Invariants

> **Scope**: Repository-wide standard operating rules for `node-pptx-templater` (v3.0.0-beta.1).

---

## 1. Pure JavaScript Runtime Mandate
- **Zero Native Binaries**: No C/C++ addons, `node-gyp`, LibreOffice, OpenOffice, or COM interop.
- **Portability**: Unconditionally isomorphic across Node.js (`>=18.x`, `20.x`, `22.x+`), AWS Lambda, Cloudflare Workers, Vercel Edge, Deno, and Bun.
- **Dual Packaging**: Identical exports and types across CommonJS (`src/index.js`) and ESM (`src/index.mjs`).

## 2. Centralized Silent Logging
- **No Console Logging**: Direct `console.log/warn/error` in `src/` is strictly prohibited.
- **Logger Factory**: All logging must route through `createLogger('ModuleName')` from `src/utils/logger.js`.
- **Production Default**: Default level is `'warn'` or `'silent'`. Debug enabled via `PPTX_LOG_LEVEL=debug` or `PPTXTemplater.setLogLevel('debug')`.

## 3. Fluent Method Chaining
- All presentation and slide mutation methods on `PPTXTemplater` must return `this` (or `Promise<PPTXTemplater>` when async).

## 4. Hardened XML & Package Security
- **XXE Prevention**: Ingested XML is validated via `safeParseXml`. Documents containing `<!DOCTYPE>` or external `SYSTEM`/`PUBLIC` entities throw `PPTXError`.
- **Entity Bomb Protection**: Rejects quadratic entity expansion and recursive custom DTDs.
- **Zip-Slip Resistance**: Extraction routines verify target paths against destination roots; path traversal (`../`) throws security violations.

## 5. Strict 8 Standardized Categories
- All documentation schemas, TypeScript declarations (`index.d.ts`), and JSDoc tags must strictly use these 8 canonical categories: `Import/Export`, `Slides`, `Text & Placeholders`, `Tables`, `Charts`, `Shapes`, `Images & Media`, `Elements`.
- Total whitelisted public methods: 130 methods. Standalone "Diagnostics & Validation" category is removed.

## 6. Documentation & Type Invariants
- **Mandatory JSDoc**: All public methods require complete JSDoc (`@param`, `@returns`, `@example`). Enforced via `npm run docs:validate`.
- **TypeScript Sync**: Type declarations in `src/index.d.ts` must stay in exact 1:1 parity with runtime signatures in `src/core/PPTXTemplater.js`.
