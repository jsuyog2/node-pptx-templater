# AGENTS.md — Master AI Operating Manual & Architecture Charter

> **Workspace Target**: `node-pptx-templater` (v3.0.0-alpha.0)  
> **Standard**: ISO/IEC 29500 (ECMA-376) OPC & PresentationML

---

## 1. Core Principles & Architectural Mandates

1. **Pure JavaScript Runtime**: Zero native binaries, C/C++ addons, `node-gyp`, LibreOffice, or COM interop. Fully portable across Node.js (`>=18.x`, `20.x`, `22.x+`), AWS Lambda, Cloudflare Workers, Vercel Edge, and Deno.
2. **Strict 18-Method I/O Whitelist**: Public I/O is locked to exactly 18 methods (7 Static, 8 Export, 3 Extraction). All deprecated aliases (`toFile`, `save`, `exportZip`, `exportXml`, `writeToBuffer`) are strictly prohibited.
3. **Silent Centralized Logging**: Direct `console.log/warn/error` in `src/` is prohibited. All telemetry routes through `createLogger('ModuleName')` from `src/utils/logger.js`.
4. **Fluent `this` Chaining**: All presentation and slide mutation methods must return `this` (or `Promise<PPTXTemplater>` when async).
5. **Hardened XML & Package Security**: Enforce XXE blocking (`safeParseXml`), Billion Laughs / entity bomb protection, and Zip-Slip traversal resistance.

---

## 2. Specialized Agent Roster

Detailed responsibilities and verification checklists are maintained in `.agent/agents/`:

| Agent Profile | Specification | Primary Focus & Domain |
| :--- | :--- | :--- |
| **`architecture-reviewer`** | `.agent/agents/architecture-reviewer.md` | Pure JS boundary, 18 I/O methods, VFS integrity, `index.d.ts` sync, fluent chaining. |
| **`openxml-debugger`** | `.agent/agents/openxml-debugger.md` | ISO/IEC 29500 schema forensics, unique `<a16:rowId>`, chart dual-sync, `<p:spTree>` order. |
| **`docs-architect`** | `.agent/agents/docs-architect.md` | Next.js static docs (`docs-site/`), decoupled JSON data sync, 100% DocCoverage. |
| **`qa-validator`** | `.agent/agents/qa-validator.md` | Vitest suites (`npm test`), package integrity audits, benchmarks, CJS/ESM exports. |

---

## 3. Directory Layout

```text
node-pptx-templater/
├── .agent/                    # Autonomous agent specifications, skills, rules & workflows
├── src/
│   ├── cli/                   # CLI (build, validate, inspect, extract, debug)
│   ├── converters/            # Format bridges (FlatOpcConverter, LegacyConverter)
│   ├── core/                  # Core engine (PPTXTemplater, PackageLoader, VFS, TemplateEngine, ValidationEngine)
│   ├── elements/              # Visual managers (TableManager, ChartManager, ShapeManager, TextManager, ImageManager)
│   ├── managers/              # Package managers (SlideManager, ZipManager, ContentTypesManager, RelationshipManager, MediaManager)
│   ├── parsers/               # Lossless XML AST parser & serializer (XMLParser.js)
│   ├── spatial/               # EMU math, conversions, and ZOrderManager
│   ├── templates/             # Blank presentation blueprints
│   ├── utils/                 # Utilities (logger, errors, xmlUtils, units, idUtils)
│   ├── writer/                # Serialization & framework adapters (Express, Fastify)
│   ├── index.js / index.mjs   # Dual CommonJS and ESM entry points
│   └── index.d.ts             # TypeScript definitions
├── docs-site/                 # Next.js static documentation portal (Cloudflare Pages)
│   ├── src/app/               # App Router pages (/docs, /features, /docs/api)
│   └── src/data/              # Decoupled JSON schemas driving documentation
├── showcase/                  # Runnable master demonstrations (input_export.js, examples.js)
└── tests/                     # Vitest test suites (unit, integration, helpers)
```

---

## 4. Verification Loop

Before concluding tasks or submitting changes, execute this verification sequence:
```bash
npm run lint && npm run format:check
npx vitest run tests/unit/PackageSetup.test.js
npm run docs:validate
npm test
node showcase/input_export.js
node showcase/examples.js
```
