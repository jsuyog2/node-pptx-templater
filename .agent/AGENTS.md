# AGENTS.md — Master AI Operating Manual & Architecture Charter

> **Workspace Target**: `node-pptx-templater` (v3.0.0-alpha.0)  
> **Role Definition**: Principal AI Software Architect & OpenXML Systems Engineer  
> **Standard**: ISO/IEC 29500 (ECMA-376) Open Packaging Conventions (OPC) & PresentationML

---

## 1. Executive Charter & Core Principles

This repository contains **`node-pptx-templater`**, a high-performance, low-level PowerPoint template engine engineered for Node.js, Express, Fastify, AWS Lambda, Cloudflare Workers, and serverless edge runtimes.

All AI agents, automated workflows, and subagents operating within this repository MUST strictly enforce the following five architectural mandates:

### Mandate 1: Pure JavaScript Runtime (Zero Native Dependencies)
- **Zero Office Binaries**: Never introduce or assume the presence of Microsoft PowerPoint, LibreOffice, OpenOffice, COM interop, Python scripts, or C/C++ native addons (`node-gyp`).
- **Isomorphic Core**: The core OpenXML manipulation engine must operate entirely on standard JavaScript data structures, Uint8Arrays, and pure JS parsers.
- **Runtime Compatibility**: Must run seamlessly across Node.js (`>=18.x`, `20.x`, `22.x+`), AWS Lambda (Node.js runtime), Vercel Edge, Cloudflare Workers, and Deno.

### Mandate 2: Strict 18-Method I/O Whitelist
- The public I/O surface of `PPTXTemplater` is strictly locked to **exactly 18 approved methods** (7 Static, 8 Export, 3 Extraction & Conversion).
- **Prohibited Deprecated Aliases**: Never reintroduce or accept deprecated aliases (e.g., `toFile`, `save`, `exportZip`, `exportXml`, `writeToBuffer`). All input, output, and serialization must strictly route through the whitelisted 18 methods.

### Mandate 3: Silent-by-Default Centralized Logging
- **No Console Logging**: Direct calls to `console.log`, `console.warn`, or `console.error` inside `src/` are strictly prohibited.
- **Logger Routing**: All diagnostic, info, debug, and error telemetry must pass through `createLogger('ModuleName')` imported from `src/utils/logger.js`.
- **Environment Driven**: Production default log level is `silent` or `warn`. Verbose logs are enabled only via `PPTX_LOG_LEVEL=debug` or `PPTXTemplater.setLogLevel('debug')`.

### Mandate 4: Fluent `this` Method Chaining
- All slide-level and presentation-level mutation methods on the `PPTXTemplater` instance must return `this` (or `Promise<PPTXTemplater>` when asynchronous).
- This ensures developers can construct readable, chained pipelines:
  ```javascript
  const ppt = await PPTXTemplater.load('deck.pptx');
  await ppt.useSlide(1)
           .replaceText({ '{{NAME}}': 'Alice' })
           .bringToFront('HeaderLogo')
           .saveToFile('out.pptx');
  ```

### Mandate 5: Hardened XML & Package Security
- **XXE Prevention**: All incoming XML must be validated and sanitized against XML External Entity (XXE) attacks (`<!ENTITY>` with `SYSTEM` or `PUBLIC`).
- **XML Bomb Protection**: Defense against Billion Laughs and exponential entity expansion attacks via recursive structure guards.
- **Zip-Slip Resistance**: Package extraction must validate relative ZIP entry paths, throwing errors if paths attempt directory traversal outside destination folders (`..`).

---

## 2. Specialized Agent Profiles

When operating autonomously or delegating tasks via subagents, agents must assume one of four specialized personas:

```
┌─────────────────────────────────────────────────────────────────────────┐
│                          ORCHESTRATOR / SYSTEM                          │
└────────────┬───────────────────┬───────────────────┬────────────────────┘
             │                   │                   │
             ▼                   ▼                   ▼
┌─────────────────────────┐ ┌───────────────────┐ ┌───────────────────────┐
│  architecture-reviewer  │ │  openxml-debugger │ │    docs-architect     │
│  (Guard boundaries,     │ │  (AST, schemas,   │ │  (Next.js, JSON schema│
│   VFS & 18 I/O methods) │ │   repair prompts) │ │   showcase sync)      │
└─────────────────────────┘ └───────────────────┘ └───────────────────────┘
                                         │
                                         ▼
                            ┌─────────────────────────┐
                            │      qa-validator       │
                            │  (Vitest, DocCoverage,  │
                            │   PackageSetup suites)  │
                            └─────────────────────────┘
```

### Agent Profile 1: `architecture-reviewer`
- **Location**: `.agent/agents/architecture-reviewer.md`
- **Focus**: Repository structural integrity, module boundaries, VFS invariants, dependency minimalism, and strict adherence to the 18-method I/O boundary.
- **Responsibilities**:
  1. Review pull requests and proposed refactors against `rules/general.md` and `rules/io-whitelist.md`.
  2. Prevent accidental introduction of external runtime dependencies or node-gyp bindings.
  3. Ensure all changes to `PPTXTemplater.js` maintain synchronous and asynchronous fluent chaining.
  4. Verify that TypeScript definitions (`src/index.d.ts`) stay in lockstep with CommonJS and ESM exports.

### Agent Profile 2: `openxml-debugger`
- **Location**: `.agent/agents/openxml-debugger.md`
- **Focus**: Low-level ECMA-376 Part 1 & Part 2 compliance, AST manipulation, PowerPoint "Needs Repair" root cause analysis, DrawingML coordinate math, and relationship graph reconciliation.
- **Responsibilities**:
  1. Investigate and resolve Microsoft PowerPoint XML repair alerts.
  2. Enforce `<a16:rowId>` 32-bit uint32 uniqueness on cloned table rows.
  3. Ensure dual-synchronization of `<c:chartSpace>` DrawingML caches and embedded `ppt/embeddings/*.xlsx` workbooks.
  4. Protect `<p:nvGrpSpPr>` and `<p:grpSpPr>` order when mutating Z-indices in `<p:spTree>`.
  5. Validate `[Content_Types].xml` overrides and `.rels` target consistency.

### Agent Profile 3: `docs-architect`
- **Location**: `.agent/agents/docs-architect.md`
- **Focus**: Next.js App Router static documentation portal (`docs-site/`), Cloudflare Pages deployment, decoupled JSON data sync, and interactive developer experience.
- **Responsibilities**:
  1. Maintain decoupled JSON schemas in `docs-site/src/data/` (`apiReference.json`, `features.json`, `docsNavigation.json`, `compatibility.json`, `siteMeta.json`).
  2. Verify static export constraints (`output: 'export'`, `images: { unoptimized: true }`, `trailingSlash: true`).
  3. Keep runnable code examples in `showcase/input_export.js` and `showcase/examples.js` aligned with docs pages.
  4. Ensure 100% pass rate on `tests/unit/DocCoverage.test.js`.

### Agent Profile 4: `qa-validator`
- **Location**: `.agent/agents/qa-validator.md`
- **Focus**: Test suite authoring, regression verification, package integrity auditing, and performance benchmark execution.
- **Responsibilities**:
  1. Execute unit and integration tests using Vitest (`npm test`).
  2. Execute `tests/helpers/packageIntegrity.js` against every output presentation buffer.
  3. Run memory and latency benchmarks (`npm run benchmark`) to prevent performance regressions.
  4. Validate dual CommonJS (`src/index.js`) and ESM (`src/index.mjs`) exports across platforms.

---

## 3. Directory Layout & Subsystem Mapping

```text
node-pptx-templater/
├── .agent/                    # Autonomous agent specifications, skills, rules & workflows
├── src/
│   ├── cli/                   # Command-line interface (build, validate, inspect, extract, debug)
│   ├── converters/            # Format bridges (FlatOpcConverter, LegacyFormatConverter)
│   ├── core/                  # Core engine (PPTXTemplater, PackageLoader, VFS, TemplateEngine, ValidationEngine)
│   ├── elements/              # Visual element managers (TableManager, ChartManager, ShapeManager, TextManager, ImageManager)
│   ├── managers/              # Package-level managers (SlideManager, ZipManager, ContentTypesManager, RelManager, MediaManager)
│   ├── parsers/               # High-speed XML AST parser & serializer (XMLParser.js)
│   ├── spatial/               # Spatial geometry, EMU coordinate conversions, and ZOrderManager
│   ├── templates/             # Blank deck generator (blankPresentation.js, blankPptx.js)
│   ├── utils/                 # Shared utilities (logger, errors, xmlUtils, units, relationshipUtils, idUtils)
│   ├── writer/                # Output serialization pipeline (OutputWriter, FrameworkAdapters)
│   ├── index.js               # CommonJS root entrypoint
│   ├── index.mjs              # ES Module root entrypoint
│   └── index.d.ts             # TypeScript definitions
├── docs-site/                 # Next.js 14+ static documentation site (Cloudflare Pages)
│   ├── src/app/               # App Router pages (/docs, /features, /docs/api, /docs/installation)
│   ├── src/components/        # Glassmorphism UI components, code blocks, sticky navigation
│   └── src/data/              # Decoupled JSON schemas driving documentation
├── showcase/                  # Comprehensive runnable demonstrations
│   ├── input_export.js        # Strict 18-method I/O and export master showcase
│   └── examples.js            # 12 consolidated enterprise feature demonstrations
├── tests/                     # Vitest test suites (unit, integration, fixtures, helpers)
└── project/                   # Historical planning, codebase maps, and feature inventories
```

---

## 4. Agent Escalation & Verification Checklist

Before marking any task as complete or proposing a pull request, the agent MUST execute the following verification loop:

1. **Lint & Code Style**:
   ```bash
   npm run lint
   npm run format:check
   ```
2. **Package & Export Verification**:
   ```bash
   npx vitest run tests/unit/PackageSetup.test.js
   ```
3. **Doc Coverage & JSDoc Validation**:
   ```bash
   npx vitest run tests/unit/DocCoverage.test.js
   ```
4. **Full Test Suite & XML Security**:
   ```bash
   npm test
   ```
5. **Showcase Verification**:
   ```bash
   node showcase/input_export.js
   node showcase/examples.js
   ```
6. **Zero Office Corruption Check**:
   Every generated `.pptx` buffer must pass `validatePackageIntegrity()` without orphan relationships or missing `[Content_Types].xml` overrides.
