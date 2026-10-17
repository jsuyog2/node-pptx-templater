# Workflow: Adding a New Templating Feature

> **Purpose**: Checklist-driven Standard Operating Procedure for developing, verifying, and documenting new features in `node-pptx-templater`.

---

## 1. Step-by-Step SOP

### Step 1: OpenXML AST & Schema Design
- [ ] Identify affected parts (`ppt/slides/slideX.xml`, `ppt/charts/chartX.xml`, etc.).
- [ ] Confirm ECMA-376 Part 1 element ordering and namespace requirements.
- [ ] Determine if relationship files (`.rels`) or Content Types overrides (`[Content_Types].xml`) need updates.

### Step 2: Subsystem Implementation
- [ ] Implement domain logic in target manager (`src/elements/` or `src/managers/`).
- [ ] Use `createLogger('SubsystemName')` for all telemetry.
- [ ] Throw specific subclasses of `PPTXError`.

### Step 3: Core Orchestrator Integration (`PPTXTemplater.js`)
- [ ] Expose public method on `PPTXTemplater` (`src/core/PPTXTemplater.js`).
- [ ] Return `this` (or `Promise<PPTXTemplater>`) for fluent chaining.
- [ ] Write complete JSDoc block (`@param`, `@returns`, `@example`).
- [ ] Synchronize TypeScript declarations in `src/index.d.ts`.

### Step 4: Test Suite Creation
- [ ] Author integration tests in `tests/integration/` testing happy paths, boundary inputs, and errors.
- [ ] Audit output presentation buffers with `validatePackageIntegrity()`.

### Step 5: Documentation & Schema Sync
- [ ] Add method schema entry to `docs-site/src/data/apiReference.json`.
- [ ] Update `docs-site/src/data/features.json` if introducing a new capability.
- [ ] Add runnable demonstration to `showcase/examples.js`.

---

## 2. Verification Commands

Run the full verification loop:
```bash
npm run lint && npm run format:check
npm run docs:validate
npx vitest run tests/unit/PackageSetup.test.js
npm test
npm run docs:build
```
