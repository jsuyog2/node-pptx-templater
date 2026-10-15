# Workflow: Adding a New Templating Feature

> **Purpose**: Step-by-step Standard Operating Procedure (SOP) for developing, testing, and documenting a new feature in `node-pptx-templater`.

---

## 1. Overview & Lifecycle Stages

Every new capability added to `node-pptx-templater` must transition through 6 mandatory stages:

```
┌──────────────────┐    ┌──────────────────┐    ┌──────────────────┐
│ Stage 1: Design  │───►│ Stage 2: Sub-Mgr │───►│ Stage 3: Engine  │
│ (ECMA-376 Schema)│    │ Implementation   │    │ Fluent Wrapper   │
└──────────────────┘    └──────────────────┘    └──────────────────┘
                                                          │
                                                          ▼
┌──────────────────┐    ┌──────────────────┐    ┌──────────────────┐
│ Stage 6: Verify  │◄───│ Stage 5: Docs    │◄───│ Stage 4: Testing │
│ (All Tests Pass) │    │ & Schema Sync    │    │ (Vitest + Integ) │
└──────────────────┘    └──────────────────┘    └──────────────────┘
```

---

## 2. Stage-by-Stage Instructions

### Stage 1: OpenXML Schema & AST Design
1. Identify the target OpenXML schema parts impacted (e.g., `ppt/slides/slideX.xml`, `ppt/charts/chartX.xml`, `ppt/tableStyles.xml`).
2. Verify element ordering and namespace rules according to ECMA-376 Part 1.
3. Determine if any relationship files (`.rels`) or Content Types overrides (`[Content_Types].xml`) are required.

### Stage 2: Subsystem Manager Implementation
1. Add the domain logic to the appropriate sub-manager:
   - Tables: `src/elements/TableManager.js`
   - Charts: `src/elements/ChartManager.js`
   - Shapes: `src/elements/ShapeManager.js`
   - Text: `src/elements/TextManager.js`
   - Images & Media: `src/elements/ImageManager.js` or `src/managers/MediaManager.js`
   - Slides: `src/managers/SlideManager.js`
   - Z-Order: `src/spatial/ZOrderManager.js`
2. Ensure all logging uses `createLogger('SubsystemName')`.
3. Throw specific domain errors inheriting from `PPTXError`.

### Stage 3: Orchestrator Integration (`PPTXTemplater.js`)
1. Expose a public method on `PPTXTemplater` in `src/core/PPTXTemplater.js`.
2. Ensure the method operates on `#selectedSlides` when slide context is set (`useSlide()`).
3. Return `this` (or `Promise<PPTXTemplater>` when async) to preserve fluent method chaining.
4. Write a comprehensive JSDoc block:
   - Clear description.
   - `@param` with accurate types.
   - `@returns {PPTXTemplater}` or `@returns {Promise<PPTXTemplater>}`.
   - `@example` code snippet.
5. Update `src/index.d.ts` with the new method signature and accurate TypeScript types.

### Stage 4: Test Suite Authoring
1. Create a dedicated integration test file in `tests/integration/` (e.g., `tests/integration/MyNewFeature.test.js`).
2. Test happy path, boundary values, and error conditions.
3. Validate the generated presentation buffer with `validatePackageIntegrity()`:
   ```javascript
   import { validatePackageIntegrity } from '../helpers/packageIntegrity.js';
   const buffer = await ppt.toBuffer();
   const errors = await validatePackageIntegrity(buffer);
   expect(errors).toHaveLength(0);
   ```

### Stage 5: Documentation & Schema Sync
1. Add the new method entry to `docs-site/src/data/apiReference.json` under the appropriate category.
2. If adding a new capability concept, update `docs-site/src/data/features.json`.
3. Add a demonstration of the feature in `showcase/examples.js`.

### Stage 6: Final Verification Checklist
Run the complete automated verification suite:
```bash
# 1. Lint and code format check
npm run lint
npm run format:check

# 2. Verify documentation coverage (must be 100%)
npm run docs:validate

# 3. Verify package setup and exports
npx vitest run tests/unit/PackageSetup.test.js

# 4. Run all unit and integration tests
npm test

# 5. Build docs site
npm run docs:build
```
