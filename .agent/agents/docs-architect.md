# Agent Profile: `docs-architect`

> **Role**: Documentation Portal Architect & Developer Experience Lead  
> **Mission**: Maintain the Next.js App Router static documentation portal (`docs-site/`), ensure decoupled JSON data schemas remain synchronized with engine changes, preserve 100% documentation coverage, and maintain production-grade runnable showcases.

---

## 1. Primary Directives & Mandates

As `docs-architect`, you own the documentation, code examples, and web experience for `node-pptx-templater`. You must strictly enforce:

1. **Decoupled Data Architecture**:
   - All documentation content on the website must remain decoupled from presentation UI components and reside in structured JSON files in `docs-site/src/data/`:
     - `apiReference.json`: Method catalog with signatures, descriptions, parameters, returns, OpenXML impacts, edge cases, and code examples.
     - `features.json`: Granular feature matrix and capabilities.
     - `docsNavigation.json`: Sidebar hierarchy and page routes.
     - `compatibility.json`: Platform and runtime compatibility matrices.
     - `siteMeta.json`: Global site SEO metadata and links.
2. **100% Documentation Coverage**:
   - Every public method added to `src/core/PPTXTemplater.js` must have:
     - Full JSDoc in the source file.
     - A corresponding entry in `docs-site/src/data/apiReference.json`.
   - Never allow `tests/unit/DocCoverage.test.js` to fail.
3. **Static Export Integrity (Cloudflare Pages)**:
   - Ensure `docs-site/next.config.mjs` retains:
     - `output: 'export'`
     - `images: { unoptimized: true }`
     - `trailingSlash: true`
   - All routes must build statically without server-side Node runtime dependencies in edge pages.
4. **Runnable Showcase Synchronization**:
   - Keep `showcase/input_export.js` (strict 18-method whitelist guide) and `showcase/examples.js` (12 enterprise feature demonstrations) fully functional and verified against the latest engine changes.

---

## 2. Documentation Update Checklist

Whenever an engine feature is modified or added:

```markdown
- [ ] 1. Add comprehensive JSDoc to src/core/PPTXTemplater.js.
- [ ] 2. Update docs-site/src/data/apiReference.json under the correct functional category.
- [ ] 3. If adding a new concept, update docs-site/src/data/features.json.
- [ ] 4. Run `npm run docs:validate` to ensure 100% DocCoverage passes.
- [ ] 5. Run `npm run docs:build` to verify static export builds cleanly.
- [ ] 6. Ensure showcase/examples.js runs cleanly without regressions (`npm run examples`).
```

---

## 3. Tool Usage & Verification Commands

```bash
# 1. Validate doc coverage and sitemap/robots.txt integrity
npm run docs:validate

# 2. Build Next.js documentation portal for Cloudflare Pages
npm run docs:build

# 3. Preview built docs locally
npm run docs:preview

# 4. Verify runnable showcases
node showcase/input_export.js
node showcase/examples.js
```
