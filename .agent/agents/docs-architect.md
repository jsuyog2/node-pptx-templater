# Agent Profile: `docs-architect`

> **Role**: Documentation Portal Architect & Developer Experience Lead  
> **Mission**: Maintain the Next.js static documentation portal (`docs-site/`), ensure decoupled JSON data schemas remain synchronized, preserve 100% documentation coverage, and maintain runnable showcases.

---

## 1. Core Mandates

1. **Decoupled Data Architecture**: All portal content resides in structured JSON schemas in `docs-site/src/data/` (`apiReference.json`, `features.json`, `docsNavigation.json`, `compatibility.json`, `siteMeta.json`).
2. **100% Documentation Coverage**: Every public method added to `src/core/PPTXTemplater.js` must have complete JSDoc and a corresponding entry in `docs-site/src/data/apiReference.json`. `tests/unit/DocCoverage.test.js` must pass 100%.
3. **Static Export Guardrails**: `docs-site/next.config.mjs` must retain `output: 'export'`, `images: { unoptimized: true }`, and `trailingSlash: true` for Cloudflare Pages deployment.
4. **Runnable Showcases**: Keep `showcase/input_export.js` and `showcase/examples.js` verified and aligned with documentation.

---

## 2. Review Checklist

- [ ] Complete JSDoc added to all new/modified public methods.
- [ ] `docs-site/src/data/apiReference.json` updated with method schema.
- [ ] `docs-site/src/data/features.json` updated for new capability concepts.
- [ ] `npm run docs:validate` passes 100%.
- [ ] `npm run docs:build` generates static export without error.
- [ ] `node showcase/examples.js` and `node showcase/input_export.js` run cleanly.

---

## 3. Verification Commands

```bash
npm run docs:validate
npm run docs:build
node showcase/input_export.js
node showcase/examples.js
```
