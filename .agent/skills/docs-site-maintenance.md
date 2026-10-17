# Skill: Documentation Portal Maintenance & Schema Synchronization

> **Domain**: Next.js App Router Static Site Generation & Cloudflare Pages Deployment  
> **Target Subsystems**: `docs-site/`, `tests/unit/DocCoverage.test.js`

---

## 1. Schema-Driven Architecture

The documentation portal (`docs-site/`) is a Next.js 14+ static export hosted on Cloudflare Pages (`https://node-pptx-templater.pages.dev`). All content is driven by decoupled JSON schemas in `docs-site/src/data/`:

- `apiReference.json`: Method catalog with signatures, descriptions, parameters, returns, and examples.
- `features.json`: Feature matrix and capability breakdowns.
- `docsNavigation.json`: Sidebar hierarchy and routing.
- `compatibility.json`: Runtime compatibility matrix (Node, Lambda, Workers, Deno).
- `siteMeta.json`: SEO titles, metadata, OpenGraph tags, and repository links.

---

## 2. API Documentation SOP

Whenever public methods in `PPTXTemplater.js` are added, modified, or deprecated:

1. **Add Full JSDoc**:
   Add `@param`, `@returns`, and `@example` tags in `src/core/PPTXTemplater.js`. Missing JSDocs will fail automated checks.
2. **Update Schema**:
   Add or update the method definition in `docs-site/src/data/apiReference.json`.
3. **Validate Coverage**:
   ```bash
   npm run docs:validate  # Runs tests/unit/DocCoverage.test.js
   ```
4. **Compile & Preview Static Export**:
   ```bash
   npm run docs:build    # Compiles static site into docs-site/out/
   npm run docs:preview  # Local verification server
   ```

---

## 3. Configuration Guardrails

- `docs-site/next.config.mjs`: Must retain `output: 'export'`, `images: { unoptimized: true }`, and `trailingSlash: true`.
- `docs-site/wrangler.toml`: Must specify `pages_build_output_dir = "out"`.
