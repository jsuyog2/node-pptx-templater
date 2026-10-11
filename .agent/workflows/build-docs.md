# Documentation Build & Verification Workflow

This workflow describes the procedure for maintaining, testing, and building the modern Next.js documentation portal (`docs-site`).

---

## 1. Prerequisites
- Node.js >= 18 (recommended >= 20.12.0)
- Dependencies installed in root and `docs-site`:
  ```bash
  npm install
  npm --prefix docs-site install
  ```

---

## 2. Content Synchronization
When modifying public APIs, features, guides, or compatibility matrix, regenerate or update the decoupled JSON schemas in `docs-site/src/data/`:
- `docs-site/src/data/features.json` (Parsed from `project/Features.md`)
- `docs-site/src/data/apiReference.json`
- `docs-site/src/data/docsNavigation.json`
- `docs-site/src/data/compatibility.json`
- `docs-site/src/data/siteMeta.json`

Run data generator:
```bash
npm run docs:data
```

---

## 3. Local Development
To launch the live Next.js documentation development server with hot reload:
```bash
npm run docs:dev
```
The site runs at `http://localhost:3000`.

---

## 4. Static Export & Cloudflare Pages Build
The portal compiles directly into pure static HTML (`output: 'export'`) for zero-runtime serverless hosting on Cloudflare Pages:
```bash
npm run docs:build
```
This produces `docs-site/out/` containing:
- `index.html` (Home landing page with GSAP animations)
- `features/index.html` (Interactive feature grid)
- `docs/index.html` (Documentation overview)
- `docs/installation/index.html` (Installation guides)
- `docs/compatibility/index.html` (Runtime & serverless compatibility matrix)
- `docs/api/index.html` (Full-featured 10-domain API catalog)
- `sitemap.xml` & `robots.txt`

---

## 5. Preview Static Build
To preview the exported static HTML locally:
```bash
npm run docs:preview
```

---

## 6. Verification Gates
Run documentation coverage tests to ensure all public methods and sitemaps remain intact:
```bash
npm run docs:validate
```
Run repository test suite:
```bash
npm test
```
