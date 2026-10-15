# Workflow: Building & Deploying the Documentation Portal

> **Purpose**: Step-by-step procedure for verifying, building, and deploying the Next.js static documentation portal to Cloudflare Pages.

---

## 1. Prerequisites & Environment

The documentation portal is located in `docs-site/` and targets Cloudflare Pages:
- **Framework**: Next.js 14+ App Router
- **Static Output**: `output: 'export'`
- **Target URL**: `https://node-pptx-templater.pages.dev`
- **Deployment Tool**: Wrangler / Cloudflare Pages Git Integration

---

## 2. Step-by-Step Build & Verification Workflow

```
┌────────────────────────────────────────────────────────┐
│ Step 1: Validate Documentation Coverage                │
│ npm run docs:validate                                  │
│ • Asserts 100% JSDoc coverage across public methods    │
│ • Asserts sitemap.xml & robots.txt valid URLs          │
└──────────────────────────┬─────────────────────────────┘
                           ▼
┌────────────────────────────────────────────────────────┐
│ Step 2: Regenerate Derived Data (Optional)             │
│ npm run docs:data                                      │
│ • Updates scripts/generate-next-data.js indices        │
└──────────────────────────┬─────────────────────────────┘
                           ▼
┌────────────────────────────────────────────────────────┐
│ Step 3: Compile Next.js Static Export                  │
│ npm run docs:build                                     │
│ • Emits static HTML/CSS/JS into docs-site/out/         │
└──────────────────────────┬─────────────────────────────┘
                           ▼
┌────────────────────────────────────────────────────────┐
│ Step 4: Preview Local Static Build                     │
│ npm run docs:preview                                   │
│ • Serves docs-site/out/ at http://localhost:3000      │
└──────────────────────────┬─────────────────────────────┘
                           ▼
┌────────────────────────────────────────────────────────┐
│ Step 5: Cloudflare Pages Deployment                    │
│ Automated via Git commit to `main` or via Wrangler:   │
│ npx wrangler pages deploy docs-site/out                │
│     --project-name node-pptx-templater-docs            │
└────────────────────────────────────────────────────────┘
```

---

## 3. Configuration Invariants

### 3.1 `docs-site/next.config.mjs`
Ensure the configuration strictly retains static export settings:
```javascript
/** @type {import('next').NextConfig} */
const nextConfig = {
  output: 'export',
  images: {
    unoptimized: true,
  },
  trailingSlash: true,
};

export default nextConfig;
```

### 3.2 `docs-site/wrangler.toml`
Cloudflare Pages configuration:
```toml
name = "node-pptx-templater-docs"
compatibility_date = "2024-09-23"
pages_build_output_dir = "out"
```

---

## 4. Troubleshooting Build Errors

1. **Undocumented Method Error during `npm run docs:validate`**:
   - Cause: A public method was added to `PPTXTemplater.js` without JSDoc.
   - Solution: Add standard JSDoc block to the method in `src/core/PPTXTemplater.js`.
2. **Missing Sitemap Entry**:
   - Verify `docs-site/public/sitemap.xml` contains all active routes and uses the canonical domain `https://node-pptx-templater.pages.dev/`.
3. **Broken Image Optimization**:
   - Cause: Next.js standard `<Image>` component fails in static exports without a custom loader or `unoptimized: true`.
   - Solution: Keep `images: { unoptimized: true }` in `next.config.mjs`.
