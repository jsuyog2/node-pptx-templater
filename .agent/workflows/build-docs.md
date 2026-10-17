# Workflow: Building & Deploying Documentation Portal

> **Purpose**: SOP for verifying, building, and deploying the Next.js static documentation portal to Cloudflare Pages.

---

## 1. Execution Steps & Commands

```bash
# Step 1: Validate JSDoc coverage & SEO metadata (sitemap.xml / robots.txt)
npm run docs:validate

# Step 2: (Optional) Rebuild derived documentation indices
npm run docs:data

# Step 3: Compile Next.js static export into docs-site/out/
npm run docs:build

# Step 4: Preview built static portal locally
npm run docs:preview

# Step 5: Deploy static export to Cloudflare Pages (via Wrangler or Git push)
npx wrangler pages deploy docs-site/out --project-name node-pptx-templater-docs
```

---

## 2. Configuration Guardrails

- `docs-site/next.config.mjs`: `output: 'export'`, `images: { unoptimized: true }`, `trailingSlash: true`.
- `docs-site/wrangler.toml`: `pages_build_output_dir = "out"`.
- `docs-site/public/sitemap.xml`: Canonical domain `https://node-pptx-templater.pages.dev/`.
