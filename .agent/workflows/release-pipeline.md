# Workflow: CI/CD Pipeline & Staged Release Architecture

> **Purpose**: Standard Operating Procedure (SOP) for automated branch push gating, release publishing across package registries, and manual documentation deployment.

---

## 1. Architectural Topology & Gating Model

The continuous integration and delivery architecture is decoupled into 3 isolated pipelines to enforce strict separation of concerns:

```
[Branch Push / PR] ──► .github/workflows/ci.yml
                      ├── Step 1: ci (lint, format-check, tests)
                      ├── Step 2: verify-package-managers (npm, pnpm, yarn, bun)
                      └── Step 3: build (compile, verify outputs, pack dry-run)
                          └── [HARD STOP: No Registry Publication]

[GitHub Release]   ──► .github/workflows/publish.yml
(or workflow_dispatch)├── Step 1: build-and-test (tarball artifact creation)
                      ├── Step 2: verify-package-managers
                      ├── Step 3: publish-github-packages (--ignore-scripts, idempotent)
                      └── Step 4: publish-npm-registry (--ignore-scripts, idempotent)

[Manual Dispatch]  ──► .github/workflows/docs.yml
                      └── build-and-deploy (Next.js static export -> Cloudflare Pages)
```

---

## 2. Pipeline Stage Breakdown

### Stage 1: Push & Pull Request Gating (`ci.yml`)
- **Trigger**: Automatic on `push` and `pull_request` targeting `main`, `develop`, or `v3.0.0`.
- **Job Flow**:
  1. `ci`: Runs ESLint (`npm run lint`), Prettier (`npm run format:check`), and full Vitest suite (`npm test`).
  2. `verify-package-managers`: Compiles artifacts and installs the packaged tarball into clean sandboxes across all 4 major JavaScript package managers:
     - `npm` (CJS + ESM import)
     - `pnpm` (CJS + ESM import)
     - `yarn` (Berry v4 with `nodeLinker: node-modules`)
     - `bun` (ESM import)
  3. `build`: Compiles `dist/index.js`, `dist/index.mjs`, validates exports (`pkg.main`, `pkg.module`, `pkg.types`), and performs `npm pack --dry-run`.
- **Hard Stop Guarantee**: Pipeline stops immediately after artifact verification. Does NOT push git tags or publish packages.

### Stage 2: Dual-Registry Release Publishing (`publish.yml`)
- **Trigger**: Strictly gated to GitHub Releases (`on: release: types: [published]`) or manual dispatch (`workflow_dispatch`).
- **Distribution Endpoints**:
  - **npm**: Canonical un-scoped package `node-pptx-templater` (`npm install node-pptx-templater`).
  - **GitHub Packages**: Scoped package `@<owner>/node-pptx-templater` (`npm install @<owner>/node-pptx-templater`).
- **Key Invariants**:
  - **`--ignore-scripts`**: Added to `npm publish` commands to prevent redundant `prepublishOnly` execution in environments where development tools (e.g. `tsup`) are omitted.
  - **Idempotent Checks**: Pre-queries `npm view` before publish. Exits cleanly with status code `0` and `::notice::` if the version already exists.

### Stage 3: Documentation Deployment (`docs.yml`)
- **Trigger**: Strictly manual `workflow_dispatch` (no automated push triggers on any branch).
- **Target Options**: `cloudflare` (production/preview) or `github-pages`.
- **Manual Execution**:
  ```bash
  # Local deployment script
  npm run docs:deploy
  ```

---

## 3. Pre-Release Verification Checklist

Before tagging a new release:
1. Ensure all local tests pass: `npm test && npm run lint`.
2. Verify package build outputs: `npm run build && npm pack --dry-run`.
3. Check version synchronicity across `package.json`, `package-lock.json`, and `CHANGELOG.md`.
4. Create a GitHub Release with tag format `vX.Y.Z-beta.N` or `vX.Y.Z` to trigger `publish.yml`.
