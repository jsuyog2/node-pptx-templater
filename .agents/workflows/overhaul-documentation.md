# Workflow: Full Documentation and Showcase Overhaul

Follow this order and do not skip the inventory phase.

## Phase 1 — Repository reconnaissance
- Check repository root, Git status, package metadata, lockfiles, CI workflows, GitHub Pages settings, and current docs scripts.
- Inventory every file in `docs/`, `scripts/`, `examples/`, and `templates/`.
- Search all tracked files for references to these paths.
- Map the real public API, feature modules, supported Node.js versions, test commands, and current documentation claims.
- Preserve uncommitted work.

## Phase 2 — Migration plan
- Create a mapping table from each old example/template path to its new path under one parent directory.
- Identify which docs content remains useful and map it into the new site.
- Identify scripts that are obsolete versus essential. Plan replacement/migration for essential build/test/release tasks.
- Choose the site source directory and build system after inspecting the repository.
- Implement rather than stopping at a plan unless an external blocker prevents progress.

## Phase 3 — Consolidate examples/templates
- Move examples and templates under a single parent such as `showcase/`.
- Keep asset dependencies with the corresponding example or in a shared assets directory.
- Update every relative path and repository reference.
- Build one single-file JavaScript showcase with embedded JSON data. Demonstrate actual supported APIs; clearly skip/document unsupported ones.
- Run the showcase and inspect the generated output.

## Phase 4 — Build new documentation
- Create the new docs site in the selected site source directory.
- Implement Tailwind styles and GSAP animations, glassmorphism, accessible light/dark toggle, responsive navigation, quick start, feature catalog, guides, API references, showcase, troubleshooting, npm/GitHub links, and footer.
- Respect reduced-motion preferences and the GitHub Pages base path.
- Migrate and rewrite docs from source evidence, not by blindly copying stale prose.

## Phase 5 — Remove old folders and update automation
- Once migration is complete and verified, delete the old `docs/` and `scripts/` folders.
- Remove or update stale package scripts, CI workflow paths, deployment configuration, README references, and links.
- Configure GitHub Pages to deploy the new site.
- Do not push or deploy remotely unless the user explicitly asks.

## Phase 6 — Validation
- Run tests, lint, showcase, docs build, and available package checks.
- Check path resolution and GitHub Pages base URL behavior.
- Inspect generated PPTX ZIP/XML and important relationships.
- Search for references to removed directories.
- Review the full diff and confirm no user changes or unrelated files were lost.

## Completion
Return a concise but evidence-based summary of completed work, changed paths, validation commands/results, deployment status, and any remaining blockers. Do not claim the docs are live unless verified.
