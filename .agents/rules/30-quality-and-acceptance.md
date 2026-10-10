# Quality Gates and Acceptance Criteria

## General
- Work incrementally, but complete the requested tasks end to end.
- Keep diffs focused and avoid unrelated API changes.
- Do not remove tests or weaken assertions to make checks pass.
- Keep package runtime code separate from documentation-site dependencies when practical.
- Update lockfiles when package dependencies change.
- Verify all package scripts and CI references after moving/deleting directories.

## Acceptance criteria
### Old folders
- [ ] Existing GitHub Pages documentation source is inventoried, useful content migrated, and old docs source removed as requested.
- [ ] Old `scripts/` folder is removed after essential functionality is migrated or obsolete scripts are confirmed unnecessary.
- [ ] No broken references remain to removed paths.
- [ ] GitHub Pages workflow points at the new site and uses valid build/artifact paths.

### Examples/templates
- [ ] All original examples and templates are inventoried and accounted for.
- [ ] A single parent directory contains consolidated examples, templates, data, and required assets.
- [ ] All references, scripts, tests, and docs use the new paths.
- [ ] Original directories are removed after migration and validation.

### Single-file showcase
- [ ] One runnable JavaScript file contains all applicable example logic and embedded JSON data.
- [ ] It uses actual exported APIs and documents which capabilities are skipped or unsupported.
- [ ] It uses stable paths relative to itself for required assets.
- [ ] It runs from a clean checkout using the documented command.
- [ ] Output PPTX is structurally checked where feasible.

### Documentation website
- [ ] Built with Tailwind CSS and GSAP.
- [ ] Glassmorphism design is consistent and readable.
- [ ] Light and dark theme toggle works.
- [ ] Responsive and keyboard accessible.
- [ ] Reduced-motion preference is respected.
- [ ] npm and GitHub links work.
- [ ] Quick start, guides, API references, showcase, feature status, and troubleshooting are present and source-verified.
- [ ] GitHub Pages base path and deployment workflow are tested.

## Validation
Discover commands from `package.json` and CI; do not assume command names. Run:
1. package tests and lint/format checks;
2. showcase execution;
3. docs build;
4. local docs preview or static output checks;
5. link/base-path and console-error checks if available;
6. PPTX ZIP/XML validation for generated output;
7. final Git diff and status review.

If a check cannot run, record why and mark it not run rather than passed.

## Final report
Summarize:
- directories/files deleted, migrated, created, and modified;
- where the new site source and consolidated showcase live;
- how to run the single-file example and build/preview docs;
- test commands and actual results;
- deployment status and remaining limitations.
