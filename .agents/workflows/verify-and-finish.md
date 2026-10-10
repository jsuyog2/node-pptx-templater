# Workflow: Verify and Finish

1. Check `git status` and inspect the complete diff.
2. Search the repository for stale references to deleted `docs/`, `scripts/`, `examples/`, and `templates/` paths. Distinguish intentional historical references from broken live references.
3. Verify every original example/template is either migrated or explicitly documented as intentionally removed with a reason.
4. Run the unified single-file example using the documented command.
5. Run tests, lint/format, docs build, and package validation commands discovered from the actual repository.
6. Inspect the new docs output for missing assets, incorrect GitHub Pages base paths, broken internal navigation, theme-toggle behavior, and reduced-motion handling.
7. Validate generated PPTX ZIP/XML parts and relationships where feasible.
8. Verify npm and GitHub links.
9. Ensure the old folders are truly removed and no workflow/package script still depends on them.
10. Do not publish, push, or deploy remotely unless explicitly requested.
11. Report each check as Pass, Fail, or Not run, with exact evidence and limitations.
