# Repository Reorganization Requirements

## Task A — remove old docs and scripts
1. Inspect the entire existing `docs/` directory, current GitHub Pages configuration, workflow files, package scripts, and all references to documentation build/deploy scripts.
2. Inspect the entire existing `scripts/` directory and determine which scripts are obsolete, which are referenced, and which contain useful functionality.
3. Preserve useful information by migrating it before deletion. Preserve required build/test/release behavior by moving or replacing essential scripts in a maintainable location and updating all callers.
4. Remove the existing GitHub Pages documentation directory and old scripts directory fully, as requested, after completing the migration.
5. Remove stale links, references, scripts, workflow paths, and configuration pointing at deleted locations.
6. If the repository has a GitHub Pages deployment workflow, update or replace it so the new docs site can be built and deployed from its new location. Do not leave broken CI references.
7. Keep the library's `src/`, package entry points, tests, lockfile, license, and package metadata safe unless changes are needed and validated.

## Task B — unify examples and templates
1. Inventory both existing `examples/` and `templates/` folders recursively.
2. Identify each example's input data, expected output, assets, dependencies, and execution command.
3. Consolidate them into one clear top-level directory, preferably `showcase/` unless the repository conventions suggest a better name.
4. Use descriptive subfolders, for example `showcase/examples/`, `showcase/templates/`, and `showcase/data/` only if this still meets the user's requirement for a unified folder. The examples and templates must be grouped under one parent.
5. Update all relative file paths, imports, README/docs links, test fixtures, package scripts, and CI references.
6. Remove the original `examples/` and `templates/` folders only after verifying that all content was migrated and references updated.
7. Avoid duplicating large binary files without reason. Keep template assets paired with their example and explain any intentional shared fixtures.

## Task C — one single-file example with JSON data
1. Create one self-contained JavaScript example file in the unified folder, e.g. `showcase/all-features.example.cjs` or `.mjs`, choosing the module format supported by the repository.
2. This single file must demonstrate the largest useful set of real, implemented public APIs across the library. Do not assume every item in the 86-feature catalog is implemented.
3. Include JSON-driven input data in a clearly marked constant or embedded JSON object. The file must run without requiring the user to manually create a separate JSON file.
4. Use the actual public API and current method signatures, verified against source and tests.
5. Where the library requires external template/media files, reference assets within the unified showcase folder using paths resolved relative to the example file, not the current working directory.
6. Include a single documented command to run the example. Generate outputs into a predictable output folder and avoid overwriting source fixtures.
7. Keep the file readable and organized into sections: setup, JSON data, load/create presentation, metadata, slides, text, shapes, images, tables, charts, hyperlinks, serialization/export, and validation—only for capabilities actually supported.
8. Add graceful capability checks or clearly explained skipped sections when an API is optional or not implemented. Do not hide genuine failures behind blanket `try/catch` blocks.
9. If “one file includes all examples” conflicts with the need for separate binary PPTX templates/media, consolidate the runnable code into one file and keep required assets inside the unified showcase folder. State this constraint clearly in the docs.
10. Ensure the example runs in a clean checkout using documented setup steps. Verify generated PPTX output structurally where possible.

## Preserve the feature catalog
Treat the supplied 86-feature list as a target checklist. For each feature, inspect implementation and either demonstrate it in the single-file example, document why it cannot be demonstrated, or mark it as unsupported/not implemented. Never fake support.
