# Workflow: Release Readiness

Do not publish, tag, push, or release unless explicitly requested. This workflow prepares and validates a candidate release locally.

## Checklist
- [ ] Confirm current branch, working tree, version, and intended release scope.
- [ ] Review `package.json` fields: name, version, exports, main/module/types, bin, files, engines, dependencies, scripts, and license.
- [ ] Confirm packed files include all runtime files and declarations required by consumers, and exclude private fixtures/secrets/unneeded development assets.
- [ ] Run the full test suite, lint, format check, package validation, docs validation, and relevant examples.
- [ ] Verify the published entry point and exported symbols match declarations and docs.
- [ ] Check CommonJS/ESM loading behavior if both are supported.
- [ ] Check Node.js engine claims against implementation and CI.
- [ ] Check ZIP/XML security-sensitive changes and dependency advisories using available local tools.
- [ ] Review README, API docs, examples, changelog, and migration notes for consistency.
- [ ] Inspect generated PPTX artifacts structurally; verify relevant relationships and content types.
- [ ] Run performance benchmarks for material hot-path changes where practical.
- [ ] Review diff and ensure no secrets, accidental binary outputs, or unrelated edits are included.
- [ ] Summarize blockers and evidence. Do not mark an unchecked item as passed.

## Report
Return a checklist with Pass / Fail / Not run, exact commands, key findings, and any release blockers.
