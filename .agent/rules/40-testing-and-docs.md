# Testing, Documentation, and Definition of Done

## Test-first risk management
- Inspect the existing test runner, test conventions, fixture organization, and scripts before adding tests.
- Add a regression test for every bug fix where feasible.
- Cover success paths, invalid inputs, empty inputs, boundary indices, Unicode, malformed files, and repeated operations where relevant.
- Include integration tests for package-level changes; unit tests alone are insufficient for relationship/serialization behavior.
- Use small deterministic fixtures. Do not overwrite canonical fixtures or generated output used by unrelated tests.
- Tests must assert semantic outcomes and package integrity, not just that a method does not throw.
- Avoid flaky tests and unnecessary timing assumptions.
- Do not weaken or delete tests simply to get a green run. If a test is obsolete, explain why and replace its coverage.

## Expected validation
Run the most relevant checks first, then broader checks when practical:
1. Focused test file(s).
2. Full test suite.
3. Lint and format checks.
4. Type declaration/API consistency checks if available.
5. Documentation validation/build/inventory scripts if available.
6. Examples relevant to the change.
7. Benchmarks for meaningful performance changes.
8. Package validation/build checks defined by the repository.

Discover exact script names in `package.json`; do not blindly assume commands exist. If a command fails due to environment/dependency constraints, report the exact command and failure.

## Documentation discipline
When a public feature or behavior changes, update all relevant surfaces:
- JSDoc and public type declarations.
- API reference and feature inventory.
- README and quick-start examples where appropriate.
- Runnable examples.
- Changelog or migration notes for notable changes.
- Error/edge-case documentation and compatibility notes.

Keep examples executable against the current source and package exports. Never copy an API signature from an older website without verifying it.

## Definition of done
A task is complete only when:
- The requested behavior is implemented, or any blocker is precisely identified.
- Relevant tests have been added/updated and run.
- The public API and declarations remain aligned.
- Documentation and examples reflect the actual behavior.
- OpenXML package consistency has been considered for every changed part.
- No unrelated changes or secrets were introduced.
- The final response summarizes changed files, behavior, tests run, results, and any remaining risks. Distinguish verified facts from assumptions.
