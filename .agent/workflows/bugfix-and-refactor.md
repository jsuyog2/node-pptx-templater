# Workflow: Bug Fix or Refactor

1. Reproduce the defect or establish a failing regression test before editing, when feasible.
2. Trace the complete path from public API through orchestration to the low-level XML/package mutation.
3. Identify the smallest root cause; avoid patching only the visible symptom.
4. Preserve public contracts, formatting, namespaces, relationship targets, and unrelated XML.
5. Add a regression test that fails before the fix and passes after it.
6. If refactoring, separate behavior-preserving changes from behavior changes where practical.
7. Run focused tests, full tests, lint/format checks, and integration validation as appropriate.
8. Review the diff for accidental API drift and unnecessary rewrites.
9. Document any changed edge-case behavior and include migration guidance if compatibility is affected.
10. Finish with a concise report of root cause, changes, validation, and residual risks.
