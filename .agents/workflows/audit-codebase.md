# Workflow: Audit the Existing Codebase

Use this workflow before major work or when entering the repository for the first time.

## Steps
1. Confirm repository root and inspect Git status. Do not discard pre-existing user changes.
2. Read `package.json`, lockfile metadata, README, changelog, contributing/security files, lint/format config, test config, CI workflows, and package entry points.
3. Inventory `src/`, `tests/`, `examples/`, `docs/`, `scripts/`, `benchmarks/`, templates, and generated files.
4. Trace public exports from entry point to implementation modules. Identify the real API and module format.
5. Map each feature area to source files, tests, docs, examples, and exports.
6. Run the repository's documented validation commands if dependencies/environment permit.
7. Compare source with the online documentation and target feature catalog. Record mismatches, stale signatures, missing coverage, and unsupported claims.
8. Inspect relationship and package-manifest logic for slide operations, tables, charts, images, and output serialization.
9. Inspect input validation and security handling for ZIP paths, XML parser configuration, oversized inputs, and error handling.
10. Produce a concise audit report, preferably in the chat first. Create/update a repository audit document only when it provides durable value and fits existing docs conventions.

## Audit output
Include:
- Actual repository structure and architectural map.
- Public API and runtime compatibility findings.
- Test/lint/docs status with exact commands and results.
- Feature catalog status by subsystem, with file/test evidence.
- High-risk correctness/security gaps.
- Documentation drift and missing tests.
- Prioritized next steps: critical, high, medium, low.

Do not implement a broad refactor during an audit unless the user asked for fixes too. Fix small, clearly safe blockers only when they are necessary to complete the audit, and disclose them.
