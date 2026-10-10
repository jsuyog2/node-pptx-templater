# Architecture, JavaScript, and API Design

## Baseline
- This is an existing JavaScript-first Node.js package. Keep implementation in JavaScript unless the repository itself establishes a different convention or the user requests a migration.
- Follow the current `package.json`, module format, export map, supported Node.js engines, lint rules, formatting configuration, and test runner. Do not infer the runtime floor from old documentation; verify it from the checked-out package metadata and CI.
- Preserve the current CommonJS/ESM compatibility strategy and public entry points. Do not change `main`, `module`, `exports`, `types`, `bin`, `files`, or `engines` casually.
- Do not add a large dependency for functionality that can be implemented safely with existing dependencies. Any dependency addition must have a clear maintenance, security, and bundle-size rationale.

## Layering
Inspect the existing architecture before reorganizing. Maintain clear responsibilities between orchestration and specialized modules, which may include ZIP/package access, XML parsing/building, relationships, slides, text, tables, charts, media, shapes, hyperlinks, validation, and output. Use the repository's actual names and boundaries rather than blindly imposing this example architecture.

- Keep low-level OpenXML manipulation separate from public fluent API orchestration.
- Centralize shared relationship, ID, path, unit-conversion, XML escaping, and package-manifest logic where the current architecture supports it.
- Avoid duplicated XML mutation logic across managers.
- Keep public methods consistent in naming, argument validation, return values, error behavior, and chainability.
- Maintain async behavior where I/O or existing public contracts require it.
- Prefer explicit, readable code over clever abstractions.
- Use private fields and encapsulation in line with existing conventions.
- Validate external inputs and report actionable errors with useful context. Do not swallow errors or silently produce malformed output.

## Public API and compatibility
- Treat exported classes, functions, CLI commands, type declarations, and documented examples as compatibility-sensitive.
- Add regression tests for fixed bugs and compatibility tests for public behavior.
- If a method is chainable, preserve that contract unless a deliberate, documented breaking change is approved by the task.
- Keep JavaScript runtime behavior and `.d.ts` declarations aligned.
- Ensure documented examples use real exported names and valid argument shapes verified against source/tests.
- Do not create fake aliases merely to make documentation examples appear to work.

## Performance
- Avoid repeatedly parsing or serializing the same XML or workbook without need.
- Preserve streaming and buffer-based pathways where supported.
- Consider large presentations, many slides, high-row-count tables, large media, and serverless memory limits.
- Measure before making performance claims. Use existing benchmarks or add targeted benchmarks when a change affects a hot path.
- Never trade away package correctness for a micro-optimization.
