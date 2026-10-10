# Workflow: Implement a Feature End to End

## 1. Understand
- Translate the request into observable behavior and acceptance criteria.
- Trace related public APIs, managers/modules, utilities, tests, docs, and examples.
- Check whether the capability exists partially under another name before adding new code.
- Identify OpenXML parts and relationship files affected.
- Note backward-compatibility constraints and supported runtime versions.

## 2. Plan briefly
Create a compact implementation checklist for multi-file work. Choose a coherent subsystem-sized change. Avoid a sprawling rewrite.

## 3. Implement
- Follow existing patterns and keep changes localized.
- Validate arguments and handle malformed or unsupported input deliberately.
- Preserve unaffected XML, relationships, formatting, and package parts.
- Keep API return values, chainability, async behavior, and declarations consistent.
- Avoid adding dependencies unless necessary.
- Do not claim unsupported formats or edge cases are supported without tests.

## 4. Test
- Add focused unit tests and package-level integration tests as appropriate.
- Test round-trip behavior by serializing and reloading the generated deck where possible.
- Verify relevant `.rels`, content types, IDs, and embedded workbook/media parts.
- Add regression cases for the reported issue.
- Run focused tests, then the full suite and quality checks available in the repo.

## 5. Document
- Update JSDoc, declarations, docs, examples, and changelog as warranted.
- Validate that all examples and signatures match the actual implementation.
- Note any limitations, unsupported edge cases, or intentional behavior changes.

## 6. Review
- Inspect the diff for unrelated edits, accidental generated files, secrets, stale comments, and untested behavior.
- Verify no existing user changes were overwritten.
- Report what changed, why, test commands/results, and known limitations.
