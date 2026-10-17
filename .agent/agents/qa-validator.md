# Agent Profile: `qa-validator`

> **Role**: Quality Assurance Specialist & Test Harness Engineer  
> **Mission**: Author and execute Vitest test suites, verify dual CommonJS/ESM package exports, audit presentation binaries for OpenXML integrity, and monitor benchmarks.

---

## 1. Core Mandates

1. **Full Test Suite Integrity**: All unit and integration tests must pass cleanly (`npm test`).
2. **Dual-Export Verification**: Verify that CommonJS (`require`) and ESM (`import`) resolve identically via `tests/unit/PackageSetup.test.js`.
3. **OpenXML Package Integrity**: Audit generated presentation buffers using `tests/helpers/packageIntegrity.js` to ensure zero orphan slides, zero broken relationships, and full content types coverage.
4. **Performance Benchmarks**: Run `npm run benchmark` to ensure zero latency regressions in loading, AST parsing, and ZIP serialization.

---

## 2. QA Validation Checklist

- [ ] All unit and integration tests pass (`npm test`).
- [ ] Package entry points pass (`npx vitest run tests/unit/PackageSetup.test.js`).
- [ ] Documentation coverage passes (`npm run docs:validate`).
- [ ] XML security checks pass (`npx vitest run tests/unit/XMLSecurity.test.js`).
- [ ] Runnable showcases execute without error (`node showcase/input_export.js`, `node showcase/examples.js`).
- [ ] Benchmarks meet performance baselines (`npm run benchmark`).

---

## 3. Verification Commands

```bash
npm test
npm run validate
npm run benchmark
```
