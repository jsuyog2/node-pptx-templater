# Contributing to pptx-templater

Thank you for your interest in contributing! This document guides you through the contribution process.

## 🚦 Quick Start

```bash
# Fork and clone the repository
git clone https://github.com/jsuyog2/pptx-templater.git
cd pptx-templater

# Install dependencies
npm install

# Run tests to verify setup
npm test

# Run with debug logging
PPTX_LOG_LEVEL=debug npm test
```

## 📐 Architecture Principles

### 1. No PowerPoint Libraries
This library is built exclusively on:
- `jszip` — ZIP archive manipulation
- `fast-xml-parser` — XML parsing/building
- `fs-extra` — File system utilities
- `crypto` — Built-in Node.js crypto (no package)

**Never add** `pptx-automizer`, `officegen`, `pptx-template`, or similar.

### 2. Separation of Concerns
Each manager has a single responsibility:
- `ZipManager` → ZIP I/O only
- `XMLParser` → XML serialization only
- `SlideManager` → Slide registry + ordering
- `ChartManager` → Chart XML updates
- etc.

### 3. Preserve Formatting
When updating content (text, charts, tables), always:
- **Replace ONLY** content nodes (`<a:t>`, `<c:v>`, etc.)
- **Preserve** all styling attributes (`<a:rPr>`, `<a:tcPr>`, etc.)
- **Never** rebuild entire elements from scratch

### 4. Relationship Safety
When adding relationships:
- Always use `RelationshipManager.addRelationship()` — never write `.rels` XML manually
- Always use `generateRelationshipId()` to avoid collisions
- Always register new parts in `[Content_Types].xml`

## 🧪 Writing Tests

### Unit tests
Place in `tests/unit/`. Test one class/function in isolation.
Mock the ZipManager to avoid file I/O:

```js
const mockZip = {
  readFile: async (path) => '<xml>...</xml>',
  writeFile: (path, content) => { /* capture writes */ },
  listFiles: () => [],
};
```

### Integration tests
Place in `tests/integration/`. Use fixture files from `tests/fixtures/`.

To add a test fixture PPTX:
```bash
# Place your .pptx file here
cp my-template.pptx tests/fixtures/sample.pptx
```

### Snapshot tests
Place in `tests/snapshot/`. Test that generated XML matches expected output:

```js
it('should generate correct slide XML', () => {
  const xml = buildNewSlideXml({ title: 'Test' }, 1);
  expect(xml).toMatchSnapshot();
});
```

## 🐛 Bug Reports

When reporting bugs, please include:

1. **Node.js version** (`node --version`)
2. **Library version** (`npm list pptx-templater`)
3. **Minimal reproduction** (ideally a code snippet + PPTX template)
4. **Error message** and full stack trace (`PPTX_LOG_LEVEL=debug`)
5. **Expected vs actual behavior**

## 📬 Pull Request Process

1. **Branch naming**: `feature/my-feature`, `fix/issue-123`, `docs/update-readme`
2. **Tests required**: All new features need unit tests
3. **JSDoc required**: All public methods need JSDoc with `@param`, `@returns`, `@example`
4. **No TypeScript**: Keep pure JavaScript (JSDoc types for IDE support)
5. **Lint clean**: `npm run lint` must pass
6. **Format**: Run `npm run format` before committing

## 🔒 Security

For security vulnerabilities, please email security@jsuyog2.com instead of opening a public issue.

## 📄 License

By contributing, you agree that your contributions will be licensed under the MIT License.
