# node-pptx-templater

[![npm version](https://img.shields.io/badge/version-3.0.0--alpha.0-blue.svg)](https://www.npmjs.com/package/node-pptx-templater)
[![Vitest Passing](https://img.shields.io/badge/tests-280%20passing-brightgreen.svg)](https://github.com/jsuyog2/node-pptx-templater/actions)
[![Node.js Support](https://img.shields.io/badge/node-%3E%3D20.12.0-blue.svg)](https://nodejs.org)
[![Zero Native Dependencies](https://img.shields.io/badge/dependencies-pure%20JS-success.svg)](https://www.npmjs.com/package/node-pptx-templater)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

> **High-performance, spatial-aware OpenXML PowerPoint manipulation, templating, and layout engine for Node.js.**

Operates with **zero external native binary dependencies** (no LibreOffice, no Java, no headless browser) and executes pure, non-corrupting in-memory OpenXML transformations.

- 📖 **Documentation Portal:** [https://jsuyog2.github.io/node-pptx-templater/](https://jsuyog2.github.io/node-pptx-templater/)
- 📦 **npm Package:** [https://www.npmjs.com/package/node-pptx-templater](https://www.npmjs.com/package/node-pptx-templater)
- 🐙 **GitHub Repository:** [https://github.com/jsuyog2/node-pptx-templater](https://github.com/jsuyog2/node-pptx-templater)

---

## ⚡ Key Highlights

- **Anti-Corruption Architecture:** Automatically synchronizes `<c:chartSpace>` chart XML and embedded `.xlsx` workbook cache parts (`ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx`) to prevent PowerPoint "Repair Presentation" warnings.
- **Run Normalization (Anti-Fragmentation):** Merges split `<a:r>` text runs inside `<a:p>` paragraphs before matching placeholder tags (`{{variable}}`), preserving formatting, colors, and fonts without brittle regexes.
- **Unified Showcase Suite:** Run all 12 production scenarios with a single command (`npm run examples` or `node showcase/examples.js`) with embedded JSON data.
- **Dynamic Table Engine:** Table updates, custom cell borders, row striping, nested arrays for `rowspan` vertical cell merges, and auto-pagination (`splitTableOnOverflow`) with repeated headers and pinned summary rows.
- **Layer Stacking & Spatial Geometry:** Z-order manipulation (`bringToFront`, `sendToBack`), spatial alignment (`alignShapes`), and uniform distribution (`distributeShapes`).
- **Serverless & Streaming Ready:** Export presentations directly in memory with `toBuffer()` for AWS Lambda, Google Cloud Functions, or Docker microservices.
- **Multi-Format Media:** Add or replace PNG, JPEG, WebP, Base64 data URIs, and TIFF images with automatic content-type management.

---

## 📦 Installation

```bash
npm install node-pptx-templater
# or
pnpm add node-pptx-templater
# or
yarn add node-pptx-templater
```

**Requirements:** Node.js `>= 20.12.0`. Supports both CommonJS (`require`) and ESM (`import`).

---

## 🚀 Quick Start

### CommonJS

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function generateReport() {
  // 1. Load template presentation
  const ppt = await PPTXTemplater.load('./showcase/templates/sample.pptx');
  ppt.useAllSlides();

  // 2. Dynamic placeholder replacement
  ppt.replaceMultiple({
    '{{title}}': 'Q3 Financial Performance',
    '{{company}}': 'Acme Global Inc.',
    '{{date}}': 'October 2026'
  });

  // 3. Update charts with automatic Excel cache synchronization
  ppt.useSlide(1);
  ppt.updateChart('chart1', {
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [
      { name: 'Cloud ARR', values: [18.4, 22.1, 26.8, 31.5] },
      { name: 'AI Solutions', values: [8.2, 12.6, 17.4, 23.0] }
    ]
  });

  // 4. Save modified presentation
  await ppt.save('./output/report.pptx');
  console.log('✓ Presentation generated successfully!');
}

generateReport();
```

### ESM

```javascript
import { PPTXTemplater } from 'node-pptx-templater';

const ppt = await PPTXTemplater.load('./showcase/templates/sample.pptx');
ppt.useSlide(1);
ppt.replaceTextByTag('{{title}}', 'AI Systems Architecture');

// Export directly to Buffer in memory (perfect for AWS Lambda)
const buffer = await ppt.toBuffer();
console.log(`✓ Exported ${buffer.length} bytes to memory Buffer`);
```

---

## 🎪 Unified Showcase Suite

All 12 supported demonstrations and sample templates are consolidated into the `showcase/` directory:

```bash
# Run all 12 showcase demonstrations
npm run examples
# or:
node showcase/examples.js
```

Generated presentations are stored in `showcase/output/`.

| Demo | Focus Area | Demonstrated APIs |
| :--- | :--- | :--- |
| `01` | **Basic Report & 16:9** | `setSlideSize('16:9')`, `replaceMultiple()`, `findText()` |
| `02` | **Chart & Excel Sync** | `updateChart()`, `updateDataLabels()` |
| `03` | **Tables, Striping & Overflow** | `updateTable()`, `addTableRow()`, `stripeRows()`, `splitTableOnOverflow()` |
| `04` | **Nested Tables & Extraction** | `addTableRow({mergeStrategy:'rowspan'})`, `getTableRows()` |
| `05` | **Shapes, Z-Order & Alignment** | `addShape()`, `bringToFront()`, `alignShapes()` |
| `06` | **Cell Shapes & Badges** | `addTableRow([{type:'circle', fill:'#10B981'}, ...])` |
| `07` | **Images: PNG, WebP & Base64** | `addImage()`, `replaceImage()` |
| `08` | **Slide Cloning & Import** | `duplicateSlide()`, `reorderSlides()`, `importSlideFrom()` |
| `09` | **Hyperlinks & Navigation** | `addHyperlink()`, `addSlideLink()`, `getHyperlinks()` |
| `10` | **Typography & Text Runs** | `addShape(textStyle)`, `appendTextRun()` |
| `11` | **OpenXML Folder Workflow** | `extractPptx()`, `saveToFolder()`, `buildPptx()` |
| `12` | **Serverless Buffer Export** | `toBuffer()` (in-memory ZIP archive) |

---

## 🛠️ Testing & Validation

```bash
# Run full Vitest test suite (280 passing tests across 42 test suites)
npm test

# Run package validation and showcase suite
npm run validate

# Run ESLint
npm run lint

# Check code formatting
npm run format:check

# Verify documentation site
npm run docs:build
```

---

## 📄 License

MIT © [node-pptx-templater contributors](LICENSE)
