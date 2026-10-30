# node-pptx-templater

<div align="center">

[![npm version](https://img.shields.io/badge/npm-v3.0.0--beta.0-blue.svg?style=flat-square)](https://www.npmjs.com/package/node-pptx-templater)
[![npm downloads](https://img.shields.io/npm/dm/node-pptx-templater.svg?style=flat-square&color=blue)](https://www.npmjs.com/package/node-pptx-templater)
[![Vitest Tests](https://img.shields.io/badge/tests-340%2B%20passing-brightgreen.svg?style=flat-square)](https://github.com/jsuyog2/node-pptx-templater/actions)
[![Bundle Size](https://img.shields.io/badge/runtime-pure%20JS%20%3C2MB-success.svg?style=flat-square)](https://www.npmjs.com/package/node-pptx-templater)
[![Serverless](https://img.shields.io/badge/serverless-AWS%20Lambda%20%7C%20Cloudflare%20Workers%20%7C%20Vercel%20Edge-purple.svg?style=flat-square)](https://node-pptx-templater.pages.dev/docs/capabilities/)
[![Node.js Support](https://img.shields.io/badge/node-%3E%3D18.x-blue.svg?style=flat-square)](https://nodejs.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square)](https://opensource.org/licenses/MIT)

**High-Performance OpenXML & Flat OPC PowerPoint Template Engine for Node.js**  
**Fastify, Express, AWS Lambda, Cloudflare Workers & Vercel Edge PowerPoint Automation**

[📖 Documentation Overview](https://node-pptx-templater.pages.dev/docs/) &nbsp;•&nbsp; [⚡ Installation Guide](https://node-pptx-templater.pages.dev/docs/installation/) &nbsp;•&nbsp; [🚀 Capabilities & Matrix](https://node-pptx-templater.pages.dev/docs/capabilities/) &nbsp;•&nbsp; [📚 API Reference](https://node-pptx-templater.pages.dev/docs/api-reference/)

</div>

---

## 🚀 Overview: Pure JavaScript PowerPoint Automation

**`node-pptx-templater`** is an enterprise-grade, low-overhead PowerPoint (PPTX) template engine and presentation generator written in **100% pure JavaScript**. It enables developers to dynamically update existing PowerPoint templates or create brand-new presentations from scratch across Node.js, Express, Fastify, AWS Lambda, Cloudflare Workers, and Vercel Edge.

### Why Choose `node-pptx-templater`?

- 🚫 **Zero Native Binaries & Zero Microsoft Office Dependencies**: No `node-gyp`, no C/C++ compilation, no LibreOffice headless daemon (500MB+ Docker containers), and no Windows COM interop.
- ⚡ **Sub-50ms Cold Starts**: In-memory Virtual File System (VFS) with streaming zip inflation and fast selective XML AST manipulation.
- 📊 **Dual-Sync Chart Engine**: Updates both DrawingML XML (`<c:chartSpace>` / `<cx:chartSpace>`) and embedded Excel binaries (`ppt/embeddings/*.xlsx`) across 17 chart types. Prevents PowerPoint "corrupted presentation" repair dialogs.
- 📄 **Dynamic Multi-Slide Table Pagination**: Automatically calculates row heights in EMUs and paginates overflowing table rows across cloned slides with repeating headers and footer row anchoring.
- 🧩 **AST Run Normalization (Fragment Healing)**: Eliminates broken template variables caused by PowerPoint splitting tokens like `{{COMPANY}}` into separate `<a:r>` runs.
- 🔗 **Unified Elements & Hyperlinks Subsystem**: Complete element introspection, z-order layer stacking, spatial collision detection, and ISO/IEC 29500 hyperlinks.
- 🛡️ **Hardened XML Security**: Built-in defenses against XXE attacks, entity expansion bombs (Billion Laughs), and Zip-Slip path traversal.

---

## 📊 Comprehensive Comparison Matrix

| Capability / Benchmark | `node-pptx-templater` | `python-pptx` | Headless LibreOffice | `pptxgenjs` | Native MS Office / COM |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Runtime Environment** | **Pure JavaScript** (Node.js, Edge, Lambda) | Python C-Extensions | Heavy C++ Binary / Docker | Pure JavaScript | Windows OS Only |
| **Existing PPTX In-Place Update** | ✅ **Full In-Place AST Mutation** | ✅ Supported | ⚠️ Brittle UNO Macro Scripting | ❌ Blank creation only | ✅ Full API |
| **Template-Free Blank Creation** | ✅ **`createBlank()` (16:9, 4:3)** | ✅ Supported | ⚠️ Macro required | ✅ Supported | ✅ Supported |
| **Chart & Excel Dual-Sync** | ✅ **XML + XLSX Binary Sync** | ⚠️ Limited chart types | ⚠️ Slow binary recalculation | ⚠️ Basic values only | ✅ Native Excel engine |
| **17 Chart Types (Inc. 2016+ Extended)** | ✅ **Waterfall, Treemap, Funnel, etc.** | ❌ Standard charts only | ⚠️ Partial formatting | ❌ Standard charts only | ✅ Supported |
| **Run Fragment Healing** | ✅ **Automatic AST Run Assembly** | ❌ Manual run splitting | ❌ Brittle regex find/replace | ❌ N/A | ❌ N/A |
| **Dynamic Table Auto-Pagination** | ✅ **Automatic Multi-Slide Splitting** | ❌ Manual math required | ❌ Unsupported | ❌ Manual math required | ❌ Unsupported |
| **Serverless Cold Start** | ⚡ **< 15ms (< 2MB overhead)** | ⚠️ 300–800ms | ❌ 10–20 seconds (500MB+ Docker) | ⚡ < 20ms | ❌ Impossible |
| **Web Streaming & Framework Adapters** | ✅ **Express, Fastify, Web Response** | ❌ File / Bytes only | ❌ Process stdout only | ⚠️ Buffer / Base64 only | ❌ COM pipes only |
| **Flat OPC Single-File XML** | ✅ **Direct Parse & Serialization** | ❌ Unsupported | ⚠️ Incomplete | ❌ Unsupported | ⚠️ File export only |

---

## 🏛️ Exact 8 Standardized Categories

Public APIs are strictly structured across 8 canonical categories:

1. **`Import/Export`** (18 Methods): Static loading (`load`, `createBlank`, `preload`, `fromCache`, `clearCache`), export channels (`saveToFile`, `toBuffer`, `toStream`, `toBase64`, `toArrayBuffer`, `toResponse`, `toExpress`, `toFastify`), and package extractions (`extractToZip`, `extractToXml`, `toFlatXml`, `extractPptx`, `buildPptx`).
2. **`Slides`** (24 Methods): Slide selection (`useSlide`, `useAllSlides`), lifecycle (`addSlide`, `duplicateSlide`, `removeSlide`, `getSlides`, `getSlide`, `importSlideFrom`, `exportSlides`, `moveSlide`, `reorderSlides`, `tagSlide`, `getSlideDimensions`, `getSlideLayout`), section grouping (`useSection`, `addSection`, `updateSection`, `removeSection`, `getSection`), visibility (`hideSlide`, `unhideSlide`, `resetSlide`), and collaborative comments (`addComment`, `removeComment`).
3. **`Text & Placeholders`** (10 Methods): Template substitution, typography, and lists (`replaceText`, `addText`, `replaceTextByTag`, `updateText`, `addLists`, `getList`, `findText`, `addStyle`, `removeStyle`, `updateStyle`).
4. **`Tables`** (30 Methods): Full table lifecycle, data hydration, rows, columns, cell styling, and merging (`addTable`, `updateTable`, `updateTableData`, `removeTable`, `getTables`, `getTable`, `cloneTable`, `addTableRow`, `addTableColumn`, `removeTableRow`, `removeTableColumn`, `updateRowData`, `updateColumnData`, `updateTableCell`, `clearRowData`, `clearColumnData`, `clearCellData`, `getTableData`, `getColumnData`, `getRowData`, `getColumn`, `getRow`, `getCell`, `getCellData`, `mergeCells`, `splitCell`, `styleTable`, `styleColumn`, `styleRow`, `styleCell`).
5. **`Charts`** (11 Methods): 17 chart types with Excel Dual-Sync (`addChart`, `updateChart`, `updateChartData`, `removeChart`, `clearCharts`, `getCharts`, `getChart`, `cloneChart`, `getDataLabels`, `getChartLabelPositions`).
6. **`Shapes`** (11 Methods): Preset vector shapes, fills, outlines, and containers (`addShape`, `updateShape`, `removeShape`, `clearShapes`, `getShapes`, `getShape`, `cloneShape`, `getGroupShapes`, `groupShapes`, `ungroupShapes`).
7. **`Images & Media`** (11 Methods): Media embedding and deduplication (`addImage`, `updateImage`, `removeImage`, `addVideo`, `updateVideo`, `removeVideo`, `addAudio`, `updateAudio`, `removeAudio`, `getMedias`, `getMedia`, `cloneMedia`).
8. **`Elements`** (15 Methods): Unified queries, layer stacking, collision detection, and hyperlinks (`getElements`, `getElement`, `removeElement`, `removeAllElements`, `setZOrder`, `bringToFront`, `sendToBack`, `bringForward`, `sendBackward`, `addHyperlink`, `updateHyperlink`, `removeHyperlink`, `getHyperlinks`, `getHyperlink`, `setHyperlink`, `checkCollision`, `isContained`).

---

## 📦 Installation

```bash
# npm (canonical package)
npm install node-pptx-templater

# or install pre-release beta:
npm install node-pptx-templater@beta

# pnpm
pnpm add node-pptx-templater

# yarn
yarn add node-pptx-templater

# bun
bun add node-pptx-templater

# or via GitHub Packages registry:
npm install @jsuyog2/node-pptx-templater --registry=https://npm.pkg.github.com
```

**Requirements:** Node.js `>= 18.x` (fully verified on `18.x`, `20.x`, `22.x+`). Dual CommonJS (`require`) and ESM (`import`) distribution.

---

## ⚡ Production Quick Starts

### 1. PPTX Template Update with AST Fragment Healing
```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function updateReport() {
  const ppt = await PPTXTemplater.load('./templates/executive_summary.pptx');
  
  // Target slide 1 and replace placeholder tokens cleanly
  ppt.useSlide(1).replaceText({
    '{{COMPANY}}': 'Acme Global Technologies Inc.',
    '{{FISCAL_YEAR}}': '2026',
    '{{REVENUE}}': '$52.4M',
    '{{GROWTH}}': '+34% YoY'
  });

  await ppt.saveToFile('./output/annual_report.pptx');
}
```

### 2. Native Charts with Excel Dual-Sync
```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function updateCharts() {
  const ppt = await PPTXTemplater.load('./templates/dashboard.pptx');

  // Synchronously updates XML DrawingML and backing Excel workbook
  ppt.useSlide(2).updateChartData('RevenueByQuarter', {
    categories: ['Q1', 'Q2', 'Q3', 'Q4'],
    series: [
      { name: 'Cloud Services', values: [18.2, 22.4, 27.1, 33.8] },
      { name: 'Consulting', values: [10.5, 11.2, 12.0, 12.8] }
    ]
  });

  await ppt.saveToFile('./output/updated_dashboard.pptx');
}
```

### 3. Dynamic Tables with Auto-Pagination & Multi-Slide Splitting
```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function createTablePresentation() {
  const ppt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' });
  ppt.useSlide(1);

  // Add 4-row by 3-column table with custom styling
  ppt.addTable(1, 4, 3, [
    ['Department', 'Headcount', 'Budget'],
    ['Engineering', '58', '$6,200,000'],
    ['Product & Design', '22', '$2,400,000'],
    ['Marketing & Sales', '35', '$4,100,000']
  ], {
    x: '10%',
    y: '15%',
    width: '80%',
    headerFill: '1E293B',
    headerColor: 'FFFFFF',
    zebraFill: 'F8FAFC'
  });

  await ppt.saveToFile('./output/department_overview.pptx');
}
```

### 4. Express.js & Fastify Serverless Streaming
```javascript
const express = require('express');
const { PPTXTemplater } = require('node-pptx-templater');

const app = express();

app.get('/api/export-deck', async (req, res) => {
  const ppt = await PPTXTemplater.load('./templates/invoice.pptx');
  ppt.replaceText({ '{{CLIENT}}': req.query.client || 'Client' });
  
  // Streams directly with correct Content-Type and Content-Disposition headers
  await ppt.toExpress(res, 'Custom_Report.pptx');
});

// Fastify
fastify.get('/api/fastify-export', async (request, reply) => {
  const ppt = await PPTXTemplater.load('./templates/invoice.pptx');
  await ppt.toFastify(reply, 'Invoice.pptx');
});
```

### 5. Elements & Hyperlinks Integration
```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function buildInteractiveDeck() {
  const ppt = await PPTXTemplater.load('./templates/portfolio.pptx');
  ppt.useSlide(1);

  // Add clickable external hyperlink to a shape with ScreenTip
  ppt.addHyperlink('DocsButton', 'https://node-pptx-templater.pages.dev', {
    tooltip: 'Visit Official Documentation'
  });

  // Reorder visual layers safely
  ppt.bringToFront('DocsButton');

  // Substring hyperlink styling within paragraph runs
  ppt.setHyperlink('SummaryCard', 'API Documentation', 'https://node-pptx-templater.pages.dev/docs/api/', {
    color: '#0563C1',
    underline: true,
    bold: true
  });

  await ppt.saveToFile('./output/interactive_deck.pptx');
}
```

---

## 🎪 CLI Usage

```bash
# Build presentation from template with JSON data
node-pptx-templater build template.pptx output.pptx --data ./data.json

# Extract all OpenXML parts into a formatted directory for git inspection
node-pptx-templater extract deck.pptx --all --format --out ./extracted-deck

# Extract to clean ZIP archive
node-pptx-templater extract deck.pptx --zip ./deck.zip

# Validate presentation archive integrity
node-pptx-templater validate deck.pptx
```

---

## 🛠️ Verification & Quality Assurance

```bash
# Run Vitest test suites
npm test

# Check documentation coverage
npm run docs:validate

# Build documentation portal
npm run docs:build
```

---

## 📄 License & Resources

- **License:** [MIT](https://opensource.org/licenses/MIT)
- **Documentation Portal:** [https://node-pptx-templater.pages.dev/](https://node-pptx-templater.pages.dev/)
- **GitHub Repository:** [https://github.com/jsuyog2/node-pptx-templater](https://github.com/jsuyog2/node-pptx-templater)
- **Issue Tracker:** [https://github.com/jsuyog2/node-pptx-templater/issues](https://github.com/jsuyog2/node-pptx-templater/issues)
