# node-pptx-templater

<div align="center">

[![npm version](https://img.shields.io/npm/v/node-pptx-templater.svg?style=flat-square&color=blue)](https://www.npmjs.com/package/node-pptx-templater)
[![npm downloads](https://img.shields.io/npm/dm/node-pptx-templater.svg?style=flat-square&color=blue)](https://www.npmjs.com/package/node-pptx-templater)
[![Vitest Tests](https://img.shields.io/badge/tests-285%20passing-brightgreen.svg?style=flat-square)](https://github.com/jsuyog2/node-pptx-templater/actions)
[![Bundle Size](https://img.shields.io/badge/bundle%20size-pure%20JS%20%3C2MB-success.svg?style=flat-square)](https://www.npmjs.com/package/node-pptx-templater)
[![Serverless](https://img.shields.io/badge/serverless-AWS%20Lambda%20%7C%20Cloudflare%20Workers-purple.svg?style=flat-square)](https://node-pptx-templater.pages.dev/guides/serverless-deployment/)
[![Node.js Support](https://img.shields.io/badge/node-%3E%3D20.12.0-blue.svg?style=flat-square)](https://nodejs.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square)](https://opensource.org/licenses/MIT)

**Pure JavaScript OpenXML template engine for Node.js.**  
**Serverless PPTX automation for AWS Lambda & Cloudflare Workers.**

[📖 Documentation Portal](https://node-pptx-templater.pages.dev/) &nbsp;•&nbsp; [⚡ 5-Minute Quick Start](https://node-pptx-templater.pages.dev/guides/quick-start/) &nbsp;•&nbsp; [📚 API Reference (101 Methods)](https://node-pptx-templater.pages.dev/api-reference/) &nbsp;•&nbsp; [⚡ Serverless Deployment](https://node-pptx-templater.pages.dev/guides/serverless-deployment/)

</div>

---

## 🚀 Overview

**`node-pptx-templater`** is a high-performance, spatial-aware PowerPoint (PPTX) OpenXML manipulation and templating engine built specifically for **Node.js (>=20.12.0)**, **AWS Lambda**, **Cloudflare Workers**, and **Vercel Edge**.

Unlike traditional tools that force you to spin up heavyweight headless browsers, LibreOffice Docker containers (500MB+ image, 15-second cold starts), or Java runtimes, `node-pptx-templater` is **100% pure JavaScript** with **zero native binary dependencies**. It mutates OpenXML presentation packages entirely in memory without corrupting PresentationML relationships, layouts, or formatting.

### 🌟 Architectural Highlights

- **Anti-Corruption Run Fragment Healing:** PowerPoint frequently splits placeholder tokens like `{{COMPANY_NAME}}` into disjoint `<a:r>` XML run nodes. Our AST paragraph normalizer merges split runs while preserving font styles, weights, colors, and layout intact.
- **Excel Dual-Sync Technology:** Updating only XML chart series causes Microsoft Office to trigger "Repair Presentation" warnings when a user clicks "Edit Data in Excel". `node-pptx-templater` updates both DrawingML `<c:chartSpace>` series and embedded binary Excel workbooks (`.xlsx`) simultaneously.
- **Dynamic Table Engine & Overflow Pagination:** Append rows, alternate row striping, merge cells vertically (`rowspan`) and horizontally (`gridspan`), and automatically paginate 50+ database rows across newly cloned slides with repeated headers via `splitTableOnOverflow()`.
- **Spatial Geometry & Z-Order Layering:** Reorder visual stack depth (`bringToFront`, `sendToBack`, `stepForward`, `stepBackward`) and calculate coordinates using English Metric Units (EMU), pixels, inches, and millimeters with zero floating-point rounding drift.
- **Serverless & Edge Native:** Instant sub-15ms cold start, memory footprint under 2MB, with zero disk I/O requirements (`toBuffer()`, `toArrayBuffer()`, and Express stream piping).

---

## 📊 Feature Comparison Matrix

| Capability / Feature | `node-pptx-templater` | `pptxgenjs` | `officegen` | `LibreOffice Headless` |
| :--- | :---: | :---: | :---: | :---: |
| **Existing PPTX Template Ingestion** | ✅ **Full In-Place Mutation** | ❌ Blank only | ❌ Blank only | ⚠️ Macro scripting |
| **Run Fragment Normalizer (Anti-Corruption)** | ✅ **AST Paragraph Healer** | ❌ N/A | ❌ N/A | ❌ Brittle regex |
| **Chart & Excel Workbook Dual-Sync** | ✅ **XML + XLSX Cache Sync** | ⚠️ Basic values only | ❌ Unsupported | ⚠️ Slow headless calc |
| **Table Overflow Auto-Pagination** | ✅ **`splitTableOnOverflow()`** | ❌ Manual math | ❌ Unsupported | ❌ Unsupported |
| **Z-Order Layer Stacking (`p:spTree`)** | ✅ **`bringToFront` / `sendToBack`** | ❌ Re-render required | ❌ Unsupported | ❌ Manual scripts |
| **Serverless Deployment (Lambda, Cloudflare)** | ✅ **<2MB Pure JS, <15ms Boot** | ✅ Pure JS | ⚠️ High RAM | ❌ Heavy 500MB+ Docker |
| **Node.js & TypeScript ESM/CJS Support** | ✅ **Dual ESM/CJS + `.d.ts`** | ✅ Supported | ⚠️ Legacy CJS | ❌ CLI process spawn |

---

## 📦 Installation

```bash
# Using npm
npm install node-pptx-templater

# Using pnpm
pnpm add node-pptx-templater

# Using yarn
yarn add node-pptx-templater
```

**Requirements:** Node.js `>= 20.12.0`. First-class support for both CommonJS (`require`) and ESM (`import`).

---

## ⚡ Quick-Copy Production Examples

### 1. Dynamic Text Replacement & Fragment Healing

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function updateText() {
  const ppt = await PPTXTemplater.load('./templates/executive-deck.pptx');

  // Replaces tokens even if split across multiple <a:r> runs by PowerPoint
  ppt.useSlide(1).replaceText({
    '{{COMPANY}}': 'Acme Global Technologies Inc.',
    '{{QUARTER}}': 'Q3 2026',
    '{{PRESENTER}}': 'Sarah Jenkins, VP Engineering'
  });

  await ppt.saveToFile('./output/updated-deck.pptx');
}
```

### 2. Native Chart Update with Excel Dual-Sync

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function updateMetrics() {
  const ppt = await PPTXTemplater.load('./templates/financial-report.pptx');

  // Synchronizes XML <c:chartSpace> AND embedded .xlsx workbook
  ppt.useSlide(2);
  await ppt.updateChart('QuarterlyRevenueChart', {
    categories: ['Q1', 'Q2', 'Q3', 'Q4 (Proj)'],
    series: [
      { name: 'Enterprise Cloud', values: [18.4, 22.1, 26.8, 31.5] },
      { name: 'AI Solutions', values: [8.2, 12.6, 17.4, 23.0] }
    ]
  });

  await ppt.saveToFile('./output/synced-charts.pptx');
}
```

### 3. Dynamic Tables, Row Merging & Overflow Auto-Pagination

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function generateTable() {
  const ppt = await PPTXTemplater.load('./templates/inventory.pptx');
  ppt.useSlide(3);

  // Append database rows
  for (const item of inventoryRecords) {
    ppt.addTableRow('InventoryTable', [item.sku, item.name, item.warehouse, item.qty]);
  }

  // Alternate row background colors
  ppt.stripeRows('InventoryTable', '#F8FAFC', '#FFFFFF');

  // Automatically split table across cloned slides if it overflows slide height
  await ppt.splitTableOnOverflow('InventoryTable', {
    bottomMargin: '1in',
    repeatHeaders: 1
  });

  await ppt.saveToFile('./output/paginated-table.pptx');
}
```

### 4. Shape Injection, Styling & Z-Order Layering

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function styleShapes() {
  const ppt = await PPTXTemplater.load('./templates/status-deck.pptx');
  ppt.useSlide(1);

  // Add geometric card
  await ppt.addShape('roundRect', {
    x: '5%', y: '10%', width: '90%', height: '20%',
    fill: '1E293B', text: 'Operational Status: ALL SYSTEMS NOMINAL'
  });

  // Adjust layer depth in visual tree
  ppt.bringToFront('CompanyBadge');
  ppt.sendToBack('BackgroundGradient');

  await ppt.saveToFile('./output/styled-layers.pptx');
}
```

### 5. Serverless In-Memory Export (AWS Lambda & Cloudflare Workers)

```javascript
import { PPTXTemplater } from 'node-pptx-templater';

// AWS Lambda / Cloudflare Workers handler
export async function handler(event) {
  const ppt = await PPTXTemplater.load(event.templateBuffer);

  ppt.useSlide('all').replaceText(event.replacements);

  // Export directly to Buffer / ArrayBuffer without filesystem disk access
  const buffer = await ppt.toBuffer();

  return {
    statusCode: 200,
    headers: {
      'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'Content-Disposition': 'attachment; filename="GeneratedDeck.pptx"'
    },
    body: buffer.toString('base64'),
    isBase64Encoded: true
  };
}
```

---

## 🎪 Unified Showcase Suite (12 Demonstrations)

All 12 supported demonstrations and sample templates are runnable via a single command:

```bash
# Run all 12 showcase demonstrations
npm run examples
# or
node showcase/examples.js
```

| Demo | Focus Area | Demonstrated APIs |
| :---: | :--- | :--- |
| `01` | **Basic Report & 16:9** | `setSlideSize('16:9')`, `replaceMultiple()`, `findText()` |
| `02` | **Chart & Excel Dual-Sync** | `updateChart()`, `updateDataLabels()` |
| `03` | **Tables, Striping & Overflow** | `updateTable()`, `addTableRow()`, `stripeRows()`, `splitTableOnOverflow()` |
| `04` | **Nested Tables & Extraction** | `addTableRow({mergeStrategy:'rowspan'})`, `getTableRows()` |
| `05` | **Shapes, Z-Order & Alignment** | `addShape()`, `bringToFront()`, `alignShapes()` |
| `06` | **Cell Shapes & Badges** | `addTableRow([{type:'circle', fill:'#10B981'}, ...])` |
| `07` | **Images: PNG, WebP & Base64** | `addImage()`, `replaceImage()` |
| `08` | **Slide Cloning & Import** | `duplicateSlide()`, `reorderSlides()`, `importSlideFrom()` |
| `09` | **Hyperlinks & Navigation** | `addHyperlink()`, `addSlideLink()`, `getHyperlinks()` |
| `10` | **Typography & Text Runs** | `addShape(textStyle)`, `appendTextRun()` |
| `11` | **OpenXML Folder Workflow** | `extractPptx()`, `saveToFolder()`, `buildPptx()` |
| `12` | **Serverless Buffer Export** | `toBuffer()`, `toArrayBuffer()` (zero disk I/O) |

---

## 🛠️ Verification & Testing

```bash
# Run complete test suite (285 passing Vitest tests)
npm test

# Run validation and showcase integration tests
npm run validate

# Build modern Cloudflare Pages documentation application
npm run docs:build
```

---

## 📄 License & Community

- **License:** [MIT License](https://opensource.org/licenses/MIT)
- **Documentation:** [https://node-pptx-templater.pages.dev/](https://node-pptx-templater.pages.dev/)
- **Repository:** [https://github.com/jsuyog2/node-pptx-templater](https://github.com/jsuyog2/node-pptx-templater)
- **Issues & Contributions:** [https://github.com/jsuyog2/node-pptx-templater/issues](https://github.com/jsuyog2/node-pptx-templater/issues)
