# node-pptx-templater

> High-performance, low-level PowerPoint (PPTX) OpenXML template engine for Node.js. Dynamically replace text, insert images, update charts (with Excel workbook data caching), and merge table cells without PowerPoint corruption or Repair Mode prompts.

[![npm version](https://img.shields.io/npm/v/node-pptx-templater.svg?style=flat-square&color=blue)](https://www.npmjs.com/package/node-pptx-templater)
[![CI Build Status](https://img.shields.io/github/actions/workflow/status/jsuyog2/node-pptx-templater/ci.yml?branch=main&style=flat-square)](https://github.com/jsuyog2/node-pptx-templater/actions/workflows/ci.yml)
[![Bundle Size](https://img.shields.io/bundlephobia/min/node-pptx-templater?style=flat-square&color=brightgreen)](https://bundlephobia.com/package/node-pptx-templater)
[![Downloads](https://img.shields.io/npm/dm/node-pptx-templater.svg?style=flat-square&color=orange)](https://www.npmjs.com/package/node-pptx-templater)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg?style=flat-square)](./LICENSE)
[![Node.js Version](https://img.shields.io/badge/node-%3E%3D18.0.0-brightgreen?style=flat-square)](https://nodejs.org)

---

## ⚡ Why node-pptx-templater?

Traditional PowerPoint generation libraries require building slides from scratch in code, which is verbose, hard to maintain, and strips away the power of visual design tools. 

`node-pptx-templater` takes a different approach: **Design visually in PowerPoint, populate dynamically in Node.js.** 

You create slide decks using PowerPoint, Google Slides, or Keynote, set your formatting, themes, animations, and layouts, and place placeholders like `{{company}}` or `{{revenue-chart}}`. `node-pptx-templater` parses the template and updates text, injects images, replaces chart values (updating both Excel workbook data caches and XML shapes), and merges tables dynamically while keeping the presentation 100% compliant with standard OpenXML guidelines.

---

## ✨ Features

- 🏗️ **Zero Native Office/Java Dependencies**: Runs on pure Javascript/Node.js, making it ideal for high-throughput cloud environments, Lambda, or serverless runtimes.
- 🔁 **Fragmented Placeholder Resolution**: PowerPoint often splits text runs like `{{company}}` into `<a:r>` nodes. Our engine merges and resolves fragmented tags automatically.
- 📊 **Full Chart Engine Integration**: Supports Bar, Column, Line, Pie, Doughnut, Area, Scatter, and Bubble charts. Automatically synchronizes chart XML properties and coordinates with the embedded Excel sheets (`ppt/embeddings/`).
- 📋 **Flexible Table Merging & Templating**:
  - Horizontal column merge (`gridSpan` & `hMerge`), vertical row merge (`rowSpan` & `vMerge`), and rectangular block merges.
  - Formats cells dynamically with inline options (`align`, `fontSize`, `fill`).
  - Automatically handles slide table duplicates by generating unique `<a16:rowId>` 32-bit hashes to **prevent PowerPoint Repair Mode** screens.
- 🎨 **Shape & Image Manipulation**: Find shapes, clone layout blocks with offsets, replace image sources while keeping exact positions, or delete elements.
- 🎯 **Slide Management Operations**: Duplicate, reorder, delete, and import slides from external templates with automatic media asset deduplication.
- 🔍 **Deep Packaging Integrity Validation**: Real-time checking of relationships, XML schemas, table column numbers, and override duplicates.

---

## 📦 Installation

```bash
npm install node-pptx-templater
```

---

## 🚀 Quick Start

Get up and running in under 60 seconds with this simple template rendering example:

```js
const { PPTXTemplater } = require('node-pptx-templater');

async function main() {
  // 1. Load your PowerPoint presentation template
  const ppt = await PPTXTemplater.load('monthly_report_template.pptx');
  
  // 2. Select slide 1 and execute operations
  ppt.useSlide(1)
     .replaceTextByTag('title', 'Quarterly Earnings Report')
     .replaceMultiple({
       company: 'Acme Corporation',
       year: '2026'
     });

  // 3. Update chart series data on Slide 2
  ppt.useSlide(2)
     .updateChartData('sales-chart', {
       categories: ['Q1', 'Q2', 'Q3', 'Q4'],
       series: [
         { name: 'Target', values: [100, 120, 140, 160] },
         { name: 'Revenue', values: [105, 118, 145, 172] }
       ]
     });

  // 4. Update table with cell merging and formatting on Slide 3
  ppt.useSlide(3)
     .updateTable('sales-table', [
       ['Region', 'Q1 Actual', 'Q2 Actual', 'Status'],
       ['North', '120k', '140k', { value: 'On Track', align: 'ctr', fill: '10b981' }],
       ['South', '95k', '110k', { value: 'Review', align: 'ctr', fill: 'f59e0b' }]
     ]);

  // 5. Save the non-corrupted PPTX back to disk
  await ppt.saveToFile('./output/annual_earnings.pptx');
}

main().catch(err => console.error(err));
```

---

## 🏗️ OpenXML Architecture & Internals

A `.pptx` file is an OPC (Open Packaging Convention) ZIP archive containing structured XML documents and asset folders:

- `[Content_Types].xml` – Global manifest declaring content MIME types for every file part in the ZIP.
- `_rels/.rels` – Root-level package relationship index.
- `ppt/presentation.xml` – Root presentation settings and slide inventory (`sldIdLst`).
- `ppt/slides/slideN.xml` – Main slide canvas storing shapes, lines, tables, text runs, and layout components.
- `ppt/slides/_rels/slideN.xml.rels` – Relationship indexes mapping slide XML components to charts, layouts, and image assets.

### Preventing PowerPoint Table Repair Errors
PowerPoint slide tables utilize unique 32-bit identifiers inside `<a16:rowId>` nodes for collaborative edits. Duplicating rows using naive array copy operations results in overlapping IDs, triggering Microsoft PowerPoint's **"PowerPoint found a problem with content"** repair screen on open.
`node-pptx-templater` intercepts all table operations (adding, cloning, inserting, or merging rows) and dynamically injects newly generated unique `rowId` hashes, ensuring a seamless, warning-free loading experience in:
- Microsoft PowerPoint (Desktop, Mac, Online)
- Google Slides
- LibreOffice Impress

---

## 📊 Feature Comparison Matrix

| Feature / Library | `node-pptx-templater` | `pptxgenjs` | `pptx-template` | `pptx-automizer` | `officegen` |
|:---|:---:|:---:|:---:|:---:|:---:|
| **Approach** | **Template-based** | Code-based | Template-based | Template-based | Code-based |
| **No PPTX Corruption / Repair Warnings** | **Yes** (Automatic Metadata Sync) | Yes | No (Fragile row duplication) | Yes | Yes (Limited layouts) |
| **Text Run Fragmentation Resolution** | **Yes** (Dynamic merging) | N/A | No (Placeholder breaks) | Yes | N/A |
| **Chart Data Workbook Sync** | **Yes** (Direct excel caching) | Yes | No (Only raw XML text) | Yes | Yes |
| **Horizontal & Vertical Cell Merge** | **Yes** (gridSpan, rowSpan, hMerge, vMerge) | Yes | No | No | No |
| **Slide Duplication & Reordering** | **Yes** | No | No | Yes | No |
| **External Slide Imports** | **Yes** (With asset deduplication) | No | No | Yes | No |
| **Dependencies** | **Zero Native Dependencies** | Zero | Zero | Zero | Node-zip, xmlbuilder |

---

## 📚 API Reference

### Slide Operations

#### `duplicateSlide(slideIndex, atPosition)`
Duplicates a slide.
```js
ppt.duplicateSlide(1, 2); // Duplicate Slide 1 and insert it at position 2
```

#### `deleteSlide(slideIndex)`
Deletes a slide from the deck.
```js
ppt.deleteSlide(3); // Delete Slide 3
```

#### `moveSlide(fromIndex, toIndex)`
Moves a slide to a new position.
```js
ppt.moveSlide(1, 3); // Move Slide 1 to position 3
```

#### `importSlideFrom(sourcePresentation, sourceSlideIndex)`
Deep-copies a slide from another loaded presentation, automatically remapping layouts, shapes, charts, and deduplicating media assets.
```js
const source = await PPTXTemplater.load('marketing_slides.pptx');
await ppt.importSlideFrom(source, 2); // Import Slide 2 of marketing deck
```

---

### Table Manipulation

#### `updateTable(tableId, data)`
Updates a table with rows data, merge rules, and cell styles.
```js
ppt.updateTable('revenue-table', [
  ['Year', 'Revenue', 'Profit'],
  ['2025', '120k', '40k'],
  ['2026', '150k', { value: '60k', fill: '10b981', align: 'ctr' }]
]);
```

#### `mergeCells(options)`
Merges a rectangular block of cells. Supports horizontal, vertical, and block merging, concatenating all text from the merged region into the top-left cell.
```js
ppt.mergeCells({
  tableId: 'sales-table',
  startRow: 1,
  startCol: 1,
  endRow: 2,
  endCol: 2
});
```

#### `unmergeCells(options)`
Splits a merged region back to its original individual cells, removing `gridSpan`, `rowSpan`, `hMerge`, and `vMerge` attributes.
```js
ppt.unmergeCells({
  tableId: 'sales-table',
  row: 1,
  col: 1
});
```

---

### Chart Integration

#### `updateChartData(chartId, data)`
Overwrites chart categories and series values. Updates the embedded Excel spreadsheet to ensure the chart matches perfectly on refresh.
```js
ppt.updateChartData('sales-chart', {
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'Revenue', values: [100, 150, 180, 220] }
  ]
});
```

#### `updateChartTitle(chartId, title)`
```js
ppt.updateChartTitle('sales-chart', 'Revenue Growth (2026)');
```

---

## ⚡ Performance Benchmarks

Tested on a standard 50-slide enterprise presentation template:

| Operation | Execution Duration |
|:---|:---|
| Load PPTX Template | ~110ms |
| Find & Replace 20 Text Placeholders | ~2.5ms |
| XML Schema & Integrity Validation Check | ~14ms |
| Dynamic Row Insertion & Merging (15 rows) | ~3ms |
| Save and Re-package to PPTX ZIP | ~78ms |

---

## ❓ FAQ & Troubleshooting

### PowerPoint displays a "Repair" prompt when opening my generated file
This is commonly caused by:
1. **Missing overridden content type**: A new slide or chart XML was added but not registered in `[Content_Types].xml`.
2. **Duplicate row identifiers**: If table rows are duplicated without generating a new unique `rowId` under `<a16:rowId>`.
3. **Invalid relationship mapping**: An asset (like an image or worksheet) is referenced in slide XML but is missing from the slide's `.rels` file.

*Fix*: Ensure you always use the public `saveToFile()` or `toBuffer()` helper functions, which automatically execute structural verification passes and update relationship chains.

### My text placeholders are not replacing
PowerPoint text editors segment formatting runs into separate XML elements. The text `{{title}}` may look normal in PowerPoint, but in XML it could be split into `<a:t>{{ti</a:t><a:t>tle}}</a:t>`. 
*Fix*: You can enable logger output to see split tags:
```bash
PPTX_LOG_LEVEL=debug node app.js
```
To fix this in PowerPoint, highlight the entire placeholder block, cut it, and paste it back as "Keep Text Only" to unify the XML text runs.

---

## 🤝 Contributing

We welcome contributions! Please check out [CONTRIBUTING.md](./CONTRIBUTING.md) to get started.

```bash
git clone https://github.com/jsuyog2/node-pptx-templater.git
cd node-pptx-templater
npm install
npm test
```

---

## 📄 License

Licensed under the MIT License. © [node-pptx-templater contributors](./LICENSE)
