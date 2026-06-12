<p align="center">
  <img src="https://raw.githubusercontent.com/jsuyog2/node-pptx-templater/main/assets/logo.png" alt="node-pptx-templater logo" width="160">
</p>

<h1 align="center">node-pptx-templater</h1>

<p align="center">
  <strong>High-performance OpenXML PowerPoint template engine for Node.js</strong>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/node-pptx-templater"><img src="https://img.shields.io/npm/v/node-pptx-templater.svg?style=flat-square&color=6366f1" alt="npm version"></a>
  <a href="https://github.com/jsuyog2/node-pptx-templater/actions/workflows/ci.yml"><img src="https://img.shields.io/github/actions/workflow/status/jsuyog2/node-pptx-templater/ci.yml?branch=main&style=flat-square&color=34d399" alt="CI"></a>
  <a href="https://www.npmjs.com/package/node-pptx-templater"><img src="https://img.shields.io/npm/dm/node-pptx-templater.svg?style=flat-square&color=a855f7" alt="Downloads"></a>
  <a href="./LICENSE"><img src="https://img.shields.io/badge/License-MIT-blue.svg?style=flat-square" alt="MIT License"></a>
  <a href="https://nodejs.org"><img src="https://img.shields.io/badge/node-%3E%3D18.0.0-brightgreen?style=flat-square" alt="Node.js 18+"></a>
</p>

<p align="center">
  <a href="https://jsuyog2.github.io/node-pptx-templater/"><strong>📖 Full Documentation →</strong></a>
</p>

---

Design your presentations visually in PowerPoint. Populate them dynamically in Node.js.

`node-pptx-templater` parses OpenXML directly — no native office dependencies, no Java, no Electron. Pure JavaScript that runs on AWS Lambda, Vercel Edge, and Cloudflare Workers.

---

## Installation

```bash
npm install node-pptx-templater
```

---

## Quick Start

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

const ppt = await PPTXTemplater.load('template.pptx');

// Replace text placeholders
ppt.useSlide(1)
   .replaceMultiple({ '{{title}}': 'Q2 Report', '{{date}}': 'June 2026' });

// Update chart data (syncs the embedded Excel workbook too)
ppt.useSlide(2)
   .updateChart('revenue-chart', {
     categories: ['April', 'May', 'June'],
     series: [{ name: 'Revenue', values: [125000, 162000, 195000] }]
   });

// Update table
ppt.useSlide(3)
   .updateTable('SalesTable', {
     rows: [['North', '$1.2M', '↑15%'], ['South', '$1.8M', '↑22%']]
   });

await ppt.saveToFile('output.pptx');
```

---

## Features

| Category | Capabilities |
|---|---|
| **Text** | Replace placeholders, search & replace, rich text, bullet lists |
| **Tables** | Update cells, add/insert/remove rows, merge cells, extract data as JSON |
| **Charts** | Update data, sync Excel workbook, custom data labels, series management |
| **Images** | Replace, add, remove images; set position, size, rotation, crop |
| **Shapes** | Add, update, clone, remove shapes; set fill, text, position |
| **Cell Shapes** | Add shapes inside table cells with automatic bounds calculation |
| **Slides** | Duplicate, move, delete, import/export, clone, reorder |
| **Layer Ordering** | Bring to front, send to back, set Z-index, swap, sort |
| **Hyperlinks** | URL links, slide navigation, shape and image links |
| **XML Folder** | Extract to folder, edit XML directly, rebuild to PPTX |
| **Validation** | Structural validation, relationship checks, auto-repair |
| **Performance** | Template caching, async streaming, profiling |

---

## Logging

The library is **completely silent by default**. Enable logging when needed:

```javascript
// At load time
const ppt = await PPTXTemplater.load('template.pptx', { logLevel: 'debug' });

// Globally (affects all instances)
PPTXTemplater.setLogLevel('debug');

// Instance shortcut
ppt.enableDebug();

// Suppress everything (explicit)
PPTXTemplater.setLogLevel('silent');
```

Supported levels: `verbose` | `debug` | `info` | `warn` (default) | `error` | `silent`

---

## Examples

The [`examples/`](./examples) directory contains runnable scripts for every major feature:

| Script | What it demonstrates |
|---|---|
| `basic-usage.js` | Text, charts, tables, links — end-to-end workflow |
| `chart-update.js` | Multi-series chart updates with Excel sync |
| `chart-data-labels.js` | Custom data label formatting and positioning |
| `table-update.js` | Cell updates, row management, merging |
| `slide-selection.js` | Slide operations — duplicate, move, delete |
| `image-operations.js` | Add, replace, remove images |
| `shape-operations.js` | Shape lifecycle — add, update, clone, delete |
| `z-order.js` | Layer ordering — bring/send, swap, set Z-index |
| `slide-import.js` | Import/export slides between presentations |
| `text-search.js` | Find text, replace multiple, update named shapes |
| `table-extraction.js` | Extract table data as JSON (all 3 modes) |
| `nested-table-rows.js` | Nested rows with rowspan / auto / none strategies |
| `xml-folder-workflow.js` | Extract → edit XML → rebuild PPTX |

Run any example:
```bash
npm run example:basic
npm run example:images
npm run example:shapes
# ... etc
```

---

## Core Concepts

### Template Design
Create your template in PowerPoint. Use `{{placeholder}}` tags in text boxes, name your shapes, tables, and charts clearly — the library finds them by name.

### Slide Targeting
```javascript
ppt.useSlide(2)         // Target slide 2
ppt.useAllSlides()      // Target all slides
```

### Chaining
Most methods return `this`, enabling fluent chains:
```javascript
ppt.useSlide(1)
   .replaceTextByTag('{{name}}', 'Alice')
   .updateShapePosition('logo', { x: 500000, y: 200000 })
   .bringToFront('logo');
```

### Table Data Extraction
```javascript
// Object array (header row = keys)
const rows = await ppt.getTableRows('SalesTable');
// → [{ region: 'North', sales: '1200' }, ...]

// Raw string arrays
const raw = await ppt.getTableRows('SalesTable', { raw: true });
// → [['North', '1200'], ...]

// With metadata
const meta = await ppt.getTableRows('SalesTable', { includeMetadata: true });
// → { rows: [...], rowCount: 5, columnCount: 3, mergedCells: [] }
```

### Table Cell Shapes
Cell shapes are overlay graphics anchored within table cells. They are positioned absolutely based on cell bounds, and **never** modify row heights, column widths, cell margins, or trigger table reflow. Offsets (`x`, `y`) are relative to the cell's top-left corner. Oversized shapes are scaled down proportionally to fit inside the cell.

```javascript
// Add a simple indicator
await ppt.addCellShape('SalesTable', 2, 1, {
  type: 'circle',
  width: 12,
  height: 12,
  fill: '#10B981' // Green status dot
});

// Add a badge with text and custom offsets
await ppt.addCellShape('SalesTable', 1, 2, {
  type: 'badge',
  text: 'Active',
  fill: '#3B82F6',
  x: 4,
  y: 2,
  width: 50,
  height: 16
});
```

### XML Folder Workflow
```javascript
// Extract PPTX to a folder of XML files (version-control friendly)
await PPTXTemplater.extractPptx('template.pptx', './pptx-xml/');

// Load, modify, rebuild
const ppt = await PPTXTemplater.load('./pptx-xml/');
ppt.useAllSlides().replaceTextByTag('{{env}}', 'Production');
await PPTXTemplater.buildPptx('./pptx-xml/', 'output.pptx');
```

---

## API Reference

Full API documentation with examples, parameter types, and advanced usage:

**👉 https://jsuyog2.github.io/node-pptx-templater/**

---

## Output Options

```javascript
await ppt.saveToFile('./output.pptx');      // File on disk
const buffer = await ppt.toBuffer();        // Node.js Buffer (for HTTP responses)
const stream = await ppt.toStream();        // Readable stream (for piping)
await ppt.saveToFolder('./output-xml/');    // Raw XML folder
```

---

## Caching & Performance

```javascript
// Preload templates for repeated generation (e.g., in a web server)
await PPTXTemplater.preload('./template.pptx');

// Load from cache — zero disk I/O on subsequent calls
const ppt = await PPTXTemplater.fromCache('./template.pptx');

// Profile performance
ppt.enablePerformanceProfile();
// ... do work ...
console.log(ppt.getPerformanceMetrics());
```

---

## Validation & Repair

```javascript
const result = await ppt.validatePresentation();
// → { valid: true, errors: [], warnings: [] }

await ppt.repair();         // Auto-fix orphan relationships
await ppt.repairCharts();   // Fix broken chart workbook references
```

---

## Migration from v1.0.x

No breaking changes. All existing APIs continue to work unchanged.

**New in v1.1.0:**
- `PPTXTemplater.load(path, { logLevel })` — configure logging at load time
- `PPTXTemplater.setLogLevel(level)` — global log level control
- `ppt.enableDebug()` — instance debug shortcut
- `ppt.getTableRows(tableId, options)` — table data extraction
- `ppt.addTableRow(tableId, data, { mergeStrategy })` — nested row support
- `setGlobalLogLevel` / `resetLogLevel` now exported from public API
- 8 new example files
- `npm run docs:validate` — documentation validation
- `npm run docs:inventory` — feature inventory generation

---

## Contributing

Contributions are welcome! Please read [CONTRIBUTING.md](./CONTRIBUTING.md) and open a pull request.

```bash
git clone https://github.com/jsuyog2/node-pptx-templater.git
cd node-pptx-templater
npm install
npm test
```

---

## License

[MIT](./LICENSE) © node-pptx-templater contributors
