<p align="center">
  <img src="https://raw.githubusercontent.com/jsuyog2/node-pptx-templater/main/docs/logo.svg" alt="PPTXForge Logo" width="550" max-width="100%">
</p>

<h1 align="center">node-pptx-templater (PPTXForge Engine)</h1>

<p align="center">
  <strong>The High-Performance, Secure OpenXML PowerPoint Template Engine for Node.js</strong>
</p>

<p align="center">
  <a href="https://www.npmjs.com/package/node-pptx-templater"><img src="https://img.shields.io/npm/v/node-pptx-templater.svg?style=flat-square&color=6366f1" alt="npm version"></a>
  <a href="https://github.com/jsuyog2/node-pptx-templater/actions/workflows/ci.yml"><img src="https://img.shields.io/github/actions/workflow/status/jsuyog2/node-pptx-templater/ci.yml?branch=main&style=flat-square&color=34d399" alt="CI Build Status"></a>
  <a href="https://bundlephobia.com/package/node-pptx-templater"><img src="https://img.shields.io/bundlephobia/min/node-pptx-templater?style=flat-square&color=ec4899" alt="Bundle Size"></a>
  <a href="https://www.npmjs.com/package/node-pptx-templater"><img src="https://img.shields.io/npm/dm/node-pptx-templater.svg?style=flat-square&color=a855f7" alt="Downloads"></a>
  <a href="./LICENSE"><img src="https://img.shields.io/badge/License-MIT-blue.svg?style=flat-square" alt="License: MIT"></a>
  <a href="https://nodejs.org"><img src="https://img.shields.io/badge/node-%3E%3D18.0.0-brightgreen?style=flat-square" alt="Node.js Version"></a>
</p>

---

## ⚡ Why PPTXForge (`node-pptx-templater`)?

Traditional slide deck automation requires assembling slides shape-by-shape in verbose code blocks. This is fragile, difficult to maintain, and strips away the power of visual design tools like Microsoft PowerPoint, Keynote, or Google Slides.

**PPTXForge takes a visual-first approach: Design visually in PowerPoint, populate dynamically in Node.js.**

You design your layouts, tables, fonts, brand styles, charts, and animations inside PowerPoint, and insert placeholders like `{{customer_name}}`, `{{revenue_chart}}`, or `{{team_table}}`. The PPTXForge engine parses the OpenXML presentation, heals broken text run segmentations, updates Excel data workbook caches, merges drawing cells, and scales slides securely in pure JavaScript with zero external office dependencies.

---

## ✨ Enterprise-Grade Core Features

* 🚀 **Zero Native/Office/Java Dependencies**: Pure CommonJS JavaScript implementation. Runs efficiently on AWS Lambda, Vercel Edge, Netlify, and Cloudflare Workers.
* 🛡️ **Hardened XML Security**: Fully immune to Billion Laughs (XML bombs), XXE (XML External Entity Injection), and parser crashes due to oversized expansions.
* 🧩 **Text Run Fragmentation Healing**: Solves the split-tag issue. PowerPoint splits `{{placeholders}}` into fragmented `<a:r>` nodes; our parser unifies and replaces them while keeping original formatting intact.
* 📊 **Excel Cache Synchronized Charting**: Supports Bar, Line, Pie, Doughnut, Area, and Scatter charts. Not only updates the visual XML coordinates but also synchronizes data points inside the underlying embedded Excel spreadsheets (`ppt/embeddings/`) to bypass PowerPoint's "Update Data" warnings.
* 📋 **DrawingML Table Cell Merging**: Easily configure horizontal spans (`gridSpan`/`hMerge`), vertical spans (`rowSpan`/`vMerge`), and rectangular blocks. Injects unique `rowId` hashes to maintain relationship integrity.
* 🥞 **Z-Order Layer Stacking**: Reorder shapes, images, charts, and tables programmatically. Simulates PowerPoint's "Bring to Front" and "Send to Back" commands directly in the slide's `<p:spTree>`.
* 🎛️ **Slide Management**: Duplicate, delete, reorder slides, or import slides from external decks with automatic media and theme deduplication.
* 🔍 **OPC Package Verification**: Comprehensive check suite for content type overrides, relationship integrity, and XML structure before final output.

---

## 📦 Installation

```bash
npm install node-pptx-templater
```

---

## 🚀 Quick Start Guide

Generate a formatted presentation report from a visually designed template in under 60 seconds:

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function generateReport() {
  // 1. Load the presentation template
  const ppt = await PPTXTemplater.load('monthly_report_template.pptx');
  
  // 2. Select slide 1 and execute standard text replacements
  ppt.useSlide(1)
     .replaceTextByTag('title', 'Q2 Business Overview')
     .replaceMultiple({
       client: 'Acme Corporation',
       analyst: 'Sarah Jenkins',
       date: 'June 2026'
     });

  // 3. Update Excel-backed chart data on Slide 2
  ppt.useSlide(2)
     .updateChartData('revenue-chart', {
       categories: ['April', 'May', 'June'],
       series: [
         { name: 'Target', values: [120000, 150000, 180000] },
         { name: 'Actual', values: [125000, 162000, 195000] }
       ]
     });

  // 4. Update table structure with cell merges and colors on Slide 3
  ppt.useSlide(3)
     .updateTable('performance-table', [
       ['Region', 'Growth Metric', { value: 'Status / Notes', colSpan: 2 }],
       ['North Region', '+12.4%', { value: 'Exceeded Target', align: 'ctr', fill: '10b981' }, ''],
       ['South Region', '-3.1%', { value: 'Under Review', align: 'ctr', fill: 'f59e0b' }, '']
     ]);

  // 5. Duplicate a layout and bring an overlay element to the front
  ppt.duplicateSlide(3, 4);
  ppt.useSlide(4)
     .bringToFront('OverlayLogo')
     .replaceTextByTag('title', 'Duplicate Region Report');

  // 6. Save package without PowerPoint Repair Warnings
  await ppt.saveToFile('./output/q2_business_report.pptx');
  console.log('Report generated successfully!');
}

generateReport().catch(console.error);
```

---

## 📋 OpenXML Presentation Architecture

A `.pptx` file is an Open Packaging Convention (OPC) ZIP package containing structured XML schemas:

```text
PPTX Archive Structure:
├── [Content_Types].xml       # MIME type registry for all zip contents
├── _rels/
│   └── .rels                 # Package-level relationship mappings
└── ppt/
    ├── presentation.xml      # Global presentation slide order and layouts
    ├── slides/
    │   ├── slide1.xml        # DrawingML elements, text runs, and shapes
    │   └── _rels/
    │       └── slide1.xml.rels  # Resource mappings (images, charts, tables)
    ├── media/                # Image and SVG database
    └── embeddings/           # Embedded Excel books backing charts
```

### Unifying Text Run Segmentations
Under the hood, slide template text is represented as runs of text (`<a:r>`). PowerPoint's parser often breaks a single tag `{{name}}` into separate segments:
```xml
<!-- Split layout generated by PowerPoint -->
<a:r><a:t>{{n</a:t></a:r>
<a:r><a:t>ame}}</a:t></a:r>
```
PPTXForge parses the XML structure, identifies layout boundaries, merges text nodes back into a single element, and applies replacements while preserving font sizes, colors, and styling rules.

### Preventing PPTX "Repair Presentation" Warnings
Naively duplicating rows in slide tables can leave duplicate `rowId` values or break cell mappings, causing PowerPoint to crash or prompt a repair screen. PPTXForge automatically re-allocates unique `rowId` hashes and formats `gridSpan` / `rowSpan` coordinates in compliance with standard Office OpenXML (OOXML) regulations.

---

## 📊 Feature Comparison Matrix

Compare PPTXForge with other popular PowerPoint automation libraries:

| Feature / Metric | **node-pptx-templater** (PPTXForge) | **pptxgenjs** | **pptx-template** | **pptx-automizer** | **officegen** |
| :--- | :---: | :---: | :---: | :---: | :---: |
| **Automation Flow** | **Template-based** | Code-based | Template-based | Template-based | Code-based |
| **No Office/Java Dependency** | **Yes** (Pure JS) | Yes | Yes | Yes | Yes |
| **Bypass Corruption/Repair Alerts** | **Yes** (RowId sync) | Yes | No (Row duplicate breaks) | Yes | Yes |
| **Heal Text Run Fragmentation** | **Yes** | N/A | No (Tags break) | Yes | N/A |
| **Synchronized Excel Chart Updates** | **Yes** (Direct sync) | Yes | No (Text updates only) | Yes | Yes |
| **Horizontal & Vertical Cell Merge** | **Yes** (gridSpan/rowSpan) | Yes | No | No | No |
| **Z-Order Layer Reordering** | **Yes** (Front/Back) | No | No | Yes | No |
| **External Slide Imports** | **Yes** (Deduplicated) | No | No | Yes | No |

---

## 📚 API Reference Guide

### Loading & Exposing presentation

#### `static async load(filePathOrBuffer)`
Loads and extracts a presentation template from disk or memory.
* **Arguments**: `string | Buffer`
* **Returns**: `Promise<PPTXTemplater>`

#### `async saveToFile(outputPath)`
Validates, re-packages, and saves the presentation to disk.
* **Arguments**: `string`
* **Returns**: `Promise<void>`

#### `async toBuffer()`
Generates the final presentation files as a zipped binary buffer.
* **Returns**: `Promise<Buffer>`

---

### Slide Layout Actions

#### `useSlide(slideIndex)`
Fluent selector that targets a slide for subsequent actions (1-based index).
* **Arguments**: `number`
* **Returns**: `PPTXTemplater`

#### `duplicateSlide(slideIndex, insertPosition)`
Duplicates a slide and places it at the target position.
* **Arguments**: `number` (source index), `number` (destination index)
* **Returns**: `PPTXTemplater`

#### `deleteSlide(slideIndex)`
Deletes a slide and removes its associated relationship linkages.
* **Arguments**: `number`
* **Returns**: `PPTXTemplater`

#### `moveSlide(fromIndex, toIndex)`
Reorders a slide to a new positions index.
* **Arguments**: `number`, `number`
* **Returns**: `PPTXTemplater`

#### `importSlideFrom(sourcePresentation, sourceSlideIndex)`
Imports a slide from an external PowerPoint file, remapping media structures.
* **Arguments**: `PPTXTemplater`, `number`
* **Returns**: `Promise<PPTXTemplater>`

---

### Text Manipulation

#### `replaceTextByTag(tag, replacement)`
Replaces all occurrences of `{{tag}}` on the active slide.
* **Arguments**: `string` (tag name without braces), `string` (replacement text)
* **Returns**: `PPTXTemplater`

#### `replaceMultiple(replacementsMap)`
Performs bulk replacements on the active slide using a key-value map.
* **Arguments**: `Object` (e.g., `{ company: 'Acme', date: '2026' }`)
* **Returns**: `PPTXTemplater`

---

### Tables (DrawingML)

#### `updateTable(tableId, rowData)`
Updates a table with a two-dimensional grid array.
* **Arguments**: `string` (table shape name/id), `Array<Array<string | Object>>`
* **Cell Styles Options**:
  ```javascript
  {
    value: 'Text content',
    colSpan: 2,           // Merge columns horizontally
    rowSpan: 3,           // Merge rows vertically
    align: 'ctr' | 'l' | 'r',  // Text alignment
    fontSize: 1400,       // Font size in hundredths of a point
    fill: 'ff5500',       // Background hex color
    bold: true
  }
  ```
* **Returns**: `PPTXTemplater`

#### `mergeCells({ tableId, startRow, startCol, endRow, endCol })`
Merges a custom rectangular boundary of cells.
* **Arguments**: `Object`
* **Returns**: `PPTXTemplater`

#### `unmergeCells({ tableId, row, col })`
Splits a merged region back to its individual component cells.
* **Arguments**: `Object`
* **Returns**: `PPTXTemplater`

---

### Excel Chart Updates

#### `updateChartData(chartId, data)`
Overwrites chart categories and series points.
* **Arguments**: `string` (chart object id), `Object`
* **Format**:
  ```javascript
  {
    categories: ['Category 1', 'Category 2'],
    series: [
      { name: 'Series 1', values: [10, 20] }
    ]
  }
  ```
* **Returns**: `PPTXTemplater`

#### `updateChartTitle(chartId, title)`
Replaces the main header title text run of the chart.
* **Arguments**: `string`, `string`
* **Returns**: `PPTXTemplater`

---

### Stacking Z-Order & Layers

#### `getObjectOrder(slideIndex)`
Lists slide layout objects sorted from bottom-most (first in DOM) to top-most.
* **Returns**: `Array<{ id: string, type: 'shape'|'image'|'chart'|'table'|'group'|'connector', zIndex: number }>`

#### `bringToFront(objectId)` / `sendToBack(objectId)`
Moves the target object to the absolute front or back of the rendering stack.
* **Returns**: `PPTXTemplater`

#### `bringForward(objectId)` / `sendBackward(objectId)`
Shifts the target object up or down by one layer index.
* **Returns**: `PPTXTemplater`

#### `setZIndex(objectId, zIndex)`
Moves the object to a specific 1-based stack index.
* **Returns**: `PPTXTemplater`

---

## 🔒 Advanced XML Security

PPTXForge uses a secure XML parser that defends your application servers against common XML vulnerabilities.

### Built-in Protections
1. **XML Bombs & Billion Laughs Mitigation**: Instantly rejects any XML payload containing `<!DOCTYPE>` or `<!ENTITY>` definitions, preventing parser lockups.
2. **XXE Injection Defenses**: Completely blocks external entities referencing local files or external resources (e.g., `SYSTEM` or `PUBLIC`).
3. **Decoupled Entity Expansion Limits**: Evaluates code points and standard characters (`&amp;`, `&lt;`, etc.) via a single-level parser, avoiding recursive parsing loops.

### Custom Security & Parsing API
Use the safe-parsing helpers to validate user-uploaded templates manually:
```javascript
const { validateXml, safeParseXml } = require('node-pptx-templater');

const schema = validateXml(userInputXml);
if (!schema.valid) {
  console.error(`Invalid XML format: ${schema.error} on Line ${schema.line}`);
}
```

---

## ⚡ Performance Benchmarks

Below are benchmark results compiled on a standard Intel Core i7 system processing a 50-slide presentation template:

| Operation | Average Duration |
| :--- | :--- |
| Load Template | ~110 ms |
| Scan & Replace 20 Text Tag Placeholders | ~2.5 ms |
| Validate XML Schemas & Relationship Integrity | ~14 ms |
| Table Row Cloning & Grid Merges (15 rows) | ~3 ms |
| Z-Order Re-indexing & Packaging Output | ~78 ms |

---

## ❓ FAQ & Developer Troubleshooting

### Q: PowerPoint displays a "Repair Presentation" alert when opening generated files.
* **Root Cause**: This is usually caused by:
  1. A missing relationship entry inside `.rels` maps (e.g., a slide refers to a chart file that isn't declared).
  2. Duplicate `<a16:rowId>` elements on a slide (caused by cloning rows raw without updating hashes).
* **Fix**: Ensure you always use the built-in `saveToFile()` or `toBuffer()` methods, which automatically execute validation checks, repair relationship indexes, and sanitize content structures.

### Q: A text placeholder is not replacing. How do I debug it?
* **Root Cause**: PowerPoint might have split your placeholder into multiple runs (e.g., `{{` and `placeholder}}`).
* **Fix**: 
  1. Enable debug logging to locate the fragment: `PPTX_LOG_LEVEL=debug node app.js`.
  2. In PowerPoint, select the placeholder text box, cut it, and paste it back using the **"Keep Text Only"** option. This unifies the run.

---

## 🤝 Contributing

We welcome contributions from the community. Please read our [CONTRIBUTING.md](./CONTRIBUTING.md) to set up the development environment, format code, and submit Pull Requests.

```bash
git clone https://github.com/jsuyog2/node-pptx-templater.git
cd node-pptx-templater
npm install
npm test
```

---

## 📄 License

Licensed under the MIT License. © node-pptx-templater contributors.
