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
* 🏷️ **Chart Data Labels & Value From Cells**: Configure custom arrays, dynamic category maps, rich templates, positions (insideEnd, bestFit, center), and custom styles (fonts, colors, weight) and serialize them directly to the underlying Excel backing sheet.
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

  // 3b. Configure custom chart data labels with templates and styling
  ppt.useSlide(2)
     .updateDataLabels('revenue-chart', {
       series: 0,
       template: '{category}: {value}',
       position: 'insideEnd',
       labelStyle: { fontFamily: 'Arial', fontSize: 10, bold: true }
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

## 📂 PowerPoint XML Folder Templates

Instead of loading and saving standard compiled `.pptx` ZIP files, the library natively supports working with uncompressed PowerPoint OpenXML directories. This is extremely useful for server environments and development setups, bypassing ZIP compression/decompression overhead and resolving relationships relative to the unzipped structure.

### 1. Load from XML Folder Template

You can load a template directly from a folder directory or `presentation.xml` entry point:

```javascript
const { PPTXTemplater, PPTXTemplate } = require('node-pptx-templater');

// Load using the directory root path (auto-detects ppt/presentation.xml)
const ppt = await PPTXTemplater.load('./monthly-template-folder');

// Load using fromPresentationXml with a configuration object
const ppt2 = await PPTXTemplate.fromPresentationXml({
  presentation: './ppt/presentation.xml',
  root: './template'
});
```

### 2. Save/Export directly to XML Folder

You can export the modified presentation back to an uncompressed folder structure on disk:

```javascript
await ppt.saveToFolder('./output-template-folder');
```

This generates:
```text
output-template-folder/
├── [Content_Types].xml
├── _rels/
├── ppt/
│   ├── presentation.xml
│   ├── _rels/
│   ├── slides/
│   ├── slideLayouts/
│   ├── slideMasters/
│   └── theme/
└── docProps/
```

### 3. Folder Mode Performance Benefits

Our benchmark results compare standard ZIP-based templates with uncompressed XML folder workflows:
* **Concurrency Throughput**: Up to **1.4x faster** under parallel request stress due to eliminated ZIP compression CPU locks.
* **Heap Memory Footprint**: Reduces memory overhead by avoiding full in-memory ZIP archives.

### 4. Validation

Ensure XML directory templates are correct and contain no orphan relations:

```javascript
const report = await ppt.validatePresentationXml();
if (!report.valid) {
  console.error('Errors found:', report.errors);
}
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

<!-- API_REFERENCE_START -->

### Tables API

#### `updateTable(tableId, rows)`
Replaces table rows with new data in the selected slide(s). Preserves borders, merged cells, fonts, colors, and alignment from the template.

* **Arguments**:
  * `tableId` (`string`): Table name or shape ID.
  * `rows` (`string[][]`): 2D array of cell values (row × col).
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).updateTable('summary-table', [
  ['Item', 'Value'],
  ['Widgets', '1,200'],
  ['Gadgets', '850']
]);
```

#### `addCellShape(tableId, rowIndex, colIndex, options)`
Dynamically adds a shape inside a table cell based on cell coordinates.

* **Arguments**:
  * `tableId` (`string`): Table name or shape ID.
  * `rowIndex` (`number`): 0-based row index.
  * `colIndex` (`number`): 0-based column index.
  * `options` (`Object`): Shape configuration options.
* **Returns**: `this` - The chainable presentation templater instance.

```javascript
await ppt.addCellShape('Table', 1, 2, { type: 'circle', fill: '#10B981' });
```

#### `updateCellShape(tableId, rowIndex, colIndex, shapeIndex, options)`
Updates an existing shape inside a table cell.

* **Arguments**:
  * `tableId` (`string`): Table name or shape ID.
  * `rowIndex` (`number`): 0-based row index.
  * `colIndex` (`number`): 0-based column index.
  * `shapeIndex` (`number`): 0-based shape index in the cell.
  * `options` (`Object`): Shape configuration properties to update.
* **Returns**: `this` - The chainable presentation templater instance.

```javascript
await ppt.updateCellShape('Table', 1, 2, 0, { fill: '#EF4444' });
```

#### `removeCellShape(tableId, rowIndex, colIndex, shapeIndex)`
Removes a shape from a table cell.

* **Arguments**:
  * `tableId` (`string`): Table name or shape ID.
  * `rowIndex` (`number`): 0-based row index.
  * `colIndex` (`number`): 0-based column index.
  * `shapeIndex` (`number`): 0-based shape index in the cell.
* **Returns**: `this` - The chainable presentation templater instance.

```javascript
await ppt.removeCellShape('Table', 1, 2, 0);
```

#### `getCellShape(tableId, rowIndex, colIndex, shapeIndex)`
Discovers and retrieves details of an existing cell shape on the targeted slide.

* **Arguments**:
  * `tableId` (`string`): Table name or shape ID.
  * `rowIndex` (`number`): 0-based row index.
  * `colIndex` (`number`): 0-based column index.
  * `shapeIndex` (`number`): 0-based shape index in the cell.
* **Returns**: `Object|null` - Shape details object, or null if not found.

```javascript
const shape = ppt.getCellShape('Table', 1, 2, 0);
```

#### `addTableRow(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).addTableRow('data-table', ['John Doe', 'Sales Manager', '$120k']);
```

#### `removeTableRow(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).removeTableRow('data-table', 2);
```

#### `insertTableRow(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).insertTableRow(());
```

#### `cloneTableRow(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).cloneTableRow(());
```

#### `updateCell(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).updateCell(());
```

#### `mergeCells(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).mergeCells('metrics-table', 1, 1, 2, 2);
```

#### `unmergeCells(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).unmergeCells('metrics-table', 1, 1, 2, 2);
```

#### `getMergedCells(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).getMergedCells(());
```

#### `validateMergeRegion(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).validateMergeRegion(());
```

#### `isMergedCell(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).isMergedCell(());
```

#### `getMergeParent(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).getMergeParent(());
```

#### `getMergeRegion(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).getMergeRegion(());
```

#### `splitMergedRegion(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).splitMergedRegion(());
```

#### `cloneMergedRegion(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).cloneMergedRegion(());
```

#### `autoFitTable(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).autoFitTable(());
```

#### `resizeTable(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).resizeTable(());
```

#### `getTables(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).getTables(());
```

---

### Charts API

#### `updateChart(chartId, data)`
Updates chart data in the selected slide(s). Finds charts by their name/ID and updates categories, series, and values. Preserves original chart styles, themes, and formatting. Supports inline custom data labels by passing objects in the format `{ data: number, label: string }` instead of numbers.

* **Arguments**:
  * `chartId` (`string`): Chart name or relationship ID.
  * `data` (`ChartData`): New chart data.
  * `data.categories` (`string[]`): Category labels (X-axis).
  * `data.series` (`SeriesData[]`): Data series array.
  * `data.series[].name` (`string`): Series name.
  * `data.series[].values` (`number[]|object[]`): Data values (numbers or label objects).
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).updateChart(chartId, data);
```

#### `validateCharts()`
Validates all charts in the presentation to ensure they are not corrupted. Checks XML, caches, and embedded workbook references.

* **Returns**: `Promise<Object>` - Validation results for charts.

```javascript
ppt.useSlide(1).validateCharts();
```

#### `repairCharts()`
Repairs common chart corruption issues such as broken caches, missing embedded workbooks, or orphan nodes.

* **Returns**: `Promise<PPTXTemplater>` - this

```javascript
ppt.useSlide(1).repairCharts();
```

#### `getChartLabelPositions(chartId)`
Retrieves the exact coordinate positions of all data labels for a chart on the active slide. Calculates absolute layout limits in EMUs (English Metric Units).

* **Arguments**:
  * `chartId` (`string`): 
* **Returns**: `Promise<Array<{series: string, category: string, seriesIndex: number, categoryIndex: number, value: number, x: number, y: number, width: number, height: number` - 

```javascript
const positions = await ppt.useSlide(1).getChartLabelPositions('SalesChart');
```

#### `getChartBarPositions(chartId)`
Retrieves the exact coordinate positions of all bars/columns for a chart on the active slide. Calculates absolute layout limits in EMUs (English Metric Units).

* **Arguments**:
  * `chartId` (`string`): 
* **Returns**: `Promise<Array<{series: string, category: string, seriesIndex: number, categoryIndex: number, value: number, x: number, y: number, width: number, height: number` - 

```javascript
const bars = await ppt.useSlide(1).getChartBarPositions('SalesChart');
```

#### `addTextAtPosition(options)`
Adds a textbox shape at a specific EMU coordinate position on targeted slides. Supports custom font styling and alignment configuration.

* **Arguments**:
  * `options` (`Object`): 
  * `options.text` (`string`): 
  * `options.x` (`number`): 
  * `options.y` (`number`): 
  * `[options.width=1200000]` (`number`): 
  * `[options.height=300000]` (`number`): 
  * `[options.style]` (`Object`): 
* **Returns**: `this` - The chainable presentation engine instance.

```javascript
ppt.useSlide(1).addTextAtPosition({
  text: 'Label',
  x: 1000000,
  y: 1000000
});
```

#### `addTextNearChartLabel(options)`
Dynamically places textboxes next to a chart's data labels with vertical collision avoidance. Textboxes are positioned either on the left or right of the chart area, vertically aligned with their corresponding label.

* **Arguments**:
  * `options` (`Object`): 
  * `options.chart` (`string`): 
  * `options.text` (`string|Function`): 
  * `[options.position='left']` (`'left'|'right'`): 
  * `[options.style]` (`Object`): 
* **Returns**: `this` - The chainable presentation engine instance.

```javascript
ppt.addTextNearChartLabel({
  chart: 'SalesChart',
  text: 'Series',
  position: 'left'
});
```

#### `updateChartData(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).updateChartData('sales-chart', {
  categories: ['Q1', 'Q2'],
  series: [{ name: 'Actual', values: [100, 150] }]
});
```

#### `replaceChartSeries(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).replaceChartSeries(());
```

#### `updateChartTitle(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).updateChartTitle('sales-chart', 'Quarterly Metrics Overview');
```

#### `updateChartCategories(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).updateChartCategories(());
```

#### `updateDataLabels(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).updateDataLabels('SalesChart', {
  series: 0,
  labels: ['Excellent', 'Good', 'Poor']
});
```

#### `getDataLabels(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
const labels = await ppt.useSlide(1).getDataLabels('SalesChart', { series: 0 });
console.log(labels); // [{ point: 0, value: 'Excellent' }, ...]
```

#### `validateDataLabels(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
const result = await ppt.useSlide(1).validateDataLabels('SalesChart', {
  labels: ['High', 'Low']
});
console.log(result.valid);
```

#### `validateChartLabels(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
const result = await ppt.useSlide(1).validateChartLabels('SalesChart', {
  labels: ['High', 'Low']
});
console.log(result.valid);
```

#### `validateSeriesNameLabels(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
const result = await ppt.useSlide(1).validateSeriesNameLabels('SalesChart', {
  enabled: true,
  position: 'left'
});
console.log(result.valid);
```

#### `getCharts(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).getCharts(());
```

---

### Slides API

#### `useSlide(...slideRefs)`
Selects one or more slides to work on. All subsequent operations (replaceText, updateChart, etc.) apply to these slides. If not called, operations apply to ALL slides.

* **Arguments**:
  * `slideRefs` (`...number|string`): Slide numbers (1-based), IDs, or tags.
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).useSlide(...slideRefs);
```

#### `useAllSlides()`
Selects all slides.

* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).useAllSlides();
```

#### `addSlide(options = {})`
Adds a new slide to the presentation. Automatically generates required XML and relationship entries.

* **Arguments**:
  * `options` (`NewSlideOptions`): Slide definition.
  * `[options.title]` (`string`): Slide title text.
  * `[options.layout]` (`string`): Layout name to use (default: 'blank').
  * `[options.elements]` (`SlideElement[]`): Elements to add to the slide.
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).addSlide(options = {});
```

#### `cloneSlide(sourceSlideNumber, atPosition)`
Clones an existing slide and appends it to the end (or at a position).

* **Arguments**:
  * `sourceSlideNumber` (`number`): 1-based source slide number.
  * `[atPosition]` (`number`): Optional position to insert (1-based). Default: append.
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).cloneSlide(sourceSlideNumber, atPosition);
```

#### `removeSlide(slideNumber)`
Removes a slide from the presentation.

* **Arguments**:
  * `slideNumber` (`number`): 1-based slide number to remove.
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).removeSlide(slideNumber);
```

#### `reorderSlides(order)`
Reorders slides in the presentation.

* **Arguments**:
  * `order` (`number[]`): Array of 1-based slide numbers in desired order.
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).reorderSlides(order);
```

#### `tagSlide(slideNumber, tag)`
Tags a slide with a custom string identifier for later selection.

* **Arguments**:
  * `slideNumber` (`number`): 1-based slide number.
  * `tag` (`string`): Custom tag string.
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).tagSlide(slideNumber, tag);
```

#### `exportSlides(...slideNumbers)`
Exports selected slides to a new standalone PPTX engine. Useful for creating "slide decks" from a master template.

* **Arguments**:
  * `slideNumbers` (`...number`): 1-based slide numbers to export.
* **Returns**: `Promise<PPTXTemplater>` - New engine with only the selected slides.

```javascript
ppt.useSlide(1).exportSlides(...slideNumbers);
```

#### `importSlideFrom(sourceEngine, slideRef)`
Imports a single slide from another PPTXTemplater instance into this presentation. Preserves all slide layouts, charts, relationships, and embedded media.

* **Arguments**:
  * `sourceEngine` (`PPTXTemplater`): Source PPTXTemplater instance.
  * `slideRef` (`number|string`): Slide index (1-based), ID, or custom tag.
* **Returns**: `Promise<PPTXTemplater>` - this (chainable)

```javascript
const source = await PPTXTemplater.load('template2.pptx');
await ppt.importSlideFrom(source, 1);
```

#### `importSlides(slideIndices)`
Imports selected slides from the current template, discarding the rest. The remaining slides are reordered to match the provided array. Preserves all layouts, themes, relationships, and embedded media.

* **Arguments**:
  * `slideIndices` (`number[]`): Array of 1-based slide indices to keep.
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).importSlides(slideIndices);
```

#### `duplicateSlide(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.duplicateSlide(1, 2);
```

#### `deleteSlide(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).deleteSlide(());
```

#### `moveSlide(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).moveSlide(());
```

#### `insertSlide(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).insertSlide(());
```

#### `getSlides(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).getSlides(());
```

---

### Text API

#### `replaceText(replacements)`
Replaces template placeholders (e.g., {{key}}) with values in the selected slides. Works inside text boxes, titles, grouped shapes, tables, and shapes.

* **Arguments**:
  * `replacements` (`Object.<string, string>`): Map of placeholder → replacement value.
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).replaceText(replacements);
```

#### `addHyperlink(options)`
Adds or replaces a hyperlink on a text run or shape.

* **Arguments**:
  * `options` (`HyperlinkOptions`): Hyperlink configuration.
  * `options.text` (`string`): Text to find and make clickable.
  * `options.url` (`string`): Target URL.
  * `[options.tooltip]` (`string`): Optional tooltip.
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).addHyperlink(options);
```

#### `addSlideLink(options)`
Adds an inter-slide hyperlink to a specific text element.

* **Arguments**:
  * `options` (`Object`): Link configuration.
  * `options.sourceSlide` (`number`): Source slide number (1-based).
  * `options.targetSlide` (`number`): Destination slide number (1-based).
  * `options.element` (`string`): Text element to make clickable.
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).addSlideLink(options);
```

#### `addImageLink(options)`
Adds an inter-slide hyperlink to an image.

* **Arguments**:
  * `options` (`Object`): 
  * `options.slide` (`number`): Source slide number.
  * `options.imageId` (`string`): Image name/id to make clickable.
  * `options.targetSlide` (`number`): Destination slide number.
* **Returns**: `PPTXTemplater` - this

```javascript
ppt.useSlide(1).addImageLink(options);
```

#### `addShapeLink(options)`
Adds an inter-slide hyperlink to a shape.

* **Arguments**:
  * `options` (`Object`): 
  * `options.slide` (`number`): Source slide number.
  * `options.shapeId` (`string`): Shape name/id to make clickable.
  * `options.targetSlide` (`number`): Destination slide number.
* **Returns**: `PPTXTemplater` - this

```javascript
ppt.useSlide(1).addShapeLink(options);
```

#### `addTextNavigationLink(options)`
Adds a special navigation link (next, previous, first, last slide) to a text element.

* **Arguments**:
  * `options` (`Object`): 
  * `options.slide` (`number`): Source slide number (1-based).
  * `options.element` (`string`): Text element to make clickable.
  * `options.action` (`'next'|'previous'|'first'|'last'`): Navigation action type.
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).addTextNavigationLink(options);
```

#### `addShapeNavigationLink(options)`
Adds a special navigation link (next, previous, first, last slide) to a shape or image.

* **Arguments**:
  * `options` (`Object`): 
  * `options.slide` (`number`): Source slide number (1-based).
  * `options.shapeId` (`string`): Shape name/id to make clickable.
  * `options.action` (`'next'|'previous'|'first'|'last'`): Navigation action type.
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).addShapeNavigationLink(options);
```

#### `updateText(tag, data)`
Updates shape text or list content by placeholder tag or shape name/ID. Supports bullet lists, numbered lists, nested lists, and custom styling.

* **Arguments**:
  * `tag` (`string`): Placeholder tag (e.g. '{{name}}' or 'name') or shape name/ID.
  * `data` (`string|Object`): String value or list configuration object.
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).updateText('Features', {
  list: ['Point A', 'Point B', 'Point C']
});
```

#### `getList(tag)`
Retrieves list items from a shape or text box by name or placeholder tag.

* **Arguments**:
  * `tag` (`string`): Shape name/ID or placeholder tag.
* **Returns**: `Array` - Nested list structure of items.

```javascript
const items = ppt.useSlide(1).getList('Features');
console.log(items); // ['A', { text: 'B', children: [...] }]
```

#### `validateList(data)`
Validates a list structure and values.

* **Arguments**:
  * `data` (`Object|Array`): List config object or array of items.
* **Returns**: `Object` - Report containing validation result.

```javascript
const result = ppt.validateList(['Valid string', 'Another item']);
console.log(result.valid);
```

#### `replaceTextByTag(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).replaceTextByTag('company', 'Acme Corp');
```

#### `replaceMultiple(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).replaceMultiple(());
```

#### `findText(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).findText(());
```

#### `getTextElements(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).getTextElements(());
```

---

### Images API

#### `replaceImage(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
await ppt.useSlide(1).replaceImage('logo-img', './new-logo.png');
```

#### `addImage(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).addImage(());
```

#### `removeImage(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).removeImage(());
```

#### `getImages(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).getImages(());
```

---

### Shapes API

#### `updateShapePosition(shapeId, options = {})`
Updates the position and/or dimensions of an existing shape on targeted slides.

* **Arguments**:
  * `shapeId` (`string`): 
  * `options` (`Object`): 
  * `[options.x]` (`number`): 
  * `[options.y]` (`number`): 
  * `[options.width]` (`number`): 
  * `[options.height]` (`number`): 
* **Returns**: `this` - The chainable presentation engine instance.

```javascript
ppt.useSlide(1).updateShapePosition('TitleShape', { x: 1000000, y: 1500000 });
```

#### `updateTextBoxPosition(textBoxId, options = {})`
Updates the position and/or dimensions of an existing textbox on targeted slides.

* **Arguments**:
  * `textBoxId` (`string`): 
  * `options` (`Object`): 
  * `[options.x]` (`number`): 
  * `[options.y]` (`number`): 
  * `[options.width]` (`number`): 
  * `[options.height]` (`number`): 
* **Returns**: `this` - The chainable presentation engine instance.

```javascript
ppt.useSlide(1).updateTextBoxPosition('TextBox 2', { x: 1000000, y: 1500000 });
```

#### `validateShape(options)`
Validates shape options configuration.

* **Arguments**:
  * `options` (`Object`): 
* **Returns**: `string[]` - List of validation error messages.

```javascript
const errors = ppt.validateShape(shapeOptions);
```

#### `addShape(options)`
Adds a new shape dynamically to the targeted slide(s).

* **Arguments**:
  * `options` (`Object`): 
* **Returns**: `this` - The chainable presentation templater instance.

```javascript
await ppt.useSlide(1).addShape({
  type: 'rectangle',
  id: 'sales-box',
  x: 100,
  y: 100,
  width: 200,
  height: 100,
  fill: '#2563EB'
});
```

#### `updateShape(shapeId, options)`
Updates an existing shape in-place.

* **Arguments**:
  * `shapeId` (`string`): 
  * `options` (`Object`): 
* **Returns**: `this` - The chainable presentation templater instance.

```javascript
await ppt.useSlide(1).updateShape('sales-box', { fill: '#10B981' });
```

#### `removeShape(shapeId)`
Removes a shape from the targeted slide(s).

* **Arguments**:
  * `shapeId` (`string`): 
* **Returns**: `this` - The chainable presentation templater instance.

```javascript
await ppt.useSlide(1).removeShape('sales-box');
```

#### `getShape(shapeId)`
Discovers and retrieves details of an existing shape on the targeted slides.

* **Arguments**:
  * `shapeId` (`string`): 
* **Returns**: `Object|null` - Shape details object, or null if not found.

```javascript
const shape = ppt.getShape('sales-box');
```

#### `updateShapeText(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).updateShapeText(());
```

#### `cloneShape(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).cloneShape('card-bg', 'card-bg-2');
```

#### `deleteShape(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).deleteShape(());
```

#### `getShapes(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).getShapes(());
```

---

### Layer Stacking (Z-Order)

#### `bringForward(optionsOrId)`
Moves slide element one layer forward.


```javascript
ppt.useSlide(1).bringForward(optionsOrId);
```

#### `sendBackward(optionsOrId)`
Moves slide element one layer backward.


```javascript
ppt.useSlide(1).sendBackward(optionsOrId);
```

#### `bringToFront(optionsOrId)`
Moves slide element above all other objects.


```javascript
ppt.useSlide(1).bringToFront('OverlayLogo');
```

#### `sendToBack(optionsOrId)`
Moves slide element behind all other objects.


```javascript
ppt.useSlide(1).sendToBack(optionsOrId);
```

#### `setZIndex(optionsOrId, zIndex)`
Moves slide element to the specific 1-based stacking position.


```javascript
ppt.useSlide(1).setZIndex(optionsOrId, zIndex);
```

#### `moveObjectBefore(optionsOrId, targetId)`
Moves slide element directly before (below) a target element.


```javascript
ppt.useSlide(1).moveObjectBefore(optionsOrId, targetId);
```

#### `moveObjectAfter(optionsOrId, targetId)`
Moves slide element directly after (above) a target element.


```javascript
ppt.useSlide(1).moveObjectAfter(optionsOrId, targetId);
```

#### `reorderObjects(optionsOrOrder)`
Reorders slide objects exactly as specified in the array.


```javascript
ppt.useSlide(1).reorderObjects(optionsOrOrder);
```

#### `getObjectOrder(slideIndex)`
Gets the ordered metadata of all objects on the slide.


```javascript
ppt.useSlide(1).getObjectOrder(slideIndex);
```

#### `applyZOrder(slideOrConfigs, configsOption)`
Applies bulk template configurations for slide elements stacking layers.


```javascript
ppt.useSlide(1).applyZOrder(slideOrConfigs, configsOption);
```

#### `getTopMostObject(slideIndex)`
Retrieves the info of the top-most object on the slide.


```javascript
ppt.useSlide(1).getTopMostObject(slideIndex);
```

#### `getBottomMostObject(slideIndex)`
Retrieves the info of the bottom-most object on the slide.


```javascript
ppt.useSlide(1).getBottomMostObject(slideIndex);
```

#### `swapObjects(slideIndexOrId1, id1OrId2, id2)`
Swaps stacking positions of two slide objects.


```javascript
ppt.useSlide(1).swapObjects(slideIndexOrId1, id1OrId2, id2);
```

#### `sortObjects(slideIndexOrCompareFn, compareFnOption)`
Sorts stacking order using a custom comparison function.


```javascript
ppt.useSlide(1).sortObjects(slideIndexOrCompareFn, compareFnOption);
```

#### `normalizeZOrder(slideIndex)`
Cleans up and normalizes stacking order consistency.


```javascript
ppt.useSlide(1).normalizeZOrder(slideIndex);
```

---

### Utilities & Validation

#### `const()`
This is the primary public API. It coordinates all sub-managers (ZipManager, SlideManager, ChartManager, etc.) and exposes a fluent, chainable interface for template manipulation. OpenXML PPTX Structure: ├── [Content_Types].xml      — lists all parts and their MIME types ├── _rels/.rels              — root relationships (points to presentation) ├── ppt/ │   ├── presentation.xml     — slide order, slide masters references │   ├── _rels/presentation.xml.rels │   ├── slides/ │   │   ├── slide1.xml       — individual slide content │   │   └── _rels/slide1.xml.rels │   ├── slideLayouts/        — layout templates (title, content, etc.) │   ├── slideMasters/        — master slide designs │   ├── theme/               — color/font themes │   ├── charts/              — embedded chart XML │   └── media/               — embedded images/videos └── docProps/ ├── core.xml             — author, title, etc. └── app.xml              — application metadata


```javascript
ppt.useSlide(1).const();
```

#### `class()`



```javascript
ppt.useSlide(1).class();
```

#### `static()`
Loads a PPTX template from a file path or buffer. @static @throws {PPTXError} If the file cannot be read or is not a valid PPTX.

* **Arguments**:
  * `source` (`string|Buffer`): Path to PPTX file or Buffer containing PPTX data.
* **Returns**: `Promise<PPTXTemplater>` - Initialized engine instance.

```javascript
ppt.useSlide(1).static();
```

#### `getInfo()`
Returns presentation metadata (title, author, slide count, etc.)

* **Returns**: `PresentationInfo` - Metadata object.

```javascript
ppt.useSlide(1).getInfo();
```

#### `validate()`
Validates the XML structure of the current PPTX. Reports issues with relationship IDs, missing parts, etc.

* **Returns**: `ValidationResult` - Object with `valid`, `errors`, and `warnings` arrays.

```javascript
ppt.useSlide(1).validate();
```

#### `repair()`
Repairs corrupted OpenXML structure, relationships, and content types. Removes orphan relationships, rebuilds slide references, and fixes missing entries.

* **Returns**: `Promise<PPTXTemplater>` - this (chainable)

```javascript
ppt.useSlide(1).repair();
```

#### `debugRelationships()`
Logs all relationships across the presentation to the console for debugging.

* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).debugRelationships();
```

#### `inspectSlide(slideIndex)`
Inspects a specific slide's structure and relationships.

* **Arguments**:
  * `slideIndex` (`number`): 1-based slide index.
* **Returns**: `PPTXTemplater` - this (chainable)

```javascript
ppt.useSlide(1).inspectSlide(slideIndex);
```

#### `inspectXML(xmlPath)`
Inspects and logs the raw XML of any file in the ZIP.

* **Arguments**:
  * `xmlPath` (`string`): Path inside the ZIP (e.g., 'ppt/slides/slide1.xml')
* **Returns**: `Promise<PPTXTemplater>` - this (chainable)

```javascript
ppt.useSlide(1).inspectXML(xmlPath);
```

#### `inspectChart(chartId)`
Inspects a specific chart's metadata and structure.

* **Arguments**:
  * `chartId` (`string`): 

```javascript
ppt.useSlide(1).inspectChart(chartId);
```

#### `inspectChartXML(chartFileName)`
Inspects and logs the raw XML of a chart file.

* **Arguments**:
  * `chartFileName` (`string`): 

```javascript
ppt.useSlide(1).inspectChartXML(chartFileName);
```

#### `debugChartRelationships()`
Logs all chart relationships.


```javascript
ppt.useSlide(1).debugChartRelationships();
```

#### `saveToFile(filePath, options = {})`
Saves the modified PPTX to a file on disk.

* **Arguments**:
  * `filePath` (`string`): Output file path.
  * `[options]` (`Object`): Save options.
  * `[options.strict=false]` (`boolean`): Throw error on validation failure.
* **Returns**: `Promise<void>` - 

```javascript
ppt.useSlide(1).saveToFile(filePath, options = {});
```

#### `save(filePath, options = {})`
Saves the presentation. Equivalent to saveToFile.

* **Arguments**:
  * `filePath` (`string`): Output file path.
  * `[options]` (`Object`): Save options.
* **Returns**: `Promise<void>` - 

```javascript
await ppt.save('output.pptx');
```

#### `saveXml(folderPath)`
Saves the modified presentation XML structures directly to a folder.

* **Arguments**:
  * `folderPath` (`string`): Target directory path.
* **Returns**: `Promise<void>` - 

```javascript
ppt.useSlide(1).saveXml(folderPath);
```

#### `saveToFolder(folderPath)`
Saves the modified presentation XML structures directly to a folder.

* **Arguments**:
  * `folderPath` (`string`): Target directory path.
* **Returns**: `Promise<void>` - 

```javascript
await ppt.saveToFolder('./output-template');
```

#### `toBuffer(options = {})`
Returns the PPTX content as a Node.js Buffer.

* **Arguments**:
  * `[options]` (`Object`): Save options.
* **Returns**: `Promise<Buffer>` - 

```javascript
ppt.useSlide(1).toBuffer(options = {});
```

#### `toStream(options = {})`
Returns the PPTX content as a readable Node.js Stream.

* **Arguments**:
  * `[options]` (`Object`): Save options.
* **Returns**: `Promise<NodeJS.ReadableStream>` - 

```javascript
ppt.useSlide(1).toStream(options = {});
```

#### `saveToStream(writableOrOptions, options = {})`
Saves the presentation to a readable stream or pipes it to a writable stream.

* **Arguments**:
  * `[writableOrOptions]` (`NodeJS.WritableStream|Object`): Writable stream to pipe to, or options object.
  * `[options]` (`Object`): Save options if writable stream was passed first.
* **Returns**: `Promise<NodeJS.ReadableStream|void>` - 

```javascript
ppt.useSlide(1).saveToStream(writableOrOptions, options = {});
```

#### `validatePresentationXml()`
Performs validation specifically on PowerPoint XML folder contents/relationships.

* **Returns**: `Promise<{valid: boolean, errors: string[], warnings: string[]` - 

```javascript
const report = await ppt.validatePresentationXml();
if (!report.valid) console.error(report.errors);
```

#### `slideCount()`
Returns the total number of slides in the loaded presentation. @type {number}


```javascript
ppt.useSlide(1).slideCount();
```

#### `function()`
OpenXML relationship IDs follow the format rId1, rId2, rId3, ... They must be unique within each .rels file. These utilities generate collision-free IDs when adding new relationships. / /** Generates the next available relationship ID given an array of existing IDs. Always uses the format "rId{N}" where N is the next integer after the max.

* **Arguments**:
  * `existingIds` (`string[]`): Array of existing rId strings (e.g., ['rId1', 'rId2']).
* **Returns**: `string` - New relationship ID (e.g., 'rId3').

```javascript
ppt.useSlide(1).function();
```

#### `preload(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).preload(());
```

#### `cache(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).cache(());
```

#### `fromCache(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).fromCache(());
```

#### `clearCache(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).clearCache(());
```

#### `enablePerformanceProfile(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).enablePerformanceProfile(());
```

#### `getPerformanceMetrics(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).getPerformanceMetrics(());
```

#### `fromPresentationXml(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
const ppt = await PPTXTemplate.fromPresentationXml('./template-folder');
```

#### `validatePresentation(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).validatePresentation(());
```

#### `validateSlide(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).validateSlide(());
```

#### `validateTable(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).validateTable(());
```

#### `validateArchive(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).validateArchive(());
```

#### `enableDebugZip(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).enableDebugZip(());
```

#### `validateRelationships(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).validateRelationships(());
```

#### `zipManager(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).zipManager(());
```

#### `xmlParser(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).xmlParser(());
```

#### `contentTypesManager(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).contentTypesManager(());
```

#### `relationshipManager(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).relationshipManager(());
```

#### `slideManager(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).slideManager(());
```

#### `chartManager(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).chartManager(());
```

#### `tableManager(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).tableManager(());
```

#### `shapeManager(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).shapeManager(());
```

#### `imageManager(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).imageManager(());
```

#### `textManager(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).textManager(());
```

#### `hyperlinkManager(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).hyperlinkManager(());
```

#### `mediaManager(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).mediaManager(());
```

#### `load(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
const ppt = await PPTXTemplater.load('./my_template.pptx');
```

#### `create(())`
Delegates core actions to slide element sub-managers.

* **Returns**: `PPTXTemplater` - The fluent engine instance.

```javascript
ppt.useSlide(1).create(());
```

---

<!-- API_REFERENCE_END -->

---

## 📊 Chart Data Labels & Value From Cells

PPTXForge supports advanced, PowerPoint-compatible chart data labels. You can customize label data sources, templates, positioning, and visual styles. 

### 1. Value From Cells (Excel Synchronization)
Pull dynamic data labels directly from worksheet range cells inside the backing Excel spreadsheet. This is Excel's native "Value From Cells" feature, reconstructed programmatically inside the `.pptx` XML and `.xlsx` worksheets.

```javascript
ppt.useSlide(1).updateDataLabels('SalesChart', {
  series: 0,
  labelsFromCells: 'Sheet1!D2:D5'
});
```

### 2. Literal Arrays
Override data labels for specific points in a series with a static list of strings:

```javascript
ppt.useSlide(1).updateDataLabels('SalesChart', {
  series: 0,
  labels: ['Top Performance', 'Met Target', 'Action Required', 'At Risk']
});
```

### 3. Dynamic Label Templates
Combine values, category names, percentages, and custom label strings to format annotations:

```javascript
ppt.useSlide(1).updateDataLabels('SalesChart', {
  series: 0,
  template: '{category}: {value} ({percentage}%)'
});
```
* **Variables**: `{category}`, `{value}`, `{percentage}`, `{series}`, `{customLabel}`

### 4. Label Positions
Set label alignment to any of PowerPoint's standard values:

```javascript
ppt.useSlide(1).updateDataLabels('SalesChart', {
  series: 0,
  position: 'insideEnd' // 'center', 'insideEnd', 'insideBase', 'outsideEnd', 'bestFit', 'left', 'right', 'top', 'bottom'
});
```

### 5. Custom Label Styling
Format your data labels using custom fonts, sizes, colors, and weight properties:

```javascript
ppt.useSlide(1).updateDataLabels('SalesChart', {
  series: 0,
  labels: ['High', 'Medium', 'Low'],
  labelStyle: {
    fontFamily: 'Century Gothic',
    fontSize: 12,
    color: '#0055A5',
    bold: true,
    italic: true,
    underline: true
  }
});
```

### 6. Inline Custom Data Labels (`updateChart`)
You can define custom data labels inline within the standard `updateChart()` series `values` array using objects in the format `{ data: number, label: string }`. This avoids calling separate styling or update methods:

```javascript
ppt.useSlide(1).updateChart('RevenueChart', {
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    {
      name: 'Product A',
      values: [
        { data: 145, label: 'Q1: 145 (Low)' },
        { data: 210, label: 'Q2: 210 (Med)' },
        { data: 190, label: 'Q3: 190 (Med)' },
        { data: 250, label: 'Q4: 250 (High)' }
      ]
    }
  ]
});
```

To preserve PowerPoint integrity, the engine ensures that if one value contains a label, all values in that series must have labels, and the label properties must be string values.

### 7. Label Style Inheritance & Series Names Inside Bars

#### Style Inheritance
When custom labels (inline or via `updateDataLabels`) are generated, they inherit the styling properties (font family, font size, bold, italic, color, and alignment) defined in the template's `<c:txPr>` tag for the series. This ensures your custom labels match the branding and layout design from your PowerPoint template file.

#### Series Names Inside Bars (`showSeriesNameInBar`)
For stacked bar charts, you can show the series name inside each segment (typically centered) to make them easily readable. To enable this, set `showSeriesNameInBar: true` in the chart options (globally) or on individual series:

```javascript
ppt.useSlide(1).updateChart('RevenueChart', {
  showSeriesNameInBar: true, // Show series name labels globally
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'Product A', values: [100, 200, 300, 400] },
    { name: 'Product B', values: [150, 250, 350, 450], showSeriesNameInBar: true } // Or per-series
  ]
});
```

#### Chart Labels Validation
You can programmatically validate that the chart labels configured in a chart conform to your template structure:

```javascript
const report = await ppt.useSlide(1).validateChartLabels('RevenueChart', {
  labels: ['Custom label 1', 'Custom label 2'],
  showSeriesNameInBar: true
});

if (!report.valid) {
  console.log('Errors:', report.errors);
  console.log('Warnings:', report.warnings);
}
```

#### External Series Name Labels (`seriesNameLabels`)
You can position the series names outside the chart area as separate text boxes (`<p:sp>`) aligned with each corresponding bar or stack. These external labels automatically inherit the styling (font family, font size, bold, italic, color) defined in the template's series properties.

Supported features:
* **Position**: Can be set to `'left'` or `'right'` to align text boxes to the left or right of the chart.
* **Auto-fit & Height Wrapping**: Automatically wraps text and calculates height if labels are longer than the available margin space.
* **Collision Detection**: Prevent overlapping with slide bounds by shrinking the chart, and resolve vertical overlaps of labels by shifting them apart.
* **Idempotency**: Safely clean up previous labels on multiple chart updates.

Example usage:
```javascript
ppt.useSlide(1).updateChart('RevenueChart', {
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'Product A', values: [100, 200, 300, 400] },
    { name: 'Product B', values: [150, 250, 350, 450] }
  ],
  seriesNameLabels: {
    enabled: true,
    position: 'left', // 'left' or 'right'
    autoFit: true     // Automatically wrap and shrink layout if needed (default: true)
  }
});
```

To validate your configuration and check for boundary or collision warning/errors:
```javascript
const report = await ppt.useSlide(1).validateSeriesNameLabels('RevenueChart', {
  enabled: true,
  position: 'left',
  autoFit: true
});

if (!report.valid) {
  console.error('Validation errors:', report.errors);
  console.warn('Validation warnings:', report.warnings);
}
```

---

## 📋 Native Lists (Bullet & Numbered Lists)

PPTXForge supports native PowerPoint bullet lists and numbered lists across text placeholders, shapes, text boxes, table cells, and grouped shapes. When generating lists, the engine preserves run styles, custom bullet characters, indentation, and color overlays, generating valid OpenXML/DrawingML without repair alerts.

### 1. Basic Bullet List
Update a shape or text placeholder to be a bullet list:

```javascript
ppt.useSlide(1).updateText('Features', {
  list: [
    'Fast PPTX generation',
    'OpenXML based',
    'Chart updates',
    'Table updates'
  ]
});
```

### 2. Numbered / Ordered List
Use the `ordered` flag to convert the list to a numbered format:

```javascript
ppt.useSlide(1).updateText('Steps', {
  ordered: true,
  list: [
    'Import template',
    'Update data',
    'Generate PPTX'
  ]
});
```
* Custom numbering systems can be specified with `style.numberType` (e.g., `arabicPeriod`, `alphaLcParen`, `romanUcPeriod`).

### 3. Nested / Multi-Level Lists
Construct hierarchy by passing objects containing a `text` string and a `children` array:

```javascript
ppt.useSlide(1).updateText('Requirements', {
  list: [
    'Frontend',
    {
      text: 'Backend Development',
      children: [
        'Node.js',
        {
          text: 'Databases',
          children: ['MS SQL', 'PostgreSQL']
        }
      ]
    }
  ]
});
```

### 4. Custom List Styling
Customize bullet characters, colors, sizes, and font properties:

```javascript
ppt.useSlide(1).updateText('KPIs', {
  list: ['Revenue Up', 'Margins Normal'],
  style: {
    fontFamily: 'Arial',
    fontSize: 18,
    color: '#0055AA',
    bulletColor: '#FF5500',
    bulletChar: '✦',
    bulletSize: 120 // Percentage relative to text (e.g. 120%)
  }
});
```

### 5. Table Cell Lists
You can also generate list hierarchies directly inside a cell of a DrawingML table:

```javascript
ppt.useSlide(3).updateTable('sales-table', [
  ['Category', 'Performance details'],
  ['North Region', '{{CellPlaceholder}}']
]);

ppt.updateText('CellPlaceholder', {
  list: ['Table Bullet 1', 'Table Bullet 2']
});
```

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

## ⚡ Performance Optimization & Caching APIs

The library provides first-class support for memory optimization, template caching, lazy loading, and streaming saves.

### 1. In-Memory Template Caching (IIS & Server Environments)
Instead of loading and parsing the PPTX ZIP structure from disk on every request, preload the template once. Subsequent templates can be instantiated from the cache in **0ms**:

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

// Preload templates into memory cache at server startup
await PPTXTemplater.preload('./templates/report.pptx');

// Load from cache instantly inside request handlers
app.post('/generate-report', async (req, res) => {
  const ppt = await PPTXTemplater.fromCache('./templates/report.pptx');
  
  ppt.useSlide(1).replaceText({ '{{title}}': req.body.title });
  
  const buffer = await ppt.toBuffer();
  res.send(buffer);
});

// Clear cache if templates change
PPTXTemplater.clearCache();
```

### 2. Performance Profiling
Expose timing and memory metrics across the generation pipeline:

```javascript
const ppt = await PPTXTemplater.load('report.pptx');
ppt.enablePerformanceProfile();

// Perform modifications...
ppt.useSlide(1).replaceText({ '{{title}}': 'Performance' });

await ppt.toBuffer();

// Retrieve timing statistics (in milliseconds)
const metrics = ppt.getPerformanceMetrics();
console.log(metrics);
/*
Output:
{
  enabled: true,
  templateLoadMs: 12.5,
  parseMs: 45.2,
  chartUpdateMs: 0,
  imageUpdateMs: 0,
  zipGenerationMs: 65.8,
  totalMs: 125.4,
  memoryUsedMB: 38.45
}
*/
```

### 3. Configurable ZIP Compression
Balance CPU execution time and file size when saving the presentation. Compression options support `'none' | 'fast' | 'balanced' | 'maximum'`:

```javascript
// balanced is the default (level 6 DEFLATE)
await ppt.save('output.pptx', { compression: 'balanced' });

// maximum compression (level 9 DEFLATE) - best file size, slightly slower
await ppt.save('output.pptx', { compression: 'maximum' });

// fast compression (level 1 DEFLATE) - fast packaging, good compression
await ppt.save('output.pptx', { compression: 'fast' });

// none / store (0% compression) - extremely fast, skips compression entirely
const fastBuffer = await ppt.toBuffer({ compression: 'none' });
```

### 4. Streaming Save & Streaming Image Input
Avoid buffering large output files in memory by saving directly to readable/writable streams. You can also pass Readable streams (like `fs.createReadStream`) directly to image APIs:

```javascript
const fs = require('fs');

const ppt = await PPTXTemplater.load('report.pptx');

// Stream image from file path without loading into memory buffer
const imageStream = fs.createReadStream('large-image.png');
await ppt.useSlide(1).replaceImage('placeholder-img', imageStream);

// Stream final PPTX directly to file disk or HTTP response
const writeStream = fs.createWriteStream('output.pptx');
await ppt.saveToStream(writeStream);
```

---

## 🌐 IIS & Windows Server Deployment Guide

When deploying the library on **Windows Server** with **IIS** using `httpPlatformHandler` or `iisnode`, follow these production-ready recommendations:

1. **Preload Large Templates**: Always call `await PPTXTemplater.preload(templatePath)` during application startup. This avoids high IIS request queue concurrency from competing for file handles or causing disk bottlenecks.
2. **Use Streaming Saves**: For concurrent routes serving large PPTX outputs, use `saveToStream()` to stream data straight into the HTTP response stream rather than buffering the output as large Node buffers.
3. **Optimize Compression**: If CPU cycles are a bottleneck on the IIS worker process, set `{ compression: 'fast' }` or `{ compression: 'none' }` on your save options.
4. **Increase httpPlatformHandler Request Limits**: Ensure the `requestTimeout` and `maxConnections` settings in your IIS `web.config` are set appropriately to allow long-running file streaming tasks.

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
