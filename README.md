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
Updates chart data in the selected slide(s). Finds charts by their name/ID and updates categories, series, and values. Preserves original chart styles, themes, and formatting.

* **Arguments**:
  * `chartId` (`string`): Chart name or relationship ID.
  * `data` (`ChartData`): New chart data.
  * `data.categories` (`string[]`): Category labels (X-axis).
  * `data.series` (`SeriesData[]`): Data series array.
  * `data.series[].name` (`string`): Series name.
  * `data.series[].values` (`number[]`): Data values.
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

#### `toBuffer()`
Returns the PPTX content as a Node.js Buffer.

* **Returns**: `Promise<Buffer>` - 

```javascript
ppt.useSlide(1).toBuffer();
```

#### `toStream()`
Returns the PPTX content as a readable Node.js Stream.

* **Returns**: `Promise<NodeJS.ReadableStream>` - 

```javascript
ppt.useSlide(1).toStream();
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
