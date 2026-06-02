# node-pptx-templater

> A low-level PowerPoint OpenXML templating engine for Node.js that generates and edits PPTX files directly through XML manipulation without relying on PowerPoint generation libraries.

[![npm version](https://img.shields.io/npm/v/node-pptx-templater.svg)](https://www.npmjs.com/package/node-pptx-templater)
[![CI](https://github.com/jsuyog2/node-pptx-templater/actions/workflows/ci.yml/badge.svg)](https://github.com/jsuyog2/node-pptx-templater/actions/workflows/ci.yml)
[![Coverage](https://img.shields.io/codecov/c/github/jsuyog2/node-pptx-templater)](https://codecov.io/gh/jsuyog2/node-pptx-templater)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](./LICENSE)
[![Node.js](https://img.shields.io/badge/node-%3E%3D18.0.0-brightgreen)](https://nodejs.org)

---

## ✨ Features

- 🏗️ **Zero PPTX Library Dependencies**: Operates entirely via low-level XML/ZIP manipulation.
- 🔁 **Advanced Text Replacement**: Handles fragmented runs seamlessly (`{{placeholders}}`).
- 📊 **Comprehensive Chart Support**: Bar, Column, Line, Pie, Doughnut, Area, Scatter—preserves original themes.
- 📋 **Safe Table Updates**: Fully manages row insertion, deletion, merging, resizing, and cell formatting while maintaining PowerPoint integrity (prevents repair warnings by generating unique row IDs).
- 🎨 **Shape & Image Manipulation**: Update text, clone, position, replace, or delete slide components.
- 🎯 **Slide Operations**: Duplicate, reorder, insert, or delete slide parts on the fly.
- 🔍 **Deep Integrity Validation**: Live validation of relationships, XML schema consistency, table columns, and duplicates prior to saving.

---

## 📦 Installation

```bash
npm install node-pptx-templater
```

---

## 🚀 Quick Start

```js
const { PPTXTemplater } = require('node-pptx-templater');

async function run() {
  // Load presentation template
  const ppt = await PPTXTemplater.load('template.pptx');
  
  // Use Slide 1 and replace text placeholders
  ppt.useSlide(1)
     .replaceTextByTag('title', 'Annual Sales Report')
     .replaceMultiple({ company: 'Google DeepMind', year: '2026' });

  // Update chart data on Slide 2
  ppt.useSlide(2)
     .updateChartData('sales-chart', {
       categories: ['Q1', 'Q2', 'Q3', 'Q4'],
       series: [{ name: 'Revenue', values: [100, 150, 180, 220] }]
     });

  // Save the result (validates structural integrity automatically)
  await ppt.saveToFile('output.pptx');
}
run();
```

---

## 🏗️ Architecture & OpenXML Internals

A PPTX file is an OPC (Open Packaging Convention) ZIP containing XML parts:

- `/ppt/presentation.xml` – The slides inventory (`sldIdLst`) and masters.
- `/ppt/slides/slideN.xml` – Slide elements (shapes, images, tables).
- `/ppt/slides/_rels/slideN.xml.rels` – Relationship links (`rId`).
- `[Content_Types].xml` – Declares MIME types for slide parts, charts, and media.

### The rowId Table Corruption Bug
Microsoft PowerPoint assigns an `<a16:rowId>` identifier to each row to facilitate collaborative editing. Duplicate row IDs trigger the PowerPoint **Repair Mode**, causing slide elements and styles to break.
`node-pptx-templater` generates unique 32-bit unsigned integers for `<a16:rowId>` whenever a table row is cloned, appended, or inserted, ensuring output documents open flawlessly in MS PowerPoint, Google Slides, and LibreOffice.

---

## 📚 API Reference

Here is the complete reference of all public APIs.

### Slide Features

#### `duplicateSlide(slideIndex, atPosition)`
Duplicates a slide.
```js
ppt.duplicateSlide(1, 2); // Duplicate Slide 1 and insert it as Slide 2
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

#### `insertSlide(slideIndex, options)`
Inserts a new blank slide at a specific position.
```js
ppt.insertSlide(2, { title: 'New Layout Slide' });
```

#### `getSlides()`
Returns metadata about all slides in the deck.
```js
const slides = ppt.getSlides();
console.log(slides);
```

---

### Table Features

#### `addTableRow(tableId, rowData)`
Appends a new row to the table.
```js
ppt.addTableRow('sales-table', ['Q4', '150', '210', '190', '250']);
```

#### `removeTableRow(tableId, rowIndex)`
Removes a table row.
```js
ppt.removeTableRow('sales-table', 1); // Delete 2nd row (0-based)
```

#### `insertTableRow(tableId, rowIndex, rowData)`
Inserts a row at a specific index.
```js
ppt.insertTableRow('sales-table', 2, ['Q3', '100', '120', '140', '160']);
```

#### `cloneTableRow(tableId, sourceRowIndex, targetRowIndex)`
Clones a row style/data and inserts it.
```js
ppt.cloneTableRow('sales-table', 1, 3);
```

#### `updateCell(tableId, rowIndex, colIndex, value, options)`
Updates a specific cell value and styling.
```js
ppt.updateCell('sales-table', 1, 0, 'New Product Name', {
  fill: 'FF0000',     // HEX background color
  align: 'ctr',       // Text alignment: 'l', 'ctr', 'r', 'just'
  fontSize: 14        // Font size in points
});
```

#### `mergeCells(options)` or `mergeCells(tableId, startRow, startCol, endRow, endCol)`
Merges a rectangular range of cells. Supports horizontal, vertical, and block merges, moving cell texts to the origin cell (concatenated with newlines) and setting correct OpenXML `gridSpan` and `rowSpan` attributes.
```js
// Config object signature (Recommended)
ppt.mergeCells({
  slide: 3,
  tableId: 'sales-table',
  startRow: 1,
  startCol: 1,
  endRow: 3,
  endCol: 3
});

// Legacy positional signature
ppt.mergeCells('sales-table', 1, 1, 3, 3);
```

#### `unmergeCells(options)` or `unmergeCells(tableId, startRow, startCol, endRow, endCol)`
Unmerges cells, restoring original XML cell structures.
```js
// Cell-coordinate coordinate signature (Recommended)
ppt.unmergeCells({
  slide: 3,
  tableId: 'sales-table',
  row: 2,
  col: 2
});

// Legacy positional signature
ppt.unmergeCells('sales-table', 1, 1, 3, 3);
```

#### `getMergedCells(tableId)`
Scans the slide table and returns all active merged regions.
```js
const merges = ppt.getMergedCells('sales-table');
// Output: [ { startRow: 1, startCol: 1, endRow: 3, endCol: 3 } ]
```

#### `validateMergeRegion(tableId, startRow, startCol, endRow, endCol)`
Checks bounds and overlaps, returning detailed validation errors.
```js
const report = ppt.validateMergeRegion('sales-table', 1, 1, 3, 3);
console.log(report.valid); // true or false
console.log(report.errors); // Array of error strings
```

#### `isMergedCell(tableId, row, col)`
Returns `true` if the cell at `(row, col)` is part of any merged region.
```js
const merged = ppt.isMergedCell('sales-table', 2, 2);
```

#### `getMergeParent(tableId, row, col)`
Returns the coordinates `{ row, col }` of the top-left origin cell of the merge region containing `(row, col)`.
```js
const parent = ppt.getMergeParent('sales-table', 2, 2); // { row: 1, col: 1 }
```

#### `getMergeRegion(tableId, row, col)`
Returns the merged region object `{ startRow, startCol, endRow, endCol }` containing `(row, col)`.
```js
const region = ppt.getMergeRegion('sales-table', 2, 2);
```

#### `splitMergedRegion(tableId, row, col)`
Splits the merged region containing cell `(row, col)`.
```js
ppt.splitMergedRegion('sales-table', 2, 2);
```

#### `cloneMergedRegion(tableId, row, col, targetRow, targetCol)`
Clones a merged region to another starting position, preserving cell text and formatting.
```js
ppt.cloneMergedRegion('sales-table', 1, 1, 4, 1);
```

#### Template-driven cell merges
Support cell merging inside template updates dynamically via `colSpan`/`rowSpan` or `merge` arrays:
```js
ppt.updateTable('sales-table', {
  rows: [
    ['Header 1', 'Header 2', 'Header 3'],
    ['Row 1 Col 1', { value: 'Spanned Cell', colSpan: 2 }],
    ['Row 2 Col 1', 'Row 2 Col 2', { value: 'Spanned V', rowSpan: 2 }]
  ],
  merge: [
    { startRow: 0, startCol: 0, endRow: 0, endCol: 2 }
  ]
});
```

#### `autoFitTable(tableId)`
Resizes columns to fit text width.
```js
ppt.autoFitTable('sales-table');
```

#### `resizeTable(tableId, width, height)`
Resizes the bounding frame of the table. Width/height can be in EMUs or inches.
```js
ppt.resizeTable('sales-table', 8.5, 4.2); // Dimensions in inches
```

#### `getTables()`
Lists tables in the current slide.
```js
const tables = ppt.getTables();
```

---

### Chart Features

#### `updateChartData(chartId, data)`
Updates a chart's categories, series, values, and embedded Excel workbook.
```js
ppt.updateChartData('sales-chart', {
  categories: ['Q1', 'Q2', 'Q3'],
  series: [{ name: 'Sales', values: [100, 120, 150] }]
});
```

#### `replaceChartSeries(chartId, seriesIndex, newSeriesData)`
Replaces values and name of a single series.
```js
ppt.replaceChartSeries('sales-chart', 0, {
  name: 'Updated Series Name',
  values: [80, 95, 110]
});
```

#### `updateChartTitle(chartId, title)`
Updates the chart's title.
```js
ppt.updateChartTitle('sales-chart', 'Quarterly Revenue Performance');
```

#### `updateChartCategories(chartId, categories)`
Updates the chart categories.
```js
ppt.updateChartCategories('sales-chart', ['Jan', 'Feb', 'Mar']);
```

#### `getCharts()`
Returns chart metadata from the slide.
```js
const charts = ppt.getCharts();
```

---

### Text Features

#### `replaceTextByTag(tag, value, options)`
Replaces placeholders with custom values.
```js
ppt.replaceTextByTag('username', 'Alice Cooper');
```

#### `replaceMultiple(replacements, options)`
Performs multiple text tag replacements.
```js
ppt.replaceMultiple({
  'date': '2026-06-02',
  'location': 'New York'
});
```

#### `findText(text)`
Searches for text in slide shape runs.
```js
const matches = ppt.findText('DeepMind');
```

#### `getTextElements()`
Gets all raw text segments in the selected slide.
```js
const elements = ppt.getTextElements();
```

---

### Shape Features

#### `updateShapeText(shapeId, text)`
Updates text inside a shapes run.
```js
ppt.updateShapeText('HeaderShape', 'Updated Slide Header');
```

#### `cloneShape(shapeId, newShapeId, options)`
Clones a shape and places it with offsets.
```js
ppt.cloneShape('HeaderShape', 'HeaderShapeCopy', {
  offsetX: 1.0,  // Offset X in inches
  offsetY: 0.5,  // Offset Y in inches
  width: 5.0,    // Resize width
  height: 2.0    // Resize height
});
```

#### `deleteShape(shapeId)`
Deletes a shape.
```js
ppt.deleteShape('HeaderShapeCopy');
```

#### `getShapes()`
Lists shapes on the slide.
```js
const shapes = ppt.getShapes();
```

---

### Image Features

#### `replaceImage(imageIdOrName, sourcePathOrBuffer)`
Replaces an image file binary, keeping the layout.
```js
await ppt.replaceImage('LogoImage', 'path/to/new-logo.png');
```

#### `addImage(sourcePathOrBuffer, options)`
Embeds a new image with layout options.
```js
await ppt.addImage('path/to/badge.png', {
  x: 2.0,       // Position X (inches)
  y: 1.5,       // Position Y (inches)
  width: 3.0,   // Width (inches)
  height: 3.0   // Height (inches)
});
```

#### `removeImage(imageIdOrName)`
Deletes an image.
```js
ppt.removeImage('Picture 1002');
```

#### `getImages()`
Lists images inside the slide.
```js
const images = ppt.getImages();
```

---

### Validation System

#### `validatePresentation()`
Audits the complete presentation.
```js
const report = await ppt.validatePresentation();
console.log('Errors:', report.errors);
console.log('Warnings:', report.warnings);
```

#### `validateSlide(slideIndex)`
Validates slide XML structure.
```js
const report = await ppt.validateSlide(1);
```

#### `validateTable(tableId)`
Audits columns and duplicate rowIds in a table.
```js
const report = await ppt.validateTable('sales-table');
```

#### `validateRelationships(partPath)`
Audits `.rels` relationship references.
```js
const report = ppt.validateRelationships('ppt/slides/slide1.xml');
```

---

## ⚡ Performance Benchmarks

Tested on a 50-slide presentation template:

| Operation | Average Duration |
|---|---|
| Load presentation | ~120ms |
| Replace 20 text tags | ~2ms |
| Audit presentation structure | ~15ms |
| Update table rows (10 rows) | ~3ms |
| Rebuild ZIP and Save | ~80ms |

---

## ❓ FAQ

#### Why use this over libraries like pptxgenjs or officegen?
Those libraries generate presentations from scratch in code. `node-pptx-templater` is a **templating** engine. You create a beautiful template inside MS PowerPoint (applying themes, layout grids, animations, or styling), and then use this library to dynamically populate slides, update tables, and refresh chart values while preserving the design.

#### How is chart styling preserved?
We update only the `<c:cat>` (categories), `<c:val>` (values) XML nodes, and the underlying data sheet inside the embedded Excel workbook. PowerPoint reads these updated values and uses the template's pre-configured colors, font layouts, labels, and axes.

---

## 🤝 Contributing

See [CONTRIBUTING.md](./CONTRIBUTING.md).

```bash
git clone https://github.com/jsuyog2/node-pptx-templater.git
cd node-pptx-templater
npm install
npm test
```

---

## 📄 License

MIT © [node-pptx-templater contributors](./LICENSE)
