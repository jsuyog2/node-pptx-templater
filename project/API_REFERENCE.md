# `node-pptx-templater` v3.0.0-alpha.0 API Reference

This document comprehensively catalogs the public classes, methods, and helper functions available in the `node-pptx-templater` engine.

---

## 1. Ingestion (PackageLoader)

function: `static PPTXTemplater.load(source: string | Buffer | ReadableStream | ArrayBuffer, options?: LoadOptions): Promise<PPTXTemplater>`
usage: Initializes the template engine using the robust `PackageLoader`. It automatically detects and supports Direct PPTX buffers, file paths, Node streams, ZIP containers (via magic bytes), and raw unpacked OpenXML directory trees (loading directly into the VirtualFileSystem map).
example:
```typescript
import { PPTXTemplater } from 'node-pptx-templater';
import * as fs from 'fs';

// Load from File Path
const ppt = await PPTXTemplater.load('./template.pptx');

// Load from Buffer
const buffer = fs.readFileSync('./template.pptx');
const pptFromBuffer = await PPTXTemplater.load(buffer);

// Load from extracted XML directory
const pptFromDir = await PPTXTemplater.load('./unpacked_pptx_folder');
```

---

## 2. Presentation & Slide Management

function: `PPTXTemplater.useSlide(slideIndex: number | number[] | "all"): PPTXTemplater`
usage: Sets the active slide(s) context for subsequent chained operations. Uses 1-based indexing.
example:
```typescript
// Target slide 2
ppt.useSlide(2).replaceText({ '{{NAME}}': 'John Doe' });

// Target multiple slides
ppt.useSlide([1, 3, 5]).replaceText({ '{{DATE}}': '2026-10-10' });

// Target all slides
ppt.useSlide('all');
```

function: `PPTXTemplater.duplicateSlide(slideIndex: number, atPosition?: number): Promise<PPTXTemplater>`
usage: Deep-clones a slide (including layouts, relationships, and embedded media) and inserts it into the specified 1-based position.
example:
```typescript
// Clone slide 1 and place the new slide at index 2
await ppt.duplicateSlide(1, 2);
```

function: `PPTXTemplater.deleteSlide(slideIndex: number): PPTXTemplater`
usage: Removes a slide from the presentation and purges its orphaned relationships from the internal XML manifests.
example:
```typescript
ppt.deleteSlide(3); // Removes the 3rd slide
```

function: `PPTXTemplater.get slideCount(): number`
usage: Retrieves the total number of slides currently in the presentation.
example:
```typescript
console.log(`Total slides: ${ppt.slideCount}`);
```

---

## 3. Spatial Engine & Units

function: `toEmu(value: number | string): number`
usage: Utility function (exported from `src/utils/units.js`) that converts `pt`, `px`, `in`, `cm`, and `mm` to integer OpenXML English Metric Units (EMUs). Crucial for maintaining strict OpenXML coordinate compliance without floating-point distortion.
example:
```typescript
import { toEmu } from 'node-pptx-templater/src/utils/units';

const emuFromPx = toEmu('100px'); // 952500
const emuFromIn = toEmu('2in');   // 1828800
const emuFromNum = toEmu(1828800); // 1828800 (already EMU)
```

function: `PPTXTemplater.getBounds(objectIdOrName: string): {x: number, y: number, width: number, height: number}`
usage: Retrieves the calculated pixel bounding box (`a:off` and `a:ext`) for a specific shape, table, or chart.
example:
```typescript
ppt.useSlide(1);
const bounds = ppt.getBounds('LogoImage');
console.log(`Width in px: ${bounds.width}`);
```

function: `PPTXTemplater.setBounds(objectIdOrName: string, options: { x?: number|string, y?: number|string, width?: number|string, height?: number|string }): PPTXTemplater`
usage: Updates the bounding box dimensions of an element. Inputs are safely passed through the EMU conversion utility.
example:
```typescript
ppt.setBounds('HeaderBox', { width: '8in', height: '1.5in' });
```

function: `PPTXTemplater.move(objectIdOrName: string, options: { x: number|string, y: number|string }): PPTXTemplater`
usage: Convenience alias for `setBounds` that translates the element's position on the slide without affecting its dimensions.
example:
```typescript
ppt.move('Title', { x: '10px', y: '20px' });
```

function: `PPTXTemplater.stepForward(objectIdOrName: string): PPTXTemplater` / `PPTXTemplater.stepBackward(objectIdOrName: string): PPTXTemplater`
usage: Adjusts the internal Z-index (rendering order) of an element within the slide's `<p:spTree>` without recreating nodes.
example:
```typescript
// Bring a background watermark up one layer
ppt.stepForward('WatermarkShape');

// Push a shape backward
ppt.stepBackward('OverlayBox');
```

---

## 4. Element Tree Traversal

function: `PPTXTemplater.extractPlaceholders(): string[]`
usage: Parses the AST across active slides to return a unique array of text placeholder tokens (e.g., matching `{{ ... }}`). Merges fragmented OpenXML `<a:r>` text runs seamlessly before matching.
example:
```typescript
ppt.useSlide('all');
const tags = ppt.extractPlaceholders();
console.log(tags); // ['{{COMPANY_NAME}}', '{{FISCAL_YEAR}}']
```

function: `PPTXTemplater.getElements(): Array<{ id: string, name: string, type: string }>`
usage: Returns a flat array of all drawable objects on the targeted slide. 
example:
```typescript
const elements = ppt.useSlide(1).getElements();
const charts = elements.filter(e => e.type === 'chart');
```

---

## 5. Text & Shape Operations

function: `PPTXTemplater.replaceText(replacements: Record<string, string>): PPTXTemplater`
usage: Scans target slides and replaces template string keys with values. Utilizes the robust AST-level run normalizer to combat PowerPoint's tendency to split contiguous text across multiple `<a:r>` nodes.
example:
```typescript
ppt.useSlide(1).replaceText({
  '{{TITLE}}': 'Q3 Financial Review',
  '{{AUTHOR}}': 'Jane Smith'
});
```

---

## 6. Media & Images

function: `PPTXTemplater.replaceImage(imageIdOrName: string, sourcePathOrBuffer: string | Buffer): Promise<PPTXTemplater>`
usage: Replaces the binary target of an existing `<p:pic>` node while retaining its Z-index. It recalculates the new image's natural aspect ratio, updating the bounding box using `object-fit: contain` logic to lock constraints perfectly.
example:
```typescript
await ppt.useSlide(2).replaceImage('ProfilePic', './assets/new_avatar.png');
```

function: `PPTXTemplater.replaceShape(shapeIdOrName: string, sourcePathOrBuffer: string | Buffer): Promise<PPTXTemplater>`
usage: Destroys a standard `<p:sp>` (Shape) placeholder and seamlessly drops a new `<p:pic>` (Image) inside its exact bounding geometry, assigning a fresh `rId` relationship dynamically.
example:
```typescript
// Morph a grey rectangle into a company logo
await ppt.useSlide(1).replaceShape('RectanglePlaceholder', './logo.svg');
```

---

## 7. Table Engine

function: `PPTXTemplater.addTableRow(tableId: string, rowData: any[], options?: object): PPTXTemplater`
usage: Appends a row to a table. Extracts format templates from the preceding row and correctly establishes cell alignments.
example:
```typescript
ppt.addTableRow('Financials', ['Revenue', '$10,000', '$12,000', '+20%']);
```

function: `PPTXTemplater.mergeCells(tableId: string, startRow: number, startCol: number, endRow: number, endCol: number): PPTXTemplater`
usage: Applies `rowSpan` and `gridSpan` attributes across a bounding box of table cells, hiding the obscured cells (`vMerge`/`hMerge`) safely to prevent file corruption.
example:
```typescript
// Merge a header across 4 columns in the first row
ppt.mergeCells('DataTable', 0, 0, 0, 3);
```

---

## 8. Flagship Table Auto-Pagination

function: `PPTXTemplater.splitTableOnOverflow(tableId: string, options?: { bottomMargin?: number|string, repeatHeaders?: number }): Promise<PPTXTemplater>`
usage: Calculates physical row heights using text metrics. If the table surpasses the calculated slide boundary (SlideHeight - bottomMargin), it splices the overflowing rows, deep-clones the parent slide, injects the repeated headers, and recurses until all rows fit safely.
example:
```typescript
ppt.useSlide(3);
// Feed 50 rows into a table
for (let i = 0; i < 50; i++) {
  ppt.addTableRow('MyTable', [`Item ${i}`, i * 10, 'Active']);
}

// Automatically paginate the table across new slides
// Leaving a 1-inch bottom margin and repeating the first 1 header row
await ppt.splitTableOnOverflow('MyTable', { 
  bottomMargin: '1in', 
  repeatHeaders: 1 
});
```

---

## 9. Native Charts

function: `PPTXTemplater.updateChart(chartIdOrName: string, chartData: object): Promise<PPTXTemplater>`
usage: Updates native PowerPoint chart series and categories. Critically, it syncs the internal `<c:chartSpace>` AST mutations against the actual embedded `ppt/embeddings/Microsoft_Excel_Worksheet.xlsx` binary to prevent Microsoft Office "Repair Presentation" warnings.
example:
```typescript
await ppt.useSlide(4).updateChart('SalesChart', {
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'Hardware', values: [120, 150, 180, 200] },
    { name: 'Software', values: [80, 95, 110, 140] }
  ]
});
```

---

## 10. Hyperlinks

function: `PPTXTemplater.addHyperlink(objectIdOrName: string, targetUrl: string): PPTXTemplater`
usage: Wraps an element or text run with an external `r:id` hyperlink relationship pointing to a URL.
example:
```typescript
ppt.addHyperlink('WebsiteButton', 'https://example.com');
```

---

## 11. Export Adapters

function: `PPTXTemplater.saveToFile(filePath: string): Promise<void>`
usage: Flushes the VFS to disk as a complete ZIP PPTX binary.
example:
```typescript
await ppt.saveToFile('./output/Final_Report.pptx');
```

function: `PPTXTemplater.toBuffer(): Promise<Buffer>`
usage: Compiles the presentation natively to a Node Buffer.
example:
```typescript
const buffer = await ppt.toBuffer();
```

function: `PPTXTemplater.toArrayBuffer(): Promise<ArrayBuffer>`
usage: Standardizes the binary output for modern Web / Fetch API / Blob environments.
example:
```typescript
const ab = await ppt.toArrayBuffer();
```

function: `PPTXTemplater.toBase64(options?: { dataUri?: boolean }): Promise<string>`
usage: Renders the presentation payload into Base64 (optionally attaching the `data:application/vnd.openxmlformats...` URI scheme).
example:
```typescript
const uri = await ppt.toBase64({ dataUri: true });
```

function: `PPTXTemplater.toExpress(res: Response, filename?: string): Promise<void>`
usage: Convenience adapter that pipes the presentation Stream directly into an Express HTTP Response with correct disposition and MIME type headers.
example:
```typescript
app.get('/download', async (req, res) => {
  const ppt = await PPTXTemplater.load('template.pptx');
  ppt.useSlide(1).replaceText({ '{{USER}}': req.user.name });
  
  await ppt.toExpress(res, 'Custom_Report.pptx');
});
```
