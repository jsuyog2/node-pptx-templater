# node-pptx-templater

<div align="center">

[![npm version](https://img.shields.io/npm/v/node-pptx-templater.svg?style=flat-square&color=blue)](https://www.npmjs.com/package/node-pptx-templater)
[![npm downloads](https://img.shields.io/npm/dm/node-pptx-templater.svg?style=flat-square&color=blue)](https://www.npmjs.com/package/node-pptx-templater)
[![Vitest Tests](https://img.shields.io/badge/tests-338%20passing-brightgreen.svg?style=flat-square)](https://github.com/jsuyog2/node-pptx-templater/actions)
[![Bundle Size](https://img.shields.io/badge/bundle%20size-pure%20JS%20%3C2MB-success.svg?style=flat-square)](https://www.npmjs.com/package/node-pptx-templater)
[![Serverless](https://img.shields.io/badge/serverless-AWS%20Lambda%20%7C%20Cloudflare%20Workers-purple.svg?style=flat-square)](https://node-pptx-templater.pages.dev/guides/serverless-deployment/)
[![Node.js Support](https://img.shields.io/badge/node-%3E%3D20.12.0-blue.svg?style=flat-square)](https://nodejs.org)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg?style=flat-square)](https://opensource.org/licenses/MIT)

**Low-Level OpenXML PowerPoint Template Engine for Node.js**  
**Serverless PPTX Generation for AWS Lambda, Vercel & Cloudflare**

[📖 Documentation Portal](https://node-pptx-templater.pages.dev/) &nbsp;•&nbsp; [⚡ 5-Minute Quick Start](https://node-pptx-templater.pages.dev/docs/installation/) &nbsp;•&nbsp; [📚 API Reference](https://node-pptx-templater.pages.dev/docs/api/)

</div>

---

## 🚀 Low-Level OpenXML & Flat OPC PowerPoint Engine for Node.js

**`node-pptx-templater`** is an enterprise-grade, high-performance, spatial-aware PowerPoint OpenXML manipulation and templating engine built for **Node.js (>=20.12.0)**, **AWS Lambda**, **Cloudflare Workers**, **Vercel Edge**, and containerless cloud environments.

Unlike legacy solutions that require spinning up heavyweight LibreOffice Docker containers (500MB+ images, 15-second cold starts), headless browser instances, or native Java runtimes, `node-pptx-templater` is **100% pure JavaScript with zero native binary dependencies**. It mutates and synthesizes OpenXML presentation packages entirely in memory without corrupting PresentationML relationships, layouts, or formatting.

### 🌟 Architectural Highlights

- **Multi-Format Polymorphic Ingestion:** Seamlessly loads `.pptx`, `.pptm`, `.potx`, `.potm`, `.ppsx`, `.ppsm`, `.ppam`, legacy binary `.ppt`, OpenDocument `.odp`, raw ZIP buffers, uncompressed XML directories, remote HTTP/HTTPS URLs.
- **Microsoft Flat OPC XML Support (`<pkg:package>`):** Ingest and export single-file XML presentations (`sample.xml`) directly with zero ZIP overhead. Decodes embedded base64 media and workbooks into the in-memory Virtual File System (VFS).
- **Template-Free Blank Deck Generation:** Create pristine, valid OpenXML presentations from scratch with exact aspect ratios (`16:9`, `4:3`, `16:10`) or custom EMU/pixel dimensions without any template fixtures.
- **Anti-Corruption Run Fragment Healing:** PowerPoint frequently splits placeholder tokens like `{{COMPANY_NAME}}` into disjoint `<a:r>` XML run nodes. Our AST paragraph normalizer merges split runs while preserving font styles, weights, colors, and layout intact.
- **Excel Dual-Sync Technology & 17 Chart Types:** Support for all 17 chart types across standard OpenXML (`<c:chartSpace>`) and Office 2016+ extended schemas (`<cx:chartSpace>`): Column, Line, Pie, Bar, Area, Scatter, Map, Stock, Surface, Radar, Treemap, Sunburst, Histogram, Box & Whisker, Waterfall, Funnel, and Combo. Strict 11-method whitelist (`addChart`, `updateChart`, `updateChartData`, `removeChart`, `clearCharts`, `getCharts`, `getChart`, `cloneChart`, `getChartLabel`, `setChartLabel`, `updateChartLabel`). Simultaneously updates DrawingML chart caches and embedded binary Excel workbooks (`ppt/embeddings/*.xlsx`) with table cleansing to eliminate Microsoft PowerPoint repair prompts.
- **Spatial Geometry & Percentage Coordinates (0–100%):** Position shapes, tables, charts, images, and text boxes using intuitive percentage strings (`x: '10%'`, `y: '50%'`, `width: '80%'`) or numeric percentages (`unit: '%'`) mapped precisely across canvas boundaries without rounding drift.
- **Non-Blocking In-Memory Lifecycles:** Agile presentation synthesis without artificial validation hurdles. `validatePresentation()` serves as an on-demand, non-blocking diagnostic tool.
- **Spatial Geometry & Z-Order Layering:** Reorder visual stack depth (`bringToFront`, `sendToBack`, `stepForward`, `stepBackward`) and calculate coordinates using English Metric Units (EMU), pixels, inches, and millimeters with zero floating-point rounding drift.
- **Multi-Target Export & Web Framework Streaming:** Write to disk, compile to Node.js `Buffer`, `ReadableStream`, `ArrayBuffer`, Base64 / Data URI, export to Flat OPC XML string, stream directly to Express.js / Fastify responses, or emit Web Standard `Response` objects for Next.js Route Handlers and Edge workers.
- **Direct XML & ZIP Extraction:** Unpack presentations to raw directory trees with optional 2-space XML formatting (`formatXml: true`) for git inspection, or package directly to clean compressed ZIP archives.

---

### 📁 Input & Output Format Matrix

| Format / Source Type | Extension / Signature | Ingestion Support | Export Support | Zero Native Dependencies |
| :--- | :--- | :---: | :---: | :---: |
| **OpenXML Presentation** | `.pptx` | ✅ Native VFS | ✅ `.saveToFile()`, `.toBuffer()` | ✅ Pure JS |
| **Macro-Enabled Presentation** | `.pptm` | ✅ Native VFS | ✅ Preserved Macros | ✅ Pure JS |
| **PowerPoint Template** | `.potx`, `.potm` | ✅ Native VFS | ✅ Direct OpenXML compile | ✅ Pure JS |
| **PowerPoint Slide Show** | `.ppsx`, `.ppsm` | ✅ Native VFS | ✅ Direct OpenXML compile | ✅ Pure JS |
| **Microsoft Flat OPC XML** | `.xml`, `<pkg:package>` | ✅ Direct XML Parse (No ZIP) | ✅ `.toFlatXml()` | ✅ Pure JS |
| **Uncompressed XML Folder** | Directory (`[Content_Types].xml`, `ppt/`) | ✅ VFS Direct Mount | ✅ `.extractToXml()` | ✅ Pure JS |
| **Raw ZIP Buffer / Stream** | `Buffer`, `Uint8Array`, `ReadableStream` | ✅ In-Memory Streaming | ✅ `.toBuffer()`, `.toStream()` | ✅ Pure JS |
| **Remote HTTP/HTTPS URL** | Remote Endpoint | ✅ `PPTXTemplater.load(url)` | ✅ Direct Web Response | ✅ Pure JS |
| **Legacy & Alternative Formats** | `.ppt`, `.odp` | ✅ Header Detection & OPC Ingestion | ✅ Normalizes to OpenXML | ✅ Pure JS |

---

## ⚡ Flat OPC XML & Serverless PPTX Generator for AWS Lambda & Cloudflare

Engineered specifically for serverless and edge architectures:
- **Instant Cold Boot:** < 15ms initialization time.
- **Micro Memory Footprint:** < 2MB runtime overhead.
- **Zero Disk Thrashing:** 100% in-memory Virtual File System (VFS) operations.
- **Isomorphic Output:** Standard Web `Response` and `ArrayBuffer` ready for Cloudflare Workers, Next.js Route Handlers, Remix, and Vercel Edge.

---

## 📊 Feature Comparison Matrix

| Capability / Feature | `node-pptx-templater` | `pptxgenjs` | `officegen` | `LibreOffice Headless` |
| :--- | :---: | :---: | :---: | :---: |
| **Existing PPTX Template Ingestion** | ✅ **Full In-Place Mutation** | ❌ Blank only | ❌ Blank only | ⚠️ Macro scripting |
| **Multi-Format (.pptm, .potx, .ppt, .odp)** | ✅ **Native Auto-Conversion** | ❌ Unsupported | ❌ Unsupported | ⚠️ External binary |
| **Remote Web URL** | ✅ **`PPTXTemplater.load()`** | ❌ Unsupported | ❌ Unsupported | ❌ Unsupported |
| **Template-Free Blank Generation** | ✅ **`createBlank()` (16:9, 4:3)** | ✅ Supported | ⚠️ Basic | ⚠️ Macro required |
| **Run Fragment Normalizer (Anti-Corruption)** | ✅ **AST Paragraph Healer** | ❌ N/A | ❌ N/A | ❌ Brittle regex |
| **Chart & Excel Workbook Dual-Sync** | ✅ **XML + XLSX Cache Sync** | ⚠️ Basic values only | ❌ Unsupported | ⚠️ Slow headless calc |
| **Table Overflow Auto-Pagination** | ✅ **`splitTableOnOverflow()`** | ❌ Manual math | ❌ Unsupported | ❌ Unsupported |
| **Z-Order Layer Stacking (`p:spTree`)** | ✅ **`bringToFront` / `sendToBack`** | ❌ Re-render required | ❌ Unsupported | ❌ Manual scripts |
| **Web Framework & Edge Streaming** | ✅ **Express, Fastify, Web Response**| ⚠️ Buffer only | ⚠️ Stream only | ❌ CLI process spawn |
| **Direct XML Extraction with Formatting** | ✅ **`extractToXml({formatXml})`**| ❌ Unsupported | ❌ Unsupported | ❌ Unsupported |
| **Serverless Deployment (Lambda, Cloudflare)** | ✅ **<2MB Pure JS, <15ms Boot** | ✅ Pure JS | ⚠️ High RAM | ❌ Heavy 500MB+ Docker |

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

### 1. Multi-Format Import (PPTX, PPTM, POTX, ODP, Binary PPT, ZIP & Folders)

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function importDecks() {
  // 1. Standard OpenXML presentation or macro-enabled template
  const pptx = await PPTXTemplater.load('./templates/annual-report.pptx');
  const pptm = await PPTXTemplater.load('./templates/macros.pptm');

  // 2. Legacy Microsoft PowerPoint (.ppt) or OpenDocument (.odp)
  const legacyPpt = await PPTXTemplater.load('./templates/legacy-presentation.ppt');
  const openDoc = await PPTXTemplater.load('./templates/slides.odp');

  // 3. Raw ZIP buffer or uncompressed directory tree
  const bufferPpt = await PPTXTemplater.load(zipBuffer);
  const folderPpt = await PPTXTemplater.load('./unpacked-presentation-folder');

  // 4. Microsoft Flat OPC XML package (single-file XML, zero ZIP overhead)
  const flatXmlPpt = await PPTXTemplater.load('./templates/sample.xml');

  console.log(`Loaded ${pptx.slideCount} slides successfully.`);
}
```

### 2. Flat OPC Single-File XML Ingestion & Export (Zero ZIP Overhead)

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function handleFlatOpc() {
  // Ingest directly from Flat OPC XML file, XML string, or Buffer
  const ppt = await PPTXTemplater.load('./templates/sample.xml');

  ppt.useSlide(1).replaceText({ '{{METRIC}}': '$1.2M ARR' });

  // Export to standard .pptx binary OR serialize to Flat OPC XML (disk or string)
  await ppt.saveToFile('./output/converted-deck.pptx');
  await ppt.toFlatXml('./output/converted-deck.xml');
  const flatXml = await ppt.toFlatXml();
}
```

### 3. Template-Free Blank Presentation Creation

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function createFromScratch() {
  // Create 16:9 widescreen presentation (12192000 × 6858000 EMU)
  const ppt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' });

  // Or create with 4:3 standard, or custom pixel dimensions
  // const ppt = await PPTXTemplater.createBlank({ width: 1920, height: 1080, unit: 'px' });

  ppt.useSlide(1);
  await ppt.addShape('roundedRectangle', {
    id: 'heroCard',
    x: 100, y: 100, width: 600, height: 200,
    fill: '#2563EB', text: 'Synthesized with node-pptx-templater'
  });

  await ppt.saveToFile('./output/brand-new-deck.pptx');
}
```

### 4. Fastify & Express.js PowerPoint Streaming

```javascript
import { PPTXTemplater } from 'node-pptx-templater';

const ppt = await PPTXTemplater.load('./templates/report.pptx');
ppt.useSlide(1).replaceText({ '{{STATUS}}': 'Approved' });

// 1. Commit directly to local disk or unpacked XML folder
await ppt.saveToFile('./output/ApprovedReport.pptx');
await ppt.extractToXml('./output/unpacked-tree');

// 2. Compile to Node.js Buffer or ReadableStream
const buffer = await ppt.toBuffer();
const stream = await ppt.toStream();

// 3. Web-standard ArrayBuffer or Base64 / Data URI (Edge runtimes)
const arrayBuffer = await ppt.toArrayBuffer();
const dataUri = await ppt.toBase64({ dataUri: true });

// 4. Express.js HTTP response streaming
app.get('/api/report/express', async (req, res) => {
  await ppt.toExpress(res, 'ApprovedReport.pptx');
});

// 5. Fastify reply streaming
fastify.get('/api/report/fastify', async (request, reply) => {
  await ppt.toFastify(reply, 'ApprovedReport.pptx');
});

// 6. Next.js App Router / Cloudflare Workers Web API Response
export async function GET(request) {
  return await ppt.toResponse({ filename: 'ApprovedReport.pptx' });
}

// 7. High-concurrency template caching (< 1ms clone)
await PPTXTemplater.preload('./templates/report.pptx');
const fastPpt = await PPTXTemplater.fromCache('./templates/report.pptx');
```

### 5. XML & ZIP Direct Extraction for Git Inspection

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function inspectPresentation() {
  const ppt = await PPTXTemplater.load('./presentation.pptx');

  // Extract all XML parts into a folder with 2-space formatted XML for git diffs
  await ppt.extractToXml('./inspection-folder', { formatXml: true });

  // Export clean, compressed ZIP archive
  await ppt.extractToZip('./presentation-clean.zip');
}
```

### 6. Dynamic Text Replacement, Fragment Healing & Text Box ID Targeting

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function replacePresentationText() {
  const ppt = await PPTXTemplater.load('./templates/executive-deck.pptx');

  // 1. In-place template token replacement with AST fragment run healing
  ppt.useSlide(1).replaceText({
    '{{COMPANY}}': 'Acme Global Technologies Inc.',
    '{{QUARTER}}': 'Q3 2026',
    '{{PRESENTER}}': 'Sarah Jenkins, VP Engineering'
  });

  // 2. Full text box replacement by shape ID or name with formatting preservation
  ppt.replaceTextByTag('SummaryBox', 'Key Takeaways:\n• Revenue up 38%\n• Retention healthy', {
    bold: true
  });

  await ppt.saveToFile('./output/updated-deck.pptx');
}
```

### 7. Native Charts Engine (11 Whitelisted Methods, 17 Chart Types & Excel Dual-Sync)

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function chartOperations() {
  const ppt = await PPTXTemplater.load('./templates/financial-report.pptx');

  // 1. Add standard or Office 2016+ extended charts (e.g. Column, Waterfall, Combo)
  ppt.addChart(1, 'waterfall', {
    categories: ['Starting', 'Revenue', 'COGS', 'Taxes', 'Ending'],
    series: [{ name: 'Cash Flow', values: [1000, 450, -300, -100, 1050] }]
  }, { name: 'CashWaterfall', title: 'Enterprise Cash Waterfall' });

  // 2. Synchronize XML cache AND embedded .xlsx workbook simultaneously
  ppt.useSlide(2).updateChartData('QuarterlyRevenueChart', {
    categories: ['Q1', 'Q2', 'Q3', 'Q4 (Proj)'],
    series: [
      { name: 'Enterprise Cloud', values: [18.4, 22.1, 26.8, 31.5] },
      { name: 'AI Solutions', values: [8.2, 12.6, 17.4, 23.0] }
    ]
  });

  // 3. Clone charts deeply with unique XML parts and workbooks
  ppt.useSlide(2).cloneChart('QuarterlyRevenueChart', 'RevenueChart_Copy', { x: '5in', y: '2in' });

  // 4. Data labels manipulation
  ppt.useSlide(2).setChartLabel('RevenueChart_Copy', true, { position: 'outsideEnd' });

  // 5. Slide-wide clean evacuation while preserving text boxes, layout shapes, and tables
  ppt.useSlide(3).clearCharts();

  await ppt.saveToFile('./output/synced-charts.pptx');
}
```

### 8. Dynamic Tables, Nested Rows & Overflow Auto-Pagination

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function generateTable() {
  const ppt = await PPTXTemplater.load('./templates/inventory.pptx');
  ppt.useSlide(3);

  // 1. Append simple database rows
  for (const item of inventoryRecords) {
    ppt.addTableRow('InventoryTable', [item.sku, item.name, item.warehouse, item.qty]);
  }

  // 2. Append nested sub-rows with vertical merge strategy
  await ppt.addTableRow('InventoryTable', ['West Region', ['Electronics', '450'], ['Hardware', '120'], 'In Stock'], {
    mergeStrategy: 'rowspan'
  });

  // 3. Extract table rows as JSON with full layout metadata
  const meta = await ppt.getTableRows('InventoryTable', { includeMetadata: true });
  console.log(`Total rows: ${meta.rowCount}, Merges: ${meta.mergedCells.length}`);

  // 4. Alternate row background colors
  ppt.stripeRows('InventoryTable', '#F8FAFC', '#FFFFFF');

  // 5. Automatically split table across cloned slides if it overflows slide height
  await ppt.splitTableOnOverflow('InventoryTable', {
    bottomMargin: '1in',
    repeatHeaders: 1
  });

  await ppt.saveToFile('./output/paginated-table.pptx');
}
```

### 9. Slide Lifecycle, Section Management & Collaborative Comments (Strict 23 Methods)

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function manageSlideDeck() {
  const ppt = await PPTXTemplater.load('./templates/annual-deck.pptx');

  // 1. Fluent slide selection & metadata inspection
  ppt.useSlide(1);
  const slides = ppt.getSlides();
  const slide1 = ppt.getSlide(1);
  const dims = ppt.getSlideDimensions();
  const layout = ppt.getSlideLayout(1);

  // 2. Lifecycle operations: add, duplicate, reorder, move, tag & remove
  await ppt.addSlide({ title: 'Executive Summary', layout: 'Title and Content' }, 2);
  await ppt.duplicateSlide(2, 3);
  ppt.moveSlide(3, 1);
  ppt.reorderSlides([1, 2, 3]);
  ppt.tagSlide(1, 'reviewed');

  // 3. Section organization (<p14:sectionLst>)
  await ppt.addSection('Executive Overview', 1);
  ppt.useSection('Executive Overview');
  ppt.updateSection('Executive Overview', { name: 'Leadership Overview' });
  const section = ppt.getSection('Leadership Overview');

  // 4. Slide visibility & canvas reset
  ppt.hideSlide(2);
  ppt.unhideSlide(2);
  ppt.resetSlide(3, { preservePlaceholders: true });

  // 5. Collaborative review comments (ppt/commentAuthors.xml & ppt/comments/commentN.xml)
  await ppt.addComment('Please verify Q3 ARR metrics prior to board presentation.', {
    author: 'Sarah Jenkins',
    initials: 'SJ',
    x: 100, y: 150
  });
  await ppt.removeComment(0);

  // 6. Cross-deck slide import & selective export
  const sourceDeck = await PPTXTemplater.load('./templates/appendix.pptx');
  await ppt.importSlideFrom(sourceDeck, 1, 4);
  const exportedBuffer = await ppt.exportSlides(1, 2);

  await ppt.saveToFile('./output/final-presentation.pptx');
}
```

### 10. Images & Media Lifecycle (Strict 12 Methods)

```javascript
const { PPTXTemplater } = require('node-pptx-templater');

async function manageMedia() {
  const ppt = await PPTXTemplater.load('./templates/multimedia.pptx');

  // 1. Image Lifecycle: Add, Update, Remove
  await ppt.AddImage(1, './assets/logo.png', {
    id: 'companyLogo',
    x: '5%', y: '5%', width: '20%', height: '10%',
    altText: 'Company Logo',
    transparency: 10
  });

  await ppt.UpdateImage('companyLogo', {
    x: '10%',
    rotation: 45,
    crop: { left: 5, top: 5, right: 5, bottom: 5 }
  });

  await ppt.RemoveImage('companyLogo');

  // 2. Video Lifecycle: Add, Update, Remove
  await ppt.AddVideo(1, './assets/intro.mp4', {
    id: 'introVideo',
    x: 100, y: 150, width: 400, height: 250,
    autoplay: true,
    loop: false,
    posterFrame: './assets/intro-poster.png'
  });

  await ppt.UpdateVideo('introVideo', {
    volume: 80,
    startTime: 2000,
    endTime: 15000
  });

  await ppt.RemoveVideo('introVideo');

  // 3. Audio Lifecycle: Add, Update, Remove
  await ppt.AddAudio(1, './assets/narration.mp3', {
    id: 'slideNarration',
    x: 50, y: 50, width: 48, height: 48,
    autoplay: true,
    loop: true,
    hideDuringShow: true
  });

  await ppt.UpdateAudio('slideNarration', {
    volume: 50
  });

  await ppt.RemoveAudio('slideNarration');

  // 4. Querying & Deduplicated Cloning
  const allMedia = await ppt.GetMedias(1);
  const singleMedia = await ppt.GetMedia('introVideo');
  await ppt.CloneMedia('companyLogo', 'companyLogoClone', {
    x: '75%', y: '5%'
  });

  await ppt.saveToFile('./output/multimedia-presentation.pptx');
}
```

---

## 🎪 CLI Usage

```bash
# Build a PPTX from a template with data from JSON
node-pptx-templater build template.pptx output.pptx --data ./data.json

# Extract all OpenXML parts into a formatted folder for git inspection
node-pptx-templater extract deck.pptx --all --format --out ./extracted-deck

# Extract to clean ZIP archive
node-pptx-templater extract deck.pptx --zip ./deck.zip

# Validate PPTX package integrity
node-pptx-templater validate deck.pptx
```

---

## 🛠️ Verification & Testing

```bash
# Run complete test suite (338 passing Vitest tests)
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
