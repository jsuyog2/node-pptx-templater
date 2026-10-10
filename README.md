# node-pptx-templater

[![npm version](https://img.shields.io/badge/version-3.0.0--alpha.0-blue.svg)](https://www.npmjs.com/package/node-pptx-templater)
[![Vitest Passing](https://img.shields.io/badge/tests-255%20passing-brightgreen.svg)](https://github.com/jsuyog2/node-pptx-templater/actions)
[![Zero Sync I/O](https://img.shields.io/badge/sync%20I%2FO-zero-success.svg)](#1-in-memory-virtual-file-system-vfs)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

> **Enterprise-grade, spatial-aware OpenXML PowerPoint manipulation, templating, and layout engine for Node.js.**

Operates with **zero external binary dependencies** (no LibreOffice, no Java, no headless browser) and executes pure non-blocking, in-memory OpenXML transformations.

---

## 🏛️ Architectural Overview

`node-pptx-templater` v3.0.0-alpha.0 is engineered from the ground up for mission-critical cloud backends, microservices, and serverless environments:

### 1. In-Memory Virtual File System (VFS)
- **Zero Synchronous / Blocking I/O:** Pure non-blocking architecture using `Uint8Array`, Node.js `stream.Readable`, and in-memory byte arrays.
- All internal transformations run against an in-memory Virtual File System (`Map<string, Uint8Array>`), eliminating disk thrashing and race conditions under concurrent load.

### 2. Strict OpenXML AST Normalizer (No Brittle Regex)
- PowerPoint fragments placeholder tokens across distinct `<a:r>` runs and `<a:t>` text nodes during user edits (e.g., `{{` in one node, variable name in another).
- The engine uses an AST-level run normalizer that merges adjacent text runs inside paragraph nodes (`<a:p>`) before token matching while strictly preserving formatting, fonts, colors, and OpenXML namespaces (`p:`, `a:`, `r:`, `c:`, `w:`).

### 3. Precise Spatial Coordinate Engine (EMUs)
- All coordinates and dimensions are strictly computed in OpenXML English Metric Units (EMUs):
  - $1\text{ pt} = 12,700\text{ EMUs}$
  - $1\text{ in} = 914,400\text{ EMUs}$
  - $1\text{ px (96 DPI)} = 9,525\text{ EMUs}$
- Integer arithmetic (`Math.round()`) is enforced when writing `<a:off>` and `<a:ext>` tags to prevent micro-pixel rounding artifacts.

### 4. Relationship (.rels) Isolation
- When duplicating slides, the engine generates isolated relationship files (`slideX.xml.rels`) with fresh `rId` keys to eliminate relationship collisions and PowerPoint repair warnings.

---

## ⚡ Feature Matrix

| Feature Domain | Capabilities |
| :--- | :--- |
| **Ingestion Pipeline** | Ingest from File Path, Buffer, `ReadableStream`, PKZip archives, or raw uncompressed XML folder trees. |
| **Elements Manipulation** | Placeholders, text styling, shape creation/updates, image swaps with aspect-ratio locks, and Z-index layering (`bringToFront`, `sendToBack`). |
| **Flagship Table Engine** | Table cell CRUD, `gridSpan`/`rowSpan` merging, cell formatting, nested cell shapes, and dynamic row height calculation. |
| **Table Auto-Pagination** | Recursively splits overflowing tables across duplicated cloned slides against $Y_{\text{limit}} = \text{SlideHeight} - \text{BottomMargin}$, carrying over repeated headers. |
| **Native Chart Updates** | Mutates `<c:chartSpace>` series/categories while automatically synchronizing the embedded Excel binary cache (`Microsoft_Excel_Worksheet*.xlsx`). |
| **Export Adapters** | 6 non-blocking targets: `toFile`, `toStream`, `toBuffer`, `toArrayBuffer`, `toBase64`, and `toExpress`. |

---

## 📦 Installation

```bash
npm install node-pptx-templater
# or
pnpm add node-pptx-templater
# or
yarn add node-pptx-templater
```

---

## 💻 Runnable TypeScript Examples

### 1. Ingestion: Raw Unpacked XML Folders vs. PPTX Buffers

```typescript
import { PPTXTemplater, PackageLoader } from 'node-pptx-templater'
import * as fs from 'fs/promises'

async function runIngestionDemo() {
  // A. Ingest directly from an unpacked OpenXML directory tree (zero intermediate zip step)
  const pptFromFolder = await PPTXTemplater.load('./templates/extracted-deck/')

  // B. Ingest from an in-memory byte Buffer (e.g., loaded from database or S3)
  const templateBuffer = await fs.readFile('./templates/corporate-deck.pptx')
  const pptFromBuffer = await PPTXTemplater.load(templateBuffer)

  // C. Ingest via low-level PackageLoader
  const vfs = await PackageLoader.load(templateBuffer)
  console.log(`VFS loaded with ${vfs.entries.size} OpenXML entries in memory.`)
}
```

---

### 2. Dynamic Table Auto-Pagination with Repeated Headers

Automatically slice and paginate tables across duplicated slides when data overflows the vertical boundary ($Y_{\text{limit}} = \text{SlideHeight} - \text{BottomMargin}$):

```typescript
import { PPTXTemplater } from 'node-pptx-templater'

async function runTablePaginationDemo() {
  const ppt = await PPTXTemplater.load('./templates/invoice-template.pptx')
  ppt.useSlide(1)

  // Populate 100 dynamic rows
  for (let i = 1; i <= 100; i++) {
    ppt.addTableRow('LineItemsTable', [
      `Item #${i}`,
      `Description for product ${i}`,
      `$${(Math.random() * 100).toFixed(2)}`,
      'Completed'
    ])
  }

  // Automatically split table across cloned slides with 1-inch bottom margin
  const result = await ppt.splitTableOnOverflow('LineItemsTable', {
    bottomMargin: '1in',  // 914,400 EMUs
    repeatHeaders: 1      // carry over top header row to all created slides
  })

  console.log(`Pagination complete: created ${result.slidesCreated} new slides.`)
  await ppt.toFile('./dist/invoice-paginated.pptx')
}
```

---

### 3. Native Chart Updates with Automatic Excel Cache Synchronization

Update PowerPoint charts without triggering "PowerPoint found a problem with content" repair dialogs:

```typescript
import { PPTXTemplater } from 'node-pptx-templater'

async function runChartUpdateDemo() {
  const ppt = await PPTXTemplater.load('./templates/analytics-deck.pptx')

  // Updates <c:chartSpace> AND synchronizes ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx
  await ppt.useSlide(2).updateChart('RevenueChart', {
    categories: ['Q1 2026', 'Q2 2026', 'Q3 2026', 'Q4 2026'],
    series: [
      { name: 'Direct Sales', values: [150000, 185000, 220000, 290000] },
      { name: 'Channel Partners', values: [85000, 110000, 140000, 195000] }
    ]
  })

  await ppt.toFile('./dist/analytics-updated.pptx')
}
```

---

### 4. Streaming Presentations Directly Over Express HTTP Responses

Stream generated PPTX binaries directly to client browsers with proper OpenXML MIME headers:

```typescript
import express from 'express'
import { PPTXTemplater } from 'node-pptx-templater'

const app = express()

app.get('/download-report', async (req, res) => {
  try {
    const ppt = await PPTXTemplater.load('./templates/monthly-report.pptx')

    ppt.useSlide(1).replaceText({
      '{{CLIENT_NAME}}': req.query.client as string || 'Valued Customer',
      '{{GENERATION_DATE}}': new Date().toLocaleDateString()
    })

    // Sets Content-Type: application/vnd.openxmlformats-officedocument.presentationml.presentation
    // Sets Content-Disposition: attachment; filename="Report.pptx"
    // Pipes the readable stream directly to res with zero intermediate disk writes
    await ppt.toExpress(res, 'Custom_Report.pptx')
  } catch (err) {
    res.status(500).send('Error generating presentation')
  }
})

app.listen(3000, () => console.log('Server running on port 3000'))
```

---

## 🛠️ Verification & Test Suite

The test suite runs with [Vitest](https://vitest.dev/) across unit and full integration suites:

```bash
# Run complete test suite (255 tests)
npm test

# Build dual ESM/CJS distributions via tsup
npm run build
```

---

## 📄 License

MIT © [node-pptx-templater contributors](LICENSE)
