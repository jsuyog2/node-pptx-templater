# node-pptx-templater

> A low-level PowerPoint OpenXML templating engine for Node.js that generates and edits PPTX files directly through XML manipulation without relying on PowerPoint generation libraries.

[![npm version](https://img.shields.io/npm/v/node-pptx-templater.svg)](https://www.npmjs.com/package/node-pptx-templater)
[![CI](https://github.com/jsuyog2/node-pptx-templater/actions/workflows/ci.yml/badge.svg)](https://github.com/jsuyog2/node-pptx-templater/actions/workflows/ci.yml)
[![Coverage](https://img.shields.io/codecov/c/github/jsuyog2/node-pptx-templater)](https://codecov.io/gh/jsuyog2/node-pptx-templater)
[![License: MIT](https://img.shields.io/badge/License-MIT-green.svg)](./LICENSE)
[![Node.js](https://img.shields.io/badge/node-%3E%3D18.0.0-brightgreen)](https://nodejs.org)
[![ES Modules](https://img.shields.io/badge/ESM-only-blueviolet)](https://nodejs.org/api/esm.html)

---

## ✨ Features

| Feature | Description |
|---|---|
| 🏗️ **Zero PPTX library dependencies** | Direct OpenXML/ZIP manipulation only |
| 🔁 **Text replacement** | Handles fragmented runs (`{{placeholders}}`) |
| 📊 **Chart updates** | Bar, Line, Pie, Area, Scatter — data only, style preserved |
| 📋 **Table updates** | Replace rows while preserving all formatting |
| 🔗 **Hyperlinks** | External URLs, shape links, slide-to-slide navigation |
| ➕ **Add new slides** | Text, images, shapes with auto-generated XML |
| 🎯 **Slide selection** | By number, ID, or custom tags |
| 📤 **Multiple outputs** | `saveToFile()`, `toBuffer()`, `toStream()` |
| 🔍 **Validation** | Structure validation with error reporting |
| 🛠️ **CLI** | `build`, `validate`, `inspect`, `extract`, `debug` |
| ⚡ **Performance** | Lazy loading, media deduplication, async/await |

---

## 📦 Installation

```bash
npm install node-pptx-templater
```

**Requirements:** Node.js ≥ 18.0.0, ES Modules (`"type": "module"`)

---

## 🚀 Quick Start

```js
import { PPTXTemplater } from 'node-pptx-templater';

// Load a template PPTX
const ppt = await PPTXTemplater.load('template.pptx');

// Select slide(s) to work on (omit to work on all)
ppt.useSlide(1);

// Replace {{placeholder}} text
ppt.replaceText({
  '{{title}}':    'Quarterly Report',
  '{{year}}':     '2026',
  '{{company}}':  'Acme Corp',
});

// Update chart data
ppt.updateChart('sales-chart', {
  categories: ['Jan', 'Feb', 'Mar'],
  series: [{ name: 'Revenue', values: [120, 150, 180] }],
});

// Update table rows
ppt.updateTable('employees-table', [
  ['Name',  'Role',       'Department'],
  ['Alice', 'Engineer',   'Platform'],
  ['Bob',   'Designer',   'Product'],
]);

// Save output
await ppt.saveToFile('./output/report.pptx');

// Or get a Buffer (for HTTP responses, emails, etc.)
const buffer = await ppt.toBuffer();
```

---

## 📚 API Reference

### `PPTXTemplater`

#### Static Methods

| Method | Description |
|---|---|
| `PPTXTemplater.load(source)` | Load from file path or Buffer |
| `PPTXTemplater.create()` | Create a blank presentation |

#### Slide Selection

| Method | Description |
|---|---|
| `.useSlide(...refs)` | Select slides by number/tag to apply operations |
| `.useAllSlides()` | Reset to all slides |
| `.tagSlide(num, tag)` | Assign custom tag for later selection |

#### Content Manipulation

| Method | Description |
|---|---|
| `.replaceText(replacements)` | Replace `{{key}}` placeholders |
| `.updateChart(chartId, data)` | Update chart categories/series/values |
| `.updateTable(tableId, rows)` | Replace table row data |
| `.addHyperlink(options)` | Add/replace external hyperlink on text |
| `.linkSlideNumber(options)` | Make slide reference navigate to another slide |

#### Slide Management

| Method | Description |
|---|---|
| `.addSlide(options)` | Add a new slide with elements |
| `.cloneSlide(num, atPos?)` | Duplicate an existing slide |
| `.removeSlide(num)` | Delete a slide |
| `.reorderSlides(order)` | Reorder slides |
| `.exportSlides(...nums)` | Export subset to new engine |

#### Output

| Method | Description |
|---|---|
| `.saveToFile(path)` | Write PPTX to disk |
| `.toBuffer()` | Get PPTX as Node.js Buffer |
| `.toStream()` | Get PPTX as readable stream |

#### Utilities

| Method | Description |
|---|---|
| `.getInfo()` | Presentation metadata |
| `.validate()` | Structure validation |
| `.slideCount` | Total slide count (getter) |

---

## 🏗️ Architecture

```
node-pptx-templater/
├── PPTXTemplater     ← Main orchestrator / public API
│   ├── ZipManager         ← ZIP archive read/write (JSZip)
│   ├── XMLParser          ← XML parse/build (fast-xml-parser)
│   ├── RelationshipManager ← .rels file management
│   ├── SlideManager       ← Slide discovery, ordering, CRUD
│   ├── ChartManager       ← Chart XML data updates
│   ├── TableManager       ← Table row replacement
│   ├── HyperlinkManager   ← Hyperlink injection
│   ├── MediaManager       ← Image embedding + deduplication
│   ├── TemplateEngine     ← {{placeholder}} replacement
│   └── OutputWriter       ← File/Buffer/Stream output
```

### OpenXML PPTX Structure

A `.pptx` file is a ZIP archive (Open Packaging Convention) containing XML files:

```
presentation.pptx (ZIP)
├── [Content_Types].xml          # MIME types for all parts
├── _rels/.rels                  # Root relationships
├── ppt/
│   ├── presentation.xml         # Slide list, master references
│   ├── _rels/presentation.xml.rels
│   ├── slides/
│   │   ├── slide1.xml           # Slide content XML
│   │   ├── slide2.xml
│   │   └── _rels/slide1.xml.rels
│   ├── slideLayouts/            # Layout templates
│   ├── slideMasters/            # Master slide designs
│   ├── theme/                   # Color & font themes
│   ├── charts/                  # Embedded chart XML
│   └── media/                   # Images, videos
└── docProps/
    ├── core.xml                 # Title, author, dates
    └── app.xml                  # App metadata
```

### Relationship Flow

```
presentation.xml
  └─[rId2]──► slides/slide1.xml
                └─[rId1]──► slideLayouts/slideLayout1.xml
                └─[rId2]──► charts/chart1.xml
                              └─ (chart data XML)
                └─[rId3]──► media/image1.png
                └─[rId4]──► https://example.com  (external)
```

### Text Fragmentation Problem & Solution

PowerPoint sometimes splits `{{placeholder}}` across multiple XML text runs:

```xml
<!-- What you write in PowerPoint: -->
{{title}}

<!-- What the XML actually contains: -->
<a:r><a:t>{{ti</a:t></a:r>
<a:r><a:t>tle}}</a:t></a:r>
```

**Our solution:** The `TemplateEngine` normalizes runs within each paragraph by:
1. Concatenating all run text content
2. Detecting placeholders in the combined text
3. Merging affected runs into one (preserving the first run's formatting)
4. Injecting the replacement value

---

## 🖥️ CLI Usage

```bash
# Install globally
npm install -g node-pptx-templater

# Build a PPTX from template + JSON data
node-pptx-templater build template.pptx output.pptx --data data.json

# Validate a PPTX structure
node-pptx-templater validate presentation.pptx

# Inspect internal structure
node-pptx-templater inspect presentation.pptx --all

# Extract a slide's XML
node-pptx-templater extract presentation.pptx --slide 1 --out slide1.xml

# Debug a corrupted PPTX
node-pptx-templater debug broken.pptx --fix --out repaired.pptx
```

### Data JSON format for `build` command

```json
{
  "text": {
    "{{title}}": "Annual Report 2026",
    "{{company}}": "Acme Corp",
    "{{date}}": "January 2026"
  },
  "charts": {
    "sales-chart": {
      "categories": ["Q1", "Q2", "Q3", "Q4"],
      "series": [
        { "name": "Revenue", "values": [145, 210, 190, 250] }
      ]
    }
  },
  "tables": {
    "data-table": [
      ["Name", "Role", "Dept"],
      ["Alice", "Engineer", "Platform"]
    ]
  }
}
```

---

## 📊 Supported Chart Types

| OpenXML Element | Chart Type |
|---|---|
| `c:barChart` | Bar / Column |
| `c:lineChart` | Line |
| `c:pieChart` | Pie |
| `c:areaChart` | Area |
| `c:scatterChart` | Scatter / XY |
| `c:doughnutChart` | Doughnut |
| `c:radarChart` | Radar / Spider |
| `c:bubbleChart` | Bubble |

---

## ⚡ Performance

| Operation | Benchmark (avg) |
|---|---|
| Load 50-slide PPTX | ~120ms |
| Text replacement (20 placeholders) | ~2ms |
| Buffer generation | ~80ms |
| Chart update | ~5ms |
| Table update | ~3ms |

> Run your own: `npm run benchmark`

---

## 🐛 Troubleshooting

### Placeholders not being replaced

If `{{placeholder}}` is not replaced, the text is likely fragmented across runs.
Use the `PPTX_LOG_LEVEL=debug` environment variable to see detailed logs:

```bash
PPTX_LOG_LEVEL=debug node your-script.js
```

Extract the slide XML to inspect:
```bash
node-pptx-templater extract template.pptx --slide 1 --out slide1.xml
```

Look for `<a:t>` elements and check if the placeholder is split.

### Chart not updating

Chart names must match the shape's `cNvPr name` attribute exactly.
Use `--inspect --charts` to see all chart names:

```bash
node-pptx-templater inspect template.pptx --charts
```

### Generated PPTX fails to open

Run the debug command to check for structural issues:

```bash
node-pptx-templater debug output.pptx
```

Common causes:
- Missing content type entries for new slides
- Broken relationship IDs
- Invalid XML characters in replacement values

### File is corrupted

```bash
node-pptx-templater debug corrupted.pptx --fix --out repaired.pptx
```

The debug command attempts:
- Removing invalid XML control characters
- Fixing unescaped `&` in text content
- Repairing broken relationship IDs

---

## 🔌 Plugin System

Extend the engine by subclassing `PPTXTemplater`:

```js
import { PPTXTemplater } from 'pptx-templater';

class MyEngine extends PPTXTemplater {
  /**
   * Custom method: fills a slide from a data object.
   */
  async fillFromData(slideNum, data) {
    this.useSlide(slideNum);

    const textReplacements = {};
    for (const [key, val] of Object.entries(data.text || {})) {
      textReplacements[`{{${key}}}`] = String(val);
    }

    this.replaceText(textReplacements);

    if (data.chart) {
      this.updateChart(data.chart.id, data.chart);
    }

    return this;
  }
}

// Usage
const ppt = await MyEngine.load('template.pptx');
await ppt.fillFromData(1, {
  text: { title: 'My Report', date: '2026-01-01' },
  chart: { id: 'sales', categories: ['Q1'], series: [{ name: 'Rev', values: [100] }] },
});
await ppt.saveToFile('output.pptx');
```

---

## 🛣️ Roadmap

- [ ] SmartArt data update
- [ ] Speaker notes modification
- [ ] Slide transitions and animation metadata editing
- [ ] PPTX → HTML export (read-only)
- [ ] Password-protected PPTX support
- [ ] Native chart creation from scratch (without template)
- [ ] Watch mode for development
- [ ] Browser/WASM support (via jszip already)

---

## 🤝 Contributing

See [CONTRIBUTING.md](./CONTRIBUTING.md) for guidelines.

Quick steps:
```bash
git clone https://github.com/jsuyog2/node-pptx-templater.git
cd node-pptx-templater
npm install
npm test
```

---

## 📄 License

MIT — see [LICENSE](./LICENSE)
