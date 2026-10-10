# Showcase — node-pptx-templater

This directory consolidates all demonstrations, sample templates, and assets for [`node-pptx-templater`](https://github.com/jsuyog2/node-pptx-templater).

## Running the Demonstrations

The showcase is driven by a single self-contained file with all JSON data embedded:

```bash
# From repository root
npm run examples
# or directly:
node showcase/examples.js
```

## Structure

```
showcase/
├── examples.js          # Consolidated single-file executable demonstration suite
├── templates/           # PowerPoint templates and asset fixtures
│   ├── sample.pptx      # Standard presentation fixture
│   ├── officePPT.pptx   # Corporate presentation fixture
│   ├── chess.png        # Sample image fixture
│   └── rook.png         # Sample image fixture
└── output/              # Generated PPTX presentations (gitignored)
```

## Demonstrations Included

1. **Basic Report Generation & Text Replacement**: Slide dimensions, widescreen setup (16:9), text tag substitution (`replaceMultiple`), text inspection (`findText`).
2. **Chart Data & Labels Update with Excel Sync**: Multi-series category charts, data label configuration, internal embedded `.xlsx` workbook caching.
3. **Dynamic Table Operations**: Appending rows, row striping (`stripeRows`), custom cell borders (`border` via `updateCell`), summary row pinning (`splitTableOnOverflow`).
4. **Nested Tables & Cell Merging**: Hierarchical row expansion with `mergeStrategy: 'rowspan'` and `'auto'`, plus structured table extraction via `getTableRows({ includeMetadata: true })`.
5. **Shape Operations & Z-Ordering**: Adding shapes, fills, strokes, dimensions, vertical alignment (`alignShapes`), and layer ordering (`bringToFront`, `sendToBack`).
6. **Dynamic Cell Shapes**: Anchoring status indicator shapes and badges directly inside table cells (`addTableRow` with shape descriptors).
7. **Multi-Format Media Operations**: Replacing existing slide images (`replaceImage`), adding PNG images with dimensions, and embedding WebP from Base64 Data URIs without native dependencies.
8. **Slide Management**: Duplicating slides (`duplicateSlide`), reordering (`reorderSlides`), and importing slides across presentations (`importSlideFrom`).
9. **Hyperlinks & Interactive Navigation**: Adding external URLs, inter-slide jumps (`addSlideLink`), and discovering hyperlink relationships (`getHyperlinks`).
10. **Rich Typography & Text Runs**: Advanced font styling, colors, weights, line spacing, and appending distinct text runs (`appendTextRun`).
11. **OpenXML Folder Workflow**: Unpacking presentations into raw XML folders (`extractPptx`), loading from folder, editing in-place, and rebuilding into binary PPTX (`buildPptx`).
12. **Serverless Buffer Export**: Generating presentations directly in-memory as a `Buffer` (`toBuffer`) for AWS Lambda, Google Cloud Functions, or streaming HTTP responses.
