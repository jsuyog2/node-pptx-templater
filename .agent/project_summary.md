# Project Architecture Summary: `node-pptx-templater`

> **Version**: 3.0.0-alpha.0  
> **Runtime**: Pure JavaScript (Node.js `>=18.x`, AWS Lambda, Cloudflare Workers, Vercel Edge, Deno)  
> **Standard**: ISO/IEC 29500 (ECMA-376) Parts 1 & 2 (OPC & PresentationML)  
> **API Catalog**: See `docs-site/src/data/apiReference.json` for full parameter definitions & examples

---

## 1. Architectural Blueprint & Subsystem Managers

`node-pptx-templater` is an in-memory OpenXML package manipulator operating on a Virtual File System (`Map<string, Uint8Array>`) with zero native binaries or office dependencies.

```
Ingestion (PackageLoader, FlatOpcConverter, LegacyConverter)
  └─► In-Memory VFS (VirtualFileSystem)
        └─► AST Parsing & Serializing (XMLParser: standard & ordered [Z_ORDER_SYMBOL])
              ├─► Core: TemplateEngine (run healing, lists), ValidationEngine (XXE, integrity)
              ├─► Slide Subsystems: SlideManager, ZOrderManager, CommentManager
              ├─► Element Subsystems: TextManager, TableManager, ChartManager, ShapeManager, ImageManager, MediaManager
              ├─► Package Manifests: RelationshipManager, ContentTypesManager, ZipManager
              └─► Output Pipeline: OutputWriter & FrameworkAdapters (18-method I/O boundary)
```

### Core Subsystems
- **`VirtualFileSystem`** (`src/core/VirtualFileSystem.js`): Normalized forward-slash in-memory file store.
- **`PackageLoader`** (`src/core/PackageLoader.js`): Ingests local files, buffers, streams, URLs, directories, cloud objects, and Flat OPC XML.
- **`XMLParser`** (`src/parsers/XMLParser.js`): Lossless `fast-xml-parser` with dual-mode AST parsing and Z-order preservation (`[Z_ORDER_SYMBOL]`).
- **`TemplateEngine`** (`src/core/TemplateEngine.js`): AST-level text run healing across fragmented `<a:r>`/`<a:t>` tags and native list generation.
- **`ValidationEngine`** (`src/core/ValidationEngine.js`): XXE defense, entity bomb prevention, and structural package verification.
- **`SlideManager`** (`src/managers/SlideManager.js`): 1-based slide operations, deep cloning, reordering, layout binding, sections, and comments.
- **`TableManager`** (`src/elements/TableManager.js`): Dynamic row/column mutations, cell merging (`gridSpan`/`rowSpan`/`hMerge`/`vMerge`), `<a16:rowId>` generation, and auto-pagination (`splitTableOnOverflow`).
- **`ChartManager`** (`src/elements/ChartManager.js`): Strict 11-method whitelist supporting all 17 standard (`<c:chartSpace>`) and Office 2016+ extended (`<cx:chartSpace>`) chart types with Dual-Sync updating of embedded Excel workbooks.
- **`ShapeManager`** (`src/elements/ShapeManager.js`): Vector geometry, fills, borders, text, alignment, and coordinate math.
- **`TextManager`** (`src/elements/TextManager.js`): Run substitutions, full text box replacement (`replaceTextByTag`), and 31-property DrawingML styling.
- **`ZOrderManager`** (`src/spatial/ZOrderManager.js`): Visual stacking operations while preserving `<p:nvGrpSpPr>` and `<p:grpSpPr>` header invariants.
- **`ImageManager`** & **`MediaManager`** (`src/elements/ImageManager.js`, `src/managers/MediaManager.js`): Image insertion, swap, MIME detection, aspect ratio, audio/video links.
- **`RelationshipManager`** & **`ContentTypesManager`** (`src/managers/`): `.rels` target graphs and `[Content_Types].xml` default/override mappings.
- **`ZipManager`** (`src/managers/ZipManager.js`): In-memory ZIP packaging and extraction via JSZip.
- **`OutputWriter`** (`src/writer/OutputWriter.js`): Multi-target output serialization with Express/Fastify adapters.

---

## 2. Strict 18-Method I/O Whitelist

All presentation input, output, and serialization must route through these 18 approved methods (all deprecated aliases like `toFile`, `save`, `exportZip`, `writeToBuffer` are prohibited):

- **Static Methods (7)**:
  `PPTXTemplater.load(source, options)` • `PPTXTemplater.createBlank(options)` • `PPTXTemplater.extractPptx(pptxPath, outputDir, options)` • `PPTXTemplater.buildPptx(sourceDir, outputPath, options)` • `PPTXTemplater.preload(source, options)` • `PPTXTemplater.fromCache(source)` • `PPTXTemplater.clearCache()`
- **Export Methods (8)**:
  `ppt.saveToFile(destinationPath, options)` • `ppt.toBuffer(options)` • `ppt.toStream(options)` • `ppt.toBase64(options)` • `ppt.toArrayBuffer(options)` • `ppt.toResponse(options)` • `ppt.toExpress(res, filename, options)` • `ppt.toFastify(reply, filename, options)`
- **Extraction & Conversion (3)**:
  `ppt.extractToZip(destinationPath, options)` • `ppt.extractToXml(destinationDir, options)` • `ppt.toFlatXml(destinationPathOrOptions, options)`

---

## 3. Strict 23-Method Slide Management Whitelist

All slide manipulations use 1-based indexing and strictly adhere to these 23 approved methods:

- **Selection & Scoping (2)**: `ppt.useSlide(...slideRefs)` • `ppt.useAllSlides()`
- **Creation & Duplication (2)**: `ppt.addSlide(options, position)` • `ppt.duplicateSlide(sourceIndex, atPosition)`
- **Lifecycle & Reordering (3)**: `ppt.removeSlide(position)` • `ppt.moveSlide(fromIndex, toIndex)` • `ppt.reorderSlides(order)`
- **Inspection & Metadata (5)**: `ppt.getSlides()` • `ppt.getSlide(index)` • `ppt.tagSlide(slideIndex, tag)` • `ppt.getSlideDimensions()` • `ppt.getSlideLayout(slideIndex)`
- **Cross-Deck Transfer (2)**: `ppt.importSlideFrom(sourcePpt, sourceIndex, atPosition)` • `ppt.exportSlides(...slideIndices)`
- **Section Management (5)**: `ppt.useSection(sectionRef)` • `ppt.addSection(name, startSlideIndex)` • `ppt.updateSection(sectionIdentifier, options)` • `ppt.removeSection(sectionIdentifier, removeSlides)` • `ppt.getSection(sectionRef)`
- **Visibility & Reset (3)**: `ppt.hideSlide(...slideRefs)` • `ppt.unhideSlide(...slideRefs)` • `ppt.resetSlide(slideRef, options)`
- **Comments (2)**: `ppt.addComment(text, options)` • `ppt.removeComment(commentIdentifier, options)`

---

## 4. Text & Style Engine Whitelist

- **Text & Lists (18 Methods)**: `replaceText`, `addText`, `replaceTextByTag`, `addLists`, `getList`, `getLists`, `findLists`, `replaceList`, `clearList`, `removeList`, `findText`, `getText`, `getTextElements`, `addStyle`, `updateStyle`, `removeStyle`, `clearText`, `removeText`.
- **`replaceTextByTag(tagOrId, newText, options)`**: Resolves by shape ID (`<p:cNvPr id="...">`), name (`name="..."`), or tag; clears existing runs inside `<p:txBody>`, partitions `\n` into distinct `<a:p>` paragraphs, and inherits root `<a:pPr>` and `<a:rPr>`.
- **31-Property Style Engine**: Typography (8), Casing & Spacing (6), Alignment & Direction (4), Bullets (5), Shadows & Outer Effects (6).

---

## 5. Strict 10-Method Shapes & Geometry Whitelist

All shape creation, inspection, mutation, grouping, and cleanup strictly adhere to these 10 approved methods (all non-whitelisted shape functions such as `deleteShape`, `updateShapeText`, `alignShapes`, `validateShape`, or `addTextAtPosition` are prohibited):

1. `ppt.addShape(slideId, type, options)`
2. `ppt.updateShape(targetId, options)`
3. `ppt.removeShape(shapeId)`
4. `ppt.clearShapes()`
5. `ppt.getShapes()`
6. `ppt.getShape(shapeId)`
7. `ppt.cloneShape(shapeId, clonedShapeId, clonedOptions)`
8. `ppt.getGroupShapes(groupId)`
9. `ppt.groupShapes(shapeIds)`
10. `ppt.ungroupShapes(groupId)`

---

## 6. Strict 12-Method Images & Media Whitelist

All image, video, and audio operations strictly adhere to these 12 approved methods across `PPTXTemplater`, `ImageManager`, and `MediaManager` (with PascalCase aliases maintained for backward compatibility):

- **Images (3)**:
  `ppt.addImage([slideId], source, options)` • `ppt.updateImage(imageId, options)` • `ppt.removeImage(imageId)`
- **Videos (3)**:
  `ppt.addVideo([slideId], source, options)` • `ppt.updateVideo(videoId, options)` • `ppt.removeVideo(videoId)`
- **Audio (3)**:
  `ppt.addAudio([slideId], source, options)` • `ppt.updateAudio(audioId, options)` • `ppt.removeAudio(audioId)`
- **Inspection & Cloning (3)**:
  `ppt.getMedias([slideId])` • `ppt.getMedia(mediaId)` • `ppt.cloneMedia(mediaId, clonedMediaId, clonedOptions)`

**Architecture Highlights**:
- **`useSlide()` Context Binding**: Operations seamlessly resolve the active slide from `ppt.useSlide(slideNumber)` when `slideId` is omitted.
- **Universal Input Ingestion**: Supports local file paths, remote URLs, Base64 Data URIs (`data:...;base64,...`), raw Base64 strings, Buffers, and streams.
- **Automatic In-Memory Deduplication**: Content-hash SHA-1 tracking avoids duplicate media assets in `ppt/media/`.
- **Full OpenXML Conformance**: Slide relationships (`.rels`), `<p:pic>` generation, `<p:timing>` audio/video node generation and automated orphaned timing sequence purging, percentage and EMU spatial resolution, and safe relationship deletion without breaking sibling references.

---

## 7. Strict 11-Method Charts Whitelist

All chart creation, mutation, inspection, cloning, labeling, and removal strictly adhere to these 11 approved methods (prohibiting non-whitelisted functions like `repairCharts`, `refreshChartCache`, `setChartTitle`, `formatChartSeries`):

1. `ppt.addChart(slideId, chartType, data, options)`
2. `ppt.updateChart(chartId, options)`
3. `ppt.updateChartData(chartId, data)`
4. `ppt.removeChart(chartId)`
5. `ppt.clearCharts()`
6. `ppt.getCharts(slideId)`
7. `ppt.getChart(chartId)`
8. `ppt.cloneChart(chartId, clonedChartId, clonedOptions)`
9. `ppt.getChartLabel(chartId)`
10. `ppt.setChartLabel(chartId, label, options)`
11. `ppt.updateChartLabel(chartId, options)`

**17 Supported Chart Types**:
- **Standard DrawingML (`<c:chartSpace>`)**: Column (`column`, `col`), Line (`line`), Pie (`pie`), Bar (`bar`), Area (`area`), XY Scatter (`scatter`, `xy`, `xyScatter`), Map (`map`), Stock (`stock`), Surface (`surface`), Radar (`radar`), Combo (`combo`).
- **Office 2016+ Extended (`<cx:chartSpace>`)**: Treemap (`treemap`), Sunburst (`sunburst`), Histogram (`histogram`), Box & Whisker (`boxWhisker`, `boxAndWhisker`, `box_and_whisker`), Waterfall (`waterfall`), Funnel (`funnel`).

**Core Invariants**:
- **Dual-Sync Invariant**: Every chart addition or data mutation synchronously updates the XML cache (`ppt/charts/chartN.xml` or `chartExN.xml`) and the backing embedded Excel workbook (`ppt/embeddings/Microsoft_Excel_WorksheetN.xlsx`) with table cleansing to eliminate PowerPoint repair prompts.
- **clearCharts Evacuation**: Evacuates all chart graphic frames from targeted slide(s) and purges unreferenced chart parts and workbooks while preserving layout shapes, tables, and text frames.

---

## 8. Documentation Portal (`docs-site/`)

The documentation site is a static export built with Next.js 14+ App Router (`output: 'export'`, `images: { unoptimized: true }`, `trailingSlash: true`) and deployed to Cloudflare Pages (`https://node-pptx-templater.pages.dev`). All documentation content is driven by decoupled JSON schemas in `docs-site/src/data/` (`apiReference.json`, `features.json`, `docsNavigation.json`, `compatibility.json`, `siteMeta.json`) and verified by `npm run docs:validate`.

