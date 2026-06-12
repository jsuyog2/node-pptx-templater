# Feature Inventory — node-pptx-templater

> Generated: 2026-06-12T05:42:27.622Z

**Total public methods:** 150  
**Documented:** 136  

## Core (9)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `load()` | ✓ | ✓ | ✓ | const ppt = await PPTXTemplater.load(buffer); |
| `preload()` | ✓ | ✓ | ✓ | const ppt = await PPTXTemplater.fromCache('./template.pptx'); |
| `cache()` | ✓ | ✓ | ✓ | await PPTXTemplater.cache('./template.pptx'); |
| `fromCache()` | ✓ | ✓ | ✓ | // ppt is ready — no disk I/O on the load |
| `clearCache()` | ✓ |  | ✓ | PPTXTemplater.clearCache(); // Force fresh reload on next fromCache() |
| `fromPresentationXml()` | ✓ | ✓ | ✓ | Loads a template from a PowerPoint XML Presentation format. |
| `create()` | ✓ | ✓ | ✓ | await ppt.saveToFile('./new.pptx'); |
| `extractPptx()` | ✓ | ✓ | ✓ | Extracts a PPTX file into an unzipped OpenXML folder structure. |
| `buildPptx()` | ✓ | ✓ | ✓ | Rebuilds a PPTX file from an unzipped OpenXML folder structure. |

## Slides (15)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `useSlide()` |  |  | ✓ | ppt.useSlide('intro');     // Select by custom tag |
| `useAllSlides()` |  |  | ✓ | Selects all slides. |
| `addSlide()` |  |  | ✓ | }); |
| `cloneSlide()` |  |  | ✓ | Clones an existing slide and appends it to the end (or at a position). |
| `removeSlide()` |  |  | ✓ | Removes a slide from the presentation. |
| `reorderSlides()` |  |  | ✓ | ppt.reorderSlides([3, 1, 2]); // Move slide 3 to position 1 |
| `tagSlide()` |  |  | ✓ | ppt.useSlide('intro').replaceText({ '{{title}}': 'Hello' }); |
| `exportSlides()` |  | ✓ | ✓ | await subset.saveToFile('./subset.pptx'); |
| `importSlideFrom()` |  | ✓ | ✓ | Preserves all slide layouts, charts, relationships, and embedded media. |
| `importSlides()` |  |  | ✓ | ppt.importSlides([1, 3, 5]); |
| `duplicateSlide()` |  |  | ✓ | ppt.duplicateSlide(1, 2); // Copy slide 1 and insert it as slide 2 |
| `deleteSlide()` |  |  | ✓ | ppt.deleteSlide(3); // Remove slide 3 |
| `moveSlide()` |  |  | ✓ | ppt.moveSlide(5, 1); // Move slide 5 to position 1 (make it first) |
| `insertSlide()` |  |  | ✓ | ppt.insertSlide(2, { layoutIndex: 1 }); // Insert blank slide at position 2 |
| `getSlides()` |  |  | ✓ | slides.forEach(s => console.log(`Slide ${s.index}: ${s.title}`)); |

## Tables (19)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `updateTable()` |  |  | ✓ | ]); |
| `getTableRows()` |  |  | ✓ | // → { rows: [...], rowCount: 5, columnCount: 3, mergedCells: [] } |
| `addTableRow()` |  |  | ✓ | ppt.addTableRow('SalesTable', ['Region', ['Q1', 'Q2'], '$5K'], { mergeStrategy:  |
| `removeTableRow()` |  |  | ✓ | ppt.removeTableRow('SalesTable', 2); // Remove the third row (0-based) |
| `insertTableRow()` |  |  | ✓ | ppt.insertTableRow('SalesTable', 1, ['East', '980', '8%']); // Insert at row 1 |
| `cloneTableRow()` |  |  | ✓ | ppt.cloneTableRow('SalesTable', 0, 3); // Clone row 0 to position 3 |
| `updateCell()` |  |  | ✓ | ppt.updateCell('SalesTable', 1, 2, '$9,800', { bold: true, color: '#10B981' }); |
| `mergeCells()` |  |  | ✓ | ppt.mergeCells({ tableId: 'SalesTable', startRow: 0, startCol: 0, endRow: 0, end |
| `unmergeCells()` |  |  | ✓ | ppt.unmergeCells({ tableId: 'SalesTable', row: 0, col: 0 }); |
| `getMergedCells()` |  |  | ✓ | merges.forEach(m => console.log(`Merged: row ${m.startRow}-${m.endRow}, col ${m. |
| `validateMergeRegion()` |  |  | ✓ | if (!result.valid) console.error(result.errors); |
| `isMergedCell()` |  |  | ✓ | } |
| `getMergeParent()` |  |  | ✓ | // → { row: 0, col: 0 } if cells (0,0)-(0,2) are merged |
| `getMergeRegion()` |  |  | ✓ | // → { startRow: 0, startCol: 0, endRow: 0, endCol: 2 } |
| `splitMergedRegion()` |  |  | ✓ | ppt.splitMergedRegion('SalesTable', 0, 0); // Split merge starting at row 0, col |
| `cloneMergedRegion()` |  |  | ✓ | ppt.cloneMergedRegion('SalesTable', 0, 0, 4, 0); |
| `autoFitTable()` |  |  | ✓ | ppt.autoFitTable('SalesTable'); |
| `resizeTable()` |  |  | ✓ | ppt.resizeTable('SalesTable', 6858000, 1371600); // 7.5" wide × 1.5" tall |
| `getTables()` |  |  | ✓ | tables.forEach(t => console.log(`${t.name}: ${t.rows}×${t.cols}`)); |

## Cell Shapes (6)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `addCellShape()` |  | ✓ | ✓ | Dynamically adds a shape inside a table cell based on cell coordinates. |
| `updateCellShape()` |  | ✓ | ✓ | Updates an existing shape inside a table cell. |
| `removeCellShape()` |  | ✓ | ✓ | Removes a shape from a table cell. |
| `getCellShape()` |  |  | ✓ | Discovers and retrieves details of an existing cell shape on the targeted slide. |
| `getCellBounds()` |  |  | ✓ | Retrieves final rendered bounds of a table cell in pixels. |
| `getCellPosition()` |  |  | ✓ | Retrieves final rendered position of a table cell in pixels. |

## Charts (15)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `updateChart()` |  |  | ✓ | }); |
| `updateChartData()` |  |  | ✓ | }); |
| `replaceChartSeries()` |  |  | ✓ | ppt.replaceChartSeries('revenue-chart', 0, { name: 'Revenue', values: [100, 200, |
| `updateChartTitle()` |  |  | ✓ | ppt.updateChartTitle('revenue-chart', 'Q2 2026 Revenue'); |
| `updateChartCategories()` |  |  | ✓ | ppt.updateChartCategories('revenue-chart', ['Jan', 'Feb', 'Mar', 'Apr']); |
| `updateDataLabels()` |  |  | ✓ | }); |
| `getDataLabels()` |  | ✓ | ✓ | console.log(labels.showValue, labels.position); |
| `validateDataLabels()` |  | ✓ | ✓ | if (!result.valid) console.error(result.errors); |
| `validateChartLabels()` |  | ✓ | ✓ | const result = await ppt.validateChartLabels('revenue-chart'); |
| `validateSeriesNameLabels()` |  | ✓ | ✓ | const result = await ppt.validateSeriesNameLabels('bar-chart'); |
| `getCharts()` |  |  | ✓ | charts.forEach(c => console.log(c.zipPath)); |
| `getChartLabelPositions()` |  | ✓ | ✓ | Calculates absolute layout limits in EMUs (English Metric Units). |
| `getChartBarPositions()` |  | ✓ | ✓ | Calculates absolute layout limits in EMUs (English Metric Units). |
| `addTextAtPosition()` |  |  | ✓ | Supports custom font styling and alignment configuration. |
| `addTextNearChartLabel()` |  |  | ✓ | Textboxes are positioned either on the left or right of the chart area, vertical |

## Images (4)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `replaceImage()` |  | ✓ | ✓ | await ppt.replaceImage('company-logo', './new-logo.png'); |
| `addImage()` |  | ✓ | ✓ | }); |
| `removeImage()` |  |  | ✓ | ppt.removeImage('old-logo'); |
| `getImages()` |  |  | ✓ | images.forEach(img => console.log(`${img.name}: ${img.width}×${img.height} EMUs` |

## Shapes (11)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `updateShapeText()` |  |  | ✓ | ppt.updateShapeText('CalloutBox', 'Important note here'); |
| `updateShapePosition()` |  |  | ✓ | Updates the position and/or dimensions of an existing shape on targeted slides. |
| `updateTextBoxPosition()` |  |  | ✓ | Updates the position and/or dimensions of an existing textbox on targeted slides |
| `cloneShape()` |  |  | ✓ | ppt.cloneShape('logo', 'logo-copy', { x: 2000000, y: 500000 }); |
| `deleteShape()` |  |  | ✓ | ppt.deleteShape('temp-banner'); |
| `getShapes()` |  |  | ✓ | shapes.forEach(s => console.log(`${s.name}: ${s.type} at (${s.x}, ${s.y})`)); |
| `validateShape()` |  |  | ✓ | Validates shape options configuration. |
| `addShape()` |  | ✓ | ✓ | Adds a new shape dynamically to the targeted slide(s). |
| `updateShape()` |  | ✓ | ✓ | Updates an existing shape in-place. |
| `removeShape()` |  | ✓ | ✓ | Removes a shape from the targeted slide(s). |
| `getShape()` |  |  | ✓ | Discovers and retrieves details of an existing shape on the targeted slides. |

## Text (8)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `replaceText()` |  |  | ✓ | }); |
| `updateText()` |  |  | ✓ | ppt.updateText('BulletBox', { items: ['Item A', 'Item B', 'Item C'] }); |
| `getList()` |  |  | ✓ | console.log(items); // ['Item A', ['Nested', 'Sub-item'], 'Item B'] |
| `validateList()` |  |  | ✓ | Validates a list structure and values. |
| `replaceTextByTag()` |  |  | ✓ | ppt.replaceTextByTag('{{date}}', '2026-06-12'); |
| `replaceMultiple()` |  |  | ✓ | }); |
| `findText()` |  |  | ✓ | matches.forEach(m => console.log(`Found on slide ${m.slideIndex} in "${m.shapeNa |
| `getTextElements()` |  |  | ✓ | elements.forEach(el => console.log(`Slide ${el.slideIndex}: ${el.text}`)) |

## Hyperlinks (6)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `addHyperlink()` |  |  | ✓ | ppt.addHyperlink({ text: 'Open Website', url: 'https://example.com' }); |
| `addSlideLink()` |  |  | ✓ | Adds an inter-slide hyperlink to a specific text element. |
| `addImageLink()` |  |  | ✓ | Adds an inter-slide hyperlink to an image. |
| `addShapeLink()` |  |  | ✓ | Adds an inter-slide hyperlink to a shape. |
| `addTextNavigationLink()` |  |  | ✓ | Adds a special navigation link (next, previous, first, last slide) to a text ele |
| `addShapeNavigationLink()` |  |  | ✓ | Adds a special navigation link (next, previous, first, last slide) to a shape or |

## Z-Order (15)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `bringForward()` |  |  | ✓ | Moves slide element one layer forward. |
| `sendBackward()` |  |  | ✓ | Moves slide element one layer backward. |
| `bringToFront()` |  |  | ✓ | Moves slide element above all other objects. |
| `sendToBack()` |  |  | ✓ | Moves slide element behind all other objects. |
| `setZIndex()` |  |  | ✓ | Moves slide element to the specific 1-based stacking position. |
| `moveObjectBefore()` |  |  | ✓ | Moves slide element directly before (below) a target element. |
| `moveObjectAfter()` |  |  | ✓ | Moves slide element directly after (above) a target element. |
| `reorderObjects()` |  |  | ✓ | Reorders slide objects exactly as specified in the array. |
| `getObjectOrder()` |  |  | ✓ | order.forEach(o => console.log(`[${o.zIndex}] ${o.name}`)); |
| `applyZOrder()` |  |  | ✓ | Applies bulk template configurations for slide elements stacking layers. |
| `getTopMostObject()` |  |  | ✓ | console.log('Front-most shape:', top.name); |
| `getBottomMostObject()` |  |  | ✓ | console.log('Back-most shape:', bottom.name); |
| `swapObjects()` |  |  | ✓ | ppt.swapObjects('logo', 'background');     // On active slide |
| `sortObjects()` |  |  | ✓ | ppt.sortObjects((a, b) => a.name.localeCompare(b.name)); |
| `normalizeZOrder()` |  |  | ✓ | ppt.normalizeZOrder(1); // Clean up stacking order on slide 1 |

## Output (7)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `saveToFile()` |  | ✓ | ✓ | Saves the modified PPTX to a file on disk. |
| `save()` |  | ✓ | ✓ | Saves the presentation. Equivalent to saveToFile. |
| `saveXml()` |  | ✓ | ✓ | Saves the modified presentation XML structures directly to a folder. |
| `saveToFolder()` |  | ✓ | ✓ | Saves the modified presentation XML structures directly to a folder. |
| `toBuffer()` |  | ✓ | ✓ | Returns the PPTX content as a Node.js Buffer. |
| `toStream()` |  | ✓ | ✓ | Returns the PPTX content as a readable Node.js Stream. |
| `saveToStream()` |  | ✓ | ✓ | Saves the presentation to a readable stream or pipes it to a writable stream. |

## Validation (10)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `validate()` |  |  | ✓ | Reports issues with relationship IDs, missing parts, etc. |
| `repair()` |  | ✓ | ✓ | Removes orphan relationships, rebuilds slide references, and fixes missing entri |
| `validateCharts()` |  | ✓ | ✓ | Checks XML, caches, and embedded workbook references. |
| `repairCharts()` |  | ✓ | ✓ | missing embedded workbooks, or orphan nodes. |
| `validatePresentation()` |  | ✓ | ❌ |  |
| `validatePresentationXml()` |  | ✓ | ✓ | Performs validation specifically on PowerPoint XML folder contents/relationships |
| `validateSlide()` |  | ✓ | ✓ | if (!result.valid) console.error(result.errors); |
| `validateTable()` |  | ✓ | ✓ | if (!result.valid) console.error(result.errors); |
| `validateArchive()` |  | ✓ | ❌ |  |
| `validateRelationships()` |  |  | ✓ | const result = ppt.validateRelationships('ppt/slides/slide1.xml'); |

## Debug (6)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `debugRelationships()` |  |  | ✓ | Logs all relationships across the presentation to the console for debugging. |
| `inspectSlide()` |  |  | ✓ | Inspects a specific slide's structure and relationships. |
| `inspectXML()` |  | ✓ | ✓ | Inspects and logs the raw XML of any file in the ZIP. |
| `inspectChart()` |  |  | ✓ | Inspects a specific chart's metadata and structure. |
| `inspectChartXML()` |  | ✓ | ✓ | Inspects and logs the raw XML of a chart file. |
| `debugChartRelationships()` |  |  | ✓ | Logs all chart relationships. |

## Performance (4)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `enablePerformanceProfile()` |  |  | ✓ | console.log(ppt.getPerformanceMetrics()); |
| `enableDebug()` |  |  | ✓ | ppt.enableDebug(); // Shows all debug output |
| `getPerformanceMetrics()` |  |  | ✓ | console.log(`Total: ${metrics.totalMs}ms, Memory: ${metrics.memoryUsedMB}MB`); |
| `enableDebugZip()` |  |  | ✓ | const buffer = await ppt.toBuffer(); // Logs ZIP entries to debug output |

## Info (1)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `getInfo()` |  |  | ✓ | Returns presentation metadata (title, author, slide count, etc.) |

## Utility (14)

| Method | Static | Async | Documented | Description |
|--------|--------|-------|------------|-------------|
| `setLogLevel()` | ✓ |  | ✓ | PPTXTemplater.setLogLevel('silent'); // Suppress all output |
| `slideCount()` |  |  | ✓ | Returns the total number of slides in the loaded presentation. |
| `zipManager()` |  |  | ❌ |  |
| `xmlParser()` |  |  | ❌ |  |
| `contentTypesManager()` |  |  | ❌ |  |
| `relationshipManager()` |  |  | ❌ |  |
| `slideManager()` |  |  | ❌ |  |
| `chartManager()` |  |  | ❌ |  |
| `tableManager()` |  |  | ❌ |  |
| `shapeManager()` |  |  | ❌ |  |
| `imageManager()` |  |  | ❌ |  |
| `textManager()` |  |  | ❌ |  |
| `hyperlinkManager()` |  |  | ❌ |  |
| `mediaManager()` |  |  | ❌ |  |

