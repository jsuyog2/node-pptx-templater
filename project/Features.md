# Comprehensive Feature Catalog: node-pptx-templater (v2.0)

A granular, itemized inventory of all capabilities supported by the modernized PowerPoint template manipulation and generation library.

---

## 1. Input & Ingestion Flexibility

1. **Direct PPTX File Ingestion**: Load standard `.pptx` documents directly from local filesystem paths.
2. **Binary Buffer Loading**: Ingest PPTX files from in-memory Node.js `Buffer` or `Uint8Array` instances.
3. **Stream-Based Ingestion**: Pipe incoming `.pptx` files using Node.js `ReadableStream` or web streams.
4. **ZIP Archive Ingestion**: Load archives with `.zip` or custom extensions by inspecting the PKZip container signature (`0x50 0x4B 0x03 0x04`).
5. **Unpacked OpenXML Directory Ingestion**: Ingest raw, uncompressed folder trees (containing `[Content_Types].xml`, `_rels/`, and `ppt/`) directly into memory without requiring an intermediate zip step.
6. **Scratch Deck Generation**: Initialize new, blank presentations from code without any input template.
7. **Custom Presentation Sizing**: Set custom presentation dimensions and standard aspect ratios (16:9 widescreen, 16:10, 4:3 standard) on initialization.

---

## 2. Presentation & Slide Lifecycle Management

8. **Presentation Metadata Query**: Read deck-level metadata (total slide count, slide dimensions in EMUs/pixels/points, layout schemes).
9. **Add Blank Slide**: Append or insert empty slides associated with specific master layouts.
10. **Slide Deletion**: Remove slides by zero-based index or internal slide identifier without corrupting presentation manifests.
11. **Slide Reordering**: Change the display sequence of slides dynamically.
12. **Deep Slide Duplication**: Clone existing slides with complete element trees, shapes, and layouts preserved.
13. **Isolated Slide Relationship Cloning**: Generate dedicated relationship mappings (`slideX.xml.rels`) during duplication to eliminate relationship ID (`rId`) collisions.
14. **Manifest Synchronization**: Automatically synchronize slide additions, duplicates, and deletions across `[Content_Types].xml` and `ppt/presentation.xml`.

---

## 3. Spatial Coordinate Engine & Unit System

15. **OpenXML EMU Normalization**: Standardize internal dimensions and positions in English Metric Units (1 pt = 12,700 EMUs; 1 in = 914,400 EMUs; 1 px = 9,525 EMUs).
16. **Coordinate & Dimension Extraction**: Retrieve computed bounding boxes (`x`, `y`, `width`, `height`, `rotation`) for any slide element.
17. **Programmatic Spatial Transformation**: Reposition and resize any element using points, pixels, inches, or EMUs.
18. **Z-Index Layer Management**: Reorder element rendering hierarchy (bring to front, send to back, step forward, step backward).
19. **Boundary & Collision Detection**: Calculate element bounding intersections and detect slide edge overflows.

---

## 4. Element Tree Traversal & Inspection

20. **Slide Element Enumeration**: Extract all visual elements residing on a specific slide.
21. **Type-Based Element Filtering**: Query elements filtered by type (`text`, `shape`, `image`, `table`, `chart`, `group`).
22. **Name & ID Selectors**: Query elements by their PowerPoint shape name (e.g., `Title 1`, `Picture 3`) or unique non-visual shape ID.
23. **Template Token Selectors**: Target and match text boxes, shapes, or tables bound to placeholder identifiers (e.g., `{{variable_name}}`).

---

## 5. Text & Typography Engine

24. **AST-Level Run Normalization**: Tokenize and assemble fragmented `<a:r>` and `<a:t>` nodes to eliminate broken regex replacements across XML boundaries.
25. **In-Place Text Replacement**: Replace single or multiple text values without losing surrounding paragraph formatting.
26. **Text Run Appending & Prepending**: Insert new text spans before or after existing content within any text box.
27. **Font Family Modification**: Update font family declarations across entire text frames or specific runs.
28. **Font Size Adjustment**: Modify text size in points or EMUs.
29. **Color & Fill Modification**: Apply hex colors, RGB values, or theme colors to text runs.
30. **Typography Styling**: Toggle bold, italic, underline, strikethrough, subscript, and superscript.
31. **Paragraph Alignment**: Set horizontal alignment (left, center, right, justify).
32. **Vertical Alignment**: Control text frame vertical centering, top alignment, or bottom alignment.
33. **Line & Paragraph Spacing**: Adjust line spacing, space before, and space after paragraphs.
34. **Word Wrap Control**: Enable or disable text wrapping within shape and text box boundaries.

---

## 6. Shape Engine

35. **Add Vector Shapes**: Insert basic shapes (rectangles, rounded rectangles, ellipses, triangles, arrows, callouts, lines).
36. **Solid Shape Fills**: Apply solid hex/RGB color fills to any shape.
37. **Gradient & Pattern Fills**: Apply linear/radial gradients or pattern fills to shapes.
38. **Fill Transparency & Opacity**: Adjust transparency percentages on shape fills.
39. **Shape Border & Outline Styling**: Customize outline color, stroke weight, and stroke pattern (solid, dashed, dotted).
40. **Text Inside Shapes**: Insert, style, and vertically align text inside any vector shape with custom internal padding.
41. **Remove Shape**: Delete individual shapes from slides cleanly.

---

## 7. Media & Image Handling

42. **Insert Images**: Add images from local file paths, binary Buffers, or Base64 data URIs.
43. **Multi-Format Media Support**: Support PNG, JPEG, SVG, GIF, and WebP media types.
44. **In-Place Image Swapping**: Replace template image placeholders while preserving original bounding boxes and coordinates.
45. **Aspect Ratio Preservation**: Automatically calculate and lock aspect ratios when inserting or updating images.
46. **Image Metadata Extraction**: Query image dimensions, format types, file sizes, and internal relationship references.
47. **Remove Image**: Remove media elements and clean up associated relationship pointers.

---

## 8. Table Engine & Structural Manipulation

48. **Insert Table**: Add tables specifying row counts, column counts, coordinates, and default dimensions.
49. **Table Data Hydration**: Populate entire tables directly from 2D arrays or structured JSON payloads.
50. **Row CRUD**: Add, append, delete, or duplicate table rows dynamically.
51. **Column CRUD**: Add, append, or delete table columns dynamically.
52. **Cell Target Query**: Retrieve individual cells via zero-based indices (`table.getCell(rowIndex, colIndex)`).
53. **Cell Data Modification**: Update text, font, alignment, and formatting inside individual cells.
54. **Cell Background Fills**: Set solid background colors and opacity per cell or row.
55. **Custom Cell Borders**: Set individual edge borders (top, bottom, left, right) with custom thickness and color.
56. **Horizontal Cell Merging**: Merge adjacent cells across columns using `gridSpan`.
57. **Vertical Cell Merging**: Merge adjacent cells down rows using `rowSpan`.
58. **Cell Unmerging**: Split previously merged cells back into individual grid cells.
59. **Nested Elements Inside Cells**: Embed shapes, icons, badges, and small indicators directly inside table cells.
60. **Header Row Styling**: Independently target and style table header rows.
61. **Zebra Striping**: Apply automated alternating row background colors.
62. **Table Dimensional Metadata**: Extract computed table width, cumulative height, individual column widths, and row heights.

---

## 9. Dynamic Table Pagination & Slide Splitting (Flagship)

63. **Dynamic Row Height Calculation**: Compute expected row height based on cell padding, font size, line count, and column width.
64. **Slide Overflow Detection**: Compare cumulative table height against available slide height ($Y_{\text{limit}} = \text{SlideHeight} - \text{BottomMargin}$).
65. **Automated Slide Cloning on Overflow**: Clone the host template slide (including headers, footers, and background styling) when rows exceed $Y_{\text{limit}}$.
66. **Table Row Slicing**: Truncate overflowing rows on the parent slide at the last fitting index.
67. **Header Repetition**: Re-render table header rows automatically on every newly generated overflow slide.
68. **Multi-Slide Pagination**: Recursively distribute remaining table rows across subsequent slides until all records are displayed.
69. **Summary / Footer Row Preservation**: Ensure designated summary/total rows remain anchored to the final generated slide.

---

## 10. Native Charts & Data Visualization

70. **Target Native Charts**: Select embedded PowerPoint charts by name, ID, or slide index.
71. **Supported Chart Types**: Support Bar, Column, Line, Pie, Doughnut, Area, and Scatter charts.
72. **Chart Series Data Updates**: Modify existing series data values, categories, and series labels dynamically.
73. **Synchronous Excel Cache Updating**: Update embedded `.xlsx` workbooks (`ppt/embeddings/`) alongside `<c:chartSpace>` to eliminate Microsoft PowerPoint repair prompts.
74. **Chart Spatial Adjustment**: Resize and reposition chart elements on the slide canvas.
75. **Chart Element Toggles**: Show or hide legends, data labels, category axes, and gridlines.

---

## 11. Hyperlinks & Slide Navigation

76. **External Web Hyperlinks**: Attach external URLs to specific text runs, entire text frames, shapes, or images.
77. **Internal Slide Jump Links**: Attach slide-to-slide navigation links (e.g., jump to a specific slide, first slide, last slide).
78. **Hyperlink Inspection**: Read existing destination URLs and internal slide targets.
79. **Hyperlink Removal**: Strip hyperlinks from elements while retaining the underlying text or shape styling.

---

## 12. Serialization & Multi-Target Export Pipeline

80. **Local File System Export (`toFile`)**: Asynchronously save generated presentations directly to disk.
81. **Node.js ReadableStream Export (`toStream`)**: Stream presentation bytes for memory-efficient piping to AWS S3, Google Cloud Storage, or Azure Blob.
82. **Memory Buffer Export (`toBuffer`)**: Export as a native Node.js `Buffer` for in-memory processing.
83. **ArrayBuffer Export (`toArrayBuffer`)**: Export as a standard `ArrayBuffer` for isomorphic Node.js and modern runtime compatibility.
84. **Base64 String Export (`toBase64`)**: Export as a Base64 string for JSON transfer, serverless lambdas, or webhooks.
85. **Data URI Base64 Export**: Export formatted with the presentation MIME type data URI prefix (`data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,...`).
86. **Express.js HTTP Pipe (`toExpress`)**: Dispatch presentations directly through Express `res` objects with automatic `Content-Type` and `Content-Disposition: attachment; filename="..."` headers.