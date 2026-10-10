# Feature Catalog and Implementation Verification

This document captures the target capability catalog supplied for development. It is a requirements inventory, not a claim that every item exists in the current checkout.

## Verification labels
For each requested capability, determine its actual status from source, tests, exports, docs, and examples:
- **Implemented and verified** — source exists and tests demonstrate expected behavior.
- **Implemented, coverage incomplete** — source exists but tests do not adequately establish behavior.
- **Partially implemented** — only some cases or variants work.
- **Documented only / not found** — docs mention it but source does not establish support.
- **Not implemented** — no credible implementation found.
- **Blocked / environment-dependent** — cannot be verified in the current environment.

Never mark a feature complete based only on a README, website, type declaration, feature name, or this catalog. Track gaps and link each conclusion to exact files and tests.

## Target capability inventory

### A. Input and initialization
1. Load PPTX from a local filesystem path.
2. Load from Node.js Buffer or Uint8Array.
3. Ingest Node.js Readable or Web ReadableStream inputs where supported by the actual runtime/API.
4. Detect ZIP archives by signature, including custom file extensions.
5. Load an unpacked OpenXML directory containing the expected package structure.
6. Create a blank presentation from code.
7. Configure custom slide dimensions and common aspect ratios (16:9, 16:10, 4:3).

### B. Presentation and slide lifecycle
8. Read slide count, dimensions, units, and layout metadata.
9. Add blank slides using master/layout choices.
10. Delete slides by zero-based index or internal identifier as specified by the API.
11. Reorder slides.
12. Deep-clone slides while preserving element trees and layouts.
13. Clone slide relationships with collision-free relationship IDs.
14. Synchronize content types and presentation manifests after slide mutations.

### C. Spatial engine
15. Normalize EMU conversions: 1 pt = 12,700 EMUs; 1 in = 914,400 EMUs; 1 px = 9,525 EMUs, subject to documented pixel assumptions.
16. Extract element bounds and rotation.
17. Position and resize elements using supported points, pixels, inches, or EMUs.
18. Manage z-order: bring forward/front and send backward/back.
19. Detect boundary overflow and bounding-box intersections/collisions.

### D. Element discovery
20. Enumerate visual elements on a slide.
21. Filter elements by text, shape, image, table, chart, or group type where supported.
22. Select by shape name or non-visual shape ID.
23. Select elements by template placeholder/token.

### E. Text and typography
24. Normalize fragmented DrawingML `<a:r>` / `<a:t>` text runs so placeholders split across XML runs can be matched.
25. Replace text without unnecessarily destroying surrounding formatting.
26. Append/prepend text runs.
27. Change font family at text-frame or run scope.
28. Change font size using documented units.
29. Change text color using supported hex/RGB/theme forms.
30. Apply bold, italic, underline, strike, subscript, and superscript.
31. Set horizontal paragraph alignment.
32. Set vertical text-frame alignment.
33. Configure line spacing and paragraph spacing.
34. Configure wrapping behavior.

### F. Shapes
35. Add supported vector shapes: rectangles, rounded rectangles, ellipses, triangles, arrows, callouts, lines, and other explicitly implemented types.
36. Apply solid fills.
37. Apply gradient and pattern fills where supported.
38. Configure fill transparency/opacity.
39. Configure outline color, width, and solid/dashed/dotted patterns.
40. Insert and style text within shapes, including vertical alignment and internal padding.
41. Remove shapes without leaving invalid references.

### G. Images and media
42. Insert images from paths, buffers, and Base64 data URIs where supported.
43. Handle PNG, JPEG, SVG, GIF, and WebP only where the package and target Office consumers support them; verify actual content-type behavior.
44. Replace image placeholders while preserving geometry.
45. Preserve or explicitly calculate image aspect ratio.
46. Extract image dimensions, format, byte size, and relationship information.
47. Remove image elements and safely clean relationships/media when no longer referenced.

### H. Tables
48. Insert tables with rows, columns, coordinates, and dimensions.
49. Populate tables from 2D arrays or structured objects.
50. Add, append, delete, and duplicate rows.
51. Add, append, and delete columns.
52. Query cells with zero-based row/column indices.
53. Update cell text and formatting.
54. Apply cell/row background fills and opacity.
55. Configure individual cell-edge borders.
56. Merge columns with gridSpan/hMerge structures as required by DrawingML.
57. Merge rows with rowSpan/vMerge structures as required by DrawingML.
58. Unmerge cells safely.
59. Embed supported shapes/icons/badges/indicators inside cells.
60. Style header rows independently.
61. Apply alternating row striping.
62. Extract table dimensions, row heights, column widths, and metadata.

### I. Dynamic table pagination and slide splitting
63. Estimate row heights using text, font size, wrapping, padding, and column width.
64. Detect table overflow against slide height and configured margins.
65. Clone the host slide with its required layout, header/footer, background, and relationships.
66. Split table rows at the last row that fits.
67. Repeat table headers on generated overflow slides.
68. Continue pagination until all rows are placed.
69. Preserve designated totals/summary rows on the final generated slide.

### J. Native charts
70. Find embedded charts by name, ID, or slide context.
71. Support Bar, Column, Line, Pie, Doughnut, Area, and Scatter chart types where actually implemented.
72. Update series values, categories, and labels.
73. Synchronize embedded XLSX workbook/cache with chart XML where applicable.
74. Position and resize chart elements.
75. Toggle legends, data labels, axes, and gridlines.

### K. Hyperlinks and navigation
76. Add external URLs to supported text runs, text frames, shapes, and images.
77. Add internal links to slides or navigation targets supported by the API.
78. Inspect existing external destinations and internal slide targets.
79. Remove links without damaging the linked element's remaining styling/content.

### L. Output and serialization
80. Save to a local file asynchronously.
81. Export as a Node.js readable stream where supported.
82. Export as a Node.js Buffer.
83. Export as an ArrayBuffer.
84. Export Base64.
85. Export a PowerPoint MIME-type data URI.
86. Write a presentation to an Express response with appropriate MIME and attachment headers, where supported.

## Engineering priorities
- Fix data corruption, dangling relationships, malformed XML, security defects, and broken output before cosmetic refactors.
- Preserve the documented public API and compatibility by default.
- Build missing features in coherent increments, with tests and documentation for each increment.
- Do not implement all 86 items in one giant change. Group work by subsystem and keep each change reviewable.
- For any catalog item that cannot be safely supported across all PPTX variants, document the supported subset and failure modes rather than overpromising.
- Maintain a living status report in a suitable existing documentation location or create `docs/feature-status.md` only if appropriate to the repository's documentation structure.
