# Context: Domain Dictionary & OpenXML Glossary

> **Scope**: Essential OpenXML, PresentationML, DrawingML, OPC, and EMU terminology.

---

### Standards & Packaging
- **OPC (Open Packaging Conventions)**: ISO/IEC 29500-2 / ECMA-376 Part 2 container packaging XML parts, media, and relationship graphs into a ZIP container or Flat OPC XML file.
- **`[Content_Types].xml`**: Mandatory package manifest declaring part MIME types via `<Default>` extensions and `<Override>` paths.
- **Relationships (`.rels`)**: Directed links connecting parts stored in `_rels/` folders (e.g. `ppt/_rels/presentation.xml.rels`). References use `r:id` or `r:embed`.
- **Flat OPC XML (`<pkg:package>`)**: Single-file XML encapsulation with `<pkg:xmlData>` for text/XML and Base64 `<pkg:binaryData>` for media/binaries.

### Namespaces & Schemas
- **PresentationML (`p:`)**: `http://schemas.openxmlformats.org/presentationml/2006/main` — Defines `<p:presentation>`, `<p:sld>`, `<p:cSld>`, and shape tree `<p:spTree>`.
- **DrawingML (`a:`)**: `http://schemas.openxmlformats.org/drawingml/2006/main` — Defines graphics, transforms (`<a:xfrm>`), text runs (`<a:r>`), formatting (`<a:rPr>`), and tables (`<a:tbl>`).
- **ChartML (`c:`)**: `http://schemas.openxmlformats.org/drawingml/2006/chart` — Defines `<c:chartSpace>`, series `<c:ser>`, category cache `<c:cat>`, and value cache `<c:val>`.
- **Office 2014 Extension (`a16:`)**: `http://schemas.microsoft.com/office/drawing/2014/main` — Defines `<a16:rowId>` unique 32-bit row identifiers.
- **Office 2010 Section (`p14:`)**: `http://schemas.microsoft.com/office/powerpoint/2010/main` — Defines `<p14:sectionLst>` slide grouping inside `<p:extLst>`.

### Spatial Math & Coordinate Units
- **EMU (English Metric Unit)**: Integer coordinates: $1 \text{ in} = 914,400$, $1 \text{ pt} = 12,700$, $1 \text{ px at 96 DPI} = 9,525$, $1 \text{ cm} = 360,000$.
- **16:9 Canvas Dimensions**: $12,192,000 \times 6,858,000 \text{ EMUs}$ ($13.333 \times 7.5 \text{ in}$).

### Engine Invariants
- **VFS**: In-memory `Map<string, Uint8Array>` eliminating disk I/O.
- **Run Healing**: Merges fragmented `<a:r>` runs across `{{placeholder}}` boundaries.
- **Chart Dual-Sync**: Concurrent update of DrawingML cache and embedded Excel `.xlsx` workbook.
- **Z-Order Invariant**: Visual shapes in `<p:spTree>` preceded by container headers `<p:nvGrpSpPr>` and `<p:grpSpPr>`.
