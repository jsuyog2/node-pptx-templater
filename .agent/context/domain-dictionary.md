# Context: Domain Dictionary & OpenXML Glossary

> **Scope**: Essential OpenXML, PresentationML, DrawingML, OPC, and EMU terminology.

---

### Standards & Packaging
- **OPC (Open Packaging Conventions)**: ISO/IEC 29500-2 / ECMA-376 Part 2 container packaging XML parts, media, and relationship graphs into a ZIP container or Flat OPC XML file.
- **`[Content_Types].xml`**: Mandatory root part declaring package MIME types via `<Default Extension="..." ContentType="..."/>` and `<Override PartName="..." ContentType="..."/>`.
- **Relationships (`.rels` / `r:id`)**: Directed links connecting parts. Located in `_rels/` subdirectories (e.g. `ppt/_rels/presentation.xml.rels`, `ppt/slides/_rels/slide1.xml.rels`). XML references use `r:id` or `r:embed`.
- **Flat OPC XML (`<pkg:package>`)**: Single-file XML encapsulation where parts are wrapped in `<pkg:part>` containing `<pkg:xmlData>` (text/XML) or Base64 `<pkg:binaryData>` (media, XLSX).

### Namespaces & PresentationML
- **PresentationML (`p:`)**: `http://schemas.openxmlformats.org/presentationml/2006/main` — Defines `<p:presentation>`, `<p:sldIdLst>`, `<p:sldId>` (`@_id >= 256`), `<p:sld>`, `<p:cSld>`, and the shape container `<p:spTree>`.
- **DrawingML (`a:`)**: `http://schemas.openxmlformats.org/drawingml/2006/main` — Defines graphics, transforms (`<a:xfrm>`, `<a:off>`, `<a:ext>`), paragraphs (`<a:p>`), runs (`<a:r>`), run properties (`<a:rPr>`), text (`<a:t>`), and tables (`<a:tbl>`, `<a:tr>`, `<a:tc>`).
- **ChartML (`c:`)**: `http://schemas.openxmlformats.org/drawingml/2006/chart` — Defines `<c:chartSpace>`, `<c:plotArea>`, `<c:ser>`, category cache `<c:cat>`, and value cache `<c:val>`.
- **Office 2014 Extension (`a16:`)**: `http://schemas.microsoft.com/office/drawing/2014/main` — `<a16:rowId val="...">` assigns unique 32-bit unsigned integers to table rows.
- **Office 2010 Section (`p14:`)**: `http://schemas.microsoft.com/office/powerpoint/2010/main` — `<p14:sectionLst>` groups slides within `<p:extLst>`.

### Spatial Coordinates & EMUs
- **EMU (English Metric Unit)**: Integer coordinate base ($1 \text{ in} = 914,400 \text{ EMUs}$, $1 \text{ pt} = 12,700 \text{ EMUs}$, $1 \text{ px at 96 DPI} = 9,525 \text{ EMUs}$, $1 \text{ cm} = 360,000 \text{ EMUs}$).
- **16:9 Slide Dimensions**: $12,192,000 \times 6,858,000 \text{ EMUs}$ ($13.333 \times 7.5 \text{ in}$).

### Engine Invariants
- **VFS (Virtual File System)**: In-memory `Map<string, Uint8Array>` eliminating synchronous disk I/O.
- **Run Fragment Healing**: Reconstitutes split placeholder tokens across fragmented `<a:r>` runs.
- **Shadowed Cells**: Cells hidden by merges; horizontal declare `hMerge="1"`, vertical declare `vMerge="1"`.
- **Chart Dual-Sync**: Concurrent update of `<c:chartSpace>` cache and embedded `ppt/embeddings/*.xlsx` workbook.
- **Z-Order Invariant**: Visual element sequence in `<p:spTree>` with container headers `<p:nvGrpSpPr>` and `<p:grpSpPr>` locked to the top.
