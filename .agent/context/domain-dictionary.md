# Domain Dictionary: OpenXML & PPTX Terminology

This glossary defines standard Open Packaging Conventions (OPC) and OpenXML concepts, element tags, and metrics used throughout `node-pptx-templater`.

---

## 1. Metrics & Units

- **EMU (English Metric Unit)**:
  The fundamental integer unit of measurement in OpenXML. Designed to be evenly divisible by both inches and centimeters.
  - $1 \text{ pt} = 12,700 \text{ EMUs}$
  - $1 \text{ inch} = 914,400 \text{ EMUs}$
  - $1 \text{ pixel (at 96 DPI)} = 9,525 \text{ EMUs}$
  - $1 \text{ cm} = 360,000 \text{ EMUs}$
  - All coordinates in `<a:off>` and dimensions in `<a:ext>` must be integer EMUs.
- **Points (pt)**:
  Standard typography unit ($1/72\text{ inch}$). Font sizes in DrawingML are specified in hundredths of a point (e.g. `sz="1400"` for $14\text{ pt}$).
- **Degrees (60k)**:
  Rotations in DrawingML `<a:xfrm rot="...">` are measured in $1/60,000\text{ths}$ of a degree (e.g., $90^\circ = 5,400,000$).

---

## 2. Package & OPC Conventions

- **OPC (Open Packaging Conventions)**:
  The ZIP-based container standard defining parts, content types, and relationships.
- **Part**:
  A file within the ZIP container (e.g., `ppt/slides/slide1.xml`, `[Content_Types].xml`).
- **`[Content_Types].xml`**:
  The root MIME registry. Defines default extensions (`.png`, `.xml`, `.rels`) and explicit part overrides (`Override PartName="..." ContentType="..."`).
- **Relationships (`_rels`)**:
  XML files defining directed connections between parts.
  - Package root relationships: `_rels/.rels`
  - Presentation relationships: `ppt/_rels/presentation.xml.rels`
  - Slide relationships: `ppt/slides/_rels/slideX.xml.rels`
  - Chart relationships: `ppt/charts/_rels/chartX.xml.rels`
- **Relationship ID (`rId` / `r:id`)**:
  A unique alphanumeric identifier (e.g., `rId2`) linking an XML element to an entry in a `.rels` file.

---

## 3. PresentationML Elements (`p:`)

- `<p:sld>`: The root container for a slide.
- `<p:spTree>` (Shape Tree): The container holding all visual slide elements (shapes, pictures, tables, charts, groups). Order determines visual Z-index.
- `<p:sp>`: A standard shape or text box element.
- `<p:pic>`: An image element. Links to a media file via `r:embed="rIdX"`.
- `<p:graphicFrame>`: Container for complex graphics including tables and charts.
- `<p:nvSpPr>`: Non-visual shape properties (shape ID, name, locking).
- `<p:txBody>`: Text body containing paragraphs and formatting.

---

## 4. DrawingML Elements (`a:`)

- `<a:xfrm>` (2D Transform): Defines spatial bounding box.
  - `<a:off x="..." y="..."/>`: Offset coordinates from top-left.
  - `<a:ext cx="..." cy="..."/>`: Extent width (`cx`) and height (`cy`).
- `<a:p>`: A text paragraph.
  - `<a:pPr>`: Paragraph properties (alignment, line spacing, margins).
- `<a:r>`: A text run.
  - `<a:rPr>`: Run properties (font size, bold, italic, color `<a:solidFill>`).
  - `<a:t>`: Text content node.
- `<a:br>`: Line break run.

---

## 5. Table Structures

- `<a:tbl>`: Table root element.
- `<a:tblPr>`: Table properties (style ID, band rows, first column styling).
- `<a:tblGrid>`: Column definitions container.
  - `<a:gridCol w="..."/>`: Defines the width of a column in EMUs.
- `<a:tr h="..."/>`: Table row with height in EMUs.
  - `<a16:rowId>`: Unique 32-bit unsigned integer ID required by PowerPoint 2016+.
- `<a:tc>`: Table cell.
  - `<a:tcPr>`: Cell properties (margins, fill, borders, vertical alignment).
  - `gridSpan`: Number of horizontal columns merged by this cell.
  - `rowSpan`: Number of vertical rows merged by this cell.
  - `hMerge="1"`: Indicator that this cell is horizontally covered by a preceding merged cell.
  - `vMerge="1"`: Indicator that this cell is vertically covered by an upper merged cell.

---

## 6. Chart Structures (`c:`)

- `<c:chartSpace>`: Root container of a chart part (`ppt/charts/chartX.xml`).
- `<c:chart>`: Main chart node containing title, plot area, and legend.
- `<c:plotArea>`: The area containing visual charts (`<c:barChart>`, `<c:lineChart>`, `<c:pieChart>`, `<c:areaChart>`).
- `<c:ser>`: A data series.
  - `<c:tx>`: Series title.
  - `<c:cat>`: Category axis labels (`<c:strRef>` or `<c:numRef>`).
  - `<c:val>`: Data values (`<c:numRef>`).
  - `<c:ptCount @_val="N"/>`: Total point count in cache.
  - `<c:pt @_idx="i">`: Data point element at 0-based index `i`.
- **Embedded Workbook**:
  `ppt/embeddings/Microsoft_Excel_Worksheet*.xlsx` storing the backing table queried by PowerPoint's "Edit Data in Excel" tool.

---

## 7. File Formats & MIME Types

- **`.pptx`**:
  `application/vnd.openxmlformats-officedocument.presentationml.presentation` — Standard OpenXML presentation archive.
- **`.pptm`**:
  `application/vnd.ms-powerpoint.presentation.macroEnabled.12` — Macro-enabled presentation archive.
- **`.potx`**:
  `application/vnd.openxmlformats-officedocument.presentationml.template` — OpenXML presentation template.
- **`.potm`**:
  `application/vnd.ms-powerpoint.template.macroEnabled.12` — Macro-enabled presentation template.
- **`.ppsx`**:
  `application/vnd.openxmlformats-officedocument.presentationml.slideshow` — Auto-playing slide show presentation.
- **`.ppsm`**:
  `application/vnd.ms-powerpoint.slideshow.macroEnabled.12` — Macro-enabled auto-playing slide show.
- **`.ppam`**:
  `application/vnd.ms-powerpoint.addin.macroEnabled.12` — PowerPoint Add-in package.
- **`.ppt`**:
  `application/vnd.ms-powerpoint` — Legacy Microsoft PowerPoint binary format (OLE Compound File, signature `0xD0 0xCF 0x11 0xE0 0xA1 0xB1 0x1A 0xE1`).
- **`.odp`**:
  `application/vnd.oasis.opendocument.presentation` — OpenDocument Presentation XML package.
- **`.zip`**:
  `application/zip` — PKZip container format (magic number `0x50 0x4B 0x03 0x04`).

---

## 8. Flat OPC XML Concepts & Tags

- **Flat OPC (Microsoft Flat Open Packaging Conventions)**:
  Single uncompressed XML representation of an OpenXML package, containing a `<pkg:package>` root node enclosing individual `<pkg:part>` entries.
- **`<pkg:package>`**:
  The root container element of a Flat OPC XML document (`xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"`).
- **`<pkg:part>`**:
  An element representing an OPC part within a Flat OPC document, defining attributes `pkg:name` (e.g. `"/ppt/presentation.xml"`) and `pkg:contentType` (e.g. `"application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"`).
- **`<pkg:xmlData>`**:
  Child container within a `<pkg:part>` holding raw, unescaped OpenXML markup for XML parts (slides, charts, relationships).
- **`<pkg:binaryData>`**:
  Child container within a `<pkg:part>` holding base64-encoded binary content for binary assets such as images (`ppt/media/`) or embedded Excel workbooks (`ppt/embeddings/`).
