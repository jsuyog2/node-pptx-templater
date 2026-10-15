# Context: Domain Dictionary & OpenXML Glossary

> **Scope**: Exhaustive technical glossary of OpenXML, PresentationML, DrawingML, OPC, and `node-pptx-templater` terminology.

---

## 1. Standards & Packaging Concepts

### OpenXML / ECMA-376 / ISO/IEC 29500
The international open standard defining XML formats for office applications:
- **Part 1**: Fundamentals and Markup Language Reference (PresentationML, DrawingML, SpreadsheetML).
- **Part 2**: Open Packaging Conventions (OPC), defining how parts are packaged inside ZIP archives or Flat OPC XML files.

### OPC (Open Packaging Conventions)
A container technology that packages multiple XML documents, binary media files, and relationship graphs into a single file (usually a standard PKZip archive). Every file inside an OPC container is termed a **Part**.

### `[Content_Types].xml`
The mandatory package-level part residing at the root of every OPC container. It declares the MIME type of every part in the package using two elements:
- `<Default Extension="ext" ContentType="mime/type"/>`: Maps generic file extensions (`xml`, `rels`, `png`, `jpeg`).
- `<Override PartName="/ppt/slides/slide1.xml" ContentType="mime/type"/>`: Explicitly maps specific parts to exact OpenXML schemas.

### Relationships (`.rels` / `r:id`)
OPC files never link to other parts via direct raw paths; they link via relationship identifiers (`rId`).
- Relationship files are named `[partName].rels` and live in a `_rels/` subfolder (e.g. `ppt/_rels/presentation.xml.rels`, `ppt/slides/_rels/slide1.xml.rels`).
- Each `<Relationship Id="rId1" Type="..." Target="..." [TargetMode="External"]/>` establishes a directed connection.
- References in XML use `r:id="rId1"` or `r:embed="rId1"`.

### Flat OPC XML (`<pkg:package>`)
ISO/IEC 29500-2 Part 2 Annex A representation of an entire OPC container within a single monolithic XML document:
- `<pkg:package>`: The root envelope element.
- `<pkg:part pkg:name="/..." pkg:contentType="...">`: Represents each individual part.
- `<pkg:xmlData>`: Contains inline XML for text parts (stripping the root `<?xml...?>` header).
- `<pkg:binaryData>`: Contains Base64-encoded binary content for images, fonts, and embedded workbooks.

---

## 2. PresentationML & Markup Namespaces

### PresentationML (`p:` Namespace)
Namespace: `http://schemas.openxmlformats.org/presentationml/2006/main`  
Defines PowerPoint presentation structure:
- `<p:presentation>`: Presentation root defining slide sizing, slide ID list, and slide masters.
- `<p:sldIdLst>`: The ordered list of slides in the presentation.
- `<p:sldId id="256" r:id="rId2"/>`: Represents an individual slide. The `@_id` attribute must be $\ge 256$.
- `<p:sld>`: The root element of a slide document (`ppt/slides/slideX.xml`).
- `<p:cSld>`: Common slide data wrapper containing the background and shape tree.
- `<p:spTree>`: The shape tree container holding all drawable elements on a slide.
- `<p:sp>`: Vector shape or text box element.
- `<p:pic>`: Picture/image element.
- `<p:graphicFrame>`: Container for complex visual elements (tables, charts, SmartArt).
- `<p:grpSp>`: Group shape containing nested visual shapes.

### DrawingML (`a:` Namespace)
Namespace: `http://schemas.openxmlformats.org/drawingml/2006/main`  
Defines graphics, geometry, text styling, colors, and coordinates across all Office applications:
- `<a:xfrm>`: 2D transform defining position (`<a:off x="..." y="..."/>`) and extent/dimensions (`<a:ext cx="..." cy="..."/>`).
- `<a:p>`: Text paragraph.
- `<a:r>`: Text run holding contiguous characters with identical styling.
- `<a:rPr>`: Text run properties (font size `@_sz`, bold `@_b`, italic `@_i`, Latin font `<a:latin typeface="..."/>`).
- `<a:t>`: Text content node holding raw string characters.
- `<a:tbl>`: Table element containing grid columns (`<a:gridCol>`) and rows (`<a:tr>`).
- `<a:tc>`: Table cell.

### ChartML (`c:` Namespace)
Namespace: `http://schemas.openxmlformats.org/drawingml/2006/chart`  
Defines native charts:
- `<c:chartSpace>`: Root chart element in `ppt/charts/chartX.xml`.
- `<c:plotArea>`: Container for the actual chart visualization (`<c:barChart>`, `<c:lineChart>`, `<c:pieChart>`).
- `<c:ser>`: A chart series.
- `<c:cat>`: Category data reference and cache (`<c:strRef>/<c:strCache>`).
- `<c:val>`: Values data reference and cache (`<c:numRef>/<c:numCache>`).
- `<c:dLbls>`: Data label configuration.

### DrawingML 2014 Extension (`a16:` Namespace)
Namespace: `http://schemas.microsoft.com/office/drawing/2014/main`  
Modern Office extensions:
- `<a16:rowId val="...">`: 32-bit unsigned integer uniquely identifying a table row in PowerPoint 2016+. Must be unique across all rows in a table.

---

## 3. Spatial Math & Units

### EMU (English Metric Unit)
The fundamental coordinate unit in OpenXML DrawingML. EMUs allow integer arithmetic to represent both fractional inches and millimeters without rounding errors:
- $1 \text{ pt} = 12,700 \text{ EMUs}$
- $1 \text{ in} = 914,400 \text{ EMUs}$
- $1 \text{ px (at 96 DPI)} = 9,525 \text{ EMUs}$
- $1 \text{ cm} = 360,000 \text{ EMUs}$
- $1 \text{ mm} = 36,000 \text{ EMUs}$

### Standard Slide Dimensions in EMUs:
- **16:9 Widescreen**: $12,192,000 \times 6,858,000 \text{ EMUs}$ ($13.333 \times 7.5 \text{ in}$)
- **4:3 Standard**: $9,144,000 \times 6,858,000 \text{ EMUs}$ ($10 \times 7.5 \text{ in}$)
- **16:10 Widescreen**: $12,192,000 \times 7,620,000 \text{ EMUs}$ ($13.333 \times 8.333 \text{ in}$)

---

## 4. Engine Internal Architecture Terms

### VFS (Virtual File System)
An in-memory `Map<string, Uint8Array>` implemented in `VirtualFileSystem.js`. Decouples the templating engine from disk I/O, allowing pure in-memory execution in serverless lambdas and cloud workers.

### Run Fragmentation
The phenomenon where a continuous string of text (such as a placeholder `{{customer_name}}`) is chopped across multiple `<a:r>` and `<a:t>` nodes by Microsoft PowerPoint. Handled by `TemplateEngine.js` via text run healing.

### Shadowed Cells
Table cells that are covered by an adjacent cell with `gridSpan > 1` or `rowSpan > 1`.
- Horizontally shadowed cells must declare `hMerge="1"`.
- Vertically shadowed cells must declare `vMerge="1"`.

### Visual Stacking Order (Z-Order)
The rendering hierarchy of shapes on a slide. In OpenXML, Z-order is implicit: shapes declared later in `<p:spTree>` render on top of earlier shapes. Managed via `ZOrderManager.js` and `Z_ORDER_SYMBOL`.

### Chart Dual-Sync
The requirement to simultaneously update both the DrawingML visual cache (`<c:chartSpace>`) and the underlying embedded Excel spreadsheet (`ppt/embeddings/*.xlsx`) to prevent PowerPoint repair dialogs.
