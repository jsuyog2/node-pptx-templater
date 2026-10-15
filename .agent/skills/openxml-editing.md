# Skill: OpenXML AST Editing & Schema Preservation

> **Domain**: ECMA-376 / ISO/IEC 29500 PresentationML & DrawingML Manipulation  
> **Target Subsystems**: `XMLParser.js`, `SlideManager.js`, `ShapeManager.js`, `RelationshipManager.js`

---

## 1. Context & Overview

OpenXML PPTX documents are sensitive to XML tag order, namespace definitions, and relationship consistency. Unlike HTML, OpenXML will fail with a repair prompt if elements are out of schema sequence or if relationship identifiers are duplicated or missing.

This skill provides step-by-step instructions for safely modifying the OpenXML AST inside `node-pptx-templater`.

---

## 2. Namespace Conventions

Always preserve the standard OpenXML prefixes when inspecting or creating nodes:

| Prefix | Namespace URI | Description |
| :--- | :--- | :--- |
| `p:` | `http://schemas.openxmlformats.org/presentationml/2006/main` | PresentationML (slides, layouts, masters, shapes) |
| `a:` | `http://schemas.openxmlformats.org/drawingml/2006/main` | DrawingML (text runs, geometry, fills, outlines) |
| `r:` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | Relationships (rId attribute references) |
| `c:` | `http://schemas.openxmlformats.org/drawingml/2006/chart` | ChartML (chart spaces, series, data points) |
| `a16:` | `http://schemas.microsoft.com/office/drawing/2014/main` | DrawingML 2014 Extension (e.g. `a16:rowId`) |
| `pkg:` | `http://schemas.microsoft.com/office/2006/xmlPackage` | Flat OPC XML envelope |

---

## 3. Parsing & Mutating Procedure

### Step 1: Read and Parse XML via `XMLParser`
Never use standard regex search-and-replace for structural XML edits. Always parse through `XMLParser`:
```javascript
const slideXml = slideManager.getSlideXml(slideIndex);
const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`);
```

### Step 2: Navigate the AST Hierarchy
The standard slide element tree follows this hierarchy:
```
p:sld
└── p:cSld
    └── p:spTree (Shape Tree)
        ├── p:nvGrpSpPr (Non-visual group properties)
        ├── p:grpSpPr   (Group transform properties)
        ├── p:sp        (Vector shapes, text boxes)
        ├── p:pic       (Images)
        ├── p:graphicFrame (Tables, Charts, SmartArt)
        ├── p:grpSp     (Grouped shapes)
        └── p:cxnSp     (Connectors)
```

### Step 3: Mutate Attributes and Elements
- All XML attributes are prefixed with `@_` (e.g., `item['@_id']`, `item['@_name']`).
- When modifying text, target `<a:t>` inside text runs `<a:r>`:
  ```javascript
  const p = shape['p:txBody']['a:p'];
  const runs = Array.isArray(p['a:r']) ? p['a:r'] : [p['a:r']];
  runs[0]['a:t'] = 'Updated Text Content';
  ```

### Step 4: Protect Container Order During Serialization
When serializing `<p:spTree>` or `<p:grpSp>` back to XML, the property header elements MUST remain first:
```javascript
// Header elements must appear before drawing children:
// 1. <p:nvGrpSpPr>
// 2. <p:grpSpPr>
// 3. Child shapes in Z-order
const decl = this.#xmlParser.extractDeclaration(slideXml);
const updatedXml = this.#xmlParser.build(slideObj, decl);
slideManager.setSlideXml(slideIndex, updatedXml);
```

---

## 4. Relationship Management Protocol

Every external or internal asset referenced by a slide (layouts, images, charts, hyperlinks) requires an entry in `ppt/slides/_rels/slideX.xml.rels`.

1. **Allocating a Relationship**:
   ```javascript
   const rId = relationshipManager.addSlideRelationship(
     slideIndex,
     'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
     '../media/image1.png'
   );
   ```
2. **Referencing the Relationship**:
   In `slideX.xml`, reference the relationship using `r:embed` or `r:id`:
   ```xml
   <p:pic>
     <p:blipFill>
       <a:blip r:embed="rId2"/>
     </p:blipFill>
   </p:pic>
   ```
3. **Cleaning Up on Deletion**:
   When deleting a shape or slide, always purge orphaned relationships from the corresponding `.rels` file to prevent dangling reference warnings.
