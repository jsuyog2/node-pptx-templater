# Skill: OpenXML AST Editing & Schema Preservation

> **Domain**: ECMA-376 / ISO/IEC 29500 PresentationML & DrawingML Manipulation  
> **Target Subsystems**: `XMLParser.js`, `SlideManager.js`, `ShapeManager.js`, `RelationshipManager.js`

---

## 1. Namespaces & AST Hierarchy

| Prefix | Namespace URI | Usage |
| :--- | :--- | :--- |
| `p:` | `http://schemas.openxmlformats.org/presentationml/2006/main` | PresentationML (slides, shapes, layouts) |
| `a:` | `http://schemas.openxmlformats.org/drawingml/2006/main` | DrawingML (geometry, formatting, text, tables) |
| `r:` | `http://schemas.openxmlformats.org/officeDocument/2006/relationships` | Relationships (`r:id`, `r:embed`) |
| `c:` | `http://schemas.openxmlformats.org/drawingml/2006/chart` | Native charts |
| `a16:` | `http://schemas.microsoft.com/office/drawing/2014/main` | DrawingML extensions (`<a16:rowId>`) |

```
p:sld ─► p:cSld ─► p:spTree
                     ├─► p:nvGrpSpPr (Header 1: Non-visual group properties)
                     ├─► p:grpSpPr   (Header 2: Group transform properties)
                     └─► [p:sp, p:pic, p:graphicFrame, p:grpSp] (Ordered shapes)
```

---

## 2. Safe AST Mutation SOP

1. **Parse via `XMLParser`**:
   ```javascript
   const slideXml = slideManager.getSlideXml(slideIndex);
   const slideObj = this.#xmlParser.parse(slideXml, `slide${slideIndex}.xml`);
   ```
2. **Access Attributes & Nodes**: Attributes use `@_` prefix (`item['@_id']`). Text runs are located in `<a:p>/<a:r>/<a:t>`.
3. **Preserve Container Headers**:
   - Headers `<p:nvGrpSpPr>` and `<p:grpSpPr>` MUST remain the first children in `<p:spTree>` and `<p:grpSp>`.
   - Never place visual shapes (`<p:sp>`, `<p:pic>`) ahead of container headers.
4. **Serialize with Declaration**:
   ```javascript
   const decl = this.#xmlParser.extractDeclaration(slideXml);
   const updatedXml = this.#xmlParser.build(slideObj, decl);
   slideManager.setSlideXml(slideIndex, updatedXml);
   ```

---

## 3. Relationship Management SOP

- **Allocate**: `const rId = relationshipManager.addSlideRelationship(slideIndex, type, target);`
- **Embed**: Reference `rId` in shape XML via `r:embed="rIdX"` or `r:id="rIdX"`.
- **Cleanup**: Purge orphaned relationship entries from `ppt/slides/_rels/slideX.xml.rels` when shapes or parts are removed.
