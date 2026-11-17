# Skill: Microsoft Flat OPC XML Ingestion & Serialization

> **Domain**: ISO/IEC 29500-2 Part 2 Annex A Flat OPC (`<pkg:package>`)  
> **Target Subsystems**: `FlatOpcConverter.js`, `PackageLoader.js`, `OutputWriter.js`

---

## 1. Structure Overview

Flat OPC encapsulates an entire multi-part OpenXML container into a single XML document:
- **XML Parts**: Serialized in `<pkg:part pkg:name="/..." pkg:contentType="...">` containing `<pkg:xmlData>`. Inner XML declarations (`<?xml...?>`) are stripped.
- **Binary Parts**: Serialized in `<pkg:part>` containing Base64-encoded bytes in `<pkg:binaryData>` (images, embedded XLSX).

---

## 2. Ingestion SOP (`FlatOpcConverter.convert`)

1. **Detection**: Triggers when input XML contains `<pkg:package>` or namespace `http://schemas.microsoft.com/office/2006/xmlPackage`.
2. **Path Normalization**: Strip leading slash (`/ppt/presentation.xml` -> `ppt/presentation.xml`).
3. **XML Parts**: Extract `<pkg:xmlData>` content and commit via `vfs.writeText()`.
4. **Binary Parts**: Extract `<pkg:binaryData>`, decode Base64 into `Uint8Array`, and commit via `vfs.writeBytes()`.
5. **Content Types Synthesis**: If no explicit `/[Content_Types].xml` part exists, dynamically synthesize one from the collected `pkg:contentType` attributes.

---

## 3. Export SOP (`ppt.toFlatXml`)

```javascript
// (a) Save to disk
await ppt.toFlatXml('./output/presentation.xml', { pretty: true });

// (b) Return as in-memory string
const flatXmlString = await ppt.toFlatXml({ pretty: false });
```

1. Emit `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` and `<?mso-application progid="PowerPoint.Show"?>`.
2. Open `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">`.
3. Iterate VFS parts:
   - Exclude `[Content_Types].xml` (content types are specified on each `<pkg:part>` attribute).
   - Strip `<?xml...?>` from XML parts and wrap in `<pkg:xmlData>`.
   - Base64-encode binary parts inside `<pkg:binaryData>`.
4. Close `</pkg:package>`.
