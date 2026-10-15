# Skill: Microsoft Flat OPC XML Ingestion & Serialization

> **Domain**: ISO/IEC 29500-2 Part 2 Annex A Flat OPC (`<pkg:package>`)  
> **Target Subsystems**: `FlatOpcConverter.js`, `PackageLoader.js`, `OutputWriter.js`

---

## 1. Context & Use Cases

Microsoft Flat OPC (Open Packaging Conventions) XML represents an entire multi-file OpenXML archive (normally compressed as a PKZip `.pptx` container) within a **single plain-text XML document**.

### Primary Enterprise Use Cases:
1. **Git Version Control**: Binary `.pptx` files produce un-diffable git conflicts. Flat OPC XML enables line-by-line diffing and merge resolution of slide decks in version control.
2. **Database Storage**: Store presentations directly in PostgreSQL/MySQL `TEXT` or `JSON` columns without binary BLOB overhead.
3. **Serverless & Edge Environments**: Zero decompression overhead during cold-starts on AWS Lambda, Cloudflare Workers, or Vercel Edge.
4. **Fast Text Substitution**: Perform rapid template transformations without ZIP inflation/deflation.

---

## 2. Flat OPC XML Structure

A Flat OPC presentation has the following standard anatomy:

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<?mso-application progid="PowerPoint.Show"?>
<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">
  
  <!-- 1. Relationships Part -->
  <pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">
    <pkg:xmlData>
      <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
      </Relationships>
    </pkg:xmlData>
  </pkg:part>

  <!-- 2. Presentation Part -->
  <pkg:part pkg:name="/ppt/presentation.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml">
    <pkg:xmlData>
      <p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ...>
        ...
      </p:presentation>
    </pkg:xmlData>
  </pkg:part>

  <!-- 3. Binary Part (Images, XLSX Embeddings) -->
  <pkg:part pkg:name="/ppt/media/image1.png" pkg:contentType="image/png">
    <pkg:binaryData>
      iVBORw0KGgoAAAANSUhEUgAA... (Base64 Encoded Binary Data)
    </pkg:binaryData>
  </pkg:part>

</pkg:package>
```

---

## 3. Ingestion Procedure (`FlatOpcConverter.convert`)

When `PPTXTemplater.load('deck.xml')` or `PackageLoader.load(xmlBuffer)` detects a Flat OPC package (matching `<pkg:package` or `http://schemas.microsoft.com/office/2006/xmlPackage`):

1. **Part Regex Scan**:
   Scan for `<pkg:part pkg:name="([^"]+)" pkg:contentType="([^"]+)">`.
2. **VFS Path Normalization**:
   Strip the leading slash: `/ppt/presentation.xml` becomes `ppt/presentation.xml`.
3. **XML Parts Extraction (`<pkg:xmlData>`)**:
   Extract inner XML string. Ensure proper XML declaration is prepended if absent before writing to `vfs.writeText()`.
4. **Binary Parts Extraction (`<pkg:binaryData>`)**:
   Strip all whitespace and newlines from the payload, convert from Base64 into a raw `Buffer` or `Uint8Array`, and write via `vfs.writeBytes()`.
5. **Content Types Synthesis**:
   If the Flat OPC file does not contain a separate `/[Content_Types].xml` part, synthesize one dynamically by collecting each part's `pkg:contentType` attribute into `<Default>` and `<Override>` declarations.

---

## 4. Export Procedure (`ppt.toFlatXml`)

To serialize any loaded presentation (whether loaded from `.pptx`, folder, or URL) into Flat OPC XML:

```javascript
// (a) Save directly to file path
await ppt.toFlatXml('./output/presentation.xml', { pretty: true });

// (b) Retrieve as an in-memory string
const flatXmlString = await ppt.toFlatXml({ pretty: false });
```

### Serialization Rules:
1. Include XML declaration and `<?mso-application progid="PowerPoint.Show"?>`.
2. Open `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">`.
3. Iterate over the in-memory VFS parts:
   - For binary media and workbooks, encode bytes into Base64 inside `<pkg:binaryData>`.
   - For XML parts, strip any leading `<?xml ... ?>` header and place inside `<pkg:xmlData>`.
   - Infer proper `pkg:contentType` based on file extension and part path.
4. Exclude `[Content_Types].xml` as an explicit part (since each `<pkg:part>` declares its own `pkg:contentType`).
5. Close `</pkg:package>`.
