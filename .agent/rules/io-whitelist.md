# Rule: Strict 18-Method I/O API Whitelist

> **Scope**: Enforces the public I/O surface for presentation ingestion, serialization, and extraction in `node-pptx-templater`.  
> **Source of Truth**: See `docs-site/src/data/apiReference.json` for detailed parameter schemas and usage examples.

---

## 1. Approved 18-Method I/O Whitelist

The public I/O surface of `PPTXTemplater` is strictly locked to **exactly 18 approved methods**:

| Category | Count | Whitelisted Methods | Returns |
| :--- | :---: | :--- | :--- |
| **Static Lifecycle & Ingestion** | 7 | `PPTXTemplater.load(source, options)`<br>`PPTXTemplater.createBlank(options)`<br>`PPTXTemplater.extractPptx(pptxPath, outputDir, options)`<br>`PPTXTemplater.buildPptx(sourceDir, outputPath, options)`<br>`PPTXTemplater.preload(source, options)`<br>`PPTXTemplater.fromCache(source)`<br>`PPTXTemplater.clearCache()` | `Promise<PPTXTemplater>`<br>`Promise<PPTXTemplater>`<br>`Promise<void>`<br>`Promise<void>`<br>`Promise<Map>`<br>`Promise<PPTXTemplater>`<br>`void` |
| **Export Instance Methods** | 8 | `ppt.saveToFile(destinationPath, options)`<br>`ppt.toBuffer(options)`<br>`ppt.toStream(options)`<br>`ppt.toBase64(options)`<br>`ppt.toArrayBuffer(options)`<br>`ppt.toResponse(options)`<br>`ppt.toExpress(res, filename, options)`<br>`ppt.toFastify(reply, filename, options)` | `Promise<void>`<br>`Promise<Buffer>`<br>`Promise<Readable>`<br>`Promise<string>`<br>`Promise<ArrayBuffer>`<br>`Promise<Response>`<br>`Promise<void>`<br>`Promise<void>` |
| **Extraction & Conversion** | 3 | `ppt.extractToZip(destinationPath, options)`<br>`ppt.extractToXml(destinationDir, options)`<br>`ppt.toFlatXml(destinationPathOrOptions, options)` | `Promise<void>`<br>`Promise<void>`<br>`Promise<string\|void>` |

---

## 2. Prohibited Deprecated Aliases

Never introduce or accept any deprecated aliases. All operations must route through the whitelisted 18 methods:

| Prohibited Alias | Approved Replacement |
| :--- | :--- |
| `ppt.toFile()`, `ppt.save()`, `ppt.write()` | `ppt.saveToFile(path)` |
| `ppt.exportZip()` | `ppt.extractToZip(path)` |
| `ppt.exportXml()` | `ppt.extractToXml(dir)` |
| `ppt.writeToBuffer()`, `ppt.getBuffer()` | `ppt.toBuffer()` |
| `ppt.exportStream()` | `ppt.toStream()` |
| `PPTXTemplater.fromBuffer()`, `PPTXTemplater.fromFile()`, `PPTXTemplater.fromStream()` | `PPTXTemplater.load(source)` |
