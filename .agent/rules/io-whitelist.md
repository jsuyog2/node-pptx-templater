# Rule: Strict 18-Method I/O API Whitelist

> **Scope**: Enforces the architectural boundary for presentation ingestion, serialization, and extraction in `node-pptx-templater`.  
> **Target Class**: `PPTXTemplater` (`src/core/PPTXTemplater.js`) and `OutputWriter` (`src/writer/OutputWriter.js`).

---

## 1. The Strict 18-Method Whitelist

The public I/O surface of `node-pptx-templater` is locked to **exactly 18 approved methods**. No other I/O methods or aliases may be introduced or exported.

```
┌────────────────────────────────────────────────────────────────────────────────────────┐
│                        STRICT 18-METHOD I/O WHITELIST BOUNDARY                         │
├──────────────────────────┬────────────────────────────┬────────────────────────────────┤
│   STATIC METHODS (7)     │   EXPORT METHODS (8)       │ EXTRACTION & CONVERSION (3)    │
├──────────────────────────┼────────────────────────────┼────────────────────────────────┤
│ 1. load                  │ 8.  saveToFile             │ 16. extractToZip               │
│ 2. createBlank           │ 9.  toBuffer               │ 17. extractToXml               │
│ 3. extractPptx           │ 10. toStream               │ 18. toFlatXml                  │
│ 4. buildPptx             │ 11. toBase64               │                                │
│ 5. preload               │ 12. toArrayBuffer          │                                │
│ 6. fromCache             │ 13. toResponse             │                                │
│ 7. clearCache            │ 14. toExpress              │                                │
│                          │ 15. toFastify              │                                │
└──────────────────────────┴────────────────────────────┴────────────────────────────────┘
```

---

## 2. Granular Method Contracts

### 2.1 Static Methods (7 Functions)

#### `[1/18] PPTXTemplater.load(source, options = {})`
- **Purpose**: Polymorphic ingestion of presentations into the in-memory VFS.
- **Accepted Sources**: Local file paths (`.pptx`, `.xml`, `.ppt`, `.odp`), Node.js `Buffer`, `Uint8Array`, Node/Web `ReadableStream`, uncompressed directory path, remote HTTP/HTTPS URL, or Cloud configuration object (AWS S3, GCS, Azure).
- **Returns**: `Promise<PPTXTemplater>`

#### `[2/18] PPTXTemplater.createBlank(options = {})`
- **Purpose**: Synthesizes a valid, compliant ECMA-376 blank presentation without needing an external template file.
- **Options**: `aspectRatio` (`'16:9'`, `'4:3'`, `'16:10'`), `width`, `height`, `unit` (`'emu'`, `'px'`, `'pt'`, `'in'`, `'cm'`), `title`, `creator`.
- **Returns**: `Promise<PPTXTemplater>`

#### `[3/18] PPTXTemplater.extractPptx(pptxPath, outputDir, options = {})`
- **Purpose**: Unpacks a binary `.pptx` archive into an uncompressed directory tree of OpenXML files and folders.
- **Options**: `overwrite` (boolean), `formatXml` (boolean).
- **Returns**: `Promise<void>`

#### `[4/18] PPTXTemplater.buildPptx(sourceDir, outputPath, options = {})`
- **Purpose**: Compiles an uncompressed OpenXML directory tree back into a compressed `.pptx` archive.
- **Options**: `compressionLevel` (`0` to `9`).
- **Returns**: `Promise<void>`

#### `[5/18] PPTXTemplater.preload(source, options = {})`
- **Purpose**: Pre-parses and caches an OpenXML template in memory during application startup to eliminate disk reads in hot request paths.
- **Options**: `key` (cache identifier), `logLevel`.
- **Returns**: `Promise<Map<string, Object>>`

#### `[6/18] PPTXTemplater.fromCache(source)`
- **Purpose**: Sub-millisecond cloning of a preloaded template from the in-memory cache.
- **Returns**: `Promise<PPTXTemplater>`

#### `[7/18] PPTXTemplater.clearCache()`
- **Purpose**: Purges all preloaded templates from memory.
- **Returns**: `void`

---

### 2.2 Export Instance Methods (8 Functions)

#### `[8/18] ppt.saveToFile(destinationPath, options = {})`
- **Purpose**: Writes the presentation to disk as a complete binary `.pptx` file.
- **Returns**: `Promise<void>`

#### `[9/18] ppt.toBuffer(options = {})`
- **Purpose**: Serializes the presentation to a native Node.js `Buffer`. Ideal for AWS Lambda, S3 uploads (`s3.putObject`), or database storage.
- **Returns**: `Promise<Buffer>`

#### `[10/18] ppt.toStream(options = {})`
- **Purpose**: Generates a Node.js `Readable` stream emitting chunks of the presentation archive.
- **Options**: `chunkSize` (default: 65536).
- **Returns**: `Promise<import('stream').Readable>`

#### `[11/18] ppt.toBase64(options = {})`
- **Purpose**: Renders the presentation payload into a Base64 string.
- **Options**: `dataUri` (boolean — if true, prepends `data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,`).
- **Returns**: `Promise<string>`

#### `[12/18] ppt.toArrayBuffer(options = {})`
- **Purpose**: Emits a standard Web `ArrayBuffer`. Mandatory for Cloudflare Workers, Vercel Edge Functions, and modern fetch APIs.
- **Returns**: `Promise<ArrayBuffer>`

#### `[13/18] ppt.toResponse(options = {})`
- **Purpose**: Creates a Web Standard `Response` object with OpenXML MIME headers and `Content-Disposition`.
- **Options**: `filename` (default: `'presentation.pptx'`), `status` (default: 200), `headers`.
- **Returns**: `Promise<Response>`

#### `[14/18] ppt.toExpress(res, filename = 'presentation.pptx', options = {})`
- **Purpose**: Pipes the presentation directly to an Express.js HTTP response object, setting attachment headers and disposition.
- **Returns**: `Promise<void>`

#### `[15/18] ppt.toFastify(reply, filename = 'presentation.pptx', options = {})`
- **Purpose**: Delivers the presentation payload directly via Fastify's `reply` object.
- **Returns**: `Promise<void>`

---

### 2.3 Extraction & Conversion Instance Methods (3 Functions)

#### `[16/18] ppt.extractToZip(destinationPath, options = {})`
- **Purpose**: Exports the in-memory VFS as a raw ZIP container archive on disk.
- **Options**: `compressionLevel` (`0` to `9`).
- **Returns**: `Promise<void>`

#### `[17/18] ppt.extractToXml(destinationDir, options = {})`
- **Purpose**: Dumps the in-memory VFS parts into an uncompressed directory tree with optional XML indentation for git tracking and inspection.
- **Options**: `formatXml` (boolean).
- **Returns**: `Promise<void>`

#### `[18/18] ppt.toFlatXml(destinationPathOrOptions = {}, options = {})`
- **Purpose**: Converts the multi-part presentation into a single-file Microsoft Flat OPC XML string (`<pkg:package>`) or writes it to a file.
- **Options**: `pretty` (boolean).
- **Returns**: `Promise<string>` (or `Promise<void>` if destination path is specified).

---

## 3. Prohibited Deprecated Aliases

The following method names have been completely deprecated and MUST NOT be re-introduced:

| Deprecated / Non-Standard Method | Whitelisted Replacement |
| :--- | :--- |
| `ppt.toFile(path)` | `ppt.saveToFile(destinationPath)` |
| `ppt.save(path)` | `ppt.saveToFile(destinationPath)` |
| `ppt.write(path)` | `ppt.saveToFile(destinationPath)` |
| `ppt.exportZip(path)` | `ppt.extractToZip(destinationPath)` |
| `ppt.exportXml(path)` | `ppt.extractToXml(destinationDir)` |
| `ppt.writeToBuffer()` | `ppt.toBuffer()` |
| `ppt.getBuffer()` | `ppt.toBuffer()` |
| `ppt.exportStream()` | `ppt.toStream()` |
| `PPTXTemplater.fromBuffer(buf)` | `PPTXTemplater.load(buf)` |
| `PPTXTemplater.fromFile(path)` | `PPTXTemplater.load(path)` |
| `PPTXTemplater.fromStream(stream)`| `PPTXTemplater.load(stream)` |

Any PR, commit, or generated code that uses deprecated aliases will be flagged and rejected by the `architecture-reviewer` agent.
