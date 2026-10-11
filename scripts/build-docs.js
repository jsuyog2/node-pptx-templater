/**
 * @fileoverview Build script for generating the complete, production-ready
 * standalone API Reference page (docs/api-reference.html) for node-pptx-templater.
 */

const fs = require('fs');
const path = require('path');

const ROOT_DIR = path.resolve(__dirname, '..');
const PPTX_SOURCE_PATH = path.join(ROOT_DIR, 'src', 'core', 'PPTXTemplater.js');
const OUTPUT_HTML_PATH = path.join(ROOT_DIR, 'docs', 'api-reference.html');
const SITEMAP_PATH = path.join(ROOT_DIR, 'docs', 'sitemap.xml');
const ROBOTS_PATH = path.join(ROOT_DIR, 'docs', 'robots.txt');

// Read source code to extract JSDoc comments and signatures
const pptxSource = fs.readFileSync(PPTX_SOURCE_PATH, 'utf8');

// The 10 API Modules specified in task.json and requirements
const MODULES = [
  {
    id: 'lifecycle-io',
    title: 'Lifecycle & I/O',
    icon: `<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 9l3 3-3 3m5 0h3M5 20h14a2 2 0 002-2V6a2 2 0 00-2-2H5a2 2 0 00-2 2v12a2 2 0 002 2z"></path></svg>`,
    description: 'Presentation initialization, binary loading from file paths, Buffers, and streams, template caching, and packaging serializations.',
    methods: [
      {
        name: 'load',
        isStatic: true,
        isAsync: true,
        signature: 'PPTXTemplater.load(source, options = {})',
        returnType: 'Promise<PPTXTemplater>',
        description: 'Loads an existing PowerPoint template from a file path string, in-memory Node.js Buffer, or stream. Inflates the PKZip container, parses package relationships (<code>_rels/.rels</code>), validates <code>[Content_Types].xml</code>, and mounts the PresentationML virtual file system.',
        openXml: 'Reads ZIP archive entries: <code>[Content_Types].xml</code>, <code>_rels/.rels</code>, <code>ppt/presentation.xml</code>, <code>ppt/_rels/presentation.xml.rels</code>.',
        params: [
          { name: 'source', type: 'string | Buffer', required: true, defaultVal: '—', description: 'File path on disk or Buffer containing raw PPTX binary archive.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Configuration options for parser and logger.' },
          { name: 'options.logLevel', type: "'silent' | 'info' | 'debug'", required: false, defaultVal: "'warn'", description: 'Logging verbosity level during parsing and manipulations.' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'Resolves to an initialized, chainable PPTXTemplater instance ready for templating operations.' },
        example: `const { PPTXTemplater } = require('node-pptx-templater');

// Load from disk path
const ppt = await PPTXTemplater.load('./templates/quarterly-report.pptx', {
  logLevel: 'info'
});

// Or load directly from an in-memory buffer (e.g. AWS S3 / HTTP download)
const buffer = await fetchTemplateBuffer();
const inMemoryPpt = await PPTXTemplater.load(buffer);`
      },
      {
        name: 'create',
        isStatic: true,
        isAsync: true,
        signature: 'PPTXTemplater.create()',
        returnType: 'Promise<PPTXTemplater>',
        description: 'Creates a pristine, blank PowerPoint presentation entirely from scratch in memory. Synthesizes a valid default PresentationML package structure with 16:9 widescreen dimensions, standard master slides, slide layouts, and minimal theme definitions.',
        openXml: 'Synthesizes clean <code>[Content_Types].xml</code>, default slide master (<code>ppt/slideMasters/slideMaster1.xml</code>), layout mappings, and blank <code>ppt/presentation.xml</code>.',
        params: [],
        returns: { type: 'Promise<PPTXTemplater>', description: 'Resolves to a new PPTXTemplater instance initialized with a blank presentation deck.' },
        example: `const { PPTXTemplater } = require('node-pptx-templater');

// Create brand-new presentation from scratch
const ppt = await PPTXTemplater.create();
await ppt.addSlide();
ppt.useSlide(1);
await ppt.addShape('roundRect', {
  x: '10%', y: '10%', width: '80%', height: '20%',
  fill: '2563EB', text: 'Generated with node-pptx-templater'
});
await ppt.save('./output/new-deck.pptx');`
      },
      {
        name: 'fromPresentationXml',
        isStatic: true,
        isAsync: true,
        signature: 'PPTXTemplater.fromPresentationXml(options)',
        returnType: 'Promise<PPTXTemplater>',
        description: 'Instantiates a template engine instance from pre-extracted or virtual XML parts. Useful for microservice pipelines that store and mutate unpacked OpenXML files in distributed key-value caches or object storage.',
        openXml: 'Bypasses zip extraction; mounts raw PresentationML strings into in-memory <code>VirtualFileSystem</code>.',
        params: [
          { name: 'options', type: 'Object', required: true, defaultVal: '—', description: 'Virtual file system map or configuration containing raw OpenXML parts.' },
          { name: 'options.files', type: 'Map<string, string | Buffer>', required: true, defaultVal: '—', description: 'Key-value map of OpenXML part paths to file contents.' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'Resolves to an operational PPTXTemplater instance backed by the supplied virtual XML parts.' },
        example: `const { PPTXTemplater } = require('node-pptx-templater');

const ppt = await PPTXTemplater.fromPresentationXml({
  files: virtualXmlMap
});
ppt.useSlide(1).replaceText({ '{{status}}': 'Operational' });
const buffer = await ppt.toBuffer();`
      },
      {
        name: 'preload',
        isStatic: true,
        isAsync: true,
        signature: 'PPTXTemplater.preload(source)',
        returnType: 'Promise<VirtualFileSystem>',
        description: 'Pre-reads, inflates, and caches a PPTX template into a high-speed memory cache. Subsequent calls to <code>PPTXTemplater.fromCache()</code> clone this pre-parsed memory state directly, eliminating disk reads and zip decompression overhead.',
        openXml: 'Caches inflated OpenXML parts in memory; avoids repeating <code>JSZip.loadAsync</code> on heavy batch jobs.',
        params: [
          { name: 'source', type: 'string | Buffer', required: true, defaultVal: '—', description: 'Template file path or binary Buffer to inflate and cache.' }
        ],
        returns: { type: 'Promise<VirtualFileSystem>', description: 'Resolves to the preloaded VirtualFileSystem representation stored in the global cache.' },
        example: `const { PPTXTemplater } = require('node-pptx-templater');

// Warm up template cache at service startup
await PPTXTemplater.preload('./templates/monthly-statement.pptx');
console.log('Template cache primed for sub-millisecond cloning.');`
      },
      {
        name: 'fromCache',
        isStatic: true,
        isAsync: true,
        signature: 'PPTXTemplater.fromCache(source)',
        returnType: 'Promise<PPTXTemplater>',
        description: 'Instantiates a new PPTXTemplater by deep-cloning an existing preloaded template cache. Delivers up to 10x higher throughput for concurrent report generation in high-concurrency API servers and cloud functions.',
        openXml: 'Fast deep-copies the virtual OpenXML byte tree without reading from disk or re-parsing the zip container.',
        params: [
          { name: 'source', type: 'string | Buffer', required: true, defaultVal: '—', description: 'Cache key (file path string or Buffer identifier) previously passed to preload().' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'Resolves to an isolated, ready-to-mutate PPTXTemplater instance.' },
        example: `const { PPTXTemplater } = require('node-pptx-templater');

// In high-concurrency HTTP handler
app.post('/api/generate-deck', async (req, res) => {
  const ppt = await PPTXTemplater.fromCache('./templates/monthly-statement.pptx');
  ppt.useSlide(1).replaceText(req.body.replacements);
  await ppt.toExpress(res, 'statement.pptx');
});`
      },
      {
        name: 'clearCache',
        isStatic: true,
        isAsync: false,
        signature: 'PPTXTemplater.clearCache()',
        returnType: 'void',
        description: 'Clears all cached template packages from memory. Ideal for test teardown, memory reclamation, or when template files on disk have been hot-reloaded.',
        openXml: 'Flushes global in-memory maps holding pre-inflated OpenXML byte arrays.',
        params: [],
        returns: { type: 'void', description: 'Does not return a value.' },
        example: `const { PPTXTemplater } = require('node-pptx-templater');

// Evict cached templates during maintenance or reload
PPTXTemplater.clearCache();`
      },
      {
        name: 'saveToFile',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.saveToFile(outputPath, options = {})',
        returnType: 'Promise<void>',
        description: 'Serializes the entire in-memory PresentationML package, generates a standardized PKZip archive with DEFLATE compression, and writes the output file to disk. Creates any missing parent directories automatically.',
        openXml: 'Recompresses all modified XML parts (<code>ppt/slides/slide*.xml</code>, <code>[Content_Types].xml</code>, etc.) into compliant binary PPTX container.',
        params: [
          { name: 'outputPath', type: 'string', required: true, defaultVal: '—', description: 'Target destination file path where the generated .pptx will be saved.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Export configuration options.' },
          { name: 'options.compressionLevel', type: 'number', required: false, defaultVal: '6', description: 'Zip compression level from 1 (fastest) to 9 (maximum compression).' }
        ],
        returns: { type: 'Promise<void>', description: 'Resolves once the file stream has completely flushed and closed on disk.' },
        example: `const ppt = await PPTXTemplater.load('./template.pptx');
ppt.useSlide(1).replaceText({ '{{title}}': 'Q3 Results' });

await ppt.saveToFile('./output/q3-results.pptx');
console.log('Saved presentation successfully.');`
      },
      {
        name: 'save',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.save(outputPath, options = {})',
        returnType: 'Promise<void>',
        description: 'Convenience alias for <code>saveToFile(outputPath, options)</code>. Writes the serialized presentation package to disk.',
        openXml: 'Delegates to <code>saveToFile</code>; re-packages all DrawingML and PresentationML parts into a binary archive.',
        params: [
          { name: 'outputPath', type: 'string', required: true, defaultVal: '—', description: 'Destination file path on disk.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Zip compression and writing options.' }
        ],
        returns: { type: 'Promise<void>', description: 'Resolves once writing completes.' },
        example: `const ppt = await PPTXTemplater.load('./template.pptx');
ppt.useSlide(1).replaceText({ '{{author}}': 'Jane Doe' });
await ppt.save('./output/final.pptx');`
      },
      {
        name: 'toBuffer',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.toBuffer(options = {})',
        returnType: 'Promise<Buffer>',
        description: 'Serializes the entire PowerPoint presentation into an in-memory Node.js binary Buffer. Essential for serverless architectures (AWS Lambda, Vercel Functions, Google Cloud Run) where local disk writes are prohibited or ephemeral.',
        openXml: 'Compresses all OpenXML package parts directly into memory without touching the filesystem.',
        params: [
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Serialization configuration.' },
          { name: 'options.compression', type: "'DEFLATE' | 'STORE'", required: false, defaultVal: "'DEFLATE'", description: 'Zip compression algorithm.' }
        ],
        returns: { type: 'Promise<Buffer>', description: 'Resolves to a binary Node.js Buffer containing the complete, valid .pptx file.' },
        example: `const ppt = await PPTXTemplater.load(templateBuffer);
ppt.useSlide(1).replaceText({ '{{client}}': 'Acme Corp' });

const generatedBuffer = await ppt.toBuffer();

// Upload directly to cloud object storage (e.g. AWS S3)
await s3Client.send(new PutObjectCommand({
  Bucket: 'client-reports',
  Key: 'report-2026.pptx',
  Body: generatedBuffer,
  ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
}));`
      },
      {
        name: 'toStream',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.toStream(options = {})',
        returnType: 'Promise<ReadableStream>',
        description: 'Returns a Node.js Readable stream emitting binary chunks of the compressed PPTX archive. Enables streaming responses with low memory footprints when producing massive decks.',
        openXml: 'Streams zipped binary data directly from internal archive compression pipeline.',
        params: [
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Stream configuration options.' }
        ],
        returns: { type: 'Promise<ReadableStream>', description: 'Resolves to a Readable stream emitting binary chunks of the presentation.' },
        example: `const ppt = await PPTXTemplater.load('./large-template.pptx');
const stream = await ppt.toStream();

const writeStream = fs.createWriteStream('./output/streamed.pptx');
stream.pipe(writeStream);`
      },
      {
        name: 'toBase64',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.toBase64(options = { dataUri: false })',
        returnType: 'Promise<string>',
        description: 'Serializes the presentation archive and returns it as a Base64 string or full Data URI string (<code>data:application/vnd...;base64,...</code>). Ideal for JSON API responses, client downloads, or webhook payloads.',
        openXml: 'Converts in-memory zipped binary buffer into Base64 encoding.',
        params: [
          { name: 'options', type: 'Object', required: false, defaultVal: '{ dataUri: false }', description: 'Formatting options.' },
          { name: 'options.dataUri', type: 'boolean', required: false, defaultVal: 'false', description: 'When true, prepends MIME type data URI scheme.' }
        ],
        returns: { type: 'Promise<string>', description: 'Resolves to the base64-encoded presentation string.' },
        example: `const ppt = await PPTXTemplater.load('./template.pptx');
const base64Data = await ppt.toBase64({ dataUri: true });

// Send in JSON response
res.json({
  filename: 'report.pptx',
  dataUri: base64Data
});`
      },
      {
        name: 'toExpress',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.toExpress(res, filename, options = {})',
        returnType: 'Promise<void>',
        description: 'Streams the presentation directly to an Express.js response object. Automatically sets the correct OpenXML MIME type (<code>application/vnd.openxmlformats-officedocument.presentationml.presentation</code>), <code>Content-Disposition</code> attachment header with encoded filename, and <code>Content-Length</code>.',
        openXml: 'Direct HTTP response serialization with standardized OpenXML MIME headers.',
        params: [
          { name: 'res', type: 'Express.Response', required: true, defaultVal: '—', description: 'Active Express HTTP response object.' },
          { name: 'filename', type: 'string', required: false, defaultVal: "'presentation.pptx'", description: 'Filename suggested to the user client browser for download.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Additional streaming and header options.' }
        ],
        returns: { type: 'Promise<void>', description: 'Resolves when the response transmission completes.' },
        example: `app.get('/download-deck', async (req, res) => {
  const ppt = await PPTXTemplater.load('./template.pptx');
  ppt.useSlide(1).replaceText({ '{{title}}': 'Dynamic Presentation' });
  
  // Directly sends file download to client browser
  await ppt.toExpress(res, 'Executive-Summary-2026.pptx');
});`
      },
      {
        name: 'saveToFolder',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.saveToFolder(folderPath)',
        returnType: 'Promise<void>',
        description: 'Unpacks and writes all internal OpenXML files, media assets, and relationships as formatted files on disk. Extremely useful for Git-based version control of PowerPoint templates, diffing changes, and low-level debugging.',
        openXml: 'Writes unpacked folders: <code>_rels/</code>, <code>docProps/</code>, <code>ppt/slides/</code>, <code>ppt/charts/</code>, <code>ppt/media/</code> with pretty-printed XML.',
        params: [
          { name: 'folderPath', type: 'string', required: true, defaultVal: '—', description: 'Target destination directory on disk.' }
        ],
        returns: { type: 'Promise<void>', description: 'Resolves once all XML and binary parts are extracted and written.' },
        example: `const ppt = await PPTXTemplater.load('./complex-template.pptx');
await ppt.saveToFolder('./extracted-template');
console.log('Unpacked OpenXML structure to ./extracted-template');`
      },
      {
        name: 'buildPptx',
        isStatic: true,
        isAsync: true,
        signature: 'PPTXTemplater.buildPptx(folderPath, outputPath)',
        returnType: 'Promise<void>',
        description: 'Compiles an extracted OpenXML directory back into a valid, zipped PowerPoint presentation file. Guarantees proper zip compression and relationship integrity.',
        openXml: 'Scans directory, preserves binary media bytes, and zips back to compliant PresentationML package.',
        params: [
          { name: 'folderPath', type: 'string', required: true, defaultVal: '—', description: 'Source folder containing extracted OpenXML files.' },
          { name: 'outputPath', type: 'string', required: true, defaultVal: '—', description: 'Target destination .pptx file path.' }
        ],
        returns: { type: 'Promise<void>', description: 'Resolves once archive creation finishes.' },
        example: `const { PPTXTemplater } = require('node-pptx-templater');

// Repackage modified XML folder into binary PPTX
await PPTXTemplater.buildPptx('./extracted-template', './dist/rebuilt.pptx');`
      },
      {
        name: 'extractPptx',
        isStatic: true,
        isAsync: true,
        signature: 'PPTXTemplater.extractPptx(pptxPath, destinationFolder, options = {})',
        returnType: 'Promise<void>',
        description: 'Unpacks a binary .pptx file into a folder of formatted OpenXML documents for automated inspection, security auditing, or batch transformation.',
        openXml: 'Extracts and formats <code>[Content_Types].xml</code>, <code>ppt/presentation.xml</code>, slide XML, and relationships.',
        params: [
          { name: 'pptxPath', type: 'string', required: true, defaultVal: '—', description: 'Path to source .pptx file.' },
          { name: 'destinationFolder', type: 'string', required: true, defaultVal: '—', description: 'Destination folder for unpacked OpenXML assets.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Extraction options.' }
        ],
        returns: { type: 'Promise<void>', description: 'Resolves once extraction completes.' },
        example: `const { PPTXTemplater } = require('node-pptx-templater');

await PPTXTemplater.extractPptx('./input.pptx', './debug-unpacked');`
      }
    ]
  },
  {
    id: 'slides-management',
    title: 'Slides Management',
    icon: `<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M19 11H5m14 0a2 2 0 012 2v6a2 2 0 01-2 2H5a2 2 0 01-2-2v-6a2 2 0 012-2m14 0V9a2 2 0 00-2-2M5 11V9a2 2 0 012-2m0 0V5a2 2 0 012-2h6a2 2 0 012 2v2M7 7h10"></path></svg>`,
    description: 'Targeting, slide duplication, creation, deletion, reordering, dimensions, layouts, and cross-deck slide import/export.',
    methods: [
      {
        name: 'useSlide',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.useSlide(slideNumber)',
        returnType: 'this (chainable)',
        description: 'Directs subsequent slide-level operations (replacements, shape mutations, table updates, chart refreshes) to a specific 1-based slide index.',
        openXml: 'Sets internal pointer targeting <code>ppt/slides/slide{N}.xml</code>.',
        params: [
          { name: 'slideNumber', type: 'number', required: true, defaultVal: '—', description: '1-based slide index to target.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable) for fluent method chaining.' },
        example: `ppt.useSlide(2)
   .replaceText({ '{{agenda}}': 'Financial Overview' });`
      },
      {
        name: 'useAllSlides',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.useAllSlides()',
        returnType: 'this (chainable)',
        description: 'Switches the operation scope to target every slide currently present in the presentation deck. Operations like <code>replaceText</code> will execute across all slides.',
        openXml: 'Iterates through all <code><p:sldId></code> elements listed in <code>ppt/presentation.xml</code> <code><p:sldIdLst></code>.',
        params: [],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `// Global replacement across all slides in deck
ppt.useAllSlides()
   .replaceText({
     '{{year}}': '2026',
     '{{company}}': 'Acme Corp',
     '{{confidential}}': 'CONFIDENTIAL'
   });`
      },
      {
        name: 'addSlide',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.addSlide(options = {})',
        returnType: 'Promise<{ slideIndex: number, rId: string }>',
        description: 'Creates and appends a new blank slide to the end of the presentation. Automatically assigns a new relationship ID, links the slide layout part, and updates <code>[Content_Types].xml</code>.',
        openXml: 'Generates <code>ppt/slides/slide{N}.xml</code>, <code>ppt/slides/_rels/slide{N}.xml.rels</code>, appends <code><p:sldId></code> in <code>presentation.xml</code>, and registers in <code>[Content_Types].xml</code>.',
        params: [
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Slide creation options.' },
          { name: 'options.layoutNumber', type: 'number', required: false, defaultVal: '1', description: 'Target slide layout index to link as parent.' }
        ],
        returns: { type: 'Promise<{ slideIndex: number, rId: string }>', description: 'Object containing 1-based slide index and assigned relationship ID.' },
        example: `const newSlide = await ppt.addSlide({ layoutNumber: 1 });
ppt.useSlide(newSlide.slideIndex);
await ppt.addShape('rect', { text: 'Title on New Slide' });`
      },
      {
        name: 'insertSlide',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.insertSlide(slideNumber, options = {})',
        returnType: 'Promise<{ slideIndex: number, rId: string }>',
        description: 'Inserts a new slide at a specific 1-based position within the presentation sequence, shifting subsequent slide entries accordingly.',
        openXml: 'Splices a new <code><p:sldId></code> into <code><p:sldIdLst></code> in <code>ppt/presentation.xml</code> at the requested position.',
        params: [
          { name: 'slideNumber', type: 'number', required: true, defaultVal: '—', description: '1-based index where the new slide should be inserted.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Slide options.' }
        ],
        returns: { type: 'Promise<{ slideIndex: number, rId: string }>', description: 'Metadata for inserted slide.' },
        example: `// Insert slide at position 2
await ppt.insertSlide(2);
ppt.useSlide(2).replaceText({ '{{section}}': 'Executive Summary' });`
      },
      {
        name: 'duplicateSlide',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.duplicateSlide(sourceSlideIndex, insertAfterIndex)',
        returnType: 'Promise<{ slideIndex: number, rId: string }>',
        description: 'Deep-clones an existing slide including all shapes, charts, tables, media, relationships, and layout bindings. Cloned parts receive newly generated unique relationship IDs to guarantee non-corrupted PresentationML structures.',
        openXml: 'Clones <code>ppt/slides/slide{N}.xml</code>, duplicates <code>slide{N}.xml.rels</code>, deep-copies linked charts/drawings, updates <code>[Content_Types].xml</code> and <code>presentation.xml</code>.',
        params: [
          { name: 'sourceSlideIndex', type: 'number', required: true, defaultVal: '—', description: '1-based slide index to duplicate.' },
          { name: 'insertAfterIndex', type: 'number', required: false, defaultVal: 'sourceSlideIndex', description: '1-based slide index after which the duplicate will be placed.' }
        ],
        returns: { type: 'Promise<{ slideIndex: number, rId: string }>', description: 'Metadata with newly created slide index.' },
        example: `// Duplicate slide 2 for each regional sales team
const teams = ['EMEA', 'APAC', 'Americas'];
for (const team of teams) {
  const clone = await ppt.duplicateSlide(2);
  ppt.useSlide(clone.slideIndex)
     .replaceText({ '{{team_name}}': team });
}`
      },
      {
        name: 'removeSlide',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.removeSlide(slideNumber)',
        returnType: 'Promise<this>',
        description: 'Removes a slide from the presentation by 1-based index. Deletes the corresponding <code><p:sldId></code> in <code>presentation.xml</code>, removes the relationship in <code>presentation.xml.rels</code>, and cleans up orphaned slide XML and relationship parts.',
        openXml: 'Deletes <code><p:sldId></code>, unlinks relationship in <code>presentation.xml.rels</code>, removes <code>slide{N}.xml</code> and <code>slide{N}.xml.rels</code>.',
        params: [
          { name: 'slideNumber', type: 'number', required: true, defaultVal: '—', description: '1-based slide index to delete.' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'this (chainable).' },
        example: `// Remove temporary instruction slide
await ppt.removeSlide(1);`
      },
      {
        name: 'getSlides',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.getSlides()',
        returnType: 'Array<{ slideNumber: number, rId: string, partName: string }>',
        description: 'Returns metadata for all slides currently configured in the presentation deck in their visual presentation order.',
        openXml: 'Reads <code><p:sldIdLst></code> in <code>ppt/presentation.xml</code> and resolves target slide part names.',
        params: [],
        returns: { type: 'Array<Object>', description: 'Array of slide descriptors with 1-based slideNumber, rId, and partName.' },
        example: `const slides = ppt.getSlides();
console.log(\`Presentation has \${slides.length} slides:\`, slides);`
      },
      {
        name: 'getSlideDimensions',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.getSlideDimensions()',
        returnType: '{ width: number, height: number, widthEmu: number, heightEmu: number, type: string }',
        description: 'Inspects and returns the physical slide dimensions and aspect ratio configured for the presentation in inches, points, pixels, and EMUs.',
        openXml: 'Parses <code><p:sldSz cx="..." cy="..." type="..."/></code> in <code>ppt/presentation.xml</code>.',
        params: [],
        returns: { type: 'Object', description: 'Dimensions object with width, height, widthEmu, heightEmu, and aspect ratio type (e.g. "screen16x9").' },
        example: `const dims = ppt.getSlideDimensions();
console.log(\`Slide width: \${dims.width}in, height: \${dims.height}in (type: \${dims.type})\`);`
      },
      {
        name: 'getSlideLayout',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.getSlideLayout(slideNumber)',
        returnType: '{ name: string, type: string, master: string, rId: string }',
        description: 'Resolves and returns the layout metadata associated with a targeted slide, including layout name, type, and parent slide master reference.',
        openXml: 'Follows relationship in <code>slide{N}.xml.rels</code> to <code>ppt/slideLayouts/slideLayout{M}.xml</code>.',
        params: [
          { name: 'slideNumber', type: 'number', required: false, defaultVal: 'active slide', description: '1-based slide index.' }
        ],
        returns: { type: 'Object', description: 'Layout metadata object.' },
        example: `const layout = ppt.getSlideLayout(1);
console.log(\`Slide 1 uses layout: \${layout.name} (master: \${layout.master})\`);`
      },
      {
        name: 'importSlideFrom',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.importSlideFrom(sourcePptInstance, sourceSlideNumber)',
        returnType: 'Promise<{ slideIndex: number, rId: string }>',
        description: 'Imports a complete slide from another independent PPTXTemplater instance. Resolves, clones, and deep-remaps all dependencies (slide layouts, media files, embedded charts, hyperlinks) into the target deck while preventing relationship collisions.',
        openXml: 'Transfers external slide XML, re-indexes relationships in <code>slide{N}.xml.rels</code>, syncs <code>[Content_Types].xml</code>, and transfers media parts.',
        params: [
          { name: 'sourcePptInstance', type: 'PPTXTemplater', required: true, defaultVal: '—', description: 'Source PPTXTemplater presentation instance.' },
          { name: 'sourceSlideNumber', type: 'number', required: true, defaultVal: '—', description: '1-based slide index in source presentation to import.' }
        ],
        returns: { type: 'Promise<Object>', description: 'Imported slide descriptor with new slideIndex and rId.' },
        example: `const sourceDeck = await PPTXTemplater.load('./library/appendix.pptx');
const mainDeck = await PPTXTemplater.load('./deck.pptx');

// Import appendix slide 3 into main deck
await mainDeck.importSlideFrom(sourceDeck, 3);
await mainDeck.save('./deck-with-appendix.pptx');`
      },
      {
        name: 'exportSlides',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.exportSlides(...slideNumbers)',
        returnType: 'Promise<PPTXTemplater>',
        description: 'Extracts a subset of slides from the current deck and returns a new, fully self-contained PPTXTemplater instance containing only those slides, preserving all required masters, themes, and media.',
        openXml: 'Constructs a new PresentationML package containing selected slides and pruned masters/layouts.',
        params: [
          { name: 'slideNumbers', type: '...number[]', required: true, defaultVal: '—', description: 'One or more 1-based slide indices to export.' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'New independent PPTXTemplater instance containing extracted slides.' },
        example: `// Create mini deck with slides 1, 4, and 7
const miniDeck = await ppt.exportSlides(1, 4, 7);
await miniDeck.save('./output/mini-deck.pptx');`
      }
    ]
  },
  {
    id: 'text-placeholders',
    title: 'Text & Placeholders',
    icon: `<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M11 5H6a2 2 0 00-2 2v11a2 2 0 002 2h11a2 2 0 002-2v-5m-1.414-9.414a2 2 0 112.828 2.828L11.828 15H9v-2.828l8.586-8.586z"></path></svg>`,
    description: 'Placeholder substitution across fragmented DrawingML text runs, typography styling, searching, and bullet list synthesis.',
    methods: [
      {
        name: 'replaceText',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.replaceText(replacements)',
        returnType: 'this (chainable)',
        description: 'Replaces template placeholders across targeted slides. Analyzes DrawingML AST paragraphs (<code>&lt;a:p&gt;</code>) and unifies fragmented text runs (<code>&lt;a:r&gt;</code> / <code>&lt;a:t&gt;</code>) split across spellchecks or editing sessions before substituting values.',
        openXml: 'Parses <code>&lt;a:p&gt;</code> and <code>&lt;a:r&gt;</code>, joins adjacent <code>&lt;a:t&gt;</code> text nodes, preserves original <code>&lt;a:rPr&gt;</code> formatting, and updates slide XML.',
        params: [
          { name: 'replacements', type: 'Object.<string, string | number>', required: true, defaultVal: '—', description: 'Key-value map of placeholders (e.g. <code>{\'{{title}}\': \'Q3 Report\'}</code>).' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).replaceText({
  '{{title}}': 'Executive Quarterly Business Review',
  '{{date}}': 'October 11, 2026',
  '{{arr}}': '$124.8M',
  '{{author}}': 'Engineering & Operations Team'
});`
      },
      {
        name: 'replaceMultiple',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.replaceMultiple(replacements, options = {})',
        returnType: 'this (chainable)',
        description: 'Batch-replaces multiple placeholders across all targeted slides in a single optimized pass. Automatically wraps raw keys in <code>{{...}}</code> if delimiters are omitted.',
        openXml: 'Traverses slide AST, normalizes placeholder tokens, and mutates <code>&lt;a:t&gt;</code> elements.',
        params: [
          { name: 'replacements', type: 'Object.<string, any>', required: true, defaultVal: '—', description: 'Map of keys to replacement strings or numbers.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Replacement options.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useAllSlides().replaceMultiple({
  company: 'Acme Technologies Global Inc.',
  quarter: 'Q3 2026',
  confidentiality: 'CONFIDENTIAL'
});`
      },
      {
        name: 'replaceTextByTag',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.replaceTextByTag(tag, value, options = {})',
        returnType: 'this (chainable)',
        description: 'Replaces a single placeholder tag across targeted slides while rigorously preserving original run formatting properties (font typeface, size, color, bold, italics).',
        openXml: 'Modifies targeted <code>&lt;a:t&gt;</code> run nodes under <code>&lt;a:p&gt;</code> without altering adjacent <code>&lt;a:rPr&gt;</code> tags.',
        params: [
          { name: 'tag', type: 'string', required: true, defaultVal: '—', description: 'Placeholder tag name (e.g. <code>\'{{name}}\'</code> or <code>\'name\'</code>).' },
          { name: 'value', type: 'string | number', required: true, defaultVal: '—', description: 'Replacement string or numeric value.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Options.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).replaceTextByTag('revenue', '$142.5M');`
      },
      {
        name: 'updateText',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.updateText(tag, valueOrListConfig)',
        returnType: 'this (chainable)',
        description: 'Updates text content or synthesizes hierarchical bullet/numbered lists inside a shape or text frame by shape name, shape ID, or placeholder tag.',
        openXml: 'Replaces paragraph text or synthesizes <code>&lt;a:pPr&gt;</code> bullet properties (<code>&lt;a:buChar&gt;</code>, <code>&lt;a:buAutoNum&gt;</code>, <code>@_lvl</code>, <code>@_marL</code>).',
        params: [
          { name: 'tag', type: 'string', required: true, defaultVal: '—', description: 'Shape name, shape ID, or placeholder tag.' },
          { name: 'valueOrListConfig', type: 'string | Object | Array', required: true, defaultVal: '—', description: 'Plain text string or structured list configuration object.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `// Generate formatted multi-level bulleted list
ppt.useSlide(2).updateText('key_takeaways', {
  list: [
    'Net Revenue Retention expanded to 118.4%',
    { text: 'Gross margin improved 3.4% YoY', children: ['Cloud optimization initiatives', 'Vendor renegotiations'] },
    'Zero critical infrastructure incidents in Q3'
  ],
  style: { bulletChar: '•', color: '1E293B', fontSize: 14 }
});`
      },
      {
        name: 'findText',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.findText(text)',
        returnType: 'Array<{ slideIndex: number, shapeName: string, text: string }>',
        description: 'Searches for text occurrences across targeted slides. Scans all shapes, text boxes, and tables, returning match locations with slide index and shape identifiers.',
        openXml: 'Recursively inspects <code>&lt;p:spTree&gt;</code>, <code>&lt;p:sp&gt;</code>, <code>&lt;p:graphicFrame&gt;</code>, and <code>&lt;a:t&gt;</code> nodes.',
        params: [
          { name: 'text', type: 'string', required: true, defaultVal: '—', description: 'Substring to search for.' }
        ],
        returns: { type: 'Array<Object>', description: 'List of matches with slide numbers and shape names.' },
        example: `const matches = ppt.findText('ARR');
matches.forEach(m => console.log(\`Found "\${m.text}" on slide \${m.slideIndex} in shape "\${m.shapeName}"\`));`
      },
      {
        name: 'appendTextRun',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.appendTextRun(shapeId, text, options = {})',
        returnType: 'this (chainable)',
        description: 'Appends a new styled text run (<code>&lt;a:r&gt;</code>) to the end of the text frame of an existing shape. Supports custom font family, size, solid fill color, bold, italics, underline, and strike.',
        openXml: 'Appends <code>&lt;a:r&gt;</code> with <code>&lt;a:rPr&gt;</code> (<code>@_sz</code>, <code>@_b</code>, <code>&lt;a:solidFill&gt;</code>) to the last <code>&lt;a:p&gt;</code> in <code>&lt;p:txBody&gt;</code>.',
        params: [
          { name: 'shapeId', type: 'string', required: true, defaultVal: '—', description: 'Shape ID or shape name.' },
          { name: 'text', type: 'string', required: true, defaultVal: '—', description: 'Text string to append.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Styling options (fontSize, bold, color, fontFamily).' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).appendTextRun('subtitle_box', ' (Updated October 2026)', {
  bold: true,
  color: '2563EB',
  fontSize: 12
});`
      },
      {
        name: 'prependTextRun',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.prependTextRun(shapeId, text, options = {})',
        returnType: 'this (chainable)',
        description: 'Inserts a new styled text run at the very beginning of a shape\'s text frame, before all existing paragraphs and runs.',
        openXml: 'Prepends <code>&lt;a:r&gt;</code> to the first <code>&lt;a:p&gt;</code> under <code>&lt;p:txBody&gt;</code>.',
        params: [
          { name: 'shapeId', type: 'string', required: true, defaultVal: '—', description: 'Target shape identifier.' },
          { name: 'text', type: 'string', required: true, defaultVal: '—', description: 'Text string to prepend.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Styling configuration options.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).prependTextRun('title_box', '[FINAL] ', {
  bold: true,
  color: '10B981'
});`
      },
      {
        name: 'extractPlaceholders',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.extractPlaceholders(pattern)',
        returnType: 'string[]',
        description: 'Scans all text runs across targeted slides and extracts a unique list of all placeholder tokens matching the pattern (defaults to <code>/\\{\\{([^\\{\\}]+)\\}\\}/g</code>). Useful for validating payload completeness before template rendering.',
        openXml: 'Parses all <code>&lt;a:t&gt;</code> elements across active slide XML parts.',
        params: [
          { name: 'pattern', type: 'RegExp', required: false, defaultVal: '/\\{\\{([^\\{\\}]+)\\}\\}/g', description: 'Custom regular expression for placeholder detection.' }
        ],
        returns: { type: 'string[]', description: 'Array of unique placeholder tokens found.' },
        example: `const requiredTags = ppt.useAllSlides().extractPlaceholders();
console.log('Template requires placeholders:', requiredTags);
// ['{{title}}', '{{author}}', '{{company}}']`
      },
      {
        name: 'getList',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.getList(tag)',
        returnType: 'Array<{ text: string, level: number }>',
        description: 'Extracts existing bulleted and numbered items from a shape or text box, returning each line with its hierarchical indentation level.',
        openXml: 'Reads <code>&lt;a:p&gt;</code> elements and parses <code>&lt;a:pPr @_lvl="..."&gt;</code>.',
        params: [
          { name: 'tag', type: 'string', required: true, defaultVal: '—', description: 'Shape name, shape ID, or placeholder tag.' }
        ],
        returns: { type: 'Array<Object>', description: 'Array of items with text and 0-based indentation level.' },
        example: `const items = ppt.useSlide(2).getList('action_items');
items.forEach(item => console.log(\`Level \${item.level}: \${item.text}\`));`
      },
      {
        name: 'validateList',
        isStatic: true,
        isAsync: false,
        signature: 'PPTXTemplater.validateList(listConfig)',
        returnType: '{ valid: boolean, errors: string[] }',
        description: 'Validates a structured list configuration before injecting it into slide XML. Checks list depth, text integrity, and bullet formatting parameters.',
        openXml: 'Schema verification utility for DrawingML paragraph properties.',
        params: [
          { name: 'listConfig', type: 'Object', required: true, defaultVal: '—', description: 'List configuration object with list array and optional style.' }
        ],
        returns: { type: '{ valid: boolean, errors: string[] }', description: 'Validation outcome descriptor.' },
        example: `const result = PPTXTemplater.validateList({ list: ['Item 1', 'Item 2'] });
if (!result.valid) console.error('List errors:', result.errors);`
      }
    ]
  },
  {
    id: 'tables-cell-shapes',
    title: 'Tables & Cell Shapes',
    icon: `<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 10h18M3 14h18m-9-4v8m-7 0h14a2 2 0 002-2V6a2 2 0 00-2-2H5a2 2 0 00-2 2v8a2 2 0 002 2z"></path></svg>`,
    description: 'Grid manipulation, data matrices, row additions, horizontal and vertical cell merges, autofit, and automatic slide overflow pagination.',
    methods: [
      {
        name: 'addTable',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.addTable(tableIdOrOptions, options = {})',
        returnType: 'Promise<string>',
        description: 'Creates a brand-new table on the active slide. Generates DrawingML table structure (<code>&lt;a:tbl&gt;</code>) with customizable row count, column count, column widths, coordinates, styles, and initial data matrix.',
        openXml: 'Appends <code>&lt;p:graphicFrame&gt;</code> containing <code>&lt;a:tbl&gt;</code>, <code>&lt;a:tblPr&gt;</code>, <code>&lt;a:tblGrid&gt;</code>, and <code>&lt;a:tr&gt;</code> rows.',
        params: [
          { name: 'tableIdOrOptions', type: 'string | Object', required: true, defaultVal: '—', description: 'Table identifier name or full configuration options.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Table configuration (rows, cols, x, y, width, height, data).' }
        ],
        returns: { type: 'Promise<string>', description: 'Resolves to the newly created table shape ID.' },
        example: `ppt.useSlide(1);
await ppt.addTable('projectTable', {
  x: '5%', y: '20%', width: '90%', height: '50%',
  rows: 4, cols: 3,
  colWidths: [2000000, 4000000, 2500000],
  data: [
    ['Workstream', 'Initiative Lead', 'Status'],
    ['Core Cloud', 'Sarah Jenkins', 'Completed'],
    ['AI Pipeline', 'Elena Rostova', 'In Progress'],
    ['Security', 'Alicia Gomez', 'On Track']
  ]
});`
      },
      {
        name: 'updateTable',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.updateTable(tableId, dataMatrix)',
        returnType: 'Promise<this>',
        description: 'Replaces table cell values with a 2D array matrix of strings or numbers. Matches existing rows and columns while preserving all existing cell formatting (background fills, fonts, borders, alignments).',
        openXml: 'Iterates existing <code>&lt;a:tr&gt;</code> and <code>&lt;a:tc&gt;</code> elements, updating text run values in <code>&lt;a:t&gt;</code>.',
        params: [
          { name: 'tableId', type: 'string', required: true, defaultVal: '—', description: 'Table name or shape ID.' },
          { name: 'dataMatrix', type: 'Array<Array<any>>', required: true, defaultVal: '—', description: '2D matrix of row values.' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'this (chainable).' },
        example: `await ppt.useSlide(1).updateTable('summary_table', [
  ['Category', 'Budget', 'Actual'],
  ['Engineering', '$4.2M', '$3.9M'],
  ['Operations', '$1.8M', '$1.7M']
]);`
      },
      {
        name: 'getTableRows',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.getTableRows(tableId, options = {})',
        returnType: 'Promise<Array<Object | Array<string>>>',
        description: 'Extracts data from an existing table on the active slide. Can return rows as structured JSON objects (using header row values as object keys) or as raw 2D string matrices.',
        openXml: 'Parses <code>&lt;a:tbl&gt;</code> cells and extracts text from <code>&lt;a:t&gt;</code> elements.',
        params: [
          { name: 'tableId', type: 'string', required: true, defaultVal: '—', description: 'Table name or shape ID.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Extraction options.' },
          { name: 'options.asObjects', type: 'boolean', required: false, defaultVal: 'true', description: 'When true, formats rows as JSON objects using table headers.' }
        ],
        returns: { type: 'Promise<Array>', description: 'Array of extracted table row records.' },
        example: `const rows = await ppt.useSlide(1).getTableRows('kpi_table', { asObjects: true });
console.log('Extracted table data:', rows);
// [{ Metric: 'ARR', Value: '$124.8M' }, { Metric: 'NRR', Value: '118.4%' }]`
      },
      {
        name: 'addTableRow',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.addTableRow(tableId, rowData, options = {})',
        returnType: 'Promise<this>',
        description: 'Appends a new row to an existing table. Supports nested arrays for multi-row merging (rowspan), automatic cell styling, custom background fills, and borders.',
        openXml: 'Synthesizes and appends <code>&lt;a:tr&gt;</code> containing <code>&lt;a:tc&gt;</code> cells into <code>&lt;a:tbl&gt;</code>.',
        params: [
          { name: 'tableId', type: 'string', required: true, defaultVal: '—', description: 'Table name or shape ID.' },
          { name: 'rowData', type: 'Array<any>', required: true, defaultVal: '—', description: 'Array of cell values for the new row.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Options (mergeStrategy, cellStyles).' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'this (chainable).' },
        example: `await ppt.useSlide(1).addTableRow('projectTable', [
  'Infrastructure', 'Multi-Region Failover', 'Completed'
]);`
      },
      {
        name: 'mergeCells',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.mergeCells(tableId, startRow, startCol, endRow, endCol)',
        returnType: 'Promise<this>',
        description: 'Merges a rectangular block of cells across rows and/or columns. Applies compliant OpenXML rowspan and colspan attributes (<code>rowSpan</code>, <code>gridSpan</code>, <code>hMerge</code>, <code>vMerge</code>) without corrupting PowerPoint.',
        openXml: 'Sets <code>@_rowSpan</code>, <code>@_gridSpan</code>, <code>@_vMerge="1"</code>, and <code>@_hMerge="1"</code> on <code>&lt;a:tcPr&gt;</code>.',
        params: [
          { name: 'tableId', type: 'string', required: true, defaultVal: '—', description: 'Table name or shape ID.' },
          { name: 'startRow', type: 'number', required: true, defaultVal: '—', description: '0-based start row index.' },
          { name: 'startCol', type: 'number', required: true, defaultVal: '—', description: '0-based start column index.' },
          { name: 'endRow', type: 'number', required: true, defaultVal: '—', description: '0-based end row index.' },
          { name: 'endCol', type: 'number', required: true, defaultVal: '—', description: '0-based end column index.' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'this (chainable).' },
        example: `// Merge header cells across columns 1 to 3
await ppt.useSlide(1).mergeCells('financials_table', 0, 1, 0, 3);`
      },
      {
        name: 'cloneMergedRegion',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.cloneMergedRegion(tableId, options)',
        returnType: 'Promise<this>',
        description: 'Replicates an existing merged cell topology across multiple contiguous rows or columns.',
        openXml: 'Copies <code>&lt;a:tcPr&gt;</code> merge flags across corresponding row indices in <code>&lt;a:tbl&gt;</code>.',
        params: [
          { name: 'tableId', type: 'string', required: true, defaultVal: '—', description: 'Table identifier.' },
          { name: 'options', type: 'Object', required: true, defaultVal: '—', description: 'Source and destination range configuration.' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'this (chainable).' },
        example: `await ppt.useSlide(1).cloneMergedRegion('salesTable', {
  sourceRow: 1, targetRows: [2, 3, 4]
});`
      },
      {
        name: 'autoFitTable',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.autoFitTable(tableId, options = {})',
        returnType: 'Promise<this>',
        description: 'Automatically calculates and distributes column widths in <code>&lt;a:tblGrid&gt;</code> based on text length and typography metrics, preventing awkward text wrapping.',
        openXml: 'Recalculates <code>&lt;a:gridCol @_w="..."&gt;</code> values in <code>&lt;a:tblGrid&gt;</code>.',
        params: [
          { name: 'tableId', type: 'string', required: true, defaultVal: '—', description: 'Table name or shape ID.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Autofit options (minColWidth, maxColWidth).' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'this (chainable).' },
        example: `await ppt.useSlide(1).autoFitTable('data_table', { minColWidth: '1.2in' });`
      },
      {
        name: 'splitTableOnOverflow',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.splitTableOnOverflow(tableId, options = {})',
        returnType: 'Promise<{ splitCount: number, slidesCreated: number }>',
        description: 'Automatically splits large tables across consecutive duplicate slides when table height exceeds a slide threshold. Preserves repeated headers across all generated slides and keeps summary/totals rows pinned to the final slide.',
        openXml: 'Calculates cumulative row heights in EMUs, duplicates slide XML, chunks <code>&lt;a:tr&gt;</code> rows, and syncs slide relationships.',
        params: [
          { name: 'tableId', type: 'string', required: true, defaultVal: '—', description: 'Table name or shape ID.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Pagination options.' },
          { name: 'options.bottomMargin', type: 'number | string', required: false, defaultVal: '0.5 in', description: 'Bottom margin threshold for triggering split.' },
          { name: 'options.repeatHeaders', type: 'number', required: false, defaultVal: '1', description: 'Number of header rows to repeat on each split slide.' },
          { name: 'options.summaryRows', type: 'number', required: false, defaultVal: '0', description: 'Number of summary rows to preserve on the final slide.' }
        ],
        returns: { type: 'Promise<Object>', description: 'Pagination outcome with splitCount and slidesCreated.' },
        example: `// Split 100-row transaction ledger across slides
const result = await ppt.useSlide(3).splitTableOnOverflow('ledger_table', {
  bottomMargin: '0.75in',
  repeatHeaders: 1,
  summaryRows: 1
});
console.log(\`Created \${result.slidesCreated} additional slides.\`);`
      },
      {
        name: 'getCellBounds',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.getCellBounds(tableId, rowIndex, colIndex)',
        returnType: '{ x: number, y: number, width: number, height: number }',
        description: 'Computes the exact absolute spatial bounding box (x, y, width, height) of a table cell in EMUs relative to the slide canvas. Critical for placing floating badge shapes or icons precisely inside cells.',
        openXml: 'Derives coordinates from table position (<code>&lt;a:xfrm&gt;</code>), <code>&lt;a:tblGrid&gt;</code> column widths, and row heights.',
        params: [
          { name: 'tableId', type: 'string', required: true, defaultVal: '—', description: 'Table name or shape ID.' },
          { name: 'rowIndex', type: 'number', required: true, defaultVal: '—', description: '0-based row index.' },
          { name: 'colIndex', type: 'number', required: true, defaultVal: '—', description: '0-based column index.' }
        ],
        returns: { type: 'Object', description: 'Spatial bounds in EMUs.' },
        example: `const bounds = ppt.useSlide(1).getCellBounds('status_table', 2, 3);
console.log(\`Cell spatial coordinates:\`, bounds);`
      },
      {
        name: 'addCellShape',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.addCellShape(tableId, rowIndex, colIndex, options)',
        returnType: 'Promise<string>',
        description: 'Embeds a geometric badge or indicator shape (pill, status circle, checkmark) directly inside a table cell, positioned dynamically according to cell bounds.',
        openXml: 'Inserts a positioned <code>&lt;p:sp&gt;</code> into <code>&lt;p:spTree&gt;</code> aligned with cell spatial boundaries.',
        params: [
          { name: 'tableId', type: 'string', required: true, defaultVal: '—', description: 'Table name or shape ID.' },
          { name: 'rowIndex', type: 'number', required: true, defaultVal: '—', description: '0-based row index.' },
          { name: 'colIndex', type: 'number', required: true, defaultVal: '—', description: '0-based column index.' },
          { name: 'options', type: 'Object', required: true, defaultVal: '—', description: 'Shape options (shapeType, fill, text, padding).' }
        ],
        returns: { type: 'Promise<string>', description: 'Assigned shape ID.' },
        example: `// Add green status pill inside cell (row 1, col 2)
await ppt.useSlide(1).addCellShape('status_table', 1, 2, {
  shapeType: 'roundRect',
  fill: '10B981',
  text: 'ACTIVE',
  textColor: 'FFFFFF'
});`
      },
      {
        name: 'updateCellShape',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.updateCellShape(tableId, rowIndex, colIndex, shapeIndex, options)',
        returnType: 'Promise<this>',
        description: 'Updates properties (fill color, text label, border) of an indicator shape previously placed inside a table cell.',
        openXml: 'Modifies <code>&lt;p:sp&gt;</code> properties within <code>&lt;p:spTree&gt;</code>.',
        params: [
          { name: 'tableId', type: 'string', required: true, defaultVal: '—', description: 'Table identifier.' },
          { name: 'rowIndex', type: 'number', required: true, defaultVal: '—', description: '0-based row index.' },
          { name: 'colIndex', type: 'number', required: true, defaultVal: '—', description: '0-based column index.' },
          { name: 'shapeIndex', type: 'number', required: true, defaultVal: '—', description: '0-based index of shape in cell.' },
          { name: 'options', type: 'Object', required: true, defaultVal: '—', description: 'Updated properties.' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'this (chainable).' },
        example: `await ppt.useSlide(1).updateCellShape('status_table', 1, 2, 0, {
  fill: 'EF4444',
  text: 'OFFLINE'
});`
      }
    ]
  },
  {
    id: 'charts-series-sync',
    title: 'Charts & Series Sync',
    icon: `<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 19v-6a2 2 0 00-2-2H5a2 2 0 00-2 2v6a2 2 0 002 2h2a2 2 0 002-2zm0 0V9a2 2 0 012-2h2a2 2 0 012 2v10m-6 0a2 2 0 002 2h2a2 2 0 002-2m0 0V5a2 2 0 012-2h2a2 2 0 012 2v14a2 2 0 01-2 2h-2a2 2 0 01-2-2z"></path></svg>`,
    description: 'Multi-series chart updates, category synchronization, data labels, and automated embedded Excel (.xlsx) worksheet cache syncing.',
    methods: [
      {
        name: 'updateChart',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.updateChart(chartId, data)',
        returnType: 'this (chainable)',
        description: 'Updates chart categories, series names, and numerical values on the active slide. Simultaneously synchronizes both the DrawingML ChartSpace XML (<code>ppt/charts/chart{N}.xml</code>) and the embedded Excel worksheet data cache (<code>ppt/embeddings/Microsoft_Excel_Worksheet{N}.xlsx</code>), preventing PowerPoint "Repair Workbook" warnings.',
        openXml: 'Mutates <code>&lt;c:chartSpace&gt;</code> (<code>&lt;c:ser&gt;</code>, <code>&lt;c:cat&gt;</code>, <code>&lt;c:val&gt;</code>, <code>&lt;c:numCache&gt;</code>) and updates embedded Excel sheet data.',
        params: [
          { name: 'chartId', type: 'string', required: true, defaultVal: '—', description: 'Chart title, shape name, or chart relationship ID.' },
          { name: 'data', type: 'Object', required: true, defaultVal: '—', description: 'New chart data payload.' },
          { name: 'data.categories', type: 'string[]', required: true, defaultVal: '—', description: 'Array of category labels (X-axis).' },
          { name: 'data.series', type: 'Array<{ name: string, values: number[] }>', required: true, defaultVal: '—', description: 'Array of data series with name and numerical values.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(2).updateChart('RevenueTrend', {
  categories: ['Q1 2026', 'Q2 2026', 'Q3 2026', 'Q4 2026'],
  series: [
    { name: 'Enterprise Cloud', values: [18.4, 22.1, 26.8, 31.5] },
    { name: 'AI Solutions', values: [8.2, 12.6, 17.4, 23.0] }
  ]
});`
      },
      {
        name: 'updateChartCategories',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.updateChartCategories(chartId, categories)',
        returnType: 'this (chainable)',
        description: 'Updates category labels along the chart axis without altering series data values. Updates both chart XML string cache and Excel worksheet row/column headers.',
        openXml: 'Mutates <code>&lt;c:cat&gt;</code> <code>&lt;c:strCache&gt;</code> in <code>chart{N}.xml</code> and updates embedded Excel cells.',
        params: [
          { name: 'chartId', type: 'string', required: true, defaultVal: '—', description: 'Chart identifier.' },
          { name: 'categories', type: 'string[]', required: true, defaultVal: '—', description: 'Array of category labels.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).updateChartCategories('quarterly_chart', [
  'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun'
]);`
      },
      {
        name: 'updateChartTitle',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.updateChartTitle(chartId, title)',
        returnType: 'this (chainable)',
        description: 'Updates the title text displayed on a chart while preserving existing title typography, rotation, and font styling.',
        openXml: 'Replaces <code>&lt;a:t&gt;</code> text under <code>&lt;c:title&gt;</code> in <code>chart{N}.xml</code>.',
        params: [
          { name: 'chartId', type: 'string', required: true, defaultVal: '—', description: 'Chart name or ID.' },
          { name: 'title', type: 'string', required: true, defaultVal: '—', description: 'New title text.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(2).updateChartTitle('chart1', 'Annual Recurring Revenue (2026)');`
      },
      {
        name: 'updateDataLabels',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.updateDataLabels(chartId, options)',
        returnType: 'this (chainable)',
        description: 'Controls chart data label display visibility and formatting. Toggles value display, category names, series names, percentages, and legend keys.',
        openXml: 'Mutates <code>&lt;c:dLbls&gt;</code> (<code>&lt;c:showVal&gt;</code>, <code>&lt;c:showCatName&gt;</code>, <code>&lt;c:showPercent&gt;</code>, <code>&lt;c:showLegendKey&gt;</code>).',
        params: [
          { name: 'chartId', type: 'string', required: true, defaultVal: '—', description: 'Chart identifier.' },
          { name: 'options', type: 'Object', required: true, defaultVal: '—', description: 'Display toggles.' },
          { name: 'options.showVal', type: 'boolean', required: false, defaultVal: 'true', description: 'Display numerical value.' },
          { name: 'options.showCatName', type: 'boolean', required: false, defaultVal: 'false', description: 'Display category label.' },
          { name: 'options.showPercent', type: 'boolean', required: false, defaultVal: 'false', description: 'Display percentage.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(2).updateDataLabels('market_share_pie', {
  showVal: true,
  showPercent: true,
  showCatName: false
});`
      },
      {
        name: 'repairCharts',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.repairCharts()',
        returnType: 'Promise<this>',
        description: 'Audits and repairs chart XML inconsistencies, missing Excel workbook relationships, or corrupt formula ranges across all charts in the presentation.',
        openXml: 'Re-synthesizes missing Excel parts, fixes <code>&lt;c:f&gt;</code> cell coordinate references, and syncs <code>[Content_Types].xml</code>.',
        params: [],
        returns: { type: 'Promise<PPTXTemplater>', description: 'this (chainable).' },
        example: `await ppt.repairCharts();
console.log('All chart structures validated and synchronized.');`
      },
      {
        name: 'getCharts',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.getCharts()',
        returnType: 'Array<{ id: string, name: string, rId: string, type: string }>',
        description: 'Discovers and returns descriptors for all chart elements embedded on targeted slides.',
        openXml: 'Finds <code>&lt;c:chart&gt;</code> relationship references in <code>&lt;p:graphicFrame&gt;</code> elements.',
        params: [],
        returns: { type: 'Array<Object>', description: 'List of chart descriptors.' },
        example: `const charts = ppt.useSlide(2).getCharts();
console.log('Slide 2 contains charts:', charts);`
      },
      {
        name: 'getChartLabelPositions',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.getChartLabelPositions(chartId)',
        returnType: 'Promise<Array<Object>>',
        description: 'Calculates the spatial coordinates (in EMUs) of rendered data labels and data points on a chart canvas. Allows programmatic alignment of external callouts.',
        openXml: 'Derives geometry from <code>&lt;c:plotArea&gt;</code> and chart frame transform <code>&lt;a:xfrm&gt;</code>.',
        params: [
          { name: 'chartId', type: 'string', required: true, defaultVal: '—', description: 'Target chart identifier.' }
        ],
        returns: { type: 'Promise<Array<Object>>', description: 'Array of spatial coordinates for each series point.' },
        example: `const positions = await ppt.useSlide(2).getChartLabelPositions('salesChart');
console.log('Label positions for series 0:', positions[0]);`
      },
      {
        name: 'addTextNearChartLabel',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.addTextNearChartLabel(chartId, labelIndex, text, options = {})',
        returnType: 'Promise<this>',
        description: 'Places a floating text callout or annotation badge positioned adjacent to a specific chart data point or label.',
        openXml: 'Inserts a positioned <code>&lt;p:sp&gt;</code> shape into <code>&lt;p:spTree&gt;</code> relative to chart coordinates.',
        params: [
          { name: 'chartId', type: 'string', required: true, defaultVal: '—', description: 'Target chart identifier.' },
          { name: 'labelIndex', type: 'number', required: true, defaultVal: '—', description: '0-based data point index.' },
          { name: 'text', type: 'string', required: true, defaultVal: '—', description: 'Annotation text string.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Offset and styling options.' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'this (chainable).' },
        example: `await ppt.useSlide(2).addTextNearChartLabel('salesChart', 3, '★ Record High', {
  fill: 'FEF3C7',
  textColor: 'B45309'
});`
      }
    ]
  },
  {
    id: 'shapes-geometry',
    title: 'Shapes & Geometry',
    icon: `<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2H6a2 2 0 01-2-2V6zM14 6a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2h-2a2 2 0 01-2-2V6zM4 16a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2H6a2 2 0 01-2-2v-2zM14 16a2 2 0 012-2h2a2 2 0 012 2v2a2 2 0 01-2 2h-2a2 2 0 01-2-2v-2z"></path></svg>`,
    description: 'Vector preset shapes, geometry styling, positioning transforms, geometric alignment, cloning, and spatial distribution.',
    methods: [
      {
        name: 'addShape',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.addShape(type, options = {})',
        returnType: 'Promise<string>',
        description: 'Creates and inserts a new vector shape onto the active slide. Supports preset geometries (<code>rect</code>, <code>roundRect</code>, <code>ellipse</code>, <code>triangle</code>, <code>rightArrow</code>, etc.), hex solid fills, borders, opacity, rotation, and embedded text frames.',
        openXml: 'Inserts <code>&lt;p:sp&gt;</code> containing <code>&lt;p:nvSpPr&gt;</code>, <code>&lt;p:spPr&gt;</code> (<code>&lt;a:prstGeom prst="..."&gt;</code>, <code>&lt;a:xfrm&gt;</code>, <code>&lt;a:solidFill&gt;</code>), and <code>&lt;p:txBody&gt;</code> into <code>&lt;p:spTree&gt;</code>.',
        params: [
          { name: 'type', type: 'string', required: true, defaultVal: '—', description: 'Preset geometry type (e.g. <code>\'rect\'</code>, <code>\'roundRect\'</code>, <code>\'ellipse\'</code>).' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Geometry coordinates (x, y, width, height), fill color, border, text.' }
        ],
        returns: { type: 'Promise<string>', description: 'Assigned unique shape ID.' },
        example: `ppt.useSlide(1);
const shapeId = await ppt.addShape('roundRect', {
  x: '10%', y: '15%', width: '80%', height: '12%',
  fill: '2563EB',
  text: 'Quarterly Milestone Achieved',
  textColor: 'FFFFFF',
  fontSize: 16,
  bold: true
});`
      },
      {
        name: 'updateShape',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.updateShape(shapeId, options)',
        returnType: 'Promise<this>',
        description: 'Updates spatial coordinates, fills, borders, or visibility of an existing shape by shape ID or shape name.',
        openXml: 'Mutates <code>&lt;a:xfrm&gt;</code> (<code>&lt;a:off&gt;</code>, <code>&lt;a:ext&gt;</code>) and <code>&lt;a:solidFill&gt;</code> under <code>&lt;p:spPr&gt;</code>.',
        params: [
          { name: 'shapeId', type: 'string', required: true, defaultVal: '—', description: 'Target shape ID or shape name.' },
          { name: 'options', type: 'Object', required: true, defaultVal: '—', description: 'Properties to update (x, y, width, height, fill, stroke).' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'this (chainable).' },
        example: `await ppt.useSlide(1).updateShape('banner_box', {
  fill: '10B981',
  width: '85%'
});`
      },
      {
        name: 'updateShapeText',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.updateShapeText(shapeId, text, options = {})',
        returnType: 'Promise<this>',
        description: 'Replaces the entire text body of a shape with new text while preserving or updating font styles.',
        openXml: 'Updates <code>&lt;a:p&gt;</code> and <code>&lt;a:r&gt;</code> in <code>&lt;p:txBody&gt;</code>.',
        params: [
          { name: 'shapeId', type: 'string', required: true, defaultVal: '—', description: 'Shape ID or name.' },
          { name: 'text', type: 'string', required: true, defaultVal: '—', description: 'New text content.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Typography formatting options.' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'this (chainable).' },
        example: `await ppt.useSlide(1).updateShapeText('badge', '38.2% YoY Growth', {
  bold: true,
  color: 'FFFFFF'
});`
      },
      {
        name: 'cloneShape',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.cloneShape(shapeId, options = {})',
        returnType: 'Promise<string>',
        description: 'Deep-clones a shape on the active slide, assigning a new unique shape ID and offsetting its coordinates.',
        openXml: 'Deep-clones <code>&lt;p:sp&gt;</code> node, re-indexes <code>&lt;p:cNvPr id="..."&gt;</code>, and adjusts <code>&lt;a:off&gt;</code>.',
        params: [
          { name: 'shapeId', type: 'string', required: true, defaultVal: '—', description: 'Shape ID to clone.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Coordinate offsets (dx, dy) or override properties.' }
        ],
        returns: { type: 'Promise<string>', description: 'New shape ID.' },
        example: `const cloneId = await ppt.useSlide(1).cloneShape('card_template', {
  dx: '3.2in',
  text: 'Secondary Metric'
});`
      },
      {
        name: 'deleteShape',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.deleteShape(shapeId)',
        returnType: 'Promise<this>',
        description: 'Removes a shape element from the active slide by shape ID or name.',
        openXml: 'Removes <code>&lt;p:sp&gt;</code> node from <code>&lt;p:spTree&gt;</code>.',
        params: [
          { name: 'shapeId', type: 'string', required: true, defaultVal: '—', description: 'Shape identifier to delete.' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'this (chainable).' },
        example: `await ppt.useSlide(1).deleteShape('draft_watermark');`
      },
      {
        name: 'alignShapes',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.alignShapes(shapeIds, alignment)',
        returnType: 'this (chainable)',
        description: 'Aligns multiple shapes along a common edge or centerline (<code>\'left\'</code>, <code>\'center\'</code>, <code>\'right\'</code>, <code>\'top\'</code>, <code>\'middle\'</code>, <code>\'bottom\'</code>).',
        openXml: 'Calculates bounding coordinates and mutates <code>&lt;a:off x="..." y="..."&gt;</code> on <code>&lt;a:xfrm&gt;</code>.',
        params: [
          { name: 'shapeIds', type: 'string[]', required: true, defaultVal: '—', description: 'Array of shape IDs or names to align.' },
          { name: 'alignment', type: "'left' | 'center' | 'right' | 'top' | 'middle' | 'bottom'", required: true, defaultVal: '—', description: 'Target alignment axis.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).alignShapes(['card1', 'card2', 'card3'], 'top');`
      },
      {
        name: 'addTextAtPosition',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.addTextAtPosition(text, options)',
        returnType: 'Promise<string>',
        description: 'Creates a transparent floating text box at specific slide coordinates.',
        openXml: 'Creates <code>&lt;p:sp&gt;</code> with transparent geometry (<code>&lt;a:noFill/&gt;</code>) and text body.',
        params: [
          { name: 'text', type: 'string', required: true, defaultVal: '—', description: 'Text string.' },
          { name: 'options', type: 'Object', required: true, defaultVal: '—', description: 'Position coordinates (x, y, width, height) and typography.' }
        ],
        returns: { type: 'Promise<string>', description: 'Assigned shape ID.' },
        example: `await ppt.useSlide(1).addTextAtPosition('Confidential — Internal Use Only', {
  x: '5%', y: '92%', width: '90%', height: '5%',
  color: '64748B', fontSize: 10
});`
      },
      {
        name: 'getShapes',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.getShapes()',
        returnType: 'Array<{ id: string, name: string, type: string, bounds: Object }>',
        description: 'Returns all shapes on the active slide with IDs, names, preset geometries, and bounding box coordinates.',
        openXml: 'Scans <code>&lt;p:spTree&gt;</code> children.',
        params: [],
        returns: { type: 'Array<Object>', description: 'List of shape descriptors.' },
        example: `const shapes = ppt.useSlide(1).getShapes();
console.log('Slide 1 shapes:', shapes);`
      },
      {
        name: 'validateShape',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.validateShape(shapeId)',
        returnType: '{ valid: boolean, errors: string[] }',
        description: 'Validates DrawingML schema compliance for a given shape.',
        openXml: 'Audits <code>&lt;p:sp&gt;</code> XML node.',
        params: [
          { name: 'shapeId', type: 'string', required: true, defaultVal: '—', description: 'Shape ID.' }
        ],
        returns: { type: '{ valid: boolean, errors: string[] }', description: 'Validation result.' },
        example: `const check = ppt.useSlide(1).validateShape('hero_box');
if (!check.valid) console.warn(check.errors);`
      }
    ]
  },
  {
    id: 'images-media',
    title: 'Images & Media',
    icon: `<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 16l4.586-4.586a2 2 0 012.828 0L16 16m-2-2l1.586-1.586a2 2 0 012.828 0L20 14m-6-6h.01M6 20h12a2 2 0 002-2V6a2 2 0 00-2-2H6a2 2 0 00-2 2v12a2 2 0 002 2z"></path></svg>`,
    description: 'Embedding raster images, replacing placeholder images, maintaining aspect ratios, and media extraction.',
    methods: [
      {
        name: 'addImage',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.addImage(source, options = {})',
        returnType: 'Promise<string>',
        description: 'Embeds a new image (PNG, JPEG, GIF, SVG, WebP) onto the active slide. Copies image binary data to <code>ppt/media/</code>, creates a relationship link in <code>slide{N}.xml.rels</code>, and synthesizes a compliant <code>&lt;p:pic&gt;</code> DrawingML picture element.',
        openXml: 'Writes binary file to <code>ppt/media/image{N}.ext</code>, registers relationship in <code>slide{N}.xml.rels</code>, and inserts <code>&lt;p:pic&gt;</code> with <code>&lt;a:blip r:embed="rId..."/&gt;</code>.',
        params: [
          { name: 'source', type: 'string | Buffer', required: true, defaultVal: '—', description: 'File path string or Buffer containing image binary bytes.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Position coordinates (x, y, width, height) and aspect ratio locking.' }
        ],
        returns: { type: 'Promise<string>', description: 'Assigned picture shape ID.' },
        example: `ppt.useSlide(1);
const picId = await ppt.addImage('./assets/company-logo.png', {
  x: '5%', y: '5%', width: '2.5in', height: '0.8in'
});`
      },
      {
        name: 'replaceImage',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.replaceImage(imageIdOrName, source)',
        returnType: 'Promise<this>',
        description: 'Replaces an existing image placeholder or picture element with new image binary data while perfectly preserving original bounding coordinates, dimensions, and z-ordering.',
        openXml: 'Replaces binary bytes in <code>ppt/media/</code> or updates <code>r:embed</code> target relationship.',
        params: [
          { name: 'imageIdOrName', type: 'string', required: true, defaultVal: '—', description: 'Existing image shape ID or picture name.' },
          { name: 'source', type: 'string | Buffer', required: true, defaultVal: '—', description: 'New image file path or Buffer.' }
        ],
        returns: { type: 'Promise<PPTXTemplater>', description: 'this (chainable).' },
        example: `await ppt.useSlide(1).replaceImage('headshot_placeholder', './photos/ceo.jpg');`
      },
      {
        name: 'getImages',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.getImages()',
        returnType: 'Array<{ id: string, name: string, rId: string, target: string, bounds: Object }>',
        description: 'Discovers and returns descriptors for all images embedded on the active slide, including target media file paths and spatial bounds.',
        openXml: 'Inspects <code>&lt;p:pic&gt;</code> elements and resolves relationship IDs in <code>slide{N}.xml.rels</code>.',
        params: [],
        returns: { type: 'Array<Object>', description: 'List of embedded image descriptors.' },
        example: `const images = ppt.useSlide(1).getImages();
console.log('Images on slide 1:', images);`
      }
    ]
  },
  {
    id: 'z-order-layering',
    title: 'Z-Order & Layering',
    icon: `<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M7 16V4m0 0L3 8m4-4l4 4m6 0v12m0 0l4-4m-4 4l-4-4"></path></svg>`,
    description: 'Visual stacking order within the DrawingML spTree, reordering, swapping, and normalizing z-indexes.',
    methods: [
      {
        name: 'bringToFront',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.bringToFront(objectId)',
        returnType: 'this (chainable)',
        description: 'Moves an object (shape, image, table, chart, group) to the very top of the visual stacking order on the active slide, ensuring it renders above all other elements.',
        openXml: 'Repositions the element XML node to the end of child elements in <code>&lt;p:spTree&gt;</code>.',
        params: [
          { name: 'objectId', type: 'string', required: true, defaultVal: '—', description: 'Shape ID or object name.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).bringToFront('active_badge');`
      },
      {
        name: 'sendToBack',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.sendToBack(objectId)',
        returnType: 'this (chainable)',
        description: 'Moves an object to the bottom of the visual stacking order, behind all other shapes and content on the active slide.',
        openXml: 'Repositions element XML node to the beginning of child elements in <code>&lt;p:spTree&gt;</code>.',
        params: [
          { name: 'objectId', type: 'string', required: true, defaultVal: '—', description: 'Shape ID or object name.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).sendToBack('background_card');`
      },
      {
        name: 'bringForward',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.bringForward(objectId)',
        returnType: 'this (chainable)',
        description: 'Increments the visual layer of an object by one step in the stacking hierarchy.',
        openXml: 'Swaps element node with its immediate next sibling in <code>&lt;p:spTree&gt;</code>.',
        params: [
          { name: 'objectId', type: 'string', required: true, defaultVal: '—', description: 'Object ID.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).bringForward('icon_overlay');`
      },
      {
        name: 'sendBackward',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.sendBackward(objectId)',
        returnType: 'this (chainable)',
        description: 'Decrements the visual layer of an object by one step in the stacking hierarchy.',
        openXml: 'Swaps element node with its immediate previous sibling in <code>&lt;p:spTree&gt;</code>.',
        params: [
          { name: 'objectId', type: 'string', required: true, defaultVal: '—', description: 'Object ID.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).sendBackward('shadow_shape');`
      },
      {
        name: 'setZIndex',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.setZIndex(objectId, targetIndex)',
        returnType: 'this (chainable)',
        description: 'Moves an object to an explicit 0-based position in the visual stack array.',
        openXml: 'Slices and splices object node within <code>&lt;p:spTree&gt;</code> child nodes array.',
        params: [
          { name: 'objectId', type: 'string', required: true, defaultVal: '—', description: 'Object ID.' },
          { name: 'targetIndex', type: 'number', required: true, defaultVal: '—', description: '0-based index in stacking order.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).setZIndex('watermark', 0);`
      },
      {
        name: 'reorderObjects',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.reorderObjects(orderedIds)',
        returnType: 'this (chainable)',
        description: 'Re-orders all objects on the active slide according to an explicit back-to-front array of IDs.',
        openXml: 'Re-sorts child nodes of <code>&lt;p:spTree&gt;</code> matching the order of <code>orderedIds</code>.',
        params: [
          { name: 'orderedIds', type: 'string[]', required: true, defaultVal: '—', description: 'Array of object IDs ordered from bottom-most to top-most.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).reorderObjects(['bg', 'chart1', 'overlay', 'title']);`
      },
      {
        name: 'swapObjects',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.swapObjects(objectIdA, objectIdB)',
        returnType: 'this (chainable)',
        description: 'Swaps the visual layers of two specified objects.',
        openXml: 'Exchanges position indices of two nodes within <code>&lt;p:spTree&gt;</code>.',
        params: [
          { name: 'objectIdA', type: 'string', required: true, defaultVal: '—', description: 'First object ID.' },
          { name: 'objectIdB', type: 'string', required: true, defaultVal: '—', description: 'Second object ID.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).swapObjects('shapeA', 'shapeB');`
      },
      {
        name: 'getObjectOrder',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.getObjectOrder()',
        returnType: 'string[]',
        description: 'Returns the current back-to-front visual order of object IDs on the active slide.',
        openXml: 'Maps child node IDs in <code>&lt;p:spTree&gt;</code>.',
        params: [],
        returns: { type: 'string[]', description: 'Array of object IDs in layer order.' },
        example: `const order = ppt.useSlide(1).getObjectOrder();
console.log('Stacking order:', order);`
      },
      {
        name: 'normalizeZOrder',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.normalizeZOrder()',
        returnType: 'this (chainable)',
        description: 'Normalizes and cleans shape IDs and child node ordering to prevent Office repair warnings after heavy additions and deletions.',
        openXml: 'Recompacts shape IDs and structural ordering in <code>&lt;p:spTree&gt;</code>.',
        params: [],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).normalizeZOrder();`
      }
    ]
  },
  {
    id: 'hyperlinks-actions',
    title: 'Hyperlinks & Actions',
    icon: `<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101m-.758-4.899a4 4 0 005.656 0l4-4a4 4 0 00-5.656-5.656l-1.1 1.1"></path></svg>`,
    description: 'Web URLs, slide jump links, click actions on shapes, text runs, and pictures with relationship remap.',
    methods: [
      {
        name: 'addHyperlink',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.addHyperlink(options)',
        returnType: 'this (chainable)',
        description: 'Attaches an external web URL hyperlink to a text run or shape. Creates an external relationship in <code>slide{N}.xml.rels</code> and registers the click action.',
        openXml: 'Adds <code>&lt;a:hlinkClick r:id="rId..."/&gt;</code> under <code>&lt;a:rPr&gt;</code> and registers relationship in <code>slide{N}.xml.rels</code> with <code>TargetMode="External"</code>.',
        params: [
          { name: 'options', type: 'Object', required: true, defaultVal: '—', description: 'Hyperlink configuration.' },
          { name: 'options.url', type: 'string', required: true, defaultVal: '—', description: 'Web URL destination (e.g. <code>\'https://example.com\'</code>).' },
          { name: 'options.text', type: 'string', required: false, defaultVal: '—', description: 'Specific text substring to attach link to.' },
          { name: 'options.shapeId', type: 'string', required: false, defaultVal: '—', description: 'Target shape ID.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).addHyperlink({
  shapeId: 'contact_box',
  text: 'Visit Documentation',
  url: 'https://jsuyog2.github.io/node-pptx-templater'
});`
      },
      {
        name: 'addSlideLink',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.addSlideLink(options)',
        returnType: 'this (chainable)',
        description: 'Creates an internal navigation jump link that transitions the slideshow presentation directly to another specified slide index when clicked.',
        openXml: 'Appends <code>&lt;a:hlinkClick action="ppaction://hlinksldjump" r:id="rId..."/&gt;</code> linked to target slide relationship.',
        params: [
          { name: 'options', type: 'Object', required: true, defaultVal: '—', description: 'Slide jump configuration.' },
          { name: 'options.targetSlide', type: 'number', required: true, defaultVal: '—', description: '1-based target slide index to navigate to.' },
          { name: 'options.shapeId', type: 'string', required: false, defaultVal: '—', description: 'Target shape ID.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `// Link table of contents item directly to Appendix slide 8
ppt.useSlide(2).addSlideLink({
  shapeId: 'toc_item_4',
  targetSlide: 8
});`
      },
      {
        name: 'addImageLink',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.addImageLink(imageId, url)',
        returnType: 'this (chainable)',
        description: 'Attaches an external web hyperlink action to an image or picture element on the active slide.',
        openXml: 'Attaches <code>&lt;a:hlinkClick r:id="..."/&gt;</code> under <code>&lt;p:cNvPr&gt;</code> inside <code>&lt;p:pic&gt;</code>.',
        params: [
          { name: 'imageId', type: 'string', required: true, defaultVal: '—', description: 'Image shape ID or picture name.' },
          { name: 'url', type: 'string', required: true, defaultVal: '—', description: 'Target web URL.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).addImageLink('logo_img', 'https://github.com/jsuyog2/node-pptx-templater');`
      },
      {
        name: 'addShapeLink',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.addShapeLink(shapeId, url)',
        returnType: 'this (chainable)',
        description: 'Attaches an external web click action to an entire vector shape.',
        openXml: 'Mutates <code>&lt;p:cNvPr&gt;</code> in <code>&lt;p:nvSpPr&gt;</code> and links external URL relationship.',
        params: [
          { name: 'shapeId', type: 'string', required: true, defaultVal: '—', description: 'Shape ID.' },
          { name: 'url', type: 'string', required: true, defaultVal: '—', description: 'Target web URL.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).addShapeLink('cta_button', 'https://company.com/signup');`
      },
      {
        name: 'addTextNavigationLink',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.addTextNavigationLink(shapeId, targetSlideNumber, options = {})',
        returnType: 'this (chainable)',
        description: 'Creates an internal slide jump link on a specific text run inside a shape.',
        openXml: 'Adds <code>ppaction://hlinksldjump</code> to targeted <code>&lt;a:rPr&gt;</code>.',
        params: [
          { name: 'shapeId', type: 'string', required: true, defaultVal: '—', description: 'Shape ID.' },
          { name: 'targetSlideNumber', type: 'number', required: true, defaultVal: '—', description: '1-based destination slide index.' },
          { name: 'options', type: 'Object', required: false, defaultVal: '{}', description: 'Matching options.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(2).addTextNavigationLink('agenda_box', 5, { text: 'Financials' });`
      },
      {
        name: 'addShapeNavigationLink',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.addShapeNavigationLink(shapeId, targetSlideNumber)',
        returnType: 'this (chainable)',
        description: 'Makes an entire vector shape trigger navigation to another slide when clicked.',
        openXml: 'Applies internal jump click action to shape <code>&lt;p:cNvPr&gt;</code>.',
        params: [
          { name: 'shapeId', type: 'string', required: true, defaultVal: '—', description: 'Shape ID.' },
          { name: 'targetSlideNumber', type: 'number', required: true, defaultVal: '—', description: '1-based destination slide index.' }
        ],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.useSlide(1).addShapeNavigationLink('next_button', 2);`
      },
      {
        name: 'getHyperlinks',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.getHyperlinks(slideIndex)',
        returnType: 'Array<{ type: string, target: string, text?: string, shapeId?: string }>',
        description: 'Extracts and lists all internal and external hyperlinks defined on the specified slide.',
        openXml: 'Scans <code>slide{N}.xml.rels</code> and matches <code>&lt;a:hlinkClick&gt;</code> tags.',
        params: [
          { name: 'slideIndex', type: 'number', required: false, defaultVal: 'active slide', description: '1-based slide index.' }
        ],
        returns: { type: 'Array<Object>', description: 'List of hyperlink descriptors.' },
        example: `const links = ppt.useSlide(1).getHyperlinks();
console.log('Slide 1 links:', links);`
      }
    ]
  },
  {
    id: 'validation-diagnostics',
    title: 'Validation & Diagnostics',
    icon: `<svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>`,
    description: 'OpenXML schema verification, relationship integrity auditing, automatic corruption repairs, XML inspection, and performance profiling.',
    methods: [
      {
        name: 'validate',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.validate()',
        returnType: 'Promise<{ valid: boolean, errors: string[] }>',
        description: 'Performs a comprehensive integrity check of the entire PowerPoint package. Verifies PresentationML schemas, checks all XML tags for balanced closures, audits DrawingML tree structures, and ensures no dangling references exist.',
        openXml: 'Audits <code>[Content_Types].xml</code>, <code>_rels/.rels</code>, <code>ppt/presentation.xml</code>, all slides, layouts, masters, and charts.',
        params: [],
        returns: { type: 'Promise<{ valid: boolean, errors: string[] }>', description: 'Validation outcome with list of any structural errors.' },
        example: `const result = await ppt.validate();
if (!result.valid) {
  console.error('Validation issues found:', result.errors);
} else {
  console.log('Presentation is 100% compliant and non-corrupting.');
}`
      },
      {
        name: 'validatePresentation',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.validatePresentation()',
        returnType: 'Promise<{ valid: boolean, errors: string[] }>',
        description: 'Audits presentation-level structures: verifies slide ordering, checks layout linkages, and ensures master slide relationship consistency.',
        openXml: 'Validates <code>ppt/presentation.xml</code> and <code>ppt/_rels/presentation.xml.rels</code>.',
        params: [],
        returns: { type: 'Promise<Object>', description: 'Validation outcome.' },
        example: `const check = await ppt.validatePresentation();`
      },
      {
        name: 'validateTable',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.validateTable(tableId)',
        returnType: 'Promise<{ valid: boolean, errors: string[] }>',
        description: 'Validates a table\'s DrawingML grid structures, ensuring cell counts match grid columns and merge coordinates do not overlap or exceed table boundaries.',
        openXml: 'Audits <code>&lt;a:tblGrid&gt;</code> and <code>&lt;a:tr&gt;</code> row/cell consistency.',
        params: [
          { name: 'tableId', type: 'string', required: true, defaultVal: '—', description: 'Table name or shape ID.' }
        ],
        returns: { type: 'Promise<Object>', description: 'Validation outcome.' },
        example: `const res = await ppt.useSlide(1).validateTable('projectTable');`
      },
      {
        name: 'validateRelationships',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.validateRelationships()',
        returnType: 'Promise<{ valid: boolean, orphaned: string[], missing: string[] }>',
        description: 'Scans all <code>.rels</code> files across the presentation to ensure every relationship targets an existing package part and no unreferenced orphaned assets exist.',
        openXml: 'Audits all <code>_rels/*.rels</code> files against package file entries.',
        params: [],
        returns: { type: 'Promise<Object>', description: 'Audit result with orphaned and missing target lists.' },
        example: `const relsAudit = await ppt.validateRelationships();
console.log('Orphaned parts:', relsAudit.orphaned);`
      },
      {
        name: 'repair',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.repair()',
        returnType: 'Promise<this>',
        description: 'Automatically analyzes and repairs common PowerPoint corruption triggers: fixes orphaned relationship IDs, resolves unclosed XML run tags, regenerates missing Content-Type entries, and normalizes table structures.',
        openXml: 'Corrects <code>[Content_Types].xml</code>, trims broken <code>.rels</code> links, and repairs corrupted DrawingML trees.',
        params: [],
        returns: { type: 'Promise<PPTXTemplater>', description: 'this (chainable).' },
        example: `// Automatically fix an externally modified or corrupted deck
await ppt.repair();
await ppt.save('./repaired-deck.pptx');`
      },
      {
        name: 'inspectXML',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.inspectXML(xmlPath)',
        returnType: 'Promise<string>',
        description: 'Returns the formatted, syntax-indented XML content of any internal part in the presentation package (e.g. <code>ppt/slides/slide1.xml</code>).',
        openXml: 'Reads and pretty-prints raw XML part string.',
        params: [
          { name: 'xmlPath', type: 'string', required: true, defaultVal: '—', description: 'Relative part path inside PPTX package.' }
        ],
        returns: { type: 'Promise<string>', description: 'Formatted XML string.' },
        example: `const slideXml = await ppt.inspectXML('ppt/slides/slide1.xml');
console.log(slideXml.substring(0, 500));`
      },
      {
        name: 'inspectSlide',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.inspectSlide(slideNumber)',
        returnType: 'Promise<string>',
        description: 'Returns the formatted, indented XML string of a specific slide.',
        openXml: 'Retrieves <code>ppt/slides/slide{N}.xml</code>.',
        params: [
          { name: 'slideNumber', type: 'number', required: true, defaultVal: '—', description: '1-based slide index.' }
        ],
        returns: { type: 'Promise<string>', description: 'Formatted slide XML.' },
        example: `const xml = await ppt.inspectSlide(1);`
      },
      {
        name: 'inspectChart',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.inspectChart(chartId)',
        returnType: 'Promise<string>',
        description: 'Returns the formatted XML content of a specific chart part.',
        openXml: 'Retrieves <code>ppt/charts/chart{N}.xml</code>.',
        params: [
          { name: 'chartId', type: 'string', required: true, defaultVal: '—', description: 'Chart name or ID.' }
        ],
        returns: { type: 'Promise<string>', description: 'Formatted chart XML.' },
        example: `const chartXml = await ppt.inspectChart('chart1');`
      },
      {
        name: 'debugRelationships',
        isStatic: false,
        isAsync: true,
        signature: 'ppt.debugRelationships(slideNumber)',
        returnType: 'Promise<Array<Object>>',
        description: 'Dumps all relationship entries linked to a specific slide, listing relationship IDs, schemas, and target files.',
        openXml: 'Parses <code>ppt/slides/_rels/slide{N}.xml.rels</code>.',
        params: [
          { name: 'slideNumber', type: 'number', required: true, defaultVal: '—', description: '1-based slide index.' }
        ],
        returns: { type: 'Promise<Array<Object>>', description: 'Array of relationship entries.' },
        example: `const rels = await ppt.debugRelationships(1);
console.table(rels);`
      },
      {
        name: 'enablePerformanceProfile',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.enablePerformanceProfile()',
        returnType: 'this (chainable)',
        description: 'Activates microsecond performance telemetry profiling for all subsequent XML parsing, AST manipulations, and zip serialization.',
        openXml: 'Hooks high-resolution timers around XML mutations and package compilation.',
        params: [],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.enablePerformanceProfile();
ppt.useSlide(1).replaceText({ '{{title}}': 'Fast Generation' });
const metrics = ppt.getPerformanceMetrics();
console.log('Execution timing breakdown:', metrics);`
      },
      {
        name: 'getPerformanceMetrics',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.getPerformanceMetrics()',
        returnType: 'Object',
        description: 'Returns high-resolution timing metrics for operations performed on the presentation.',
        openXml: 'Returns telemetry data map.',
        params: [],
        returns: { type: 'Object', description: 'Metrics summary object with durations in milliseconds.' },
        example: `const metrics = ppt.getPerformanceMetrics();`
      },
      {
        name: 'enableDebug',
        isStatic: false,
        isAsync: false,
        signature: 'ppt.enableDebug()',
        returnType: 'this (chainable)',
        description: 'Sets global and instance logger verbosity to debug level, printing detailed OpenXML transformation traces to stdout.',
        openXml: 'Enables debug logging for all internal managers.',
        params: [],
        returns: { type: 'PPTXTemplater', description: 'this (chainable).' },
        example: `ppt.enableDebug();`
      }
    ]
  }
];

// Helper to escape HTML characters
function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

// Light JavaScript syntax highlighter for code snippets
function highlightJs(code) {
  let esc = escapeHtml(code);

  // Comments (single line)
  esc = esc.replace(/(\/\/[^\n]*)/g, '<span class="hl-comment">$1</span>');

  // Strings (single and double quotes, backticks)
  esc = esc.replace(/('(?:\\'|[^'\n])*')/g, '<span class="hl-string">$1</span>');
  esc = esc.replace(/("(?:\\"|[^"\n])*")/g, '<span class="hl-string">$1</span>');

  // Keywords
  const keywords = ['const', 'let', 'var', 'async', 'await', 'function', 'return', 'class', 'import', 'from', 'require', 'new', 'for', 'of', 'in', 'if', 'else', 'true', 'false', 'null', 'undefined', 'static', 'try', 'catch'];
  const kwRegex = new RegExp(`\\b(${keywords.join('|')})\\b`, 'g');
  esc = esc.replace(kwRegex, '<span class="hl-keyword">$1</span>');

  // Function calls
  esc = esc.replace(/\.([a-zA-Z0-9_]+)\(/g, '.<span class="hl-fn">$1</span>(');

  // Numbers
  esc = esc.replace(/\b([0-9]+(?:\.[0-9]+)?)\b/g, '<span class="hl-number">$1</span>');

  return esc;
}

// Generate the full HTML document
function generateHtml() {
  let totalMethodsCount = 0;
  MODULES.forEach(m => totalMethodsCount += m.methods.length);

  // Build Sidebar Categories and Links
  let sidebarModulesHtml = '';
  let rightTocModulesHtml = '';
  let mainContentHtml = '';

  MODULES.forEach((mod, modIdx) => {
    // Sidebar category block
    sidebarModulesHtml += `
      <div class="sidebar-category mb-4" data-category="${mod.id}">
        <button type="button" class="category-header w-full flex items-center justify-between px-2.5 py-2 rounded-lg text-xs font-semibold text-slate-700 dark:text-slate-300 hover:bg-slate-200/50 dark:hover:bg-slate-800/50 transition-colors group" onclick="toggleCategory('${mod.id}')">
          <div class="flex items-center gap-2">
            <span class="text-slate-400 group-hover:text-blue-500 transition-colors">${mod.icon}</span>
            <span class="tracking-tight">${mod.title}</span>
          </div>
          <span class="category-badge text-[10px] font-mono px-1.5 py-0.5 rounded-full bg-slate-200 dark:bg-slate-800 text-slate-500 dark:text-slate-400">${mod.methods.length}</span>
        </button>
        <div id="cat-list-${mod.id}" class="category-methods mt-1 ml-4 pl-2 border-l border-slate-200 dark:border-slate-800 space-y-0.5">
          ${mod.methods.map(m => `
            <a href="#${m.name}" class="sidebar-method-link block px-2 py-1 text-xs text-slate-600 dark:text-slate-400 hover:text-blue-600 dark:hover:text-blue-400 rounded transition-colors truncate font-mono" data-method="${m.name}" title="${m.signature}">
              ${m.name}
            </a>
          `).join('')}
        </div>
      </div>
    `;

    // Right-hand TOC category
    rightTocModulesHtml += `
      <div class="mb-3">
        <a href="#section-${mod.id}" class="right-toc-section block text-xs font-semibold text-slate-700 dark:text-slate-300 hover:text-blue-600 dark:hover:text-blue-400 mb-1">
          ${mod.title}
        </a>
        <div class="space-y-0.5 ml-2 border-l border-slate-200 dark:border-slate-800 pl-2">
          ${mod.methods.map(m => `
            <a href="#${m.name}" class="right-toc-link block text-[11px] text-slate-500 dark:text-slate-400 hover:text-blue-500 truncate font-mono py-0.5" data-toc-target="${m.name}">
              ${m.name}
            </a>
          `).join('')}
        </div>
      </div>
    `;

    // Main Content Section
    mainContentHtml += `
      <section id="section-${mod.id}" class="module-section py-10 border-b border-slate-200 dark:border-slate-800">
        <!-- Module Header -->
        <div class="mb-8">
          <div class="flex items-center gap-2 mb-2">
            <span class="p-2 rounded-xl bg-blue-500/10 text-blue-600 dark:text-blue-400 border border-blue-500/20">
              ${mod.icon}
            </span>
            <span class="text-xs font-mono font-bold uppercase tracking-wider text-blue-600 dark:text-blue-400">Module ${modIdx + 1} of 10</span>
          </div>
          <h2 class="text-2xl sm:text-3xl font-extrabold tracking-tight text-slate-900 dark:text-white">
            ${mod.title}
          </h2>
          <p class="mt-2 text-sm sm:text-base text-slate-600 dark:text-slate-300 leading-relaxed max-w-3xl">
            ${mod.description}
          </p>
        </div>

        <!-- Function Cards -->
        <div class="space-y-8">
          ${mod.methods.map(m => renderFunctionCard(m, mod)).join('')}
        </div>
      </section>
    `;
  });

  return `<!DOCTYPE html>
<html lang="en" class="scroll-smooth">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>API Reference — node-pptx-templater</title>
  <meta name="description" content="Comprehensive API Reference for node-pptx-templater. Complete specifications, parameters, OpenXML mutations, and copy-and-paste examples for all 95 public methods.">

  <!-- Favicon / Meta -->
  <link rel="icon" type="image/png" href="./assets/logo.png">
  <link rel="canonical" href="https://node-pptx-templater.pages.dev/api-reference.html">
  <meta property="og:title" content="API Reference — node-pptx-templater">
  <meta property="og:description" content="Production-ready API Reference for node-pptx-templater with categorized navigation and 1-click copyable snippets.">
  <meta property="og:type" content="website">
  <meta property="og:url" content="https://node-pptx-templater.pages.dev/api-reference.html">

  <!-- Inter & JetBrains Mono Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500;600&display=swap" rel="stylesheet">

  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <script>
    tailwind.config = {
      darkMode: 'class',
      theme: {
        extend: {
          colors: {
            brand: {
              50: '#eff6ff',
              100: '#dbeafe',
              500: '#3b82f6',
              600: '#2563eb',
              700: '#1d4ed8',
              800: '#1e40af',
              900: '#1e3a8a',
            }
          },
          fontFamily: {
            sans: ['Inter', 'sans-serif'],
            mono: ['JetBrains Mono', 'monospace'],
          }
        }
      }
    }
  </script>

  <!-- Custom Glassmorphic Stylesheet -->
  <link rel="stylesheet" href="./css/style.css">

  <style>
    /* Custom Layout & Responsive Styles for API Reference */
    :target {
      scroll-margin-top: 5.5rem;
    }

    .api-card {
      transition: border-color 0.2s ease, box-shadow 0.2s ease;
    }
    .api-card:target {
      border-color: #3b82f6;
      box-shadow: 0 0 25px -4px rgba(59, 130, 246, 0.35);
    }

    /* Floating Mobile Menu Button */
    #mobile-fab {
      position: fixed;
      bottom: 1.5rem;
      left: 1.5rem;
      z-index: 40;
    }

    /* Widescreen 3-Column Layout */
    @media (min-width: 1440px) {
      .api-grid-layout {
        display: grid;
        grid-template-columns: 280px minmax(0, 1fr) 240px;
        gap: 2.5rem;
      }
    }

    /* Desktop 2-Column Layout */
    @media (min-width: 1024px) and (max-width: 1439px) {
      .api-grid-layout {
        display: grid;
        grid-template-columns: 260px minmax(0, 1fr);
        gap: 2rem;
      }
    }

    /* Tablet Layout */
    @media (min-width: 768px) and (max-width: 1023px) {
      .api-grid-layout {
        display: grid;
        grid-template-columns: 220px minmax(0, 1fr);
        gap: 1.5rem;
      }
    }

    /* Mobile Layout */
    @media (max-width: 767px) {
      .api-grid-layout {
        display: block;
      }
    }

    /* Parameter Table horizontal scroll styling */
    .table-scroll-container {
      overflow-x: auto;
      -webkit-overflow-scrolling: touch;
    }

    /* Code Container & Copy Button */
    .code-container {
      position: relative;
    }
    .code-container pre {
      margin: 0;
      padding: 1.15rem;
      overflow-x: auto;
      background: #090d16;
      border-radius: 0.65rem;
      font-size: 0.8125rem;
      line-height: 1.6;
      font-family: var(--font-mono);
      border: 1px solid rgba(255, 255, 255, 0.08);
      color: #e2e8f0;
    }
    .code-container .copy-btn {
      position: absolute;
      top: 0.625rem;
      right: 0.625rem;
      z-index: 10;
    }

    /* Search input clear button */
    #search-clear-btn {
      display: none;
    }
  </style>

  <!-- Anti-FOUC Theme Initializer -->
  <script>
    (function() {
      const savedTheme = localStorage.getItem('theme');
      if (savedTheme === 'dark' || (!savedTheme && window.matchMedia('(prefers-color-scheme: dark)').matches)) {
        document.documentElement.classList.add('dark');
      } else {
        document.documentElement.classList.remove('dark');
      }
    })();
  </script>
</head>
<body class="bg-slate-50 text-slate-900 dark:bg-[#0b0f19] dark:text-slate-100 transition-colors duration-200 antialiased selection:bg-brand-500 selection:text-white">

  <!-- Ambient Background Glows -->
  <div class="hero-glow-1"></div>
  <div class="hero-glow-2"></div>

  <!-- Toast Notification Container -->
  <div id="toast-container"></div>

  <!-- Mobile Backdrop -->
  <div id="mobile-backdrop" class="fixed inset-0 bg-slate-950/60 backdrop-blur-sm z-40 hidden md:hidden" onclick="closeMobileMenu()"></div>

  <!-- Floating Mobile Menu Button -->
  <div id="mobile-fab" class="md:hidden">
    <button onclick="toggleMobileMenu()" class="flex items-center gap-2 px-4 py-2.5 rounded-full bg-blue-600 hover:bg-blue-700 text-white font-medium shadow-xl shadow-blue-500/30 text-xs">
      <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M4 6h16M4 12h16M4 18h7"></path></svg>
      <span>Browse Methods</span>
    </button>
  </div>

  <!-- Mobile Navigation Drawer -->
  <div id="mobile-drawer" class="fixed top-0 bottom-0 left-0 w-80 max-w-[85vw] bg-white dark:bg-slate-900 border-r border-slate-200 dark:border-slate-800 z-50 transform -translate-x-full transition-transform duration-300 ease-in-out md:hidden flex flex-col p-5 shadow-2xl">
    <div class="flex items-center justify-between pb-4 border-b border-slate-200 dark:border-slate-800">
      <div class="flex items-center gap-2">
        <img src="./assets/logo.png" alt="Logo" class="w-6 h-6 object-contain">
        <span class="font-bold text-sm tracking-tight">API Reference</span>
      </div>
      <button onclick="closeMobileMenu()" class="p-1 rounded-lg text-slate-500 hover:text-slate-800 dark:hover:text-slate-200" aria-label="Close menu">
        <svg class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M6 18L18 6M6 6l12 12"></path></svg>
      </button>
    </div>

    <!-- Mobile Search -->
    <div class="mt-4">
      <div class="relative">
        <input type="text" id="mobile-search-input" placeholder="Search ${totalMethodsCount} methods..." oninput="handleMobileSearch(this.value)" class="w-full pl-8 pr-3 py-1.5 text-xs rounded-lg bg-slate-100 dark:bg-slate-800/80 border border-slate-300 dark:border-slate-700 focus:outline-none focus:border-blue-500 font-mono">
        <svg class="w-3.5 h-3.5 absolute left-2.5 top-2 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
      </div>
    </div>

    <!-- Mobile Sidebar Modules List -->
    <div class="mt-4 flex-1 overflow-y-auto pr-1">
      <div class="text-xs font-semibold text-slate-400 uppercase tracking-wider mb-2">Modules</div>
      <div id="mobile-sidebar-modules">
        ${sidebarModulesHtml}
      </div>
    </div>

    <!-- Mobile Footer Links -->
    <div class="pt-4 border-t border-slate-200 dark:border-slate-800 flex items-center justify-between text-xs">
      <a href="index.html" class="text-blue-600 dark:text-blue-400 font-medium">← Back to Main Guide</a>
      <a href="https://github.com/jsuyog2/node-pptx-templater" target="_blank" rel="noopener" class="text-slate-500 hover:text-blue-500">GitHub</a>
    </div>
  </div>

  <!-- Top Sticky Header -->
  <header class="sticky top-0 z-30 glass-nav">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
      <!-- Brand & Version -->
      <div class="flex items-center gap-3">
        <a href="index.html" class="flex items-center gap-2.5 group" title="Return to Main Guide">
          <img src="./assets/logo.png" alt="Logo" class="w-8 h-8 rounded-lg shadow-sm group-hover:scale-105 transition-transform">
          <span class="font-bold text-base sm:text-lg tracking-tight bg-gradient-to-r from-blue-600 to-indigo-600 dark:from-blue-400 dark:to-indigo-300 bg-clip-text text-transparent">node-pptx-templater</span>
        </a>
        <span class="hidden md:inline-flex text-xs text-slate-500 font-mono">v3.0.0</span>
      </div>

      <!-- Center / Desktop Navigation Links (Bi-directional routing) -->
      <nav class="hidden md:flex items-center gap-1 lg:gap-2">
        <!-- <a href="index.html" class="nav-link text-xs lg:text-sm text-blue-600 dark:text-blue-400 font-medium">
          ← Main Guide
        </a> -->
        <a href="index.html#getting-started" class="nav-link text-xs lg:text-sm">Getting Started</a>
        <a href="api-reference.html" class="nav-link text-xs lg:text-sm text-blue-600 dark:text-blue-400 font-semibold">API Reference</a>
        <a href="index.html#showcase" class="nav-link text-xs lg:text-sm">Showcase</a>
        <a href="index.html#templates" class="nav-link text-xs lg:text-sm">Templates</a>
        <a href="index.html#feature-catalog" class="nav-link text-xs lg:text-sm">Feature Catalog</a>
      </nav>

      <!-- Right Actions: Theme Toggle, GitHub, npm -->
      <div class="flex items-center gap-2">
        <!-- Quick Search Pill button -->
        <button onclick="focusSearchInput()" class="hidden sm:flex items-center gap-1.5 px-2.5 py-1.5 text-xs text-slate-500 dark:text-slate-400 bg-slate-200/60 dark:bg-slate-800/60 rounded-lg hover:text-slate-900 dark:hover:text-slate-100 transition-colors border border-slate-300/40 dark:border-slate-700/40">
          <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
          <span>Search</span>
          <kbd class="px-1 py-0.5 text-[10px] font-mono bg-white dark:bg-slate-900 border border-slate-300 dark:border-slate-700 rounded shadow-sm">/</kbd>
        </button>

        <button id="theme-toggle-btn" onclick="toggleTheme()" class="p-2 rounded-lg text-slate-500 hover:text-slate-900 dark:text-slate-400 dark:hover:text-slate-100 hover:bg-slate-200/60 dark:hover:bg-slate-800/60 transition-colors" aria-label="Toggle Light/Dark Theme">
          <svg id="theme-icon-light" class="w-5 h-5 hidden" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 3v1m0 16v1m9-9h-1M4 12H3m15.364 6.364l-.707-.707M6.343 6.343l-.707-.707m12.728 0l-.707.707M6.343 17.657l-.707.707M16 12a4 4 0 11-8 0 4 4 0 018 0z"></path></svg>
          <svg id="theme-icon-dark" class="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M20.354 15.354A9 9 0 018.646 3.646 9.003 9.003 0 0012 21a9.003 9.003 0 008.354-5.646z"></path></svg>
        </button>

        <a href="https://github.com/jsuyog2/node-pptx-templater" target="_blank" rel="noopener" class="p-2 rounded-lg text-slate-500 hover:text-slate-900 dark:text-slate-400 dark:hover:text-slate-100 hover:bg-slate-200/60 dark:hover:bg-slate-800/60 transition-colors" title="GitHub Repository">
          <svg class="w-5 h-5 fill-current" viewBox="0 0 24 24"><path fill-rule="evenodd" clip-rule="evenodd" d="M12 2C6.477 2 2 6.484 2 12.017c0 4.425 2.865 8.18 6.839 9.504.5.092.682-.217.682-.483 0-.237-.008-.868-.013-1.703-2.782.605-3.369-1.343-3.369-1.343-.454-1.158-1.11-1.466-1.11-1.466-.908-.62.069-.608.069-.608 1.003.07 1.53 1.032 1.53 1.032.892 1.53 2.341 1.088 2.91.832.092-.647.35-1.088.636-1.338-2.22-.253-4.555-1.113-4.555-4.951 0-1.093.39-1.988 1.029-2.688-.103-.253-.446-1.272.098-2.65 0 0 .84-.27 2.75 1.026A9.564 9.564 0 0112 6.844c.85.004 1.705.115 2.504.337 1.909-1.296 2.747-1.027 2.747-1.027.546 1.379.202 2.398.1 2.651.64.7 1.028 1.595 1.028 2.688 0 3.848-2.339 4.695-4.566 4.943.359.309.678.92.678 1.855 0 1.338-.012 2.419-.012 2.747 0 .268.18.58.688.482A10.019 10.019 0 0022 12.017C22 6.484 17.522 2 12 2z"></path></svg>
        </a>

        <a href="https://www.npmjs.com/package/node-pptx-templater" target="_blank" rel="noopener" class="p-2 rounded-lg text-slate-500 hover:text-slate-900 dark:text-slate-400 dark:hover:text-slate-100 hover:bg-slate-200/60 dark:hover:bg-slate-800/60 transition-colors" title="npm Package">
          <svg class="w-5 h-5 fill-current" viewBox="0 0 24 24"><path d="M1.763 0C.786 0 0 .786 0 1.763v20.474C0 23.214.786 24 1.763 24h20.474c.977 0 1.763-.786 1.763-1.763V1.763C24 .786 23.214 0 22.237 0zM5.13 5.13h13.74v13.74h-3.435V8.565h-3.435v10.305H5.13z"></path></svg>
        </a>
      </div>
    </div>
  </header>

  <!-- Main Container -->
  <main class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
    <div class="api-grid-layout">

      <!-- ====================================================================
           COLUMN 1: CATEGORIZED STICKY SIDEBAR (LEFT)
           ==================================================================== -->
      <aside id="api-sidebar" class="hidden md:block">
        <div class="sticky top-24 max-h-[calc(100vh-7rem)] flex flex-col glass-panel rounded-2xl p-4">
          <!-- Real-Time Search Bar -->
          <div class="mb-4 pb-3 border-b border-slate-200 dark:border-slate-800">
            <div class="relative">
              <input type="text" id="sidebar-search-input" placeholder="Search methods (/)..." oninput="handleSearch(this.value)" class="w-full pl-8 pr-7 py-2 text-xs rounded-xl bg-slate-100 dark:bg-slate-800/90 border border-slate-200 dark:border-slate-700/80 focus:outline-none focus:border-blue-500 focus:ring-1 focus:ring-blue-500 font-mono text-slate-900 dark:text-slate-100">
              <svg class="w-4 h-4 absolute left-2.5 top-2.5 text-slate-400" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 21l-6-6m2-5a7 7 0 11-14 0 7 7 0 0114 0z"></path></svg>
              <button id="search-clear-btn" onclick="clearSearch()" class="absolute right-2.5 top-2.5 text-slate-400 hover:text-slate-700 dark:hover:text-slate-200 text-xs" title="Clear search">✕</button>
            </div>
            <div class="mt-2 flex items-center justify-between text-[11px] text-slate-500">
              <span id="search-results-counter">95 methods available</span>
              <a href="#" onclick="expandAllCategories(); return false;" class="text-blue-500 hover:underline">Expand all</a>
            </div>
          </div>

          <!-- Modules List -->
          <div id="sidebar-modules-container" class="overflow-y-auto flex-1 pr-1 space-y-1">
            ${sidebarModulesHtml}
          </div>

          <!-- Sidebar Footer -->
          <div class="mt-auto pt-3 border-t border-slate-200 dark:border-slate-800 flex items-center justify-between text-[11px] text-slate-500">
            <a href="index.html" class="text-blue-500 hover:underline">← Main Guide</a>
            <span>v3.0.0-alpha.0</span>
          </div>
        </div>
      </aside>

      <!-- ====================================================================
           COLUMN 2: MAIN DOCUMENTATION FEED (CENTER)
           ==================================================================== -->
      <div id="api-content" class="min-w-0">
        <!-- Page Overview Header -->
        <div class="glass-panel rounded-2xl p-6 sm:p-8 mb-8">
          <div class="flex flex-wrap items-center justify-between gap-4">
            <div>
              <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full text-xs font-semibold bg-blue-500/10 text-blue-600 dark:text-blue-400 border border-blue-500/20 mb-3">
                <span>Production Reference</span>
                <span>•</span>
                <span>All 95 Public Methods</span>
              </div>
              <h1 class="text-3xl sm:text-4xl font-extrabold tracking-tight text-slate-900 dark:text-white">
                API Reference Manual
              </h1>
              <p class="mt-2 text-sm sm:text-base text-slate-600 dark:text-slate-300 max-w-2xl leading-relaxed">
                Complete technical specification for <code class="font-mono text-blue-600 dark:text-blue-400">node-pptx-templater</code>. Includes full method signatures, parameter tables with defaults, underlying OpenXML DrawingML/PresentationML mutations, and 1-click copyable examples.
              </p>
            </div>
            <div class="flex flex-col gap-2">
              <a href="index.html" class="inline-flex items-center justify-center gap-2 px-4 py-2 rounded-xl text-xs font-medium text-slate-700 dark:text-slate-200 bg-slate-200/60 dark:bg-slate-800/60 hover:bg-slate-200 dark:hover:bg-slate-800 transition-colors">
                <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M10 19l-7-7m0 0l7-7m-7 7h18"></path></svg>
                <span>Return to Main Guide</span>
              </a>
              <div class="text-[11px] text-slate-400 text-center font-mono">
                Pure JavaScript • Dual ESM/CJS
              </div>
            </div>
          </div>
        </div>

        <!-- Render All 10 Modules -->
        ${mainContentHtml}
      </div>

      <!-- ====================================================================
           COLUMN 3: ON THIS PAGE (RIGHT - WIDESCREEN > 1440px)
           ==================================================================== -->
      <aside id="right-toc" class="hidden xl:block">
        <div class="sticky top-24 max-h-[calc(100vh-7rem)] overflow-y-auto glass-panel rounded-2xl p-4">
          <div class="text-xs font-bold uppercase tracking-wider text-slate-400 mb-3 flex items-center justify-between">
            <span>On This Page</span>
            <span class="text-[10px] font-mono text-blue-500">10 Modules</span>
          </div>
          <div id="right-toc-container" class="space-y-4">
            ${rightTocModulesHtml}
          </div>
        </div>
      </aside>

    </div>
  </main>

  <!-- Footer -->
  <footer class="mt-20 border-t border-slate-200 dark:border-slate-800 py-10 bg-slate-100/50 dark:bg-slate-950/50">
    <div class="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 flex flex-col sm:flex-row items-center justify-between gap-4 text-xs text-slate-500">
      <div class="flex items-center gap-2">
        <img src="./assets/logo.png" alt="Logo" class="w-5 h-5 object-contain">
        <span>node-pptx-templater • MIT License</span>
      </div>
      <div class="flex items-center gap-4">
        <a href="index.html" class="hover:text-blue-500">Main Guide</a>
        <a href="api-reference.html" class="hover:text-blue-500 font-semibold text-blue-600 dark:text-blue-400">API Reference</a>
        <a href="https://github.com/jsuyog2/node-pptx-templater" target="_blank" rel="noopener" class="hover:text-blue-500">GitHub</a>
        <a href="https://www.npmjs.com/package/node-pptx-templater" target="_blank" rel="noopener" class="hover:text-blue-500">npm</a>
      </div>
    </div>
  </footer>

  <!-- ====================================================================
       INTERACTIVE CLIENT-SIDE JAVASCRIPT
       ==================================================================== -->
  <script>
    'use strict';

    // 1. One-Click Code Copy Handler (Spec Section 3)
    function copyCode(btn) {
      const container = btn.closest('.code-container');
      const codeEl = container ? container.querySelector('code') : null;
      if (!codeEl) return;

      const codeText = codeEl.innerText || codeEl.textContent;

      navigator.clipboard.writeText(codeText).then(() => {
        const originalHtml = btn.innerHTML;
        btn.classList.add('copied');
        btn.innerHTML = '<svg class="w-3.5 h-3.5 text-emerald-400 inline mr-1" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path></svg>Copied!';
        showToast('Code snippet copied to clipboard');
        setTimeout(() => {
          btn.classList.remove('copied');
          btn.innerHTML = originalHtml;
        }, 2000);
      }).catch(err => {
        console.error('Clipboard copy error:', err);
        showToast('Failed to copy snippet', 3000);
      });
    }

    // 2. Toast Notifications
    function showToast(message, duration = 2500) {
      const container = document.getElementById('toast-container');
      if (!container) return;

      const toast = document.createElement('div');
      toast.className = 'toast';
      toast.innerHTML = \`
        <svg class="w-4 h-4 text-emerald-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path>
        </svg>
        <span>\${message}</span>
      \`;

      container.appendChild(toast);
      setTimeout(() => {
        toast.style.opacity = '0';
        toast.style.transition = 'opacity 0.25s ease';
        setTimeout(() => toast.remove(), 250);
      }, duration);
    }

    // 3. Theme Management (Synchronized with localStorage)
    function applyTheme(theme) {
      const root = document.documentElement;
      const lightIcon = document.getElementById('theme-icon-light');
      const darkIcon = document.getElementById('theme-icon-dark');

      if (theme === 'dark') {
        root.classList.add('dark');
        if (lightIcon) lightIcon.classList.remove('hidden');
        if (darkIcon) darkIcon.classList.add('hidden');
      } else {
        root.classList.remove('dark');
        if (lightIcon) lightIcon.classList.add('hidden');
        if (darkIcon) darkIcon.classList.remove('hidden');
      }
    }

    function toggleTheme() {
      const isDark = document.documentElement.classList.contains('dark');
      const nextTheme = isDark ? 'light' : 'dark';
      localStorage.setItem('theme', nextTheme);
      applyTheme(nextTheme);
      showToast('Switched to ' + nextTheme + ' theme');
    }

    // 4. Real-time Client-Side Search Filtering
    function handleSearch(query) {
      const q = (query || '').toLowerCase().trim();
      const clearBtn = document.getElementById('search-clear-btn');
      if (clearBtn) clearBtn.style.display = q ? 'block' : 'none';

      const methodLinks = document.querySelectorAll('.sidebar-method-link');
      const categoryBlocks = document.querySelectorAll('.sidebar-category');
      const cards = document.querySelectorAll('.api-card');
      const counterEl = document.getElementById('search-results-counter');

      let visibleCount = 0;

      cards.forEach(card => {
        const name = (card.getAttribute('data-name') || '').toLowerCase();
        const desc = (card.getAttribute('data-desc') || '').toLowerCase();
        const matches = !q || name.includes(q) || desc.includes(q);

        if (matches) {
          card.classList.remove('hidden');
          visibleCount++;
        } else {
          card.classList.add('hidden');
        }
      });

      // Filter sidebar links
      methodLinks.forEach(link => {
        const methodName = (link.getAttribute('data-method') || '').toLowerCase();
        const matches = !q || methodName.includes(q);
        if (matches) {
          link.classList.remove('hidden');
        } else {
          link.classList.add('hidden');
        }
      });

      // Show/hide category groups based on visible children
      categoryBlocks.forEach(cat => {
        const visibleLinks = cat.querySelectorAll('.sidebar-method-link:not(.hidden)');
        const list = cat.querySelector('.category-methods');
        if (visibleLinks.length > 0) {
          cat.classList.remove('hidden');
          if (q && list) list.classList.remove('hidden');
        } else {
          cat.classList.add('hidden');
        }
      });

      if (counterEl) {
        counterEl.textContent = q ? \`\${visibleCount} matches found\` : '95 methods available';
      }
    }

    function clearSearch() {
      const searchInput = document.getElementById('sidebar-search-input');
      if (searchInput) {
        searchInput.value = '';
        handleSearch('');
      }
    }

    function handleMobileSearch(query) {
      handleSearch(query);
    }

    function focusSearchInput() {
      const input = document.getElementById('sidebar-search-input') || document.getElementById('mobile-search-input');
      if (input) {
        input.focus();
        input.select();
      }
    }

    // 5. Category Collapsible Toggles
    function toggleCategory(catId) {
      const list = document.getElementById('cat-list-' + catId);
      if (list) {
        list.classList.toggle('hidden');
      }
    }

    function expandAllCategories() {
      document.querySelectorAll('.category-methods').forEach(el => el.classList.remove('hidden'));
    }

    // 6. Mobile Drawer Toggle
    function toggleMobileMenu() {
      const drawer = document.getElementById('mobile-drawer');
      const backdrop = document.getElementById('mobile-backdrop');
      if (!drawer || !backdrop) return;

      const isClosed = drawer.classList.contains('-translate-x-full');
      if (isClosed) {
        drawer.classList.remove('-translate-x-full');
        backdrop.classList.remove('hidden');
        document.body.style.overflow = 'hidden';
      } else {
        closeMobileMenu();
      }
    }

    function closeMobileMenu() {
      const drawer = document.getElementById('mobile-drawer');
      const backdrop = document.getElementById('mobile-backdrop');
      if (drawer) drawer.classList.add('-translate-x-full');
      if (backdrop) backdrop.classList.add('hidden');
      document.body.style.overflow = '';
    }

    // 7. Active Scrollspy for Sidebar & Right TOC
    function setupScrollspy() {
      const cards = document.querySelectorAll('.api-card');
      const sidebarLinks = document.querySelectorAll('.sidebar-method-link');
      const rightTocLinks = document.querySelectorAll('.right-toc-link');

      function onScroll() {
        const scrollPos = window.scrollY + 140;
        let activeId = null;

        for (let i = cards.length - 1; i >= 0; i--) {
          const card = cards[i];
          if (card.offsetTop <= scrollPos) {
            activeId = card.id;
            break;
          }
        }

        if (activeId) {
          sidebarLinks.forEach(link => {
            if (link.getAttribute('href') === '#' + activeId) {
              link.classList.add('text-blue-600', 'dark:text-blue-400', 'font-semibold', 'bg-blue-500/10');
            } else {
              link.classList.remove('text-blue-600', 'dark:text-blue-400', 'font-semibold', 'bg-blue-500/10');
            }
          });

          rightTocLinks.forEach(link => {
            if (link.getAttribute('href') === '#' + activeId) {
              link.classList.add('text-blue-500', 'font-bold');
            } else {
              link.classList.remove('text-blue-500', 'font-bold');
            }
          });
        }
      }

      window.addEventListener('scroll', onScroll, { passive: true });
      onScroll();
    }

    // 8. Keyboard Shortcuts ('/' to search, Esc to clear/close)
    document.addEventListener('keydown', (e) => {
      if (e.key === '/' && document.activeElement.tagName !== 'INPUT' && document.activeElement.tagName !== 'TEXTAREA') {
        e.preventDefault();
        focusSearchInput();
      } else if (e.key === 'Escape') {
        clearSearch();
        closeMobileMenu();
      }
    });

    // 9. DOM Ready Initialization
    document.addEventListener('DOMContentLoaded', () => {
      const saved = localStorage.getItem('theme') || (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
      applyTheme(saved);
      setupScrollspy();

      // Close mobile drawer on link clicks
      document.querySelectorAll('#mobile-drawer a').forEach(a => {
        a.addEventListener('click', closeMobileMenu);
      });
    });
  </script>
</body>
</html>`;
}

// Function to render an individual Function Card
function renderFunctionCard(m, mod) {
  const isStatic = !!m.isStatic;
  const isAsync = !!m.isAsync;

  // Badges
  const staticBadge = isStatic ? `<span class="badge badge-info text-[11px] font-mono">static</span>` : '';
  const asyncBadge = isAsync ? `<span class="badge badge-success text-[11px] font-mono">async</span>` : '';
  const returnBadge = `<span class="badge badge-purple text-[11px] font-mono">${escapeHtml(m.returnType)}</span>`;

  // Parameters Table HTML
  let paramsTableHtml = '';
  if (m.params && m.params.length > 0) {
    paramsTableHtml = `
      <div class="mt-4">
        <div class="text-xs font-semibold text-slate-700 dark:text-slate-300 uppercase tracking-wider mb-2">Parameters</div>
        <div class="table-scroll-container rounded-xl border border-slate-200 dark:border-slate-800 bg-white/50 dark:bg-slate-900/40">
          <table class="w-full text-left text-xs">
            <thead class="bg-slate-100/70 dark:bg-slate-800/60 text-slate-600 dark:text-slate-300 border-b border-slate-200 dark:border-slate-800 font-semibold">
              <tr>
                <th class="py-2.5 px-3">Parameter</th>
                <th class="py-2.5 px-3">Type</th>
                <th class="py-2.5 px-3">Required</th>
                <th class="py-2.5 px-3">Default</th>
                <th class="py-2.5 px-3">Description</th>
              </tr>
            </thead>
            <tbody class="divide-y divide-slate-200/60 dark:divide-slate-800/60 text-slate-700 dark:text-slate-300">
              ${m.params.map(p => `
                <tr class="hover:bg-slate-100/40 dark:hover:bg-slate-800/40 transition-colors">
                  <td class="py-2 px-3 font-mono text-blue-600 dark:text-blue-400 font-medium">${escapeHtml(p.name)}</td>
                  <td class="py-2 px-3 font-mono text-purple-600 dark:text-purple-400 text-[11px]">${escapeHtml(p.type)}</td>
                  <td class="py-2 px-3">
                    ${p.required
        ? '<span class="inline-flex px-1.5 py-0.5 rounded text-[10px] font-semibold bg-rose-500/10 text-rose-600 dark:text-rose-400 border border-rose-500/20">Required</span>'
        : '<span class="inline-flex px-1.5 py-0.5 rounded text-[10px] font-semibold bg-slate-500/10 text-slate-500 dark:text-slate-400">Optional</span>'}
                  </td>
                  <td class="py-2 px-3 font-mono text-slate-500 text-[11px]">${escapeHtml(p.defaultVal)}</td>
                  <td class="py-2 px-3 text-slate-600 dark:text-slate-300 leading-relaxed">${escapeHtml(p.description)}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        </div>
      </div>
    `;
  } else {
    paramsTableHtml = `
      <div class="mt-3 text-xs text-slate-500 italic">None (no parameters required).</div>
    `;
  }

  // Returns Section HTML
  const returnsHtml = `
    <div class="mt-4 pt-3 border-t border-slate-200/70 dark:border-slate-800/70 text-xs">
      <span class="font-semibold text-slate-700 dark:text-slate-300">Returns: </span>
      <code class="font-mono text-purple-600 dark:text-purple-400 font-medium">${escapeHtml(m.returns.type)}</code>
      <span class="text-slate-600 dark:text-slate-400 ml-1.5">— ${escapeHtml(m.returns.description)}</span>
    </div>
  `;

  // Code Example HTML (Strictly complying with Section 3 snippet specification)
  const exampleCodeHighlighted = highlightJs(m.example);
  const codeSnippetHtml = `
    <div class="mt-4">
      <div class="flex items-center justify-between text-xs font-semibold text-slate-700 dark:text-slate-300 uppercase tracking-wider mb-2">
        <span>Usage Example</span>
        <span class="text-[11px] text-slate-400 font-mono font-normal">JavaScript</span>
      </div>
      <div class="code-container relative">
        <button class="copy-btn" onclick="copyCode(this)">Copy</button>
        <pre><code>${exampleCodeHighlighted}</code></pre>
      </div>
    </div>
  `;

  return `
    <div id="${m.name}" class="api-card glass-panel rounded-2xl p-5 sm:p-6" data-name="${m.name}" data-category="${mod.id}" data-desc="${escapeHtml(m.description)}">
      <!-- Card Header -->
      <div class="flex flex-wrap items-start justify-between gap-3 border-b border-slate-200 dark:border-slate-800 pb-3 mb-3">
        <div>
          <div class="flex flex-wrap items-center gap-2 mb-1.5">
            <h3 class="text-lg sm:text-xl font-bold font-mono text-slate-900 dark:text-white">
              <a href="#${m.name}" class="hover:text-blue-500 transition-colors">${m.name}</a>
            </h3>
            ${staticBadge}
            ${asyncBadge}
            ${returnBadge}
          </div>
          <div class="font-mono text-xs text-blue-600 dark:text-blue-400 break-all">
            ${escapeHtml(m.signature)}
          </div>
        </div>
        <a href="#${m.name}" class="text-slate-400 hover:text-blue-500 p-1" title="Direct link to ${m.name}">
          <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13.828 10.172a4 4 0 00-5.656 0l-4 4a4 4 0 105.656 5.656l1.102-1.101m-.758-4.899a4 4 0 005.656 0l4-4a4 4 0 00-5.656-5.656l-1.1 1.1"></path></svg>
        </a>
      </div>

      <!-- Card Description -->
      <p class="text-xs sm:text-sm text-slate-600 dark:text-slate-300 leading-relaxed">
        ${m.description}
      </p>

      <!-- OpenXML Under the Hood Mutation Badge -->
      <div class="mt-2.5 flex items-start gap-2 p-2.5 rounded-lg bg-slate-100/80 dark:bg-slate-800/50 border border-slate-200/80 dark:border-slate-800 text-[11px] text-slate-600 dark:text-slate-400">
        <span class="font-semibold text-slate-700 dark:text-slate-300 whitespace-nowrap">OpenXML Mutation:</span>
        <span class="font-mono text-slate-500 dark:text-slate-400">${m.openXml}</span>
      </div>

      <!-- Parameters Table -->
      ${paramsTableHtml}

      <!-- Return Type -->
      ${returnsHtml}

      <!-- Copyable Example -->
      ${codeSnippetHtml}
    </div>
  `;
}

// Sitemap and robots.txt generators
function generateSitemap() {
  const domain = 'https://node-pptx-templater.pages.dev';
  const today = new Date().toISOString().split('T')[0];
  return `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
  <url>
    <loc>${domain}/</loc>
    <lastmod>${today}</lastmod>
    <changefreq>weekly</changefreq>
    <priority>1.0</priority>
  </url>
  <url>
    <loc>${domain}/api-reference.html</loc>
    <lastmod>${today}</lastmod>
    <changefreq>weekly</changefreq>
    <priority>0.9</priority>
  </url>
</urlset>
`;
}

function generateRobotsTxt() {
  return `User-agent: *
Allow: /

Sitemap: https://node-pptx-templater.pages.dev/sitemap.xml
`;
}

// Build and write out file
function build() {
  console.log('Generating production API reference for node-pptx-templater...');
  const html = generateHtml();
  fs.writeFileSync(OUTPUT_HTML_PATH, html, 'utf8');
  console.log(`✓ Successfully generated docs/api-reference.html (${(html.length / 1024).toFixed(1)} KB)`);

  const sitemapXml = generateSitemap();
  fs.writeFileSync(SITEMAP_PATH, sitemapXml, 'utf8');
  console.log(`✓ Successfully generated docs/sitemap.xml`);

  const robotsTxt = generateRobotsTxt();
  fs.writeFileSync(ROBOTS_PATH, robotsTxt, 'utf8');
  console.log(`✓ Successfully generated docs/robots.txt`);
}

if (require.main === module) {
  build();
}

module.exports = { build };
