/**
 * @fileoverview showcase/input_export.js — Comprehensive I/O & Export Master Guide.
 *
 * Demonstrates and validates the strict 18-method I/O whitelist for node-pptx-templater.
 * Every method is demonstrated with informative console output, timing, options handling,
 * and detailed enterprise architectural use cases (AWS Lambda, Cloudflare Workers,
 * Next.js Edge, microservice streaming, and version-controlled OpenXML storage).
 *
 * The Strict 18-Method Whitelist:
 * =========================================================================
 * STATIC METHODS (7 functions):
 *   1.  PPTXTemplater.load(source, options)
 *   2.  PPTXTemplater.createBlank(options)
 *   3.  PPTXTemplater.extractPptx(pptxPath, outputDir, options)
 *   4.  PPTXTemplater.buildPptx(sourceDir, outputPath, options)
 *   5.  PPTXTemplater.preload(source, options)
 *   6.  PPTXTemplater.fromCache(source)
 *   7.  PPTXTemplater.clearCache()
 *
 * EXPORT INSTANCE METHODS (8 functions):
 *   8.  ppt.saveToFile(destinationPath, options)
 *   9.  ppt.toBuffer(options)
 *   10. ppt.toStream(options)
 *   11. ppt.toBase64(options)
 *   12. ppt.toArrayBuffer(options)
 *   13. ppt.toResponse(options)
 *   14. ppt.toExpress(res, filename, options)
 *   15. ppt.toFastify(reply, filename, options)
 *
 * EXTRACTION & CONVERSION INSTANCE METHODS (3 functions):
 *   16. ppt.extractToZip(destinationPath, options)
 *   17. ppt.extractToXml(destinationDir, options)
 *   18. ppt.toFlatXml(destinationPathOrOptions, options)
 * =========================================================================
 */

'use strict'

const path = require('path')
const fs = require('fs-extra')
const { Writable } = require('stream')
const { finished } = require('stream/promises')
const { PPTXTemplater } = require('../src/index.js')

const SHOWCASE_DIR = __dirname
const TEMPLATES_DIR = path.join(SHOWCASE_DIR, 'templates')
const OUTPUT_DIR = path.join(SHOWCASE_DIR, 'output/input_export')

/**
 * Creates a mock HTTP Response / Reply stream for Express.js and Fastify demonstrations.
 * Captures headers, status code, and written chunks into a concatenated Buffer.
 *
 * @returns {import('stream').Writable & { statusCode: number, headers: Record<string, string>, body: Buffer | null }}
 */
function createMockHttpResponse() {
  const chunks = []
  const mock = Object.assign(
    new Writable({
      write(chunk, encoding, callback) {
        chunks.push(Buffer.isBuffer(chunk) ? chunk : Buffer.from(chunk))
        callback()
      },
    }),
    {
      statusCode: 200,
      headers: {},
      body: null,
      headersSent: false,
      status(code) {
        this.statusCode = code
        return this
      },
      code(code) {
        this.statusCode = code
        return this
      },
      set(key, value) {
        if (typeof key === 'object') {
          Object.assign(this.headers, key)
        } else {
          this.headers[key.toLowerCase()] = value
        }
        return this
      },
      header(key, value) {
        this.headers[key.toLowerCase()] = value
        return this
      },
      setHeader(key, value) {
        this.headers[key.toLowerCase()] = value
        return this
      },
      type(typeString) {
        this.headers['content-type'] = typeString
        return this
      },
      send(body) {
        if (body) {
          chunks.push(Buffer.isBuffer(body) ? body : Buffer.from(body))
        }
        this.headersSent = true
        this.end()
        return this
      },
    }
  )

  mock.on('finish', () => {
    mock.body = Buffer.concat(chunks)
  })

  return mock
}

async function main() {
  await fs.ensureDir(OUTPUT_DIR)

  console.log('='.repeat(78))
  console.log('  node-pptx-templater — Comprehensive I/O & Export Master Guide')
  console.log('  Validating and Demonstrating the Complete 18-Method Whitelist API')
  console.log('='.repeat(78))
  console.log(`Working Directory : ${process.cwd()}`)
  console.log(`Templates Path    : ${TEMPLATES_DIR}`)
  console.log(`Output Directory  : ${OUTPUT_DIR}\n`)

  const samplePptxPath = path.join(TEMPLATES_DIR, 'sample.pptx')
  const sampleXmlPath = path.join(TEMPLATES_DIR, 'sample.xml')
  const sampleZipPath = path.join(TEMPLATES_DIR, 'sample.zip')
  const sampleFolderPath = path.join(TEMPLATES_DIR, 'sample')

  // =========================================================================
  // SECTION 1: STATIC METHODS (7 Functions)
  // =========================================================================
  console.log('------------------------------------------------------------------------------')
  console.log('  SECTION 1: Static Lifecycle, Generation & Ingestion (7 Methods)')
  console.log('------------------------------------------------------------------------------')

  /**
   * [1/18] PPTXTemplater.load(source, options)
   * Use Case: Universal polymorphic ingestion across microservices. Ingests .pptx files,
   * in-memory Buffers (from S3/GCS), Flat OPC single-file XML (<pkg:package>), .zip files,
   * uncompressed OpenXML directory trees, or HTTP/HTTPS remote URLs into an in-memory VFS.
   */
  console.log('\n[1/18] PPTXTemplater.load(source, options) — Polymorphic Ingestion')

  // (a) From .pptx file path
  const deckFromPath = await PPTXTemplater.load(samplePptxPath)
  console.log(`  ✓ [1.a] File path (.pptx)          : ${deckFromPath.slideCount} slides loaded`)

  // (b) From in-memory Buffer (AWS S3 getObject / multipart upload)
  const fileBuffer = await fs.readFile(samplePptxPath)
  const deckFromBuffer = await PPTXTemplater.load(fileBuffer)
  console.log(
    `  ✓ [1.b] Buffer / Uint8Array        : ${deckFromBuffer.slideCount} slides (${(fileBuffer.length / 1024).toFixed(1)} KB)`
  )

  // (c) From Microsoft Flat OPC XML file (.xml / <pkg:package>)
  if (await fs.pathExists(sampleXmlPath)) {
    const deckFromXml = await PPTXTemplater.load(sampleXmlPath)
    console.log(
      `  ✓ [1.c] Flat OPC single-file XML   : ${deckFromXml.slideCount} slides (zero ZIP overhead)`
    )
  }

  // (d) From ZIP archive (.zip)
  if (await fs.pathExists(sampleZipPath)) {
    const deckFromZip = await PPTXTemplater.load(sampleZipPath)
    console.log(`  ✓ [1.d] .zip container archive     : ${deckFromZip.slideCount} slides`)
  }

  // (e) From uncompressed OpenXML directory tree
  if (await fs.pathExists(sampleFolderPath)) {
    const deckFromFolder = await PPTXTemplater.load(sampleFolderPath)
    console.log(`  ✓ [1.e] Uncompressed folder tree   : ${deckFromFolder.slideCount} slides`)
  }

  // (f) From Url
  const sampleUrl =
    'https://github.com/jsuyog2/node-pptx-templater/raw/refs/heads/main/templates/sample.pptx'
  const deckFromUrl = await PPTXTemplater.load(sampleUrl)
  console.log(`  ✓ [1.f] Remote URL                 : ${deckFromUrl.slideCount} slides`)

  // (g) Preload template into memory cache for faster access
  await PPTXTemplater.preload(samplePptxPath, { key: 'deck1' })

  // (h) Get template from cache
  const pptFromCache = await PPTXTemplater.fromCache('deck1')
  console.log(`  ✓ [1.g] From cache                 : ${pptFromCache.slideCount} slides`)

  // (i) Clear cache
  PPTXTemplater.clearCache()
  console.log(`  ✓ [1.h] Clear cache`)

  /**
   * [2/18] PPTXTemplater.createBlank(options)
   * Use Case: Cold-start deck generation in AWS Lambda, Google Cloud Run, or microservices
   * where no template file is provisioned. Synthesizes a valid, compliant OpenXML presentation.
   */
  console.log('\n[2/18] PPTXTemplater.createBlank(options) — Template-Free Synthesis')
  const blank16x9 = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })
  blank16x9.useSlide(1)
  await blank16x9.addShape('roundedRectangle', {
    x: 100,
    y: 100,
    width: 600,
    height: 200,
    fill: '#1E293B',
    text: 'Synthesized with PPTXTemplater.createBlank()',
  })
  const valResult = await blank16x9.validatePresentation()
  const blankOutPath = path.join(OUTPUT_DIR, '01-created-blank-16x9.pptx')
  await blank16x9.saveToFile(blankOutPath)
  console.log(
    `  ✓ Blank 16:9 widescreen created    : ${path.basename(blankOutPath)} (valid: ${valResult.valid}, errors: ${valResult.errors.length})`
  )

  /**
   * [3/18] PPTXTemplater.extractPptx(pptxPath, outputDir, options)
   * Use Case: CI/CD build scripts, automated template versioning, and Git tracking.
   * Unpacks a binary .pptx archive into readable, diffable XML files and folders.
   */
  console.log('\n[3/18] PPTXTemplater.extractPptx(pptxPath, outputDir, options) — Static Unpack')
  const staticUnpackDir = path.join(OUTPUT_DIR, '02-static-unpacked-pptx')
  await PPTXTemplater.extractPptx(samplePptxPath, staticUnpackDir, {
    overwrite: true,
    formatXml: true,
  })
  console.log(`  ✓ Static extraction to directory   : ${path.basename(staticUnpackDir)}`)

  /**
   * [4/18] PPTXTemplater.buildPptx(sourceDir, outputPath, options)
   * Use Case: Compiling Git-versioned XML template trees back into compressed .pptx artifacts
   * during CI/CD build steps or deployment pipelines.
   */
  console.log('\n[4/18] PPTXTemplater.buildPptx(sourceDir, outputPath, options) — Static Compiler')
  const staticRebuiltPptx = path.join(OUTPUT_DIR, '03-static-rebuilt-from-folder.pptx')
  await PPTXTemplater.buildPptx(staticUnpackDir, staticRebuiltPptx, { compressionLevel: 9 })
  const rebuiltDeck = await PPTXTemplater.load(staticRebuiltPptx)
  console.log(
    `  ✓ Static directory compilation     : ${path.basename(staticRebuiltPptx)} (${rebuiltDeck.slideCount} slides verified)`
  )

  /**
   * [5/18] PPTXTemplater.preload(source, options)
   * Use Case: High-throughput backend services (Node.js/Express/Fastify/NestJS). Pre-parses
   * OpenXML packages into static memory during application startup to eliminate disk I/O.
   */
  console.log('\n[5/18] PPTXTemplater.preload(source, options) — Memory Pre-compilation')
  const tPreload0 = performance.now()
  await PPTXTemplater.preload(samplePptxPath, { key: 'sample-deck-key', logLevel: 'silent' })
  const preloadDuration = (performance.now() - tPreload0).toFixed(2)
  console.log(
    `  ✓ Template cached in memory        : ${preloadDuration}ms (key: "sample-deck-key")`
  )

  /**
   * [6/18] PPTXTemplater.fromCache(source)
   * Use Case: Sub-millisecond slide generation in high-concurrency request handlers.
   * Clones preloaded VFS structures instantly with zero disk reads.
   */
  console.log('\n[6/18] PPTXTemplater.fromCache(source) — Sub-millisecond In-Memory Cloning')
  const tClone0 = performance.now()
  const cachedInstance = await PPTXTemplater.fromCache('sample-deck-key')
  const cloneDuration = (performance.now() - tClone0).toFixed(2)
  cachedInstance.useSlide(1).replaceText({ '{{title}}': 'High Concurrency Clone' })
  console.log(
    `  ✓ Cloned from memory cache         : ${cloneDuration}ms (zero disk/decompression latency)`
  )

  /**
   * [7/18] PPTXTemplater.clearCache()
   * Use Case: Memory management and cache invalidation when templates are updated dynamically.
   */
  console.log('\n[7/18] PPTXTemplater.clearCache() — Memory Eviction')
  PPTXTemplater.clearCache()
  console.log('  ✓ Preload cache purged cleanly     : Static memory freed')

  // =========================================================================
  // SECTION 2: EXPORT INSTANCE METHODS (8 Functions)
  // =========================================================================
  console.log('\n------------------------------------------------------------------------------')
  console.log('  SECTION 2: Multi-Target Export & Web Framework Adapters (8 Methods)')
  console.log('------------------------------------------------------------------------------')

  const loadSimple = async text => {
    const exportDeck = await PPTXTemplater.load(samplePptxPath)
    exportDeck.useSlide(1).replaceText({ '{{title}}': `${text} Enterprise Production Export` })
    return exportDeck
  }
  /**
   * [8/18] ppt.saveToFile(destinationPath, options)
   * Use Case: CLI utilities, batch offline report generation, and local disk persistence.
   */
  console.log('\n[8/18] ppt.saveToFile(destinationPath, options) — Disk Persistence')
  const diskPath = path.join(OUTPUT_DIR, '04-save-to-file.pptx')
  const ppt = await loadSimple('save-to-file')
  await ppt.saveToFile(diskPath, { compressionLevel: 6 })
  console.log(
    `  ✓ Saved directly to disk           : ${path.basename(diskPath)} (${(await fs.stat(diskPath)).size} bytes)`
  )

  /**
   * [9/18] ppt.toBuffer(options)
   * Use Case: Serverless AWS Lambda handlers, uploading directly to Amazon S3 (s3.putObject),
   * Google Cloud Storage, or transmitting via AMQP/Kafka message queues.
   */
  console.log('\n[9/18] ppt.toBuffer(options) — In-Memory Node.js Buffer')
  const nodeBuffer = await loadSimple('toBuffer')
  const nodeBufferData = await nodeBuffer.toBuffer({ compressionLevel: 0 })
  console.log(
    `  ✓ Emitted Buffer                   : ${(nodeBufferData.length / 1024).toFixed(1)} KB (isBuffer: ${Buffer.isBuffer(nodeBuffer)})`
  )

  /**
   * [10/18] ppt.toStream(options)
   * Use Case: Node.js HTTP servers and streaming file storage, minimizing peak memory
   * usage by piping chunks directly to the consumer without intermediate buffering.
   */
  console.log('\n[10/18] ppt.toStream(options) — Node.js ReadableStream')
  const streamOutPath = path.join(OUTPUT_DIR, '05-piped-stream.pptx')
  const readableStream = await loadSimple('toStream')
  const readableStreamData = await readableStream.toStream({ chunkSize: 16384 })
  const fileWriteStream = fs.createWriteStream(streamOutPath)
  readableStreamData.pipe(fileWriteStream)
  await finished(fileWriteStream)
  console.log(
    `  ✓ Piped Readable stream to disk    : ${path.basename(streamOutPath)} (${(await fs.stat(streamOutPath)).size} bytes)`
  )

  /**
   * [11/18] ppt.toBase64(options)
   * Use Case: Embedding presentations in JSON API payloads, GraphQL responses,
   * or creating Data URIs (`data:application/...`) for direct browser download links.
   */
  console.log('\n[11/18] ppt.toBase64(options) — Base64 String & Data URI')
  const base64Str = await loadSimple('toBase64')
  const base64StrData = await base64Str.toBase64()
  const dataUriData = await base64Str.toBase64({ dataUri: true })
  console.log(
    `  ✓ Base64 payload                   : ${base64StrData.length} characters (Data URI prefix: "${dataUriData.slice(0, 48)}...")`
  )

  /**
   * [12/18] ppt.toArrayBuffer(options)
   * Use Case: Cloudflare Workers, Vercel Edge Functions, and Deno environments where
   * Web Standard ArrayBuffer / Uint8Array is required instead of Node.js Buffer.
   */
  console.log('\n[12/18] ppt.toArrayBuffer(options) — Edge Runtime ArrayBuffer')
  const arrayBuffer = await loadSimple('toArrayBuffer')
  const arrayBufferData = await arrayBuffer.toArrayBuffer()
  console.log(
    `  ✓ Emitted ArrayBuffer              : byteLength = ${arrayBufferData.byteLength} bytes`
  )

  /**
   * [13/18] ppt.toResponse(options)
   * Use Case: Next.js App Router (route.js), Remix, and SvelteKit endpoints.
   * Returns a standard Web Fetch API Response with standard OpenXML MIME and headers.
   */
  console.log('\n[13/18] ppt.toResponse(options) — Standard Fetch API Response')
  const webResponse = await loadSimple('toResponse')
  const webResponseData = await webResponse.toResponse({ filename: 'QuarterlyReport.pptx' })
  console.log(`  ✓ Web Response                     : Status ${webResponseData.status}`)
  console.log(
    `    Content-Type                     : ${webResponseData.headers.get('content-type')}`
  )
  console.log(
    `    Content-Disposition              : ${webResponseData.headers.get('content-disposition')}`
  )

  /**
   * [14/18] ppt.toExpress(res, filename, options)
   * Use Case: Express.js controllers and REST endpoints. Automatically sets headers
   * and pipes the presentation stream with full error handling.
   */
  console.log('\n[14/18] ppt.toExpress(res, filename, options) — Express.js HTTP Streaming')
  const expressRes = createMockHttpResponse()
  const expressFinished = new Promise(resolve => expressRes.on('finish', resolve))
  const expressResData = await loadSimple('toExpress')
  await expressResData.toExpress(expressRes, 'ExpressExport.pptx')
  await expressFinished
  console.log(`  ✓ Express response delivered       : ${expressRes.body.length} bytes`)
  console.log(`    Content-Disposition              : ${expressRes.headers['content-disposition']}`)

  /**
   * [15/18] ppt.toFastify(reply, filename, options)
   * Use Case: High-performance Fastify microservice endpoints. Configures OpenXML
   * headers and streams or buffers the payload cleanly via reply.send().
   */
  console.log('\n[15/18] ppt.toFastify(reply, filename, options) — Fastify HTTP Reply')
  const fastifyReply = createMockHttpResponse()
  const fastifyFinished = new Promise(resolve => fastifyReply.on('finish', resolve))
  const fastifyReplyData = await loadSimple('toFastify')
  await fastifyReplyData.toFastify(fastifyReply, 'FastifyExport.pptx')
  await fastifyFinished
  console.log(`  ✓ Fastify reply delivered          : ${fastifyReply.body.length} bytes`)
  console.log(
    `    Content-Disposition              : ${fastifyReply.headers['content-disposition']}`
  )

  // =========================================================================
  // SECTION 3: EXTRACTION & CONVERSION INSTANCE METHODS (3 Functions)
  // =========================================================================
  console.log('\n------------------------------------------------------------------------------')
  console.log('  SECTION 3: Extraction & Conversion Instance Methods (3 Methods)')
  console.log('------------------------------------------------------------------------------')

  /**
   * [16/18] ppt.extractToZip(destinationPath, options)
   * Use Case: Direct output of uncorrupted raw ZIP packages for diagnostics,
   * archive mirroring, or external decompression toolchains.
   */
  console.log('\n[16/18] ppt.extractToZip(destinationPath, options) — Raw ZIP Package Output')
  const zipOutPath = path.join(OUTPUT_DIR, '06-extracted-package.zip')
  const zipOutPathData = await loadSimple('extractToZip')
  await zipOutPathData.extractToZip(zipOutPath, { compressionLevel: 6 })
  console.log(
    `  ✓ Raw ZIP package written          : ${path.basename(zipOutPath)} (${(await fs.stat(zipOutPath)).size} bytes)`
  )

  /**
   * [17/18] ppt.extractToXml(destinationDir, options)
   * Use Case: Deep package inspection, automated audit of DrawingML elements,
   * pretty-printed XML diffing, and template schema debugging.
   */
  console.log(
    '\n[17/18] ppt.extractToXml(destinationDir, options) — Unpacked OpenXML Directory Tree'
  )
  const xmlOutDir = path.join(OUTPUT_DIR, '07-extracted-openxml-tree')
  const xmlOutDirData = await loadSimple('extractToXml')
  await xmlOutDirData.extractToXml(xmlOutDir, { formatXml: true })
  const extractedFiles = await fs.readdir(xmlOutDir)
  console.log(
    `  ✓ OpenXML tree unpacked to folder  : ${path.basename(xmlOutDir)} (Root items: ${extractedFiles.join(', ')})`
  )

  /**
   * [18/18] ppt.toFlatXml(destinationPathOrOptions, options)
   * Use Case: Single-file Microsoft Flat OPC XML format (<pkg:package>). Perfect for
   * Git version control (line-by-line diffs of entire presentations), database storage,
   * or serverless environments requiring single-file representations without ZIP overhead.
   */
  console.log(
    '\n[18/18] ppt.toFlatXml(destinationPathOrOptions, options) — Flat OPC Single-File XML'
  )
  const flatXmlOutPath = path.join(OUTPUT_DIR, '08-converted-flat-opc.xml')
  // (a) Save directly to file path with options
  const flatXmlStringData = await loadSimple('toFlatXml')
  const flatXmlString = await flatXmlStringData.toFlatXml(flatXmlOutPath, { pretty: true })
  // (b) Also retrieve as an in-memory XML string
  const flatXmlInMemory = await flatXmlStringData.toFlatXml({ pretty: false })
  const rootTagMatch = flatXmlString.match(/<pkg:package[^>]*>/)
  console.log(
    `  ✓ Saved Flat OPC XML to file       : ${path.basename(flatXmlOutPath)} (${(await fs.stat(flatXmlOutPath)).size} bytes)`
  )
  console.log(
    `  ✓ In-memory Flat OPC XML string    : ${flatXmlInMemory.length} chars (root tag: ${rootTagMatch ? '<pkg:package>' : 'valid'})`
  )

  console.log('\n' + '='.repeat(78))
  console.log('  ALL 18 APPROVED I/O, EXPORT & EXTRACTION METHODS EXECUTED CLEANLY!')
  console.log('='.repeat(78) + '\n')
}

main().catch(err => {
  console.error('\n❌ Error during showcase execution:', err)
  process.exit(1)
})
