/**
 * @fileoverview Comparative performance benchmark: PPTX (ZIP) vs XML Presentation Folder.
 *
 * Measures:
 *  1. Startup/Load Time
 *  2. Memory Usage (Heap allocated)
 *  3. Latency (Load -> Modify -> Export)
 *  4. Concurrent Request processing
 *
 * Run: node benchmarks/xml-vs-zip.js
 */

const { performance } = require('perf_hooks')
const path = require('path')
const fs = require('fs-extra')
const JSZip = require('jszip')
const { PPTXTemplater } = require('../src/index.js')

const RESULTS_DIR = path.resolve(__dirname, 'results')
const ZIP_TEMPLATE = path.resolve(__dirname, '../tests/fixtures/sample.pptx')
const TEMP_DIR = path.resolve(__dirname, 'temp-xml-bench')
const XML_TEMPLATE_DIR = path.resolve(TEMP_DIR, 'sample-extracted')

// Ensure temp directory exists and extract template
async function setup() {
  if (!fs.existsSync(RESULTS_DIR)) fs.mkdirSync(RESULTS_DIR, { recursive: true })
  await fs.emptyDir(TEMP_DIR)

  if (!fs.existsSync(ZIP_TEMPLATE)) {
    throw new Error(`Template not found at ${ZIP_TEMPLATE}. Run test setup first.`)
  }

  // Extract ZIP template to folder
  const data = await fs.readFile(ZIP_TEMPLATE)
  const zip = await JSZip.loadAsync(data)
  for (const [filename, file] of Object.entries(zip.files)) {
    if (file.dir) continue
    const destPath = path.resolve(XML_TEMPLATE_DIR, filename)
    await fs.ensureDir(path.dirname(destPath))
    const content = await file.async('nodebuffer')
    await fs.writeFile(destPath, content)
  }
}

async function cleanup() {
  try {
    await new Promise(resolve => setTimeout(resolve, 500))
    await fs.remove(TEMP_DIR)
  } catch (err) {
    // Ignore cleanup errors on busy files in temp directory
  }
}

function getMemoryUsage() {
  if (global.gc) {
    global.gc()
  }
  return process.memoryUsage().heapUsed / 1024 / 1024 // MB
}

async function runBenchmarks() {
  console.log('⚡ PowerPoint XML Folder vs ZIP Template Benchmarks')
  console.log('===================================================')
  
  await setup()

  const results = {
    zip: {},
    xml: {},
  }

  // ── 1. Startup/Load Time ───────────────────────────────────────────
  console.log('\n⏱️  Running startup (load template) benchmarks...')
  
  // Warmup
  for (let i = 0; i < 5; i++) {
    await PPTXTemplater.load(ZIP_TEMPLATE)
    await PPTXTemplater.load(XML_TEMPLATE_DIR)
  }

  const loadZipStart = performance.now()
  for (let i = 0; i < 50; i++) {
    await PPTXTemplater.load(ZIP_TEMPLATE)
  }
  const loadZipTime = (performance.now() - loadZipStart) / 50
  results.zip.loadTimeMs = +loadZipTime.toFixed(3)

  const loadXmlStart = performance.now()
  for (let i = 0; i < 50; i++) {
    await PPTXTemplater.load(XML_TEMPLATE_DIR)
  }
  const loadXmlTime = (performance.now() - loadXmlStart) / 50
  results.xml.loadTimeMs = +loadXmlTime.toFixed(3)

  console.log(`   ZIP Template Load Time: ${results.zip.loadTimeMs} ms`)
  console.log(`   XML Folder Load Time:   ${results.xml.loadTimeMs} ms`)
  console.log(`   ⚡ Speedup:             ${((results.zip.loadTimeMs / results.xml.loadTimeMs)).toFixed(1)}x`)

  // ── 2. Memory Usage ────────────────────────────────────────────────
  console.log('\n💾 Measuring heap memory consumption...')
  
  const memBefore = getMemoryUsage()
  const zipTemplates = []
  for (let i = 0; i < 20; i++) {
    zipTemplates.push(await PPTXTemplater.load(ZIP_TEMPLATE))
  }
  const memAfterZip = getMemoryUsage()
  results.zip.memoryUsedMb = +(memAfterZip - memBefore).toFixed(3)
  
  // Clear references
  zipTemplates.length = 0
  if (global.gc) global.gc()

  const memBeforeXml = getMemoryUsage()
  const xmlTemplates = []
  for (let i = 0; i < 20; i++) {
    xmlTemplates.push(await PPTXTemplater.load(XML_TEMPLATE_DIR))
  }
  const memAfterXml = getMemoryUsage()
  results.xml.memoryUsedMb = +(memAfterXml - memBeforeXml).toFixed(3)

  // Clear references
  xmlTemplates.length = 0
  if (global.gc) global.gc()

  console.log(`   ZIP heap overhead (20 instances): ${results.zip.memoryUsedMb} MB`)
  console.log(`   XML heap overhead (20 instances): ${results.xml.memoryUsedMb} MB`)
  console.log(`   ⚡ Memory Savings:                ${(results.zip.memoryUsedMb - results.xml.memoryUsedMb).toFixed(2)} MB`)

  // ── 3. Latency (Load -> Modify -> Export) ──────────────────────────
  console.log('\n🔄 Measuring end-to-end processing latency...')

  const zipLatStart = performance.now()
  for (let i = 0; i < 20; i++) {
    const ppt = await PPTXTemplater.load(ZIP_TEMPLATE)
    ppt.useSlide(1).replaceText({ '{{title}}': `Update ${i}` })
    await ppt.save(path.resolve(TEMP_DIR, `out-zip-${i}.pptx`))
  }
  const zipLatTime = (performance.now() - zipLatStart) / 20
  results.zip.latencyMs = +zipLatTime.toFixed(3)

  const xmlLatStart = performance.now()
  for (let i = 0; i < 20; i++) {
    const ppt = await PPTXTemplater.load(XML_TEMPLATE_DIR)
    ppt.useSlide(1).replaceText({ '{{title}}': `Update ${i}` })
    await ppt.saveToFolder(path.resolve(TEMP_DIR, `out-folder-${i}`))
  }
  const xmlLatTime = (performance.now() - xmlLatStart) / 20
  results.xml.latencyMs = +xmlLatTime.toFixed(3)

  console.log(`   ZIP (Load -> Save ZIP) Latency:   ${results.zip.latencyMs} ms`)
  console.log(`   XML (Load -> Save Folder) Latency: ${results.xml.latencyMs} ms`)
  console.log(`   ⚡ Speedup:                        ${(results.zip.latencyMs / results.xml.latencyMs).toFixed(1)}x`)

  // ── 4. Concurrent Requests ──────────────────────────────────────────
  console.log('\n🚀 Measuring concurrency throughput (10 parallel tasks)...')

  const runConcurrentZip = async (runIdx) => {
    const tasks = Array.from({ length: 10 }, async (_, i) => {
      const ppt = await PPTXTemplater.load(ZIP_TEMPLATE)
      ppt.useSlide(1).replaceText({ '{{title}}': `Concurrent ${runIdx}-${i}` })
      await ppt.save(path.resolve(TEMP_DIR, `concurrent-${runIdx}-${i}.pptx`))
    })
    const start = performance.now()
    await Promise.all(tasks)
    return performance.now() - start
  }

  const runConcurrentXml = async (runIdx) => {
    const tasks = Array.from({ length: 10 }, async (_, i) => {
      const ppt = await PPTXTemplater.load(XML_TEMPLATE_DIR)
      ppt.useSlide(1).replaceText({ '{{title}}': `Concurrent ${runIdx}-${i}` })
      await ppt.saveToFolder(path.resolve(TEMP_DIR, `concurrent-xml-${runIdx}-${i}`))
    })
    const start = performance.now()
    await Promise.all(tasks)
    return performance.now() - start
  }

  // Warmup
  await runConcurrentZip('warmup')
  await runConcurrentXml('warmup')

  const zipConTimes = []
  for (let i = 0; i < 5; i++) {
    zipConTimes.push(await runConcurrentZip(i))
  }
  results.zip.concurrentMs = +(zipConTimes.reduce((a, b) => a + b, 0) / 5).toFixed(3)

  const xmlConTimes = []
  for (let i = 0; i < 5; i++) {
    xmlConTimes.push(await runConcurrentXml(i))
  }
  results.xml.concurrentMs = +(xmlConTimes.reduce((a, b) => a + b, 0) / 5).toFixed(3)

  console.log(`   ZIP Concurrency Latency (10 requests): ${results.zip.concurrentMs} ms`)
  console.log(`   XML Concurrency Latency (10 requests): ${results.xml.concurrentMs} ms`)
  console.log(`   ⚡ Speedup:                            ${(results.zip.concurrentMs / results.xml.concurrentMs).toFixed(1)}x`)

  // ── Save Results ───────────────────────────────────────────────────
  const output = {
    timestamp: new Date().toISOString(),
    nodeVersion: process.version,
    platform: process.platform,
    results,
  }
  fs.writeFileSync(path.resolve(RESULTS_DIR, 'xml-vs-zip.json'), JSON.stringify(output, null, 2))
  console.log(`\n💾 Saved results to benchmarks/results/xml-vs-zip.json`)

  await cleanup()
}

runBenchmarks().catch(console.error)
