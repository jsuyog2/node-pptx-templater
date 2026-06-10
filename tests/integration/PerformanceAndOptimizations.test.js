import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import { existsSync, createReadStream } from 'fs'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import fsExtra from 'fs-extra'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURES_DIR = resolve(__dirname, '../fixtures')
const OUTPUT_DIR = resolve(__dirname, '../output-perf')

let PPTXTemplater

beforeAll(async () => {
  const module = await import('../../src/index.js')
  PPTXTemplater = module.PPTXTemplater

  await fsExtra.ensureDir(OUTPUT_DIR)
})

afterAll(async () => {
  await fsExtra.remove(OUTPUT_DIR)
})

describe('Performance and Optimizations Integration Tests', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')
  const runTests = existsSync(fixtureFile)

  if (!runTests) {
    it.skip('Fixture file not found', () => {})
    return
  }

  it('should enable performance profiling and retrieve metrics', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    ppt.enablePerformanceProfile()

    // Perform some updates
    ppt.replaceText({ '{{title}}': 'Perf Test' })

    // Save to buffer
    const buffer = await ppt.toBuffer()
    expect(buffer).toBeDefined()

    const metrics = ppt.getPerformanceMetrics()
    expect(metrics.enabled).toBe(true)
    expect(metrics.templateLoadMs).toBeGreaterThanOrEqual(0)
    expect(metrics.parseMs).toBeGreaterThanOrEqual(0)
    expect(metrics.zipGenerationMs).toBeGreaterThan(0)
    expect(metrics.totalMs).toBeGreaterThan(0)
    expect(metrics.memoryUsedMB).toBeGreaterThan(0)
  })

  it('should support in-memory template caching', async () => {
    // Clear cache first
    PPTXTemplater.clearCache()

    // Preload
    const cacheMap = await PPTXTemplater.preload(fixtureFile)
    expect(cacheMap).toBeDefined()
    expect(cacheMap.size).toBeGreaterThan(0)

    // Load from cache 1
    const ppt1 = await PPTXTemplater.fromCache(fixtureFile)
    expect(ppt1.slideManager.slideCount).toBeGreaterThan(0)

    // Load from cache 2
    const ppt2 = await PPTXTemplater.fromCache(fixtureFile)

    // Edit ppt1
    ppt1.useSlide(1)
    ppt1.replaceText({ '{{title}}': 'Modified ppt1' })

    // Verify ppt2 is unaffected
    const slide1Xml1 = ppt1.slideManager.getSlideXml(1)
    const slide1Xml2 = ppt2.slideManager.getSlideXml(1)

    expect(slide1Xml1).toContain('Modified ppt1')
    expect(slide1Xml2).not.toContain('Modified ppt1')

    PPTXTemplater.clearCache()
  })

  it('should support configurable compression levels', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)

    // Balanced compression (default)
    const bufBalanced = await ppt.toBuffer({ compression: 'balanced' })

    // None / Store compression
    const bufNone = await ppt.toBuffer({ compression: 'none' })

    // Maximum compression
    const bufMax = await ppt.toBuffer({ compression: 'maximum' })

    expect(bufBalanced.length).toBeLessThan(bufNone.length)
    expect(bufMax.length).toBeLessThanOrEqual(bufBalanced.length)
  })

  it('should support saving to streams', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)

    // 1. Get readable stream
    const readable = await ppt.toStream()
    expect(readable).toBeDefined()
    expect(typeof readable.pipe).toBe('function')

    // 2. Save to stream (returns readable)
    const streamRes = await ppt.saveToStream()
    expect(streamRes).toBeDefined()
    expect(typeof streamRes.pipe).toBe('function')

    // 3. Save to stream piping to writable
    const outputPath = resolve(OUTPUT_DIR, 'stream-out.pptx')
    const writeStream = fsExtra.createWriteStream(outputPath)
    await ppt.saveToStream(writeStream)

    expect(existsSync(outputPath)).toBe(true)
    const stat = fsExtra.statSync(outputPath)
    expect(stat.size).toBeGreaterThan(0)
  })

  it('should support streaming image replacements', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    const imagePath = resolve(FIXTURES_DIR, 'media__1780653961880.png')

    if (existsSync(imagePath)) {
      const imgStream = createReadStream(imagePath)

      ppt.useSlide(1)
      const imagesBefore = ppt.getImages()
      if (imagesBefore.length > 0) {
        const firstImageId = imagesBefore[0].id
        await ppt.replaceImage(firstImageId, imgStream)

        const buf = await ppt.toBuffer()
        expect(buf).toBeDefined()
      }
    }
  })
})
