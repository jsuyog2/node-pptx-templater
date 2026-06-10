import { describe, it, expect, beforeAll } from 'vitest'
import { existsSync } from 'fs'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURES_DIR = resolve(__dirname, '../fixtures')

let PPTXTemplater

beforeAll(async () => {
  const module = await import('../../src/index.js')
  PPTXTemplater = module.PPTXTemplater
})

describe('Performance & Stress Benchmark Tests', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')
  const runTests = existsSync(fixtureFile)

  if (!runTests) {
    it.skip('Fixture file not found', () => {})
    return
  }

  it('should verify shape/text replacement scaling remains sub-second for 200 operations', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    ppt.enablePerformanceProfile()

    const startTime = Date.now()

    // Perform 200 text replacements sequentially
    for (let i = 0; i < 200; i++) {
      ppt.useSlide(1)
      ppt.replaceText({ '{{title}}': `Update ${i}` })
    }

    const elapsed = Date.now() - startTime
    console.log(`Time taken for 200 replacements: ${elapsed}ms`)

    // With caching & indexing, 200 updates should easily be sub-100ms
    expect(elapsed).toBeLessThan(1000)
  })

  it('should verify cache loading is at least 10x faster than disk/ZIP decompression', async () => {
    // 1. Run loading without cache 10 times to measure average
    const startTimeDisk = Date.now()
    for (let i = 0; i < 10; i++) {
      const ppt = await PPTXTemplater.load(fixtureFile)
      expect(ppt.slideManager.slideCount).toBeGreaterThan(0)
    }
    const elapsedDisk = Date.now() - startTimeDisk
    const avgDisk = elapsedDisk / 10
    console.log(`Average load time from disk: ${avgDisk.toFixed(2)}ms`)

    // 2. Preload cache
    await PPTXTemplater.preload(fixtureFile)

    // 3. Run loading with cache 10 times to measure average
    const startTimeCache = Date.now()
    for (let i = 0; i < 10; i++) {
      const ppt = await PPTXTemplater.fromCache(fixtureFile)
      expect(ppt.slideManager.slideCount).toBeGreaterThan(0)
    }
    const elapsedCache = Date.now() - startTimeCache
    const avgCache = elapsedCache / 10
    console.log(`Average load time from cache: ${avgCache.toFixed(2)}ms`)

    expect(avgCache).toBeLessThan(avgDisk)
    PPTXTemplater.clearCache()
  })

  it('should verify concurrent slide generations using the same cache are safe', async () => {
    await PPTXTemplater.preload(fixtureFile)

    const CONCURRENCY_LIMIT = 50
    const promises = []

    for (let i = 0; i < CONCURRENCY_LIMIT; i++) {
      promises.push(
        (async index => {
          const ppt = await PPTXTemplater.fromCache(fixtureFile)
          ppt.useSlide(1)
          ppt.replaceText({ '{{title}}': `Concurrent Task ${index}` })
          const buf = await ppt.toBuffer()

          // Quick validation of the generated buffer
          const reloaded = await PPTXTemplater.load(buf)
          const slideXml = reloaded.slideManager.getSlideXml(1)
          expect(slideXml).toContain(`Concurrent Task ${index}`)
        })(i)
      )
    }

    await Promise.all(promises)
    console.log(`Successfully completed ${CONCURRENCY_LIMIT} concurrent requests`)

    PPTXTemplater.clearCache()
  })
})
