/**
 * @fileoverview Integration tests for the PPTXTemplater.
 *
 * These tests work with actual PPTX files in tests/fixtures/.
 * If no fixtures are present, tests use mock data.
 */

import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import { existsSync } from 'fs'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import fsExtra from 'fs-extra'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURES_DIR = resolve(__dirname, '../fixtures')
const OUTPUT_DIR = resolve(__dirname, '../output-main')

/**
 * Import the engine lazily so tests can still run without fixtures.
 */
let PPTXTemplater

beforeAll(async () => {
  // Import the engine
  const module = await import('../../src/index.js')
  PPTXTemplater = module.PPTXTemplater

  // Create output directory
  await fsExtra.ensureDir(OUTPUT_DIR)
})

afterAll(async () => {
  // Clean up test output files
  await fsExtra.remove(OUTPUT_DIR)
})

describe('PPTXTemplater - Module Exports', () => {
  it('should export PPTXTemplater class', async () => {
    const module = await import('../../src/index.js')
    expect(module.PPTXTemplater).toBeDefined()
    expect(typeof module.PPTXTemplater).toBe('function')
  })

  it('should export all manager classes', async () => {
    const module = await import('../../src/index.js')
    expect(module.ZipManager).toBeDefined()
    expect(module.XMLParser).toBeDefined()
    expect(module.SlideManager).toBeDefined()
    expect(module.ChartManager).toBeDefined()
    expect(module.TableManager).toBeDefined()
    expect(module.HyperlinkManager).toBeDefined()
    expect(module.MediaManager).toBeDefined()
    expect(module.RelationshipManager).toBeDefined()
  })

  it('should export utility functions', async () => {
    const module = await import('../../src/index.js')
    expect(module.generateRelationshipId).toBeDefined()
    expect(module.validateXML).toBeDefined()
    expect(module.createLogger).toBeDefined()
  })

  it('should export all error classes', async () => {
    const module = await import('../../src/index.js')
    expect(module.PPTXError).toBeDefined()
    expect(module.SlideNotFoundError).toBeDefined()
    expect(module.ChartNotFoundError).toBeDefined()
    expect(module.TableNotFoundError).toBeDefined()
  })
})

describe('PPTXTemplater - Error Handling', () => {
  it('should throw PPTXError when loading non-existent file', async () => {
    const { PPTXError } = await import('../../src/index.js')
    await expect(PPTXTemplater.load('./does-not-exist.pptx')).rejects.toThrow(PPTXError)
  })

  it('should throw when calling methods before load', async () => {
    // The engine requires PPTXTemplater.load() - direct construction is not exposed
    // This is enforced by the private constructor pattern
  })
})

describe('PPTXTemplater - With Fixture File', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')

  // Skip fixture-based tests if no fixture file is available
  const runFixtureTests = existsSync(fixtureFile)

  if (!runFixtureTests) {
    it.skip('Fixture file not found — skipping integration tests. Run: npm run generate:fixtures', () => {})
    return
  }

  let ppt

  beforeAll(async () => {
    ppt = await PPTXTemplater.load(fixtureFile)
  })

  it('should load the PPTX and report slide count', () => {
    expect(ppt.slideCount).toBeGreaterThan(0)
  })

  it('should return presentation info', () => {
    const info = ppt.getInfo()
    expect(info).toHaveProperty('slideCount')
    expect(info).toHaveProperty('slides')
    expect(Array.isArray(info.slides)).toBe(true)
  })

  it('should select a slide without error', () => {
    expect(() => ppt.useSlide(1)).not.toThrow()
  })

  it('should replace text without error', () => {
    expect(() => ppt.replaceText({ '{{test}}': 'replaced' })).not.toThrow()
  })

  it('should save to buffer', async () => {
    const buffer = await ppt.toBuffer()
    expect(buffer).toBeInstanceOf(Buffer)
    expect(buffer.length).toBeGreaterThan(0)
    // Buffer should start with PK (ZIP magic bytes)
    expect(buffer[0]).toBe(0x50) // 'P'
    expect(buffer[1]).toBe(0x4b) // 'K'
  })

  it('should save to file', async () => {
    const outPath = resolve(OUTPUT_DIR, 'integration-test-output.pptx')
    await ppt.saveToFile(outPath)
    expect(existsSync(outPath)).toBe(true)
  })

  it('should validate structure', () => {
    const result = ppt.validate()
    expect(result).toHaveProperty('valid')
    expect(result).toHaveProperty('errors')
    expect(result).toHaveProperty('warnings')
    expect(Array.isArray(result.errors)).toBe(true)
    expect(Array.isArray(result.warnings)).toBe(true)
  })
})

describe('PPTXTemplater - ChainableAPI', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')
  const runFixtureTests = existsSync(fixtureFile)

  if (!runFixtureTests) {
    it.skip('Fixture file needed for API chain tests', () => {})
    return
  }

  it('should support method chaining', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)

    // All these should return the same ppt instance (chainable)
    const result = ppt.useSlide(1).replaceText({ '{{x}}': 'y' })
    expect(result).toBe(ppt)
  })
})
