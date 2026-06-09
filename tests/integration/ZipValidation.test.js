import { describe, it, expect, beforeAll } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURES_DIR = resolve(__dirname, '../fixtures')

describe('PPTXTemplater - ZIP Validation and Debug Mode', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')
  const hasFixture = existsSync(fixtureFile)

  let PPTXTemplater
  let ppt

  beforeAll(async () => {
    const module = await import('../../src/index.js')
    PPTXTemplater = module.PPTXTemplater
    if (hasFixture) {
      ppt = await PPTXTemplater.load(fixtureFile)
    }
  })

  it('should successfully run validateArchive on a valid PPTX presentation', async () => {
    if (!hasFixture) {
      it.skip('Fixture file sample.pptx not found — skipping validation tests.', () => {})
      return
    }

    // validateArchive should resolve cleanly (no throws) for a valid presentation
    await expect(ppt.validateArchive()).resolves.toBe(ppt)
  })

  it('should throw an error on validateArchive when a critical file is removed', async () => {
    if (!hasFixture) {
      it.skip('Fixture file sample.pptx not found — skipping validation tests.', () => {})
      return
    }

    // load a fresh instance
    const badPpt = await PPTXTemplater.load(fixtureFile)
    
    // Remove a critical file to simulate corruption
    badPpt.zipManager.removeFile('[Content_Types].xml')

    // validateArchive should now fail and throw a PPTXError
    await expect(badPpt.validateArchive()).rejects.toThrow('ZIP archive validation failed')
  })

  it('should support enableDebugZip method chaining and debug output flag', async () => {
    if (!hasFixture) {
      it.skip('Fixture file sample.pptx not found — skipping validation tests.', () => {})
      return
    }

    const debugPpt = await PPTXTemplater.load(fixtureFile)
    
    // Should be chainable
    const returned = debugPpt.enableDebugZip()
    expect(returned).toBe(debugPpt)
  })
})
