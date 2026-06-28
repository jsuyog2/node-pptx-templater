import { describe, it, expect, beforeAll } from 'vitest'
import { existsSync } from 'fs'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { validatePackageIntegrity } from '../helpers/packageIntegrity.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const TEMPLATE = resolve(__dirname, '../../templates/officePPT.pptx')

let PPTXTemplater

beforeAll(async () => {
  const module = await import('../../src/index.js')
  PPTXTemplater = module.PPTXTemplater
})

describe('PPTXTemplater - Corporate Template (officePPT.pptx) Package Integrity', () => {
  const runFixtureTests = existsSync(TEMPLATE)

  if (!runFixtureTests) {
    it.skip('Corporate template officePPT.pptx not found — skipping tests.', () => {})
    return
  }

  async function assertValidBuffer(buffer, label) {
    const errors = await validatePackageIntegrity(buffer)
    if (errors.length > 0) {
      throw new Error(`${label} failed with OOXML package integrity errors:\n${errors.map(e => `  - ${e}`).join('\n')}`)
    }
  }

  it('should maintain package integrity after simple load and save', async () => {
    const ppt = await PPTXTemplater.load(TEMPLATE)
    const buf = await ppt.toBuffer()
    await assertValidBuffer(buf, 'simple load and save')
  })

  it('should maintain package integrity after duplicating slide 1', async () => {
    const ppt = await PPTXTemplater.load(TEMPLATE)
    await ppt.duplicateSlide(1, 2)
    const buf = await ppt.toBuffer()
    await assertValidBuffer(buf, 'duplicateSlide(1, 2)')
  })

  it('should maintain package integrity after removing slide 2', async () => {
    const ppt = await PPTXTemplater.load(TEMPLATE)
    ppt.removeSlide(2)
    const buf = await ppt.toBuffer()
    await assertValidBuffer(buf, 'removeSlide(2)')
  })

  it('should maintain package integrity after a complex sequence of operations', async () => {
    const ppt = await PPTXTemplater.load(TEMPLATE)

    // Duplicate slide 1 to 2
    await ppt.duplicateSlide(1, 2)

    // Remove slide 3 (original slide 2)
    ppt.removeSlide(3)

    // Move slide 3 to 1
    ppt.moveSlide(3, 1)

    const buf = await ppt.toBuffer()
    await assertValidBuffer(buf, 'complex sequence (duplicate, remove, move)')
  })
})
