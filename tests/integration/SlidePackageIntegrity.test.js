import { describe, it, expect, beforeAll, afterAll } from 'vitest'
import { existsSync } from 'fs'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import fsExtra from 'fs-extra'
import { validatePackageIntegrity } from '../helpers/packageIntegrity.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURES_DIR = resolve(__dirname, '../fixtures')
const OUTPUT_DIR = resolve(__dirname, '../output-integrity')
const TEMPLATE = resolve(__dirname, '../../templates/sample.pptx')

let PPTXTemplater

beforeAll(async () => {
  const module = await import('../../src/index.js')
  PPTXTemplater = module.PPTXTemplater
  await fsExtra.ensureDir(OUTPUT_DIR)

  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')
  if (!existsSync(fixtureFile) && existsSync(TEMPLATE)) {
    await fsExtra.copy(TEMPLATE, fixtureFile)
  }
})

afterAll(async () => {
  await fsExtra.remove(OUTPUT_DIR)
})

describe('PPTXTemplater - Slide package structural integrity', () => {
  const fixtureFile = resolve(FIXTURES_DIR, 'sample.pptx')
  const runFixtureTests = existsSync(fixtureFile)

  if (!runFixtureTests) {
    it.skip('Fixture file not found — skipping slide integrity tests.', () => {})
    return
  }

  async function assertValidBuffer(buffer, label) {
    const errors = await validatePackageIntegrity(buffer)
    if (errors.length > 0) {
      throw new Error(`${label}:\n${errors.map(e => `  - ${e}`).join('\n')}`)
    }
  }

  it('should produce a valid package after removing a slide', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    ppt.removeSlide(2)
    await assertValidBuffer(await ppt.toBuffer(), 'removeSlide(2)')
  })

  it('should produce a valid package after adding a new slide', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    ppt.addSlide({ title: 'New Slide' })
    await assertValidBuffer(await ppt.toBuffer(), 'addSlide')
  })

  it('should produce a valid package when duplicating the first slide', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    await ppt.duplicateSlide(1)
    await assertValidBuffer(await ppt.toBuffer(), 'duplicateSlide(1)')
  })

  it('should produce a valid package when duplicating a middle slide', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    const middle = Math.ceil(ppt.slideCount / 2)
    await ppt.duplicateSlide(middle)
    await assertValidBuffer(await ppt.toBuffer(), `duplicateSlide(${middle})`)
  })

  it('should produce a valid package when duplicating the last slide', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    await ppt.duplicateSlide(ppt.slideCount)
    await assertValidBuffer(await ppt.toBuffer(), 'duplicateSlide(last)')
  })

  it('should produce a valid package when inserting a duplicate at another position', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    const slideCount = ppt.slideCount
    if (slideCount >= 5) {
      await ppt.duplicateSlide(5, 6)
    } else {
      await ppt.duplicateSlide(1, 1)
    }
    await assertValidBuffer(await ppt.toBuffer(), 'duplicateSlide at position')
  })

  it('should produce a valid package after remove, duplicate, and reorder (sections sync)', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    ppt.removeSlide(2)
    ppt.addSlideLink({
      element: 'Hello {{title}}',
      sourceSlide: 2,
      targetSlide: 1,
    })
    await ppt.duplicateSlide(1, 3)
    await assertValidBuffer(await ppt.toBuffer(), 'remove + link + duplicate')
  })

  it('should produce a valid package after move and insert operations', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    await ppt.duplicateSlide(1)
    ppt.moveSlide(ppt.slideCount, 1)
    ppt.insertSlide(2, { title: 'Inserted' })
    await assertValidBuffer(await ppt.toBuffer(), 'move + insert')
  })

  it('should produce a valid package after basic-usage-style slide mutations', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    if (ppt.slideCount >= 2) {
      ppt.removeSlide(2)
    }
    ppt.useSlide(Math.min(2, ppt.slideCount))
    ppt.replaceText({ '{{title}}': 'Updated Title' })

    const totalTeamSlides = 2
    let slideNumber = Math.min(5, ppt.slideCount)
    for (let index = 1; index <= totalTeamSlides; index++) {
      if (slideNumber <= ppt.slideCount) {
        await ppt.duplicateSlide(slideNumber, slideNumber + index)
      }
    }

    await assertValidBuffer(await ppt.toBuffer(), 'basic-usage-style mutations')
  })

  it('should keep sldId r:id values aligned with presentation.xml.rels', async () => {
    const ppt = await PPTXTemplater.load(fixtureFile)
    ppt.removeSlide(1)
    await ppt.duplicateSlide(1, 1)
    await ppt.duplicateSlide(ppt.slideCount)

    const buffer = await ppt.toBuffer()
    const errors = await validatePackageIntegrity(buffer)
    expect(errors).toEqual([])
  })
})
