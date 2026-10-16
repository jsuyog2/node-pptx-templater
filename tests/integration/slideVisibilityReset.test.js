import { describe, it, expect } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { PPTXTemplater } from '../../src/index.js'
import { validatePackageIntegrity } from '../helpers/packageIntegrity.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Slide Visibility & Canvas Reset (hideSlide, unhideSlide, resetSlide)', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')

  it('should hide and unhide slides via show="0|1" attribute on root <p:sld>', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Initially slide 1 is visible
    expect(ppt.getSlide(1).hidden).toBe(false)

    // Hide slide 1
    ppt.hideSlide(1)
    expect(ppt.getSlide(1).hidden).toBe(true)

    // Unhide slide 1
    ppt.unhideSlide(1)
    expect(ppt.getSlide(1).hidden).toBe(false)

    // Verify package integrity
    const buf = await ppt.toBuffer()
    const errors = await validatePackageIntegrity(buf)
    expect(errors).toEqual([])
  })

  it('should reset canvas by clearing user shapes and preserving placeholders', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.resetSlide(1, { preservePlaceholders: true })

    const buf = await ppt.toBuffer()
    const errors = await validatePackageIntegrity(buf)
    expect(errors).toEqual([])
  })

  it('should reset canvas completely when preservePlaceholders is false', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    ppt.resetSlide(1, { preservePlaceholders: false })

    const buf = await ppt.toBuffer()
    const errors = await validatePackageIntegrity(buf)
    expect(errors).toEqual([])
  })
})
