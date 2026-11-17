import { describe, it, expect } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { PPTXTemplater } from '../../src/index.js'
import { validatePackageIntegrity } from '../helpers/packageIntegrity.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Collaborative Comments (addComment & removeComment)', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')

  it('should add comments to slides with author and coordinates', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.useSlide(1).addComment('Review financial projections', {
      author: 'Alice Analyst',
      initials: 'AA',
      x: '1in',
      y: '1.5in',
    })

    const buf = await ppt.toBuffer()
    const errors = await validatePackageIntegrity(buf)
    expect(errors).toEqual([])
  })

  it('should add multiple comments and remove comment by index or text match', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.useSlide(1).addComment('First review comment', {
      author: 'Reviewer A',
      initials: 'RA',
    })

    await ppt.useSlide(1).addComment('Second note to delete', {
      author: 'Reviewer B',
      initials: 'RB',
    })

    // Remove comment matching text query
    await ppt.useSlide(1).removeComment('delete')

    const buf = await ppt.toBuffer()
    const errors = await validatePackageIntegrity(buf)
    expect(errors).toEqual([])
  })
})
