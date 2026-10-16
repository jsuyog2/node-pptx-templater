import { describe, it, expect } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { PPTXTemplater } from '../../src/index.js'
import { validatePackageIntegrity } from '../helpers/packageIntegrity.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('TASK-SLIDES-001: Bypass Mandatory validatePresentation Workflow', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')

  it('should run createBlank() -> addShape() -> importSlideFrom() workflow without mandatory validation halt', async () => {
    // 1. Create a blank deck from scratch
    const blankDeck = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })

    // 2. Add shapes using fluent chaining
    await blankDeck.addShape('rect', {
      id: 'HeroBox',
      x: 100,
      y: 100,
      width: 400,
      height: 200,
      fill: '#2563EB',
      text: 'In-Memory Blank Deck',
    })

    // 3. Import slide from source deck into position 2
    const sourceDeck = await PPTXTemplater.load(samplePath)
    await blankDeck.importSlideFrom(sourceDeck, 1, 2)

    expect(blankDeck.slideCount).toBe(2)

    // 4. Export without requiring validatePresentation to be invoked beforehand
    const buffer = await blankDeck.toBuffer()
    expect(buffer).toBeDefined()
    expect(buffer.length).toBeGreaterThan(0)

    // 5. Output package integrity must be valid
    const errors = await validatePackageIntegrity(buffer)
    expect(errors).toEqual([])

    // 6. User can still call validatePresentation() on-demand for diagnostic analysis
    const diagnostic = await blankDeck.validatePresentation()
    expect(diagnostic).toHaveProperty('valid')
    expect(diagnostic).toHaveProperty('errors')
    expect(diagnostic).toHaveProperty('warnings')
  })
})
