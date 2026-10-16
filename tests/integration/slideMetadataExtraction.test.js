import { describe, it, expect } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('SlideManager - Rich Slide Metadata Extraction', () => {
  it('should extract comprehensive element geometry, element types, and slide metadata', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'HeroCard',
      x: '10%',
      y: '10%',
      width: '80%',
      height: '30%',
      fill: '#0066CC',
      text: 'Hero Header',
    })

    ppt.addTable(
      [
        ['Metric', 'Score'],
        ['Speed', '99%'],
      ],
      {
        x: '10%',
        y: '50%',
        width: '80%',
        height: '30%',
      }
    )

    const slideInfo = ppt.getSlide(1)
    expect(slideInfo).toBeDefined()
    expect(slideInfo.index).toBe(1)
    expect(slideInfo.dimensions).toBeDefined()
    expect(slideInfo.dimensions.width).toBe(12192000)
    expect(slideInfo.dimensions.height).toBe(6858000)

    // Check elements array
    expect(Array.isArray(slideInfo.elements)).toBe(true)
    expect(slideInfo.elements.length).toBeGreaterThanOrEqual(2)

    const heroShape = slideInfo.elements.find(el => el.id === 'HeroCard' || el.name === 'HeroCard')
    expect(heroShape).toBeDefined()
    expect(heroShape.type).toBe('shape')
    expect(heroShape.bounds).toBeDefined()
    expect(heroShape.bounds.x).toBe(1219200) // 10%
    expect(heroShape.boundsPx).toBeDefined()
    expect(heroShape.boundsPx.x).toBe(128) // 1280 * 0.1
    expect(heroShape.text).toContain('Hero Header')

    const tableEl = slideInfo.elements.find(el => el.type === 'table')
    expect(tableEl).toBeDefined()
    expect(tableEl.bounds).toBeDefined()
    expect(tableEl.boundsPx).toBeDefined()

    // Check getSlides()
    const allSlides = ppt.getSlides()
    expect(Array.isArray(allSlides)).toBe(true)
    expect(allSlides.length).toBe(1)
    expect(allSlides[0].index).toBe(1)
    expect(allSlides[0].elements.length).toBe(slideInfo.elements.length)

    // Ensure JSON serializability (no circular references)
    expect(() => JSON.stringify(slideInfo)).not.toThrow()
    expect(() => JSON.stringify(allSlides)).not.toThrow()
  })
})
