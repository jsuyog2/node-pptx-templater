import { describe, it, expect } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('SlideManager - Positional importSlideFrom', () => {
  it('should import slide at position 1 (first)', async () => {
    const source = await PPTXTemplater.createBlank()
    await source
      .useSlide(1)
      .addShape('rect', { id: 'SourceCard', x: '10%', y: '10%', width: '20%', height: '20%' })

    const target = await PPTXTemplater.createBlank()
    target.addSlide() // now target has 2 slides
    expect(target.getSlides().length).toBe(2)

    // Import into position 1
    await target.importSlideFrom(source, 1, 1)
    expect(target.getSlides().length).toBe(3)

    const slides = target.getSlides()
    expect(slides[0].index).toBe(1)
    // The imported slide has SourceCard shape
    expect(slides[0].elements.some(el => el.name === 'SourceCard')).toBe(true)
  })

  it('should import slide at intermediate position (e.g. position 2)', async () => {
    const source = await PPTXTemplater.createBlank()
    await source
      .useSlide(1)
      .addShape('rect', { id: 'ImportedMiddle', x: '10%', y: '10%', width: '20%', height: '20%' })

    const target = await PPTXTemplater.createBlank()
    target.addSlide() // slide 2
    target.addSlide() // slide 3
    expect(target.getSlides().length).toBe(3)

    // Insert at position 2
    await target.importSlideFrom(source, 1, 2)
    expect(target.getSlides().length).toBe(4)

    const slides = target.getSlides()
    expect(slides[1].elements.some(el => el.name === 'ImportedMiddle')).toBe(true)
  })

  it('should clamp position <= 0 to 1 and position > slideCount to end', async () => {
    const source = await PPTXTemplater.createBlank()
    await source
      .useSlide(1)
      .addShape('rect', { id: 'ClampedStart', x: '10%', y: '10%', width: '20%', height: '20%' })

    const target = await PPTXTemplater.createBlank()
    target.addSlide() // 2 slides

    // Test <= 0 clamping to position 1
    await target.importSlideFrom(source, 1, 0)
    expect(target.getSlides().length).toBe(3)
    expect(target.getSlides()[0].elements.some(el => el.name === 'ClampedStart')).toBe(true)

    // Test position > count clamping to end
    const sourceEnd = await PPTXTemplater.createBlank()
    await sourceEnd
      .useSlide(1)
      .addShape('rect', { id: 'ClampedEnd', x: '10%', y: '10%', width: '20%', height: '20%' })
    await target.importSlideFrom(sourceEnd, 1, 9999)
    expect(target.getSlides().length).toBe(4)
    expect(target.getSlides()[3].elements.some(el => el.name === 'ClampedEnd')).toBe(true)
  })

  it('should append to the end if atPosition is undefined', async () => {
    const source = await PPTXTemplater.createBlank()
    await source
      .useSlide(1)
      .addShape('rect', { id: 'AppendedDefault', x: '10%', y: '10%', width: '20%', height: '20%' })

    const target = await PPTXTemplater.createBlank()
    target.addSlide() // 2 slides
    await target.importSlideFrom(source, 1)
    expect(target.getSlides().length).toBe(3)
    expect(target.getSlides()[2].elements.some(el => el.name === 'AppendedDefault')).toBe(true)
  })
})
