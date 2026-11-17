import { describe, it, expect } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('SlideManager - addSection', () => {
  it('should create valid <p14:sectionLst> with xmlns:p14 namespace and GUID', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.addSlide()
    ppt.addSlide() // Total 3 slides

    // Test fluent chaining
    const res = ppt.addSection('Introduction', 1)
    expect(res).toBe(ppt)

    ppt.addSection('Details', 2)

    const presentationXml = await ppt.zipManager.readFile('ppt/presentation.xml')
    expect(presentationXml).toContain(
      'xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"'
    )
    expect(presentationXml).toContain('p14:sectionLst')
    expect(presentationXml).toContain('name="Introduction"')
    expect(presentationXml).toContain('name="Details"')
    // Match GUID pattern {XXXXXXXX-XXXX-4XXX-YXXX-XXXXXXXXXXXX}
    expect(presentationXml).toMatch(
      /id="\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-4[0-9A-Fa-f]{3}-[89ABab][0-9A-Fa-f]{3}-[0-9A-Fa-f]{12}\}"/
    )

    // Check getSection
    const introSection = ppt.getSection('Introduction')
    expect(introSection).not.toBeNull()
    expect(introSection.name).toBe('Introduction')
    expect(introSection.slideIndices).toContain(1)

    const detailsSection = ppt.getSection('Details')
    expect(detailsSection).not.toBeNull()
    expect(detailsSection.name).toBe('Details')
    expect(detailsSection.slideIndices).toContain(2)
    expect(detailsSection.slideIndices).toContain(3)
  })

  it('should clamp startSlideIndex within [1, slideCount]', async () => {
    const ppt = await PPTXTemplater.createBlank() // 1 slide
    ppt.addSlide() // 2 slides

    ppt.addSection('ClampedStart', 0)
    const sec1 = ppt.getSection('ClampedStart')
    expect(sec1.startSlideIndex).toBe(1)

    ppt.addSection('ClampedEnd', 999)
    const sec2 = ppt.getSection('ClampedEnd')
    expect(sec2.startSlideIndex).toBe(2)
  })
})
