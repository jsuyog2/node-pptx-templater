import { describe, it, expect } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'
import { validatePackageIntegrity } from '../helpers/packageIntegrity.js'
import { XMLParser } from '../../src/parsers/XMLParser.js'

describe('SlideManager - addSection OpenXML Validation', () => {
  it('enforces contiguous partitioning across ALL slides without orphan slides', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.addSlide() // 2
    ppt.addSlide() // 3
    ppt.addSlide() // 4 slides total

    // Add section at slide 3 without preceding sections
    const res = ppt.addSection('Financials', 3)
    expect(res).toBe(ppt) // Fluent chaining

    const sections = ppt.slideManager.getSections()
    expect(sections).toHaveLength(2)

    // 1. First section should be auto-created 'Default Section' covering slides 1 to 2
    const defaultSec = ppt.getSection('Default Section')
    expect(defaultSec).not.toBeNull()
    expect(defaultSec.slideIndices).toEqual([1, 2])
    expect(defaultSec.startSlideIndex).toBe(1)

    // 2. Second section is 'Financials' covering slides 3 to 4
    const financialsSec = ppt.getSection('Financials')
    expect(financialsSec).not.toBeNull()
    expect(financialsSec.slideIndices).toEqual([3, 4])
    expect(financialsSec.startSlideIndex).toBe(3)

    // Verify all presentation slide IDs are accounted for in order
    const allAssignedSlideIds = sections.flatMap(s => s.slideIds)
    const xml = await ppt.zipManager.readFile('ppt/presentation.xml')
    const sldIdLstMatches = [...xml.matchAll(/<p:sldId[^>]*id="(\d+)"[^>]*>/g)].map(m => m[1])
    expect(allAssignedSlideIds).toEqual(sldIdLstMatches)

    // Adding a 3rd section at slide 4 further subdivides contiguously
    ppt.addSection('Appendix', 4)
    const updatedSections = ppt.slideManager.getSections()
    expect(updatedSections).toHaveLength(3)

    expect(ppt.getSection('Default Section').slideIndices).toEqual([1, 2])
    expect(ppt.getSection('Financials').slideIndices).toEqual([3])
    expect(ppt.getSection('Appendix').slideIndices).toEqual([4])
  })

  it('uses explicit numerical slide IDs in <p14:sldId id="..."/> matching <p:sldIdLst>', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.addSlide() // 2 slides total

    ppt.addSection('Overview', 1)

    const xml = await ppt.zipManager.readFile('ppt/presentation.xml')
    const parser = new XMLParser()
    const presObj = parser.parse(xml)

    // Extract numerical IDs from <p:sldIdLst>
    const sldIdLst = presObj['p:presentation']['p:sldIdLst']['p:sldId']
    const mainSldIds = (Array.isArray(sldIdLst) ? sldIdLst : [sldIdLst]).map(s => String(s['@_id']))

    // Extract IDs from <p14:sectionLst>
    const extLst = presObj['p:presentation']['p:extLst']
    const exts = Array.isArray(extLst['p:ext']) ? extLst['p:ext'] : [extLst['p:ext']]
    const sectionExt = exts.find(e => e['p14:sectionLst'])
    const sections = sectionExt['p14:sectionLst']['p14:section']
    const sectionList = Array.isArray(sections) ? sections : [sections]

    const sectionSlideIds = []
    for (const sec of sectionList) {
      const slds = sec['p14:sldIdLst']['p14:sldId']
      const sldArr = Array.isArray(slds) ? slds : [slds]
      for (const s of sldArr) {
        const id = String(s['@_id'])
        // Must be purely digits (e.g. 256, 257)
        expect(id).toMatch(/^\d+$/)
        // Must NOT be relationship IDs like rId2
        expect(id).not.toMatch(/^rId\d+$/)
        sectionSlideIds.push(id)
      }
    }

    expect(sectionSlideIds).toEqual(mainSldIds)
  })

  it('registers uppercase GUIDs wrapped in braces for every section', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.addSlide()
    ppt.addSlide()

    ppt.addSection('Sec1', 1)
    ppt.addSection('Sec2', 2)

    const xml = await ppt.zipManager.readFile('ppt/presentation.xml')
    const parser = new XMLParser()
    const presObj = parser.parse(xml)

    const exts = presObj['p:presentation']['p:extLst']['p:ext']
    const extList = Array.isArray(exts) ? exts : [exts]
    const sectionExt = extList.find(e => e['p14:sectionLst'])
    const sections = sectionExt['p14:sectionLst']['p14:section']

    const guidRegex = /^\{[0-9A-F]{8}-[0-9A-F]{4}-4[0-9A-F]{3}-[89AB][0-9A-F]{3}-[0-9A-F]{12}\}$/
    for (const sec of sections) {
      const id = sec['@_id']
      expect(id).toMatch(guidRegex)
    }
  })

  it('declares exact extension URI {521415D9-36F7-43E2-AB2F-B90AF26B5E84} and xmlns:p14 namespace', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.addSection('StandardURISection', 1)

    const xml = await ppt.zipManager.readFile('ppt/presentation.xml')

    // Root presentation namespace
    expect(xml).toContain('xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"')

    // Section list extension URI (MS-PPTX 2.16.1 specification)
    expect(xml).toContain('uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"')

    // Inner sectionLst xmlns
    expect(xml).toContain('p14:sectionLst')
  })

  it('ensures <p:extLst> is written as the final child before </p:presentation>', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.addSlide()
    ppt.addSection('TerminalSection', 1)

    const xml = await ppt.zipManager.readFile('ppt/presentation.xml')

    // <p:extLst> must be immediately before </p:presentation>
    const extLstAtEndPattern = /<\/p:extLst>\s*<\/p:presentation>/
    expect(xml).toMatch(extLstAtEndPattern)

    // Verify presentation package integrity
    const buffer = await ppt.toBuffer()
    const integrityErrors = await validatePackageIntegrity(buffer)
    expect(integrityErrors).toEqual([])
  })
})
