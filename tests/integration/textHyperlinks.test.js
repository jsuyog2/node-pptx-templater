import { describe, it, expect } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'
import JSZip from 'jszip'

describe('Substring Text Hyperlink & Styling API Integration Tests (TASK-HYPERLINK-004)', () => {
  it('1. should perform AST run splitting on single-run text and attach styled hyperlink', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    // Add a text shape with a full sentence
    await ppt.addShape('rect', {
      id: 'SentenceBox',
      name: 'SentenceBox',
      x: 100,
      y: 100,
      width: 400,
      height: 60,
      text: 'Please visit OpenAI for AI research papers.',
    })

    // Substring link on "OpenAI"
    ppt.setHyperlink('SentenceBox', 'OpenAI', 'https://openai.com', {
      color: '#0066CC',
      underline: true,
      bold: true,
      tooltip: 'OpenAI Website',
    })

    // Verify hyperlink query
    const links = ppt.getHyperlinks(1)
    const openAiLink = links.find(l => l.text === 'OpenAI')
    expect(openAiLink).toBeDefined()
    expect(openAiLink.url).toBe('https://openai.com')
    expect(openAiLink.type).toBe('textRun')
    expect(openAiLink.tooltip).toBe('OpenAI Website')

    // Inspect underlying AST runs
    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')
    const relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')

    // Run splitting verification
    expect(slideXml).toContain('Please visit ')
    expect(slideXml).toContain('<a:t>OpenAI</a:t>')
    expect(slideXml).toContain(' for AI research papers.')

    // Style overrides verification: color (#0066CC), underline (u="sng"), bold (b="1")
    expect(slideXml).toContain('u="sng"')
    expect(slideXml).toContain('b="1"')
    expect(slideXml).toContain('val="0066CC"')
    expect(slideXml).toContain('tooltip="OpenAI Website"')

    // Relationship verification
    expect(relsXml).toContain('Target="https://openai.com"')
    expect(relsXml).toContain('TargetMode="External"')
  })

  it('2. should perform cross-run healing and splitting when substring spans multiple runs', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    // Manually inject a fractured paragraph with runs: "Download ", "Rep", "ort now"
    const slideObj = ppt.slideManager.getSlideObj(1)
    const spTree = slideObj['p:sld']['p:cSld']['p:spTree']

    spTree['p:sp'] = [
      {
        'p:nvSpPr': {
          'p:cNvPr': { '@_id': '10', '@_name': 'FracturedTextBox' },
          'p:cNvSpPr': {},
          'p:nvPr': {},
        },
        'p:spPr': {
          'a:xfrm': {
            'a:off': { '@_x': '100000', '@_y': '100000' },
            'a:ext': { '@_cx': '2000000', '@_cy': '500000' },
          },
        },
        'p:txBody': {
          'a:bodyPr': {},
          'a:p': [
            {
              'a:r': [
                { 'a:rPr': { '@_lang': 'en-US' }, 'a:t': 'Download the ' },
                { 'a:rPr': { '@_lang': 'en-US' }, 'a:t': 'Rep' },
                { 'a:rPr': { '@_lang': 'en-US' }, 'a:t': 'ort now' },
              ],
            },
          ],
        },
      },
    ]
    ppt.slideManager.markSlideObjDirty(1)

    // Link target text "Report" which was split across runs "Rep" and "ort now"
    ppt.setHyperlink('FracturedTextBox', 'Report', 'https://example.com/report.pdf', {
      color: '#FF5500',
      italic: true,
      tooltip: 'Annual Report PDF',
    })

    const links = ppt.getHyperlinks(1)
    const reportLink = links.find(l => l.text === 'Report')
    expect(reportLink).toBeDefined()
    expect(reportLink.url).toBe('https://example.com/report.pdf')

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')

    expect(slideXml).toContain('Download the ')
    expect(slideXml).toContain('<a:t>Report</a:t>')
    expect(slideXml).toContain(' now')
    expect(slideXml).toContain('val="FF5500"')
    expect(slideXml).toContain('i="1"')
  })

  it('3. should support internal slide jump hyperlinks on text substrings', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.addSlide() // Slide 2
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'NavTextShape',
      name: 'NavTextShape',
      x: 100,
      y: 200,
      width: 300,
      height: 50,
      text: 'See details on Slide 2 for data.',
    })

    ppt.setHyperlink('NavTextShape', 'Slide 2', 2, {
      tooltip: 'Go to Slide 2',
    })

    const link = ppt.getHyperlink('Slide 2')
    expect(link).toBeDefined()
    expect(link.isInternal).toBe(true)
    expect(link.targetSlide).toBe(2)

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')
    const relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')

    expect(slideXml).toContain('action="ppaction://hlinksldjump"')
    expect(relsXml).toContain('Target="slide2.xml"')
  })
})
