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

  it('4. should apply text color via options.color and options.style.color with strict ECMA-376 schema sequence', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'ColorTestBox',
      name: 'ColorTestBox',
      x: 100,
      y: 100,
      width: 500,
      height: 80,
      text: 'Explore Alpha, Beta, and Gamma documentation.',
    })

    // Substring 1: Root options.color with hex '#'
    ppt.setHyperlink('ColorTestBox', 'Alpha', 'https://example.com/alpha', {
      color: '#E11D48',
      underline: true,
      bold: true,
    })

    // Substring 2: Nested options.style.color
    ppt.setHyperlink('ColorTestBox', 'Beta', 'https://example.com/beta', {
      style: {
        color: '#2563EB',
        italic: true,
      },
    })

    // Substring 3: Named color string 'red'
    ppt.setHyperlink('ColorTestBox', 'Gamma', 'https://example.com/gamma', {
      color: 'red',
    })

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')

    // Verify hex colors normalized without '#' and uppercase
    expect(slideXml).toContain('val="E11D48"')
    expect(slideXml).toContain('val="2563EB"')
    expect(slideXml).toContain('val="FF0000"') // 'red' normalized to 'FF0000', NOT 'RREEDD'

    // Verify ECMA-376 sequence: <a:solidFill> MUST appear before <a:hlinkClick> in each run
    const alphaRunMatch = slideXml.match(
      /<a:rPr[^>]*>[\s\S]*?<a:solidFill>[\s\S]*?val="E11D48"[\s\S]*?<\/a:solidFill>[\s\S]*?<a:hlinkClick[\s\S]*?<\/a:rPr>/
    )
    expect(alphaRunMatch).not.toBeNull()

    const betaRunMatch = slideXml.match(
      /<a:rPr[^>]*>[\s\S]*?<a:solidFill>[\s\S]*?val="2563EB"[\s\S]*?<\/a:solidFill>[\s\S]*?<a:hlinkClick[\s\S]*?<\/a:rPr>/
    )
    expect(betaRunMatch).not.toBeNull()

    const gammaRunMatch = slideXml.match(
      /<a:rPr[^>]*>[\s\S]*?<a:solidFill>[\s\S]*?val="FF0000"[\s\S]*?<\/a:solidFill>[\s\S]*?<a:hlinkClick[\s\S]*?<\/a:rPr>/
    )
    expect(gammaRunMatch).not.toBeNull()
  })

  it('5. should support custom colors on slideshow action hyperlinks without corruption', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.addSlide() // Slide 2
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'ActionTextBox',
      name: 'ActionTextBox',
      x: 100,
      y: 100,
      width: 400,
      height: 60,
      text: 'Click Next to proceed to the following slide.',
    })

    ppt.setHyperlink('ActionTextBox', 'Next', 'nextSlide', {
      color: '#10B981',
      bold: true,
      underline: true,
    })

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')

    expect(slideXml).toContain('val="10B981"')
    expect(slideXml).toContain('b="1"')
    expect(slideXml).toContain('u="sng"')
    expect(slideXml).toContain('action="ppaction://hlinkshowjump?jump=nextslide"')

    // Ensure solidFill precedes hlinkClick
    const actionRunMatch = slideXml.match(
      /<a:rPr[^>]*>[\s\S]*?<a:solidFill>[\s\S]*?val="10B981"[\s\S]*?<\/a:solidFill>[\s\S]*?<a:hlinkClick[\s\S]*?<\/a:rPr>/
    )
    expect(actionRunMatch).not.toBeNull()
  })

  it('6. should format rose hyperlink (#F43F5E) with synchronized underline fill preceding hlinkClick', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'RoseLinkBox',
      name: 'RoseLinkBox',
      x: 100,
      y: 100,
      width: 450,
      height: 70,
      text: 'Run the ValidationEngine suite.',
    })

    // Exact invocation pattern from showcase/hyperlinks_advanced.js lines 121-127
    ppt.setHyperlink('RoseLinkBox', 'ValidationEngine', 'first', {
      action: 'first',
      color: '#F43F5E',
      underline: true,
      italic: true,
      tooltip: 'Slideshow jump to First Slide',
    })

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')

    // 1. Text color sanitized to 6-digit hex without #
    expect(slideXml).toContain('val="F43F5E"')
    expect(slideXml).not.toContain('val="#F43F5E"')

    // 2. Both solidFill and uFill present with F43F5E
    expect(slideXml).toContain('<a:solidFill><a:srgbClr val="F43F5E"')
    expect(slideXml).toContain('<a:uFill><a:solidFill><a:srgbClr val="F43F5E"')

    // 3. Strict schema order: solidFill -> uFill -> hlinkClick
    const sequenceRegex =
      /<a:rPr[^>]*>[\s\S]*?<a:solidFill>[\s\S]*?val="F43F5E"[\s\S]*?<\/a:solidFill>[\s\S]*?<a:uFill>[\s\S]*?val="F43F5E"[\s\S]*?<\/a:uFill>[\s\S]*?<a:hlinkClick[^>]*action="ppaction:\/\/hlinkshowjump\?jump=firstslide"[\s\S]*?<\/a:rPr>/
    expect(slideXml).toMatch(sequenceRegex)
  })
})
