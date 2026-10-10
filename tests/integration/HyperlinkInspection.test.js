import { describe, it, expect } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')

describe('Hyperlink Inspection and Management Integration Tests', () => {
  const runTests = existsSync(FIXTURE_FILE)

  if (!runTests) {
    it.skip('Skipping: sample.pptx fixture not found', () => {})
    return
  }

  it('1. should inspect external hyperlinks added to text runs', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1)

    // Add a shape with known text
    await ppt.addShape({
      id: 'TextShape1',
      type: 'rect',
      text: 'Visit OpenAI for AI research',
      x: 50,
      y: 50,
      width: 250,
      height: 60,
    })

    // Add external hyperlink to 'OpenAI'
    ppt.addHyperlink({
      text: 'OpenAI',
      url: 'https://openai.com',
      tooltip: 'OpenAI Homepage',
    })

    const links = ppt.getHyperlinks(1)
    expect(links.length).toBeGreaterThan(0)

    const openAiLink = links.find(l => l.url === 'https://openai.com')
    expect(openAiLink).toBeDefined()
    expect(openAiLink.type).toBe('external')
    expect(openAiLink.tooltip).toBe('OpenAI Homepage')
    expect(openAiLink.text).toBe('OpenAI')

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeDefined()
    expect(buffer.length).toBeGreaterThan(0)
  })

  it('2. should inspect slide-to-slide hyperlinks and resolve target slide index', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)

    // Add text on slide 1
    ppt.useSlide(1)
    await ppt.addShape({
      id: 'SlideLinkShape',
      type: 'rect',
      text: 'Go to Slide 3',
      x: 100,
      y: 100,
      width: 200,
      height: 50,
    })

    // Link text to slide 3
    ppt.addSlideLink({
      sourceSlide: 1,
      targetSlide: 3,
      element: 'Go to Slide 3',
    })

    const links = ppt.getHyperlinks(1)
    const slideLink = links.find(l => l.type === 'slide')
    expect(slideLink).toBeDefined()
    expect(slideLink.targetSlide).toBe(3)

    const buffer = await ppt.toBuffer()
    expect(buffer.length).toBeGreaterThan(0)
  })

  it('3. should inspect special navigation links on shapes', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1)

    await ppt.addShape({
      id: 'NavNextBtn',
      type: 'rectangle',
      x: 200,
      y: 200,
      width: 100,
      height: 40,
    })

    ppt.addShapeNavigationLink({
      slide: 1,
      shapeId: 'NavNextBtn',
      action: 'next',
    })

    const links = ppt.getHyperlinks(1)
    const navLink = links.find(l => l.type === 'navigation' && l.shapeName === 'NavNextBtn')
    expect(navLink).toBeDefined()
    expect(navLink.target).toBe('nextslide')

    const buffer = await ppt.toBuffer()
    expect(buffer.length).toBeGreaterThan(0)
  })

  it('4. should remove a hyperlink and verify it is no longer listed in getHyperlinks', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1)

    await ppt.addShape({
      id: 'GithubShape',
      type: 'rect',
      text: 'Check Github repository',
      x: 50,
      y: 300,
      width: 220,
      height: 50,
    })
    ppt.addHyperlink({
      text: 'Github',
      url: 'https://github.com',
    })

    let links = ppt.getHyperlinks(1)
    expect(links.some(l => l.url === 'https://github.com')).toBe(true)

    // Remove the hyperlink
    ppt.removeHyperlink('Github')

    links = ppt.getHyperlinks(1)
    expect(links.some(l => l.url === 'https://github.com')).toBe(false)

    const buffer = await ppt.toBuffer()
    expect(buffer.length).toBeGreaterThan(0)
  })
})
