import { describe, it, expect, beforeAll } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync } from 'fs'

const __dirname = dirname(fileURLToPath(import.meta.url))
const FIXTURE_FILE = resolve(__dirname, '../fixtures/sample.pptx')

let PPTXTemplater
let ValidationEngine

beforeAll(async () => {
  const module = await import('../../src/index.js')
  PPTXTemplater = module.PPTXTemplater
  ValidationEngine = module.ValidationEngine
})

describe('PPTX Lists Integration Tests', () => {
  const runTests = existsSync(FIXTURE_FILE)

  if (!runTests) {
    it.skip('Skipping: sample.pptx fixture not found', () => {})
    return
  }

  it('should validate lists correctly', () => {
    // Valid lists
    expect(ValidationEngine.validateList(['Item A', 'Item B']).valid).toBe(true)
    expect(
      ValidationEngine.validateList({
        list: ['Item A', { text: 'Item B', children: ['Item B.1'] }],
      }).valid
    ).toBe(true)

    // Invalid levels
    const invalidLevel = ValidationEngine.validateList({
      list: [
        {
          text: 'Item A',
          children: [
            {
              text: 'Item B',
              children: [
                {
                  text: 'Item C',
                  children: [
                    {
                      text: 'Item D',
                      children: [
                        {
                          text: 'Item E',
                          children: [
                            {
                              text: 'Item F',
                              children: [
                                {
                                  text: 'Item G',
                                  children: [
                                    {
                                      text: 'Item H',
                                      children: [
                                        {
                                          text: 'Item I',
                                          children: ['Level 9 - out of bounds'],
                                        },
                                      ],
                                    },
                                  ],
                                },
                              ],
                            },
                          ],
                        },
                      ],
                    },
                  ],
                },
              ],
            },
          ],
        },
      ],
    })
    expect(invalidLevel.valid).toBe(false)
    expect(invalidLevel.errors[0]).toContain('out of supported range')

    // Empty list item
    expect(ValidationEngine.validateList(['']).valid).toBe(false)
    expect(ValidationEngine.validateList({ list: [{ text: '' }] }).valid).toBe(false)

    // Invalid style property
    expect(
      ValidationEngine.validateList({
        list: ['Item A'],
        style: { fontSize: -12 },
      }).valid
    ).toBe(false)

    expect(
      ValidationEngine.validateList({
        list: ['Item A'],
        style: { color: 'invalid-hex' },
      }).valid
    ).toBe(false)
  })

  it('should support replacing simple bulleted lists inside shapes', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1).updateText('Title', {
      list: ['Bullet Point A', 'Bullet Point B'],
    })

    const buffer = await ppt.toBuffer()
    expect(buffer).toBeDefined()

    const ppt2 = await PPTXTemplater.load(buffer)
    const list = ppt2.useSlide(1).getList('Title')
    expect(list.length).toBe(2)
    expect(list[0]).toBe('Bullet Point A')
    expect(list[1]).toBe('Bullet Point B')
  })

  it('should support numbered/ordered lists', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1).updateText('Title', {
      ordered: true,
      list: ['Step 1', 'Step 2'],
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)
    const list = ppt2.useSlide(1).getList('Title')
    expect(list).toEqual(['Step 1', 'Step 2'])

    const xml = await ppt2.zipManager.readFile('ppt/slides/slide1.xml')
    expect(xml).toContain('buAutoNum')
  })

  it('should support multi-level nested lists', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1).updateText('Title', {
      list: [
        'Root item',
        {
          text: 'Second level parent',
          children: [
            'Child level 1',
            {
              text: 'Child level 2',
              children: ['Grandchild level 3'],
            },
          ],
        },
      ],
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)
    const list = ppt2.useSlide(1).getList('Title')

    expect(list.length).toBe(2)
    expect(list[0]).toBe('Root item')
    expect(list[1]).toEqual({
      text: 'Second level parent',
      children: [
        'Child level 1',
        {
          text: 'Child level 2',
          children: ['Grandchild level 3'],
        },
      ],
    })
  })

  it('should support styled lists with fonts, colors, bullet custom character, and size overrides', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(1).updateText('Title', {
      list: ['Point Alpha', 'Point Beta'],
      style: {
        fontFamily: 'Arial',
        fontSize: 22,
        color: '#FF0000',
        bulletColor: '#0000FF',
        bulletChar: '✦',
        bulletSize: 120,
      },
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)

    const xml = await ppt2.zipManager.readFile('ppt/slides/slide1.xml')
    expect(xml).toContain('typeface="Arial"')
    expect(xml).toContain('sz="2200"') // fontSize * 100
    expect(xml).toContain('val="FF0000"') // color hex
    expect(xml).toContain('val="0000FF"') // bulletColor hex
    expect(xml).toContain('char="✦"') // bulletChar
    expect(xml).toContain('val="120000"') // bulletSize * 1000
  })

  it('should support replacing placeholder tags inside table cells', async () => {
    const ppt = await PPTXTemplater.load(FIXTURE_FILE)
    ppt.useSlide(3) // Slide 3 contains the table 'Table'

    ppt.updateTable('Table', [
      ['Header 1', 'Header 2', 'Header 3', 'Header 4', 'Header 5'],
      ['{{CellPlaceholder}}', 'Cell Value', '', '', ''],
    ])

    ppt.updateText('CellPlaceholder', {
      list: ['Table Bullet 1', 'Table Bullet 2'],
    })

    const buffer = await ppt.toBuffer()
    const ppt2 = await PPTXTemplater.load(buffer)

    const list = ppt2.useSlide(3).getList('Table Bullet 1') // Find cell containing the text
    expect(list).toEqual(['Table Bullet 1', 'Table Bullet 2'])
  })
})
