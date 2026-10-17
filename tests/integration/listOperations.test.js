import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('TASK-TEXT-004: List Manipulation & Style Management Methods Integration Tests', () => {
  const outputPath = resolve(__dirname, '../fixtures/output_list_operations.pptx')

  beforeEach(() => {
    if (existsSync(outputPath)) {
      try {
        unlinkSync(outputPath)
      } catch (_) {}
    }
  })

  afterEach(() => {
    if (existsSync(outputPath)) {
      try {
        unlinkSync(outputPath)
      } catch (_) {}
    }
  })

  it('1. addLists: supports multi-level nested hierarchies with indent levels (lvl=0, 1, 2)', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addLists(
      [
        'Level 0 Item 1',
        {
          text: 'Level 0 Item 2 (Parent)',
          children: [
            'Level 1 Child A',
            {
              text: 'Level 1 Child B (Parent)',
              children: ['Level 2 Grandchild I', 'Level 2 Grandchild II'],
            },
          ],
        },
      ],
      {
        id: 'NestedListBox',
        x: 60,
        y: 60,
        width: 450,
        height: 250,
      }
    )

    const list = ppt.getList('NestedListBox')
    expect(list.length).toBe(2)
    expect(list[0]).toBe('Level 0 Item 1')
    expect(list[1]).toEqual({
      text: 'Level 0 Item 2 (Parent)',
      children: [
        'Level 1 Child A',
        {
          text: 'Level 1 Child B (Parent)',
          children: ['Level 2 Grandchild I', 'Level 2 Grandchild II'],
        },
      ],
    })

    const xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('lvl="0"')
    expect(xml).toContain('lvl="1"')
    expect(xml).toContain('lvl="2"')
  })

  it('2. addLists: manages bullet glyphs (<a:buChar>), colors (<a:buClr>), and sizing overrides', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addLists(['Custom Bullet Item A', 'Custom Bullet Item B'], {
      id: 'StyledBullets',
      x: 100,
      y: 100,
      width: 400,
      height: 120,
      bulletChar: '◆',
      bulletColor: '#E11D48',
      bulletFontSize: 110,
    })

    const xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('char="◆"')
    expect(xml).toContain('val="E11D48"')
    expect(xml).toContain('buSzPct')
    expect(xml).toContain('val="110000"')
  })

  it('3. getLists: discovers and returns all lists across targeted slides', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addLists(['List 1 Item A', 'List 1 Item B'], { id: 'ListOne', x: 50, y: 50 })
    ppt.addLists(['List 2 Item X', 'List 2 Item Y'], { id: 'ListTwo', x: 50, y: 200 })

    const allLists = ppt.getLists()
    expect(allLists.length).toBeGreaterThanOrEqual(2)

    const l1 = allLists.find(l => l.shapeName === 'ListOne')
    expect(l1).toBeDefined()
    expect(l1.items).toEqual(['List 1 Item A', 'List 1 Item B'])

    const l2 = allLists.find(l => l.shapeName === 'ListTwo')
    expect(l2).toBeDefined()
    expect(l2.items).toEqual(['List 2 Item X', 'List 2 Item Y'])
  })

  it('4. findLists: filters lists matching query across shape names and item contents', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addLists(['Architecture Guidelines', 'Memory Safety'], {
      id: 'ArchitectureDoc',
      x: 50,
      y: 50,
    })
    ppt.addLists(['Financial Ledger', 'Quarterly Revenue'], { id: 'FinanceDoc', x: 50, y: 200 })

    const archMatches = ppt.findLists('Architecture')
    expect(archMatches.length).toBe(1)
    expect(archMatches[0].shapeName).toBe('ArchitectureDoc')

    const finMatches = ppt.findLists('Revenue')
    expect(finMatches.length).toBe(1)
    expect(finMatches[0].shapeName).toBe('FinanceDoc')

    const notFound = ppt.findLists('NonExistentQuery')
    expect(notFound.length).toBe(0)
  })

  it('5. replaceList: updates existing list items non-destructively', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addLists(['Old Item 1', 'Old Item 2'], { id: 'DynamicList', x: 50, y: 50 })
    expect(ppt.getList('DynamicList')).toEqual(['Old Item 1', 'Old Item 2'])

    // Replace list with new items
    ppt.replaceList('DynamicList', ['New Item Alpha', 'New Item Beta', 'New Item Gamma'])

    const updated = ppt.getList('DynamicList')
    expect(updated).toEqual(['New Item Alpha', 'New Item Beta', 'New Item Gamma'])
    const xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('New Item Alpha')
    expect(xml).not.toContain('Old Item 1')
  })

  it('6. clearList: clears list items while retaining shape position and bullet properties', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addLists(['Item to clear A', 'Item to clear B'], {
      id: 'ClearableList',
      x: 1500000,
      y: 1200000,
      width: 3000000,
      height: 1500000,
      bulletChar: '✦',
    })

    let xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('Item to clear A')
    expect(xml).toContain('char="✦"')
    expect(xml).toContain('x="1500000"')

    // Clear list
    ppt.clearList('ClearableList')
    xml = ppt.slideManager.getSlideXml(1)

    expect(xml).not.toContain('Item to clear A')
    expect(xml).not.toContain('Item to clear B')
    // Shape position and bullet settings remain intact
    expect(xml).toContain('ClearableList')
    expect(xml).toContain('x="1500000"')
    expect(xml).toContain('y="1200000"')
    expect(xml).toContain('char="✦"')
  })

  it('7. removeList: removes list container shape from slide', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addLists(['Dead Item 1', 'Dead Item 2'], { id: 'TrashList', x: 50, y: 50 })
    let xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('TrashList')

    ppt.removeList('TrashList')
    xml = ppt.slideManager.getSlideXml(1)
    expect(xml).not.toContain('TrashList')
    expect(xml).not.toContain('Dead Item 1')
  })

  it('8. addStyle, updateStyle, and removeStyle: cleanly merge and mutate list and text styles', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addLists(['Styled Bullet 1', 'Styled Bullet 2'], {
      id: 'StyleList',
      x: 50,
      y: 50,
      fontSize: 16,
      color: '#333333',
    })

    // addStyle: apply bold and color
    ppt.addStyle('StyleList', { bold: true, color: '#2563EB' })
    let xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('b="1"')
    expect(xml).toContain('val="2563EB"')

    // updateStyle: change font size and alignment
    ppt.updateStyle('StyleList', { fontSize: 20, alignment: 'right' })
    xml = ppt.slideManager.getSlideXml(1)
    expect(xml).toContain('sz="2000"')
    expect(xml).toContain('algn="r"')

    // removeStyle: strip bold and alignment
    ppt.removeStyle('StyleList', ['bold', 'alignment'])
    xml = ppt.slideManager.getSlideXml(1)
    expect(xml).not.toContain('b="1"')
    expect(xml).not.toContain('algn="r"')

    // Verify round-trip file output
    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    const reloadedList = reloaded.getList('StyleList')
    expect(reloadedList).toEqual(['Styled Bullet 1', 'Styled Bullet 2'])
  })
})
