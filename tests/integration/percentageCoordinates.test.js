import { describe, it, expect } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'

describe('CoordinateUtils - Global Percentage Coordinates (0-100%)', () => {
  it('should correctly position shapes using percentage coordinates and 4 corners', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    // Top-Left corner: (0, 0)
    await ppt.addShape('rect', {
      id: 'TopLeft',
      x: '0%',
      y: '0%',
      width: '10%',
      height: '10%',
    })

    // Top-Right corner: (100%, 0)
    await ppt.addShape('rect', {
      id: 'TopRight',
      x: '90%',
      y: '0%',
      width: '10%',
      height: '10%',
    })

    // Bottom-Left corner: (0, 100%)
    await ppt.addShape('rect', {
      id: 'BottomLeft',
      x: 0,
      y: 90,
      unit: '%',
      width: 10,
      height: 10,
    })

    // Bottom-Right corner: (100%, 100%)
    await ppt.addShape('rect', {
      id: 'BottomRight',
      x: '90%',
      y: '90%',
      width: '10%',
      height: '10%',
    })

    const slideInfo = ppt.getSlide(1)
    const elements = slideInfo.elements

    const tl = elements.find(e => e.id === 'TopLeft' || e.name === 'TopLeft')
    expect(tl).toBeDefined()
    expect(tl.bounds.x).toBe(0)
    expect(tl.bounds.y).toBe(0)

    const tr = elements.find(e => e.id === 'TopRight' || e.name === 'TopRight')
    expect(tr).toBeDefined()
    // 90% of 12192000 = 10972800
    expect(tr.bounds.x).toBe(10972800)
    expect(tr.bounds.y).toBe(0)

    const bl = elements.find(e => e.id === 'BottomLeft' || e.name === 'BottomLeft')
    expect(bl).toBeDefined()
    expect(bl.bounds.x).toBe(0)
    // 90% of 6858000 = 6172200
    expect(bl.bounds.y).toBe(6172200)

    const br = elements.find(e => e.id === 'BottomRight' || e.name === 'BottomRight')
    expect(br).toBeDefined()
    expect(br.bounds.x).toBe(10972800)
    expect(br.bounds.y).toBe(6172200)
  })

  it('should support percentage positioning in addTable and addTextAtPosition', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    ppt.addTable(
      [
        ['Header 1', 'Header 2'],
        ['Val 1', 'Val 2'],
      ],
      {
        x: '20%',
        y: '20%',
        width: '60%',
        height: '40%',
      }
    )

    ppt.addTextAtPosition('Sample Title', {
      x: '10%',
      y: '5%',
      width: '80%',
      height: '10%',
      fontSize: 24,
    })

    await ppt.zipManager.waitForPendingWrites()

    const slideInfo = ppt.getSlide(1)
    const tbl = slideInfo.elements.find(e => e.type === 'table')
    expect(tbl).toBeDefined()
    // 20% of 12192000 = 2438400
    expect(tbl.bounds.x).toBe(2438400)
    // 20% of 6858000 = 1371600
    expect(tbl.bounds.y).toBe(1371600)

    const txt = slideInfo.elements.find(e => e.type === 'textBox' || e.type === 'shape')
    expect(txt).toBeDefined()
  })

  it('should support percentage bounds in setBounds', async () => {
    const ppt = await PPTXTemplater.createBlank()
    ppt.useSlide(1)

    await ppt.addShape('rect', {
      id: 'TargetCard',
      x: 100,
      y: 100,
      width: 200,
      height: 200,
    })

    ppt.setBounds('TargetCard', {
      x: '50%',
      y: '50%',
      width: '25%',
      height: '25%',
    })

    const slideInfo = ppt.getSlide(1)
    const card = slideInfo.elements.find(e => e.id === 'TargetCard' || e.name === 'TargetCard')
    expect(card).toBeDefined()
    // 50% of 12192000 = 6096000
    expect(card.bounds.x).toBe(6096000)
    // 50% of 6858000 = 3429000
    expect(card.bounds.y).toBe(3429000)
  })
})
