import { describe, it, expect } from 'vitest'
const { getImageMetadata } = require('../../src/utils/imageMetadata.js')

describe('imageMetadata.js', () => {
  it('should parse PNG metadata correctly', async () => {
    const pngHeader = Buffer.alloc(24)
    pngHeader.writeUInt32BE(0x89504e47, 0)
    pngHeader.writeUInt32BE(0x0d0a1a0a, 4)
    pngHeader.writeUInt32BE(13, 8) // chunk length
    pngHeader.writeUInt32BE(0x49484452, 12) // IHDR chunk type
    pngHeader.writeUInt32BE(800, 16) // width
    pngHeader.writeUInt32BE(600, 20) // height

    const meta = await getImageMetadata(pngHeader)
    expect(meta.type).toBe('png')
    expect(meta.width).toBe(800)
    expect(meta.height).toBe(600)
    expect(meta.aspectRatio).toBe(800 / 600)
  })

  it('should parse GIF metadata correctly', async () => {
    const gifHeader = Buffer.alloc(10)
    gifHeader.write('GIF89a', 0)
    gifHeader.writeUInt16LE(320, 6) // width
    gifHeader.writeUInt16LE(240, 8) // height

    const meta = await getImageMetadata(gifHeader)
    expect(meta.type).toBe('gif')
    expect(meta.width).toBe(320)
    expect(meta.height).toBe(240)
    expect(meta.aspectRatio).toBe(320 / 240)
  })

  it('should parse BMP metadata correctly', async () => {
    const bmpHeader = Buffer.alloc(26)
    bmpHeader.write('BM', 0)
    bmpHeader.writeInt32LE(100, 18) // width
    bmpHeader.writeInt32LE(200, 22) // height

    const meta = await getImageMetadata(bmpHeader)
    expect(meta.type).toBe('bmp')
    expect(meta.width).toBe(100)
    expect(meta.height).toBe(200)
    expect(meta.aspectRatio).toBe(100 / 200)
  })

  it('should parse JPEG metadata correctly', async () => {
    const jpegHeader = Buffer.alloc(30)
    jpegHeader.writeUInt16BE(0xffd8, 0) // SOI
    jpegHeader.writeUInt8(0xff, 2) // Marker start
    jpegHeader.writeUInt8(0xe0, 3) // APP0 Marker
    jpegHeader.writeUInt16BE(16, 4) // Segment Length
    // Pad App0 segment with zeros
    for (let i = 6; i < 20; i++) jpegHeader.writeUInt8(0, i)

    // SOF0 Marker
    jpegHeader.writeUInt8(0xff, 20)
    jpegHeader.writeUInt8(0xc0, 21) // SOF0
    jpegHeader.writeUInt16BE(11, 22) // SOF Length
    jpegHeader.writeUInt8(8, 24) // precision
    jpegHeader.writeUInt16BE(480, 25) // height
    jpegHeader.writeUInt16BE(640, 27) // width

    const meta = await getImageMetadata(jpegHeader)
    expect(meta.type).toBe('jpg')
    expect(meta.width).toBe(640)
    expect(meta.height).toBe(480)
    expect(meta.aspectRatio).toBe(640 / 480)
  })

  it('should parse SVG metadata correctly', async () => {
    const svgStr = '<svg width="1024" height="768" viewBox="0 0 1024 768"></svg>'
    const svgBuffer = Buffer.from(svgStr)

    const meta = await getImageMetadata(svgBuffer)
    expect(meta.type).toBe('svg')
    expect(meta.width).toBe(1024)
    expect(meta.height).toBe(768)
    expect(meta.aspectRatio).toBe(1024 / 768)
  })

  it('should fallback to viewBox for SVG if width/height are missing', async () => {
    const svgStr = '<svg viewBox="0 0 1600 900"></svg>'
    const svgBuffer = Buffer.from(svgStr)

    const meta = await getImageMetadata(svgBuffer)
    expect(meta.type).toBe('svg')
    expect(meta.width).toBe(1600)
    expect(meta.height).toBe(900)
  })
})
