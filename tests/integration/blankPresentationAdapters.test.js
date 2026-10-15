import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'
import fs from 'fs-extra'
import path from 'path'
import { Writable } from 'stream'

const TEMP_DIR = path.resolve(__dirname, '../temp-blank-adapters')

describe('Blank Presentation Export Integrity Across All Adapters', () => {
  beforeEach(async () => {
    await fs.remove(TEMP_DIR)
    await fs.ensureDir(TEMP_DIR)
  })

  afterEach(async () => {
    await fs.remove(TEMP_DIR)
  })

  it('should export cleanly via saveToFile() and validate with zero errors or warnings', async () => {
    const ppt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })
    ppt.useSlide(1)
    await ppt.addShape('roundedRectangle', {
      x: 100,
      y: 100,
      width: 400,
      height: 200,
      fill: '#0284C7',
      text: 'Export Adapter: saveToFile',
    })

    const filePath = path.join(TEMP_DIR, 'adapter-saveToFile.pptx')
    await ppt.saveToFile(filePath)

    expect(await fs.pathExists(filePath)).toBe(true)
    const reloaded = await PPTXTemplater.load(filePath)
    expect(reloaded.slideCount).toBe(1)

    const val = await reloaded.validatePresentation()
    expect(val).toEqual({ valid: true, errors: [], warnings: [] })
  })

  it('should export cleanly via toBuffer() and re-ingest with valid presentation schema', async () => {
    const ppt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })
    ppt.useSlide(1)
    await ppt.addShape('rectangle', {
      x: 50,
      y: 50,
      width: 300,
      height: 100,
      text: 'Export Adapter: toBuffer',
    })

    const buffer = await ppt.toBuffer()
    expect(Buffer.isBuffer(buffer)).toBe(true)
    expect(buffer.length).toBeGreaterThan(2000)

    const reloaded = await PPTXTemplater.load(buffer)
    expect(reloaded.slideCount).toBe(1)
    const val = await reloaded.validatePresentation()
    expect(val).toEqual({ valid: true, errors: [], warnings: [] })
  })

  it('should export cleanly via toStream() to a writable stream destination', async () => {
    const ppt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })
    ppt.useSlide(1)
    await ppt.addShape('rectangle', {
      x: 100,
      y: 100,
      width: 200,
      height: 100,
      text: 'Export Adapter: toStream',
    })

    const stream = await ppt.toStream()
    const chunks = []
    const writable = new Writable({
      write(chunk, encoding, callback) {
        chunks.push(chunk)
        callback()
      },
    })

    await new Promise((resolve, reject) => {
      stream.pipe(writable)
      writable.on('finish', resolve)
      writable.on('error', reject)
      stream.on('error', reject)
    })

    const streamBuffer = Buffer.concat(chunks)
    expect(streamBuffer.length).toBeGreaterThan(2000)

    const reloaded = await PPTXTemplater.load(streamBuffer)
    expect(reloaded.slideCount).toBe(1)
    const val = await reloaded.validatePresentation()
    expect(val).toEqual({ valid: true, errors: [], warnings: [] })
  })

  it('should export cleanly via toArrayBuffer() with matching byteLength and validity', async () => {
    const ppt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })
    const arrayBuffer = await ppt.toArrayBuffer()

    expect(arrayBuffer).toBeInstanceOf(ArrayBuffer)
    expect(arrayBuffer.byteLength).toBeGreaterThan(2000)

    const buffer = Buffer.from(arrayBuffer)
    const reloaded = await PPTXTemplater.load(buffer)
    const val = await reloaded.validatePresentation()
    expect(val).toEqual({ valid: true, errors: [], warnings: [] })
  })

  it('should stream to Express.js response object without corrupt ZIP or schema warnings', async () => {
    const ppt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })
    ppt.useSlide(1)
    await ppt.addShape('rectangle', {
      x: 100,
      y: 100,
      width: 300,
      height: 150,
      text: 'Export Adapter: toExpress',
    })

    const chunks = []
    const headers = {}
    const mockRes = new Writable({
      write(chunk, encoding, callback) {
        chunks.push(chunk)
        callback()
      },
    })
    mockRes.setHeader = vi.fn((k, v) => {
      headers[k.toLowerCase()] = v
    })

    await ppt.toExpress(mockRes, 'express-deck.pptx')

    await new Promise(resolve => mockRes.on('finish', resolve))

    expect(mockRes.setHeader).toHaveBeenCalledWith(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )
    expect(headers['content-disposition']).toBe('attachment; filename="express-deck.pptx"')

    const exportedBuffer = Buffer.concat(chunks)
    const reloaded = await PPTXTemplater.load(exportedBuffer)
    expect(reloaded.slideCount).toBe(1)

    const val = await reloaded.validatePresentation()
    expect(val).toEqual({ valid: true, errors: [], warnings: [] })
  })

  it('should send via Fastify reply adapter without corrupt ZIP or schema warnings', async () => {
    const ppt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' })
    ppt.useSlide(1)
    await ppt.addShape('rectangle', {
      x: 100,
      y: 100,
      width: 300,
      height: 150,
      text: 'Export Adapter: toFastify',
    })

    let sentBuffer = null
    const headers = {}
    const mockReply = {
      header: vi.fn((k, v) => {
        headers[k.toLowerCase()] = v
        return mockReply
      }),
      send: vi.fn(val => {
        sentBuffer = val
        return mockReply
      }),
    }

    await ppt.toFastify(mockReply, 'fastify-deck.pptx')

    expect(mockReply.header).toHaveBeenCalledWith(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )
    expect(headers['content-disposition']).toBe('attachment; filename="fastify-deck.pptx"')
    expect(mockReply.send).toHaveBeenCalled()
    expect(sentBuffer).not.toBeNull()

    const reloaded = await PPTXTemplater.load(sentBuffer)
    expect(reloaded.slideCount).toBe(1)

    const val = await reloaded.validatePresentation()
    expect(val).toEqual({ valid: true, errors: [], warnings: [] })
  })
})
