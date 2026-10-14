import { describe, it, expect, vi } from 'vitest'
import { PPTXTemplater } from '../../src/index.js'
import fs from 'fs-extra'
import path from 'path'

const TEMP_DIR = path.resolve(__dirname, '../temp-output-writer')

describe('OutputWriter & Framework Adapters', () => {
  it('should export to Node.js Buffer and ArrayBuffer', async () => {
    const ppt = await PPTXTemplater.createBlank()
    const buffer = await ppt.toBuffer()
    expect(Buffer.isBuffer(buffer)).toBe(true)
    expect(buffer.length).toBeGreaterThan(1000)

    const arrayBuffer = await ppt.toArrayBuffer()
    expect(arrayBuffer).toBeInstanceOf(ArrayBuffer)
    expect(arrayBuffer.byteLength).toBe(buffer.length)
  })

  it('should export to standard Base64 string and Data URI', async () => {
    const ppt = await PPTXTemplater.createBlank()
    const b64 = await ppt.toBase64()
    expect(typeof b64).toBe('string')
    expect(b64.length).toBeGreaterThan(1000)

    const dataUri = await ppt.toBase64({ dataUri: true })
    expect(dataUri.startsWith('data:application/vnd.openxmlformats-officedocument.presentationml.presentation;base64,')).toBe(true)
  })

  it('should export to Web Standard Response for Edge / Serverless runtimes', async () => {
    const ppt = await PPTXTemplater.createBlank()
    const response = await ppt.toResponse({ filename: 'custom-export.pptx' })

    expect(response).toBeInstanceOf(Response)
    expect(response.status).toBe(200)
    expect(response.headers.get('Content-Type')).toBe(
      'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )
    expect(response.headers.get('Content-Disposition')).toContain('custom-export.pptx')

    const arrayBuffer = await response.arrayBuffer()
    expect(arrayBuffer.byteLength).toBeGreaterThan(1000)
  })

  it('should stream to Express.js response object with correct headers', async () => {
    const ppt = await PPTXTemplater.createBlank()
    const headers = {}
    const mockRes = {
      setHeader: vi.fn((k, v) => {
        headers[k] = v
      }),
      on: vi.fn(),
      once: vi.fn(),
      emit: vi.fn(),
      write: vi.fn(),
      end: vi.fn(),
    }

    await ppt.toExpress(mockRes, 'monthly-deck.pptx')
    expect(mockRes.setHeader).toHaveBeenCalledWith(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )
    expect(mockRes.setHeader).toHaveBeenCalledWith(
      'Content-Disposition',
      'attachment; filename="monthly-deck.pptx"'
    )
  })

  it('should send via Fastify reply adapter', async () => {
    const ppt = await PPTXTemplater.createBlank()
    const headers = {}
    const mockReply = {
      header: vi.fn((k, v) => {
        headers[k] = v
      }),
      send: vi.fn(val => val),
    }

    await ppt.toFastify(mockReply, 'fastify-deck.pptx')
    expect(mockReply.header).toHaveBeenCalledWith(
      'Content-Type',
      'application/vnd.openxmlformats-officedocument.presentationml.presentation'
    )
    expect(mockReply.header).toHaveBeenCalledWith(
      'Content-Disposition',
      'attachment; filename="fastify-deck.pptx"'
    )
    expect(mockReply.send).toHaveBeenCalled()
  })

  it('should extract raw presentation parts to folder with formatted XML', async () => {
    await fs.remove(TEMP_DIR)
    await fs.ensureDir(TEMP_DIR)

    const ppt = await PPTXTemplater.createBlank()
    const extractFolder = path.join(TEMP_DIR, 'unpacked')

    await ppt.extractToXml(extractFolder, { formatXml: true })

    const presentationXmlPath = path.join(extractFolder, 'ppt/presentation.xml')
    expect(await fs.pathExists(presentationXmlPath)).toBe(true)

    const xmlContent = await fs.readFile(presentationXmlPath, 'utf8')
    // Check that indentation exists in formatted XML
    expect(xmlContent).toContain('\n')
    expect(xmlContent).toMatch(/^\s+<p:/m)

    // Check Content_Types exists
    expect(await fs.pathExists(path.join(extractFolder, '[Content_Types].xml'))).toBe(true)
    expect(await fs.pathExists(path.join(extractFolder, 'ppt/slides/slide1.xml'))).toBe(true)

    await fs.remove(TEMP_DIR)
  })

  it('should package to ZIP archive via extractToZip', async () => {
    await fs.remove(TEMP_DIR)
    await fs.ensureDir(TEMP_DIR)

    const ppt = await PPTXTemplater.createBlank()
    const zipPath = path.join(TEMP_DIR, 'extracted.zip')

    await ppt.extractToZip(zipPath)
    expect(await fs.pathExists(zipPath)).toBe(true)

    // Load back the extracted zip to verify validity
    const loaded = await PPTXTemplater.load(zipPath)
    expect(loaded.slideCount).toBe(1)

    await fs.remove(TEMP_DIR)
  })

  it('should support toFlatXml to disk file and string', async () => {
    await fs.ensureDir(TEMP_DIR)
    const ppt = await PPTXTemplater.createBlank()
    const xmlPath = path.join(TEMP_DIR, 'out.xml')

    const xml = await ppt.toFlatXml(xmlPath)
    expect(typeof xml).toBe('string')
    expect(await fs.pathExists(xmlPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(xmlPath)
    expect(reloaded.slideCount).toBe(1)
    await fs.remove(TEMP_DIR)
  })
})
