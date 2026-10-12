import { describe, it, expect, vi } from 'vitest'
import {
  PackageLoader,
  LegacyFormatConverter,
  isLegacyPpt,
  isOdp,
  isPkZip,
  VirtualFileSystem,
  PPTXTemplater,
} from '../../src/index.js'
import { Readable } from 'stream'

describe('PackageLoader & LegacyFormatConverter', () => {
  it('should detect PKZip magic bytes correctly', () => {
    const validZip = Buffer.from([0x50, 0x4b, 0x03, 0x04, 0x00, 0x00])
    const invalidZip = Buffer.from([0x00, 0x01, 0x02, 0x03])
    expect(isPkZip(validZip)).toBe(true)
    expect(isPkZip(invalidZip)).toBe(false)
  })

  it('should detect legacy PPT (OLE CFBF) magic bytes correctly', () => {
    const oleBytes = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0x00])
    expect(isLegacyPpt(oleBytes)).toBe(true)
    expect(LegacyFormatConverter.isLegacyPpt(oleBytes)).toBe(true)
    expect(isLegacyPpt(Buffer.from([0x50, 0x4b, 0x03, 0x04]))).toBe(false)
  })

  it('should normalize legacy binary PPT into an OpenXML VirtualFileSystem', async () => {
    // Construct mock binary PPT with CFBF header and text
    const oleHeader = Buffer.from([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])
    const textData = Buffer.from('Legacy Quarterly Presentation\x00Revenue Highlights\x00Growth Strategy\x00')
    const pptBuffer = Buffer.concat([oleHeader, textData])

    expect(await LegacyFormatConverter.supports(pptBuffer)).toBe(true)
    expect(await LegacyFormatConverter.detectFormat(pptBuffer)).toBe('ppt')

    const vfs = await LegacyFormatConverter.convert(pptBuffer)
    expect(vfs).toBeInstanceOf(VirtualFileSystem)
    expect(vfs.has('[Content_Types].xml')).toBe(true)
    expect(vfs.has('ppt/presentation.xml')).toBe(true)
    expect(vfs.has('ppt/slides/slide1.xml')).toBe(true)

    const slide1Xml = vfs.readText('ppt/slides/slide1.xml')
    expect(slide1Xml).toContain('Legacy Quarterly Presentation')

    // Ingest directly via PPTXTemplater.load
    const ppt = await PPTXTemplater.load(pptBuffer)
    expect(ppt.slideCount).toBe(1)
  })

  it('should detect and normalize ODP into an OpenXML VirtualFileSystem', async () => {
    const JSZip = (await import('jszip')).default
    const zip = new JSZip()
    zip.file('mimetype', 'application/vnd.oasis.opendocument.presentation')
    zip.file(
      'content.xml',
      `<?xml version="1.0" encoding="UTF-8"?>
<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0">
  <office:body>
    <office:presentation>
      <draw:page draw:name="Slide 1">
        <draw:frame><draw:text-box><text:p>ODP Slide Title</text:p><text:p>First Bullet Point</text:p></draw:text-box></draw:frame>
      </draw:page>
      <draw:page draw:name="Slide 2">
        <draw:frame><draw:text-box><text:p>Second Slide Title</text:p><text:p>Second Bullet Point</text:p></draw:text-box></draw:frame>
      </draw:page>
    </office:presentation>
  </office:body>
</office:document-content>`
    )

    const odpBuffer = await zip.generateAsync({ type: 'nodebuffer' })
    expect(await isOdp(odpBuffer)).toBe(true)
    expect(await LegacyFormatConverter.detectFormat(odpBuffer)).toBe('odp')

    const vfs = await LegacyFormatConverter.convert(odpBuffer)
    expect(vfs.has('[Content_Types].xml')).toBe(true)
    expect(vfs.has('ppt/slides/slide1.xml')).toBe(true)
    expect(vfs.has('ppt/slides/slide2.xml')).toBe(true)

    // Load via PPTXTemplater.load
    const ppt = await PPTXTemplater.load(odpBuffer)
    expect(ppt.slideCount).toBe(2)
  })

  it('should ingest presentations via Node.js Readable stream', async () => {
    const blank = await PPTXTemplater.createBlank()
    const buffer = await blank.toBuffer()

    const createStream = () => {
      const s = new Readable()
      s.push(buffer)
      s.push(null)
      return s
    }

    const vfs = await PackageLoader.load(createStream())
    expect(vfs.has('[Content_Types].xml')).toBe(true)
    expect(vfs.has('ppt/presentation.xml')).toBe(true)

    const ppt = await PPTXTemplater.load(createStream())
    expect(ppt.slideCount).toBeGreaterThanOrEqual(1)
  })

  it('should support fromUrl loading with mock fetch', async () => {
    const blank = await PPTXTemplater.createBlank()
    const buffer = await blank.toBuffer()

    const mockFetch = vi.fn().mockResolvedValue({
      ok: true,
      status: 200,
      statusText: 'OK',
      arrayBuffer: async () => buffer,
    })

    const originalFetch = globalThis.fetch
    globalThis.fetch = mockFetch

    try {
      const ppt = await PPTXTemplater.fromUrl('https://example.com/remote-presentation.pptx')
      expect(ppt).toBeInstanceOf(PPTXTemplater)
      expect(ppt.slideCount).toBe(1)
      expect(mockFetch).toHaveBeenCalledWith(
        'https://example.com/remote-presentation.pptx',
        expect.any(Object)
      )
    } finally {
      globalThis.fetch = originalFetch
    }
  })

  it('should support fromCloud ingestion with S3 client mock', async () => {
    const blank = await PPTXTemplater.createBlank()
    const buffer = await blank.toBuffer()

    const mockS3Client = {
      send: vi.fn().mockResolvedValue({
        Body: buffer,
      }),
    }

    const ppt = await PPTXTemplater.fromCloud({
      provider: 's3',
      client: mockS3Client,
      bucket: 'corporate-templates',
      key: 'pitch-deck.pptx',
    })

    expect(ppt).toBeInstanceOf(PPTXTemplater)
    expect(ppt.slideCount).toBe(1)
    expect(mockS3Client.send).toHaveBeenCalled()
  })

  it('should support fromCloud ingestion with GCS storage mock', async () => {
    const blank = await PPTXTemplater.createBlank()
    const buffer = await blank.toBuffer()

    const mockGcsStorage = {
      bucket: vi.fn().mockReturnValue({
        file: vi.fn().mockReturnValue({
          download: vi.fn().mockResolvedValue([buffer]),
        }),
      }),
    }

    const ppt = await PPTXTemplater.fromCloud({
      provider: 'gcs',
      storage: mockGcsStorage,
      bucket: 'my-gcs-bucket',
      key: 'report.pptx',
    })

    expect(ppt).toBeInstanceOf(PPTXTemplater)
    expect(ppt.slideCount).toBe(1)
  })

  it('should support fromCloud ingestion with Azure blobClient mock', async () => {
    const blank = await PPTXTemplater.createBlank()
    const buffer = await blank.toBuffer()

    const mockBlobClient = {
      downloadToBuffer: vi.fn().mockResolvedValue(buffer),
    }

    const ppt = await PPTXTemplater.fromCloud({
      provider: 'azure',
      blobClient: mockBlobClient,
    })

    expect(ppt).toBeInstanceOf(PPTXTemplater)
    expect(ppt.slideCount).toBe(1)
  })
})
