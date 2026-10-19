import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'
import { getImageMetadata } from '../../src/utils/imageMetadata.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Images & Media: Base64, WebP, Metadata Extraction', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_media.pptx')

  // Minimal 1x1 transparent PNG as Base64 data URI
  const base64Png =
    'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

  // Minimal 1x1 WebP (VP8 lossless)
  const webpBuffer = Buffer.from([
    0x52, 0x49, 0x46, 0x46, 0x1a, 0x00, 0x00, 0x00, 0x57, 0x45, 0x42, 0x50, 0x56, 0x50, 0x38, 0x4c,
    0x0d, 0x00, 0x00, 0x00, 0x2f, 0x00, 0x00, 0x00, 0x00, 0x07, 0x00, 0x00, 0x00, 0x00,
  ])

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

  it('should parse metadata from Base64 data URIs and WebP buffers', async () => {
    const pngMeta = await getImageMetadata(base64Png)
    expect(pngMeta.type).toBe('png')
    expect(pngMeta.width).toBe(1)
    expect(pngMeta.height).toBe(1)

    const webpMeta = await getImageMetadata(webpBuffer)
    expect(webpMeta.type).toBe('webp')
    expect(webpMeta.width).toBe(1)
    expect(webpMeta.height).toBe(1)
  })

  it('should embed an image from a Base64 data URI and save presentation', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.AddImage(1, base64Png, {
      x: 50,
      y: 50,
      width: 100,
      height: 100,
      name: 'Base64Logo',
    })

    const images = ppt.GetMedias(1)
    expect(images.length).toBeGreaterThan(0)
    const embedded = images.find(img => img.name === 'Base64Logo')
    expect(embedded).toBeDefined()
    expect(embedded.format).toBe('png')
    expect(embedded.mimeType).toBe('image/png')
    expect(embedded.rId).toMatch(/^rId/i)

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    const reloadedImages = reloaded.GetMedias(1)
    const found = reloadedImages.find(img => img.name === 'Base64Logo')
    expect(found).toBeDefined()
  })

  it('should embed WebP image buffer and extract complete image metadata', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.AddImage(1, webpBuffer, {
      x: 100,
      y: 100,
      width: 80,
      height: 80,
      name: 'WebpImage',
    })

    const images = ppt.GetMedias(1)
    const webpImg = images.find(img => img.name === 'WebpImage')
    expect(webpImg).toBeDefined()
    expect(webpImg.format).toBe('webp')
    expect(webpImg.mimeType).toBe('image/webp')
    expect(webpImg.size).toBe(webpBuffer.length)
  })
})
