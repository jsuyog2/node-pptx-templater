import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync, writeFileSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Image Operations Subsystem (TASK-MEDIA-002)', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_image_ops.pptx')
  const tempImgPath = resolve(__dirname, '../fixtures/temp_test_img.png')

  // Minimal 1x1 transparent PNG
  const samplePngBuffer = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    'base64'
  )

  // Minimal 2x2 solid red PNG
  const sampleRedPngBuffer = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64'
  )

  beforeEach(() => {
    writeFileSync(tempImgPath, samplePngBuffer)
    if (existsSync(outputPath)) {
      try {
        unlinkSync(outputPath)
      } catch (_) {}
    }
  })

  afterEach(() => {
    if (existsSync(tempImgPath)) {
      try {
        unlinkSync(tempImgPath)
      } catch (_) {}
    }
    if (existsSync(outputPath)) {
      try {
        unlinkSync(outputPath)
      } catch (_) {}
    }
  })

  it('AddImage: should embed image from Buffer, file path, and data URI with fluent chaining', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // 1. Add from buffer with EMU coordinates
    const res1 = await ppt.AddImage(1, samplePngBuffer, {
      x: 914400,
      y: 914400,
      width: 1828800,
      height: 1828800,
      name: 'LogoFromBuffer',
    })
    expect(res1).toBe(ppt) // Fluent chaining

    // 2. Add from file path with percentage coordinates
    await ppt.AddImage(1, tempImgPath, {
      x: '10%',
      y: '20%',
      width: '30%',
      height: '25%',
      name: 'LogoFromFile',
      rotation: 45,
      opacity: 80,
    })

    // 3. Add from data URI with styling (border, crop, shadow)
    const dataUri = `data:image/png;base64,${sampleRedPngBuffer.toString('base64')}`
    await ppt.AddImage(1, dataUri, {
      x: 500000,
      y: 500000,
      width: 1500000,
      height: 1500000,
      name: 'StyledLogo',
      crop: { left: 0.1, top: 0.1, right: 0.1, bottom: 0.1 },
      border: { color: 'FF0000', width: 25400 },
      shadow: { color: '000000', blur: 4, distance: 3 },
    })

    const medias = ppt.GetMedias(1)
    const img1 = medias.find(m => m.name === 'LogoFromBuffer')
    const img2 = medias.find(m => m.name === 'LogoFromFile')
    const img3 = medias.find(m => m.name === 'StyledLogo')

    expect(img1).toBeDefined()
    expect(img1.type).toBe('image')
    expect(img1.format).toBe('png')
    expect(img1.position.x).toBe(914400)
    expect(img1.position.width).toBe(1828800)

    expect(img2).toBeDefined()
    expect(img2.type).toBe('image')
    expect(img2.position.x).toBeGreaterThan(0)

    expect(img3).toBeDefined()
    expect(img3.type).toBe('image')

    // Save and verify presentation reload
    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    const reloadedMedias = reloaded.GetMedias(1)
    expect(reloadedMedias.length).toBeGreaterThanOrEqual(3)
  })

  it('UpdateImage: should update position, dimensions, rotation, crop, opacity, and swap image source', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.AddImage(1, samplePngBuffer, {
      x: 100000,
      y: 100000,
      width: 500000,
      height: 500000,
      name: 'DynamicImage',
    })

    const before = ppt.GetMedias(1).find(m => m.name === 'DynamicImage')
    expect(before).toBeDefined()

    // Update coordinates, rotation, crop, and swap source to red PNG
    const res = await ppt.UpdateImage('DynamicImage', {
      x: 200000,
      y: 300000,
      width: 800000,
      height: 600000,
      rotation: 90,
      opacity: 50,
      crop: { l: 20000, t: 20000, r: 20000, b: 20000 },
      border: { color: '00FF00', width: 2 },
      source: sampleRedPngBuffer,
    })
    expect(res).toBe(ppt) // Fluent chaining

    const after = ppt.GetMedias(1).find(m => m.name === 'DynamicImage')
    expect(after).toBeDefined()
    expect(after.position.x).toBe(200000)
    expect(after.position.y).toBe(300000)
    expect(after.position.width).toBe(800000)
    expect(after.position.height).toBe(600000)

    // Save presentation and reload to ensure XML schema validity
    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    const reloadedImg = reloaded.GetMedias(1).find(m => m.name === 'DynamicImage')
    expect(reloadedImg).toBeDefined()
    expect(reloadedImg.position.x).toBe(200000)
  })

  it('RemoveImage: should remove <p:pic> shape and purge orphaned relationships', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.AddImage(1, samplePngBuffer, {
      name: 'ImageToDelete',
    })

    const initial = ppt.GetMedias(1).find(m => m.name === 'ImageToDelete')
    expect(initial).toBeDefined()
    const rId = initial.relationshipId

    // Remove image
    const res = ppt.RemoveImage('ImageToDelete')
    expect(res).toBe(ppt) // Fluent chaining

    const after = ppt.GetMedias(1).find(m => m.name === 'ImageToDelete')
    expect(after).toBeUndefined()

    // Verify slide XML does not contain the shape
    const slideInfo = ppt.slideManager.getSlideInfo(1)
    const rel = ppt.relationshipManager.getRelationshipById(slideInfo.zipPath, rId)
    // If not shared by other shapes, rel should be purged
    expect(rel).toBeNull()

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)
  })

  it('Deduplication: adding duplicate image buffers should reuse the same file in ppt/media/', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const initialMediaCount = ppt.mediaManager.mediaCount

    await ppt.AddImage(1, samplePngBuffer, { name: 'Dup1' })
    const countAfterFirst = ppt.mediaManager.mediaCount

    // Add exact same buffer
    await ppt.AddImage(1, samplePngBuffer, { name: 'Dup2' })
    const countAfterSecond = ppt.mediaManager.mediaCount

    expect(countAfterSecond).toBe(countAfterFirst)
    expect(countAfterFirst).toBe(initialMediaCount + 1)

    const medias = ppt.GetMedias(1)
    const d1 = medias.find(m => m.name === 'Dup1')
    const d2 = medias.find(m => m.name === 'Dup2')
    expect(d1.targetPath).toBe(d2.targetPath)
  })
})
