import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync, writeFileSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Strict 12 Images & Media Functions Suite (TASK-MEDIA-006)', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_images_media_suite.pptx')
  const tempVideo = resolve(__dirname, '../fixtures/temp_suite_vid.mp4')
  const tempAudio = resolve(__dirname, '../fixtures/temp_suite_aud.mp3')

  const samplePngBuffer = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    'base64'
  )

  const mockMp4 = Buffer.from([
    0x00, 0x00, 0x00, 0x18, 0x66, 0x74, 0x79, 0x70, 0x69, 0x73, 0x6f, 0x6d, 0x00, 0x00, 0x02, 0x00,
    0x69, 0x73, 0x6f, 0x6d, 0x6d, 0x70, 0x34, 0x32,
  ])

  const mockMp3 = Buffer.from([
    0x49, 0x44, 0x33, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x0a, 0x54, 0x49, 0x54, 0x32, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00,
  ])

  beforeEach(() => {
    writeFileSync(tempVideo, mockMp4)
    writeFileSync(tempAudio, mockMp3)
    if (existsSync(outputPath)) {
      try {
        unlinkSync(outputPath)
      } catch (_) {}
    }
  })

  afterEach(() => {
    if (existsSync(tempVideo)) {
      try {
        unlinkSync(tempVideo)
      } catch (_) {}
    }
    if (existsSync(tempAudio)) {
      try {
        unlinkSync(tempAudio)
      } catch (_) {}
    }
    if (existsSync(outputPath)) {
      try {
        unlinkSync(outputPath)
      } catch (_) {}
    }
  })

  it('Happy Path: executes complete lifecycle for Images, Videos, and Audio with fluent chaining', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // 1. AddImage
    await ppt.AddImage(1, samplePngBuffer, {
      x: 0,
      y: 0,
      width: 1000000,
      height: 1000000,
      name: 'LifeCycleImage',
    })

    // 2. UpdateImage
    await ppt.UpdateImage('LifeCycleImage', {
      x: 200000,
      y: 200000,
      rotation: 180,
    })

    // 3. AddVideo
    await ppt.AddVideo(1, tempVideo, {
      x: '20%',
      y: '20%',
      width: '40%',
      height: '30%',
      name: 'LifeCycleVideo',
      autoplay: true,
      loop: true,
    })

    // 4. UpdateVideo
    await ppt.UpdateVideo('LifeCycleVideo', {
      mute: true,
      startTime: 1.0,
      endTime: 5.0,
    })

    // 5. AddAudio
    await ppt.AddAudio(1, tempAudio, {
      x: '70%',
      y: '70%',
      width: '10%',
      height: '10%',
      name: 'LifeCycleAudio',
      volume: 60,
    })

    // 6. UpdateAudio
    await ppt.UpdateAudio('LifeCycleAudio', {
      volume: 90,
      loop: true,
    })

    // 7. GetMedias
    const medias = ppt.GetMedias(1)
    expect(medias.length).toBeGreaterThanOrEqual(3)

    // 8. GetMedia
    const foundImg = ppt.GetMedia('LifeCycleImage')
    expect(foundImg).toBeDefined()
    expect(foundImg.type).toBe('image')

    const foundVid = ppt.GetMedia('LifeCycleVideo')
    expect(foundVid).toBeDefined()
    expect(foundVid.type).toBe('video')
    expect(foundVid.mute).toBe(true)

    const foundAud = ppt.GetMedia('LifeCycleAudio')
    expect(foundAud).toBeDefined()
    expect(foundAud.type).toBe('audio')
    expect(foundAud.volume).toBe(90)

    // 9. CloneMedia
    ppt.CloneMedia('LifeCycleImage', 'ClonedLifeCycleImage', {
      x: 500000,
      y: 500000,
    })
    const cloned = ppt.GetMedia('ClonedLifeCycleImage')
    expect(cloned).toBeDefined()
    expect(cloned.position.x).toBe(500000)

    // 10. RemoveVideo
    ppt.RemoveVideo('LifeCycleVideo')
    expect(ppt.GetMedia('LifeCycleVideo')).toBeNull()

    // 11. RemoveAudio
    ppt.RemoveAudio('LifeCycleAudio')
    expect(ppt.GetMedia('LifeCycleAudio')).toBeNull()

    // 12. RemoveImage
    ppt.RemoveImage('LifeCycleImage')
    expect(ppt.GetMedia('LifeCycleImage')).toBeNull()

    // Ensure cloned copy still survives
    expect(ppt.GetMedia('ClonedLifeCycleImage')).toBeDefined()

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    expect(reloaded.GetMedia('ClonedLifeCycleImage')).toBeDefined()
  })

  it('Edge Case: missing media query and removal handles gracefully without throwing', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Get non-existent media
    const missing = ppt.GetMedia('DefinitelyDoesNotExist12345')
    expect(missing).toBeNull()

    // Update non-existent media should return ppt without error
    const resUpdateImg = await ppt.UpdateImage('NoSuchImg', { x: 100 })
    expect(resUpdateImg).toBe(ppt)

    const resUpdateVid = await ppt.UpdateVideo('NoSuchVid', { x: 100 })
    expect(resUpdateVid).toBe(ppt)

    const resUpdateAud = await ppt.UpdateAudio('NoSuchAud', { x: 100 })
    expect(resUpdateAud).toBe(ppt)

    // Remove non-existent media should return ppt without error
    const resRemImg = ppt.RemoveImage('NoSuchImg')
    expect(resRemImg).toBe(ppt)

    const resRemVid = ppt.RemoveVideo('NoSuchVid')
    expect(resRemVid).toBe(ppt)

    const resRemAud = ppt.RemoveAudio('NoSuchAud')
    expect(resRemAud).toBe(ppt)
  })

  it('Edge Case: boundary coordinates (0, 100%, 0%, large values)', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // 0 coordinates
    await ppt.AddImage(1, samplePngBuffer, {
      x: '0%',
      y: '0%',
      width: '100%',
      height: '100%',
      name: 'FullCanvasImage',
    })

    const fullImg = ppt.GetMedia('FullCanvasImage')
    expect(fullImg).toBeDefined()
    expect(fullImg.position.x).toBe(0)
    expect(fullImg.position.y).toBe(0)
    expect(fullImg.position.width).toBeGreaterThan(5000000)

    // Very large coordinates clamp properly and build valid XML
    await ppt.AddImage(1, samplePngBuffer, {
      x: 10000000,
      y: 5000000,
      width: 2000000,
      height: 1000000,
      name: 'OffCanvasImage',
    })

    const offImg = ppt.GetMedia('OffCanvasImage')
    expect(offImg).toBeDefined()
    expect(offImg.position.x).toBe(10000000)

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)
  })

  it('Edge Case: Deduplication preserves media count across multiple clones and duplicate adds', async () => {
    const ppt = await PPTXTemplater.load(samplePath)
    const initialCount = ppt.mediaManager.mediaCount

    await ppt.AddImage(1, samplePngBuffer, { name: 'SeedImage' })
    const countAfterAdd = ppt.mediaManager.mediaCount
    expect(countAfterAdd).toBe(initialCount + 1)

    // Clone 3 times
    ppt.CloneMedia('SeedImage', 'Clone1')
    ppt.CloneMedia('SeedImage', 'Clone2')
    ppt.CloneMedia('SeedImage', 'Clone3')

    // Media count must remain constant
    expect(ppt.mediaManager.mediaCount).toBe(countAfterAdd)

    const medias = ppt.GetMedias(1)
    const seed = medias.find(m => m.name === 'SeedImage')
    const c1 = medias.find(m => m.name === 'Clone1')
    const c2 = medias.find(m => m.name === 'Clone2')
    const c3 = medias.find(m => m.name === 'Clone3')

    expect(seed).toBeDefined()
    expect(c1).toBeDefined()
    expect(c2).toBeDefined()
    expect(c3).toBeDefined()

    expect(c1.targetPath).toBe(seed.targetPath)
    expect(c2.targetPath).toBe(seed.targetPath)
    expect(c3.targetPath).toBe(seed.targetPath)
  })

  it('useSlide() Context: camelCase methods work seamlessly without slideId parameter', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Fluent chained useSlide(1) context
    await ppt.useSlide(1).addImage(samplePngBuffer, {
      x: '10%',
      y: '10%',
      width: '20%',
      height: '20%',
      name: 'ContextualImage',
    })

    await ppt.useSlide(1).addVideo(tempVideo, {
      x: '35%',
      y: '10%',
      width: '30%',
      height: '25%',
      name: 'ContextualVideo',
      autoplay: true,
      loop: true,
    })

    await ppt.useSlide(1).addAudio(tempAudio, {
      x: '70%',
      y: '10%',
      width: '10%',
      height: '10%',
      name: 'ContextualAudio',
      volume: 75,
    })

    // Query medias on active slide
    const activeMedias = ppt.useSlide(1).getMedias()
    expect(activeMedias.some(m => m.name === 'ContextualImage')).toBe(true)
    expect(activeMedias.some(m => m.name === 'ContextualVideo')).toBe(true)
    expect(activeMedias.some(m => m.name === 'ContextualAudio')).toBe(true)

    // Update methods
    await ppt.useSlide(1).updateImage('ContextualImage', { rotation: 45 })
    await ppt.useSlide(1).updateVideo('ContextualVideo', { mute: true })
    await ppt.useSlide(1).updateAudio('ContextualAudio', { volume: 50 })

    const updatedImg = ppt.getMedia('ContextualImage')
    expect(updatedImg.rotation).toBe(45)

    const updatedVid = ppt.getMedia('ContextualVideo')
    expect(updatedVid.mute).toBe(true)

    const updatedAud = ppt.getMedia('ContextualAudio')
    expect(updatedAud.volume).toBe(50)

    // Clone and removal
    ppt.useSlide(1).cloneMedia('ContextualImage', 'ContextualImageClone', { x: '50%' })
    expect(ppt.getMedia('ContextualImageClone')).toBeDefined()

    ppt.useSlide(1).removeVideo('ContextualVideo')
    expect(ppt.getMedia('ContextualVideo')).toBeNull()

    ppt.useSlide(1).removeAudio('ContextualAudio')
    expect(ppt.getMedia('ContextualAudio')).toBeNull()

    ppt.useSlide(1).removeImage('ContextualImage')
    expect(ppt.getMedia('ContextualImage')).toBeNull()

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)
  })

  it('Base64 Input Support: embeds media from Data URIs and raw Base64 strings', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    const dataUriPng = `data:image/png;base64,${samplePngBuffer.toString('base64')}`
    const rawBase64Mp4 = mockMp4.toString('base64')
    const dataUriMp3 = `data:audio/mp3;base64,${mockMp3.toString('base64')}`

    await ppt.useSlide(1).addImage(dataUriPng, {
      x: '5%',
      y: '5%',
      width: '15%',
      height: '15%',
      name: 'Base64Image',
    })

    await ppt.useSlide(1).addVideo(rawBase64Mp4, {
      x: '25%',
      y: '5%',
      width: '25%',
      height: '25%',
      name: 'Base64Video',
    })

    await ppt.useSlide(1).addAudio(dataUriMp3, {
      x: '55%',
      y: '5%',
      width: '10%',
      height: '10%',
      name: 'Base64Audio',
    })

    const b64Img = ppt.getMedia('Base64Image')
    expect(b64Img).toBeDefined()
    expect(b64Img.format).toBe('png')

    const b64Vid = ppt.getMedia('Base64Video')
    expect(b64Vid).toBeDefined()
    expect(b64Vid.format).toBe('mp4')

    const b64Aud = ppt.getMedia('Base64Audio')
    expect(b64Aud).toBeDefined()
    expect(b64Aud.format).toBe('mp3')

    // Test source swapping with base64
    const newPng = Buffer.from(
      'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAADklEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
      'base64'
    )
    await ppt.useSlide(1).updateImage('Base64Image', { source: newPng })
    const swappedImg = ppt.getMedia('Base64Image')
    expect(swappedImg).toBeDefined()

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)
  })
})
