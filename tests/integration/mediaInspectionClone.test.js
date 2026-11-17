import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync, writeFileSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Media Inspection & Cloning (TASK-MEDIA-005)', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_media_inspect_clone.pptx')
  const tempVideoPath = resolve(__dirname, '../fixtures/sample_inspect_video.mp4')
  const tempAudioPath = resolve(__dirname, '../fixtures/sample_inspect_audio.mp3')

  const samplePngBuffer = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    'base64'
  )

  const mockMp4Buffer = Buffer.from([
    0x00, 0x00, 0x00, 0x18, 0x66, 0x74, 0x79, 0x70, 0x69, 0x73, 0x6f, 0x6d, 0x00, 0x00, 0x02, 0x00,
    0x69, 0x73, 0x6f, 0x6d, 0x6d, 0x70, 0x34, 0x32,
  ])

  const mockMp3Buffer = Buffer.from([
    0x49, 0x44, 0x33, 0x03, 0x00, 0x00, 0x00, 0x00, 0x00, 0x0a, 0x54, 0x49, 0x54, 0x32, 0x00, 0x00,
    0x00, 0x00, 0x00, 0x00,
  ])

  beforeEach(() => {
    writeFileSync(tempVideoPath, mockMp4Buffer)
    writeFileSync(tempAudioPath, mockMp3Buffer)
    if (existsSync(outputPath)) {
      try {
        unlinkSync(outputPath)
      } catch (_) {}
    }
  })

  afterEach(() => {
    if (existsSync(tempVideoPath)) {
      try {
        unlinkSync(tempVideoPath)
      } catch (_) {}
    }
    if (existsSync(tempAudioPath)) {
      try {
        unlinkSync(tempAudioPath)
      } catch (_) {}
    }
    if (existsSync(outputPath)) {
      try {
        unlinkSync(outputPath)
      } catch (_) {}
    }
  })

  it('GetMedias: should query all media elements (image, video, audio) with comprehensive metadata', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.AddImage(1, samplePngBuffer, {
      x: 100000,
      y: 100000,
      width: 500000,
      height: 400000,
      name: 'TestImage_Inspect',
    })

    await ppt.AddVideo(1, tempVideoPath, {
      x: 700000,
      y: 100000,
      width: 1500000,
      height: 900000,
      name: 'TestVideo_Inspect',
      loop: true,
      mute: true,
      startTime: 0,
      endTime: 5.0,
    })

    await ppt.AddAudio(1, tempAudioPath, {
      x: 700000,
      y: 1100000,
      width: 400000,
      height: 400000,
      name: 'TestAudio_Inspect',
      volume: 80,
      loop: true,
    })

    const medias = ppt.GetMedias(1)
    expect(medias.length).toBeGreaterThanOrEqual(3)

    const img = medias.find(m => m.name === 'TestImage_Inspect')
    expect(img).toBeDefined()
    expect(img.type).toBe('image')
    expect(img.format).toBe('png')
    expect(img.position.x).toBe(100000)
    expect(img.position.y).toBe(100000)
    expect(img.position.width).toBe(500000)
    expect(img.position.height).toBe(400000)

    const vid = medias.find(m => m.name === 'TestVideo_Inspect')
    expect(vid).toBeDefined()
    expect(vid.type).toBe('video')
    expect(vid.format).toBe('mp4')
    expect(vid.loop).toBe(true)
    expect(vid.mute).toBe(true)
    expect(vid.startTime).toBe(0)
    expect(vid.endTime).toBe(5.0)

    const aud = medias.find(m => m.name === 'TestAudio_Inspect')
    expect(aud).toBeDefined()
    expect(aud.type).toBe('audio')
    expect(aud.format).toBe('mp3')
    expect(aud.volume).toBe(80)
    expect(aud.loop).toBe(true)

    // Global inspection without slideId
    const allMedias = ppt.GetMedias()
    expect(allMedias.length).toBeGreaterThanOrEqual(medias.length)
  })

  it('GetMedia: should locate a specific media element by ID, name, or relationship ID, and return null if not found', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.AddImage(1, samplePngBuffer, {
      name: 'FindMeImage',
    })

    // By name
    const byName = ppt.GetMedia('FindMeImage')
    expect(byName).toBeDefined()
    expect(byName.name).toBe('FindMeImage')
    expect(byName.type).toBe('image')

    // By shape ID
    const byId = ppt.GetMedia(byName.id)
    expect(byId).toBeDefined()
    expect(byId.name).toBe('FindMeImage')

    // By relationship ID
    const byRel = ppt.GetMedia(byName.relationshipId)
    expect(byRel).toBeDefined()
    expect(byRel.name).toBe('FindMeImage')

    // Non-existent ID returns null
    const notFound = ppt.GetMedia('NonExistentMediaElement999')
    expect(notFound).toBeNull()
  })

  it('CloneMedia: should duplicate media element with overrides and deduplicate binary streams', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.AddImage(1, samplePngBuffer, {
      x: 100000,
      y: 100000,
      width: 400000,
      height: 300000,
      name: 'SourceImageToClone',
    })

    const initialMediaCount = ppt.mediaManager.mediaCount

    // Clone the image
    const res = ppt.CloneMedia('SourceImageToClone', 'ClonedImageCopy', {
      x: 600000,
      y: 700000,
      width: 800000,
      height: 600000,
      rotation: 30,
    })
    expect(res).toBe(ppt) // Fluent chaining

    // Deduplication check: media count should not increase
    expect(ppt.mediaManager.mediaCount).toBe(initialMediaCount)

    const cloned = ppt.GetMedia('ClonedImageCopy')
    expect(cloned).toBeDefined()
    expect(cloned.name).toBe('ClonedImageCopy')
    expect(cloned.type).toBe('image')
    expect(cloned.position.x).toBe(600000)
    expect(cloned.position.y).toBe(700000)
    expect(cloned.position.width).toBe(800000)
    expect(cloned.position.height).toBe(600000)

    // Save and reload to verify package validity
    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    const reloadedClone = reloaded.GetMedia('ClonedImageCopy')
    expect(reloadedClone).toBeDefined()
    expect(reloadedClone.position.x).toBe(600000)
  })

  it('CloneMedia: should duplicate video and preserve playback settings', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.AddVideo(1, tempVideoPath, {
      name: 'SourceVideoToClone',
      loop: true,
      mute: true,
      startTime: 1.0,
      endTime: 4.0,
    })

    const initialMediaCount = ppt.mediaManager.mediaCount

    ppt.CloneMedia('SourceVideoToClone', 'ClonedVideoCopy', {
      x: 500000,
      y: 500000,
    })

    // Deduplication verification
    expect(ppt.mediaManager.mediaCount).toBe(initialMediaCount)

    const clonedVideo = ppt.GetMedia('ClonedVideoCopy')
    expect(clonedVideo).toBeDefined()
    expect(clonedVideo.type).toBe('video')
    expect(clonedVideo.loop).toBe(true)
    expect(clonedVideo.mute).toBe(true)
    expect(clonedVideo.startTime).toBe(1.0)
    expect(clonedVideo.endTime).toBe(4.0)
    expect(clonedVideo.position.x).toBe(500000)
  })
})
