import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync, writeFileSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Video Subsystem (TASK-MEDIA-003)', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_video_ops.pptx')
  const tempVideoPath = resolve(__dirname, '../fixtures/sample_test_video.mp4')
  const tempPosterPath = resolve(__dirname, '../fixtures/sample_poster.png')

  // Minimal mock MP4 (ftyp box)
  const mockMp4Buffer = Buffer.from([
    0x00,
    0x00,
    0x00,
    0x18,
    0x66,
    0x74,
    0x79,
    0x70, // size 24, 'ftyp'
    0x69,
    0x73,
    0x6f,
    0x6d,
    0x00,
    0x00,
    0x02,
    0x00, // 'isom', minor 512
    0x69,
    0x73,
    0x6f,
    0x6d,
    0x6d,
    0x70,
    0x34,
    0x32, // compatible brands
  ])

  // Minimal 1x1 PNG poster
  const posterPngBuffer = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==',
    'base64'
  )

  beforeEach(() => {
    writeFileSync(tempVideoPath, mockMp4Buffer)
    writeFileSync(tempPosterPath, posterPngBuffer)
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
    if (existsSync(tempPosterPath)) {
      try {
        unlinkSync(tempPosterPath)
      } catch (_) {}
    }
    if (existsSync(outputPath)) {
      try {
        unlinkSync(outputPath)
      } catch (_) {}
    }
  })

  it('AddVideo: should embed video with poster, timing, playback settings and fluent chaining', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // Add video with poster, loop, mute, autoplay, startTime, endTime
    const res = await ppt.AddVideo(1, tempVideoPath, {
      x: '15%',
      y: '20%',
      width: '60%',
      height: '50%',
      name: 'HeroVideo',
      poster: tempPosterPath,
      loop: true,
      mute: true,
      autoplay: true,
      startTime: 1.5,
      endTime: 10.0,
    })
    expect(res).toBe(ppt) // Fluent chaining

    const medias = ppt.GetMedias(1)
    const video = medias.find(m => m.name === 'HeroVideo')

    expect(video).toBeDefined()
    expect(video.type).toBe('video')
    expect(video.format).toBe('mp4')
    expect(video.mimeType).toBe('video/mp4')
    expect(video.loop).toBe(true)
    expect(video.mute).toBe(true)
    expect(video.startTime).toBe(1.5)
    expect(video.endTime).toBe(10)
    expect(video.position.x).toBeGreaterThan(0)

    // Save and verify reloaded integrity
    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    const reloadedMedias = reloaded.GetMedias(1)
    const reloadedVideo = reloadedMedias.find(m => m.name === 'HeroVideo')
    expect(reloadedVideo).toBeDefined()
    expect(reloadedVideo.type).toBe('video')
    expect(reloadedVideo.loop).toBe(true)
  })

  it('UpdateVideo: should update bounds, playback properties, and poster frame', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.AddVideo(1, mockMp4Buffer, {
      name: 'PromoVideo',
      x: 100000,
      y: 100000,
      width: 1000000,
      height: 600000,
    })

    const initial = ppt.GetMedias(1).find(m => m.name === 'PromoVideo')
    expect(initial).toBeDefined()

    const res = await ppt.UpdateVideo('PromoVideo', {
      x: 200000,
      y: 250000,
      width: 1200000,
      height: 750000,
      loop: true,
      mute: true,
      autoplay: true,
      startTime: 2,
      endTime: 8,
      poster: tempPosterPath,
    })
    expect(res).toBe(ppt) // Fluent chaining

    const updated = ppt.GetMedias(1).find(m => m.name === 'PromoVideo')
    expect(updated).toBeDefined()
    expect(updated.position.x).toBe(200000)
    expect(updated.position.y).toBe(250000)
    expect(updated.position.width).toBe(1200000)
    expect(updated.position.height).toBe(750000)
    expect(updated.loop).toBe(true)
    expect(updated.mute).toBe(true)
    expect(updated.startTime).toBe(2)
    expect(updated.endTime).toBe(8)

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)
  })

  it('RemoveVideo: should remove video shape, timing sequences, and clean relationships', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.AddVideo(1, mockMp4Buffer, {
      name: 'VideoToRemove',
      autoplay: true,
    })

    const before = ppt.GetMedias(1).find(m => m.name === 'VideoToRemove')
    expect(before).toBeDefined()
    const rId = before.relationshipId

    const res = ppt.RemoveVideo('VideoToRemove')
    expect(res).toBe(ppt) // Fluent chaining

    const after = ppt.GetMedias(1).find(m => m.name === 'VideoToRemove')
    expect(after).toBeUndefined()

    // Verify relationship was cleaned up
    const slideInfo = ppt.slideManager.getSlideInfo(1)
    const rel = ppt.relationshipManager.getRelationshipById(slideInfo.zipPath, rId)
    expect(rel).toBeNull()

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    const reloadedVideo = reloaded.GetMedias(1).find(m => m.name === 'VideoToRemove')
    expect(reloadedVideo).toBeUndefined()
  })
})
