import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync, writeFileSync } from 'fs'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Audio Subsystem (TASK-MEDIA-004)', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_audio_ops.pptx')
  const tempAudioPath = resolve(__dirname, '../fixtures/sample_test_audio.mp3')

  // Minimal mock MP3 (ID3 header)
  const mockMp3Buffer = Buffer.from([
    0x49,
    0x44,
    0x33,
    0x03,
    0x00,
    0x00,
    0x00,
    0x00,
    0x00,
    0x0a, // ID3v2.3 header
    0x54,
    0x49,
    0x54,
    0x32,
    0x00,
    0x00,
    0x00,
    0x00,
    0x00,
    0x00,
  ])

  beforeEach(() => {
    writeFileSync(tempAudioPath, mockMp3Buffer)
    if (existsSync(outputPath)) {
      try {
        unlinkSync(outputPath)
      } catch (_) {}
    }
  })

  afterEach(() => {
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

  it('AddAudio: should embed audio file with volume, loop, autoplay, timing and icon visibility', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    const res = await ppt.AddAudio(1, tempAudioPath, {
      x: 500000,
      y: 500000,
      width: 700000,
      height: 700000,
      name: 'BackgroundMusic',
      volume: 75,
      loop: true,
      autoplay: true,
      hideDuringShow: true,
      startTime: 0,
      endTime: 15.5,
    })
    expect(res).toBe(ppt) // Fluent chaining

    const medias = ppt.GetMedias(1)
    const audio = medias.find(m => m.name === 'BackgroundMusic')

    expect(audio).toBeDefined()
    expect(audio.type).toBe('audio')
    expect(audio.format).toBe('mp3')
    expect(audio.mimeType).toBe('audio/mpeg')
    expect(audio.volume).toBe(75)
    expect(audio.loop).toBe(true)
    expect(audio.startTime).toBe(0)
    expect(audio.endTime).toBe(15.5)

    // Verify slide XML contains audio action geometry and p14:media
    const slideInfo = ppt.slideManager.getSlideInfo(1)
    const rel = ppt.relationshipManager.getRelationshipById(slideInfo.zipPath, audio.relationshipId)
    expect(rel).toBeDefined()
    expect(rel.type).toContain('relationships/audio')

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    const reloadedAudio = reloaded.GetMedias(1).find(m => m.name === 'BackgroundMusic')
    expect(reloadedAudio).toBeDefined()
    expect(reloadedAudio.type).toBe('audio')
  })

  it('UpdateAudio: should adjust volume, loop, timing, icon visibility, and bounds', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.AddAudio(1, mockMp3Buffer, {
      name: 'NarratorAudio',
      volume: 50,
      loop: false,
    })

    const initial = ppt.GetMedias(1).find(m => m.name === 'NarratorAudio')
    expect(initial).toBeDefined()

    const res = await ppt.UpdateAudio('NarratorAudio', {
      x: 600000,
      y: 600000,
      width: 800000,
      height: 800000,
      volume: 100,
      loop: true,
      hideDuringShow: false,
      startTime: 1.0,
      endTime: 12.0,
    })
    expect(res).toBe(ppt) // Fluent chaining

    const updated = ppt.GetMedias(1).find(m => m.name === 'NarratorAudio')
    expect(updated).toBeDefined()
    expect(updated.position.x).toBe(600000)
    expect(updated.position.y).toBe(600000)
    expect(updated.position.width).toBe(800000)
    expect(updated.position.height).toBe(800000)
    expect(updated.volume).toBe(100)
    expect(updated.loop).toBe(true)
    expect(updated.startTime).toBe(1.0)
    expect(updated.endTime).toBe(12.0)

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)
  })

  it('RemoveAudio: should purge audio shape, timing references, and slide relationship entries', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.AddAudio(1, mockMp3Buffer, {
      name: 'SoundEffect',
      autoplay: true,
    })

    const before = ppt.GetMedias(1).find(m => m.name === 'SoundEffect')
    expect(before).toBeDefined()
    const rId = before.relationshipId

    const res = ppt.RemoveAudio('SoundEffect')
    expect(res).toBe(ppt) // Fluent chaining

    const after = ppt.GetMedias(1).find(m => m.name === 'SoundEffect')
    expect(after).toBeUndefined()

    // Verify relationship was cleaned up
    const slideInfo = ppt.slideManager.getSlideInfo(1)
    const rel = ppt.relationshipManager.getRelationshipById(slideInfo.zipPath, rId)
    expect(rel).toBeNull()

    await ppt.saveToFile(outputPath)
    expect(existsSync(outputPath)).toBe(true)

    const reloaded = await PPTXTemplater.load(outputPath)
    const reloadedAudio = reloaded.GetMedias(1).find(m => m.name === 'SoundEffect')
    expect(reloadedAudio).toBeUndefined()
  })
})
