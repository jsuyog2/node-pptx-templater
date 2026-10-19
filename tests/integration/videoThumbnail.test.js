import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import { existsSync, unlinkSync } from 'fs'
import JSZip from 'jszip'
import { PPTXTemplater } from '../../src/index.js'
import { getDefaultVideoPoster } from '../../src/utils/mediaUtils.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Video Thumbnail & Green Box Prevention (TASK-VIDEO-THUMB)', () => {
  const samplePath = resolve(__dirname, '../fixtures/sample.pptx')
  const outputPath = resolve(__dirname, '../fixtures/output_video_thumb.pptx')

  // Valid minimal mock MP4
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

  // Custom 2x2 solid red PNG poster
  const customPoster1 = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
    'base64'
  )

  // Custom 2x2 solid blue PNG poster for update
  const customPoster2 = Buffer.from(
    'iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAYAAABytg0kAAAAFElEQVR42mNkYPj/HwADBwGAAYGQegAAAABJRU5ErkJggg==',
    'base64'
  )

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

  it('TASK-VIDEO-THUMB-001: addVideo without poster embeds opaque default poster (no green box)', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.useSlide(1).addVideo(mockMp4Buffer, {
      x: 100000,
      y: 100000,
      width: 4000000,
      height: 3000000,
      name: 'DefaultPosterVideo',
    })

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)

    // 1. Verify [Content_Types].xml contains png default
    const contentTypesXml = await zip.file('[Content_Types].xml').async('text')
    expect(contentTypesXml).toContain('Extension="png"')

    // 2. Verify slide1.xml.rels has an image relationship for the poster
    const relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')
    expect(relsXml).toContain('relationships/image')
    expect(relsXml).toContain('relationships/video')
    expect(relsXml).toContain('relationships/media')

    // 3. Inspect slide1.xml: check <p:blipFill><a:blip> does NOT have cstate="print"
    const slide1Xml = await zip.file('ppt/slides/slide1.xml').async('text')
    expect(slide1Xml).not.toContain('cstate="print"')
    expect(slide1Xml).toContain('<a:blip')

    // 4. Verify getMedias exposes posterPath and posterRId
    const medias = ppt.useSlide(1).getMedias()
    const video = medias.find(m => m.name === 'DefaultPosterVideo')
    expect(video).toBeDefined()
    expect(video.posterRId).toBeDefined()
    expect(video.posterPath).toBeDefined()
    expect(video.posterPath).toContain('ppt/media/')

    // 5. Verify the default poster was saved in ppt/media/ and matches defaultVideoPoster length
    const defaultPosterBuf = getDefaultVideoPoster()
    expect(defaultPosterBuf.length).toBeGreaterThan(500)
    const embeddedPosterData = await zip.file(video.posterPath).async('nodebuffer')
    expect(embeddedPosterData.length).toBe(defaultPosterBuf.length)
  })

  it('TASK-VIDEO-THUMB-001: addVideo with custom poster binds image relationship correctly', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.useSlide(1).addVideo(mockMp4Buffer, {
      x: '10%',
      y: '10%',
      width: '50%',
      height: '40%',
      name: 'CustomPosterVideo',
      poster: customPoster1,
    })

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)

    const medias = ppt.useSlide(1).getMedias()
    const video = medias.find(m => m.name === 'CustomPosterVideo')
    expect(video).toBeDefined()
    expect(video.posterPath).toBeDefined()

    const posterData = await zip.file(video.posterPath).async('nodebuffer')
    expect(posterData.equals(customPoster1)).toBe(true)
  })

  it('TASK-VIDEO-THUMB-002: updateVideo poster replacement swaps thumbnail and purges orphan binary', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    // 1. Add video with initial customPoster1
    await ppt.useSlide(1).addVideo(mockMp4Buffer, {
      name: 'UpdatableVideo',
      poster: customPoster1,
    })

    let medias = ppt.useSlide(1).getMedias()
    let video = medias.find(m => m.name === 'UpdatableVideo')
    const initialPosterPath = video.posterPath
    const initialPosterRId = video.posterRId
    expect(initialPosterPath).toBeDefined()

    // 2. Update video with customPoster2
    await ppt.useSlide(1).updateVideo('UpdatableVideo', {
      poster: customPoster2,
    })

    medias = ppt.useSlide(1).getMedias()
    video = medias.find(m => m.name === 'UpdatableVideo')
    expect(video.posterRId).not.toBe(initialPosterRId)
    expect(video.posterPath).not.toBe(initialPosterPath)

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)

    // 3. New poster is present with correct content
    const newPosterData = await zip.file(video.posterPath).async('nodebuffer')
    expect(newPosterData.equals(customPoster2)).toBe(true)

    // 4. Old poster binary was purged from zip (no orphaned files)
    expect(zip.file(initialPosterPath)).toBeNull()

    // 5. Old relationship ID is removed from slide1.xml.rels
    const relsXml = await zip.file('ppt/slides/_rels/slide1.xml.rels').async('text')
    expect(relsXml).not.toContain(`Id="${initialPosterRId}"`)
    expect(relsXml).toContain(`Id="${video.posterRId}"`)
  })

  it('TASK-VIDEO-THUMB-002: removeVideo purges video, poster, and timings cleanly', async () => {
    const ppt = await PPTXTemplater.load(samplePath)

    await ppt.useSlide(1).addVideo(mockMp4Buffer, {
      name: 'VideoToRemove',
      poster: customPoster1,
      autoplay: true,
    })

    const medias = ppt.useSlide(1).getMedias()
    const video = medias.find(m => m.name === 'VideoToRemove')
    const posterPath = video.posterPath
    const videoPath = video.targetPath

    // Remove the video
    ppt.useSlide(1).removeVideo('VideoToRemove')

    const remaining = ppt.useSlide(1).getMedias()
    expect(remaining.find(m => m.name === 'VideoToRemove')).toBeUndefined()

    const buffer = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buffer)

    // Video and poster binaries should be purged since they were only used by this shape
    expect(zip.file(posterPath)).toBeNull()
    expect(zip.file(videoPath)).toBeNull()

    // Slide XML should no longer contain video shape
    const slideXml = await zip.file('ppt/slides/slide1.xml').async('text')
    expect(slideXml).not.toContain('VideoToRemove')
  })
})
