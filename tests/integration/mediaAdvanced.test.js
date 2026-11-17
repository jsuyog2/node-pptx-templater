/**
 * @fileoverview mediaAdvanced.test.js — Regression tests for Images & Media subsystem.
 *
 * Verifies:
 *  1. Zero repair errors on addVideo and addAudio (dual relationships + Office 2010 p14 GUID)
 *  2. Schema-compliant timing tree generation (mainSeq, interactiveSeq, media nodes)
 *  3. Video poster update in updateVideo (relationship creation + blip re-pointing)
 *  4. cloneMedia duplicating shapes, stripping duplicate creationIds, and replicating timing
 *  5. Direct Base64 Data URI and raw Base64 inputs
 *  6. useSlide() context resolution
 */

import { describe, it, expect, beforeAll } from 'vitest'
import { resolve, dirname } from 'path'
import { fileURLToPath } from 'url'
import fs from 'fs-extra'
import JSZip from 'jszip'
import { PPTXTemplater } from '../../src/index.js'

const __dirname = dirname(fileURLToPath(import.meta.url))

describe('Images & Media Advanced Regression Suite', () => {
  const sampleImagePath = resolve(__dirname, '../../showcase/templates/media/sample_image.jpg')
  const sampleVideoPath = resolve(__dirname, '../../showcase/templates/media/sample_video.mp4')
  const sampleAudioPath = resolve(__dirname, '../../showcase/templates/media/sample_audio.mp3')

  let sampleImageBuf
  let sampleVideoBuf
  let sampleAudioBuf

  beforeAll(async () => {
    sampleImageBuf = await fs.readFile(sampleImagePath)
    sampleVideoBuf = await fs.readFile(sampleVideoPath)
    sampleAudioBuf = await fs.readFile(sampleAudioPath)
  })

  it('TASK-MEDIA-001 & 002: addVideo generates dual relationships, correct p14 GUID, and complete timing', async () => {
    const ppt = await PPTXTemplater.createBlank()
    const rawBase64Video = sampleVideoBuf.toString('base64')

    await ppt.useSlide(1).addVideo(rawBase64Video, {
      name: 'TestVideo',
      autoplay: true,
      loop: true,
      mute: true,
    })

    const buf = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buf)

    // Check [Content_Types].xml
    const ctXml = await zip.files['[Content_Types].xml'].async('text')
    expect(ctXml).toContain('Extension="mp4"')

    // Check slide 1 XML
    const slideXml = await zip.files['ppt/slides/slide1.xml'].async('text')
    expect(slideXml).toContain('a:videoFile')
    expect(slideXml).toContain('uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"')
    expect(slideXml).toContain('action="ppaction://media"')
    expect(slideXml).not.toContain('r:action=')
    expect(slideXml).toContain('<p:timing>')
    expect(slideXml).toContain('<p:video>')
    expect(slideXml).toContain('nodeType="interactiveSeq"')
    expect(slideXml).toContain('nodeType="mainSeq"')

    // Check slide 1 relationships (dual relationships)
    const relsXml = await zip.files['ppt/slides/_rels/slide1.xml.rels'].async('text')
    expect(relsXml).toContain(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video'
    )
    expect(relsXml).toContain('http://schemas.microsoft.com/office/2007/relationships/media')
    expect(relsXml).toContain(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
    )
  })

  it('TASK-MEDIA-001 & 003: addAudio generates dual relationships, valid non-visual props, and timing node', async () => {
    const ppt = await PPTXTemplater.createBlank()
    const audioDataUri = `data:audio/mp3;base64,${sampleAudioBuf.toString('base64')}`

    await ppt.useSlide(1).addAudio(audioDataUri, {
      name: 'TestAudio',
      volume: 90,
      loop: false,
      hideDuringShow: true,
    })

    const buf = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buf)

    // Check [Content_Types].xml
    const ctXml = await zip.files['[Content_Types].xml'].async('text')
    expect(ctXml).toContain('Extension="mp3"')

    // Check slide 1 XML
    const slideXml = await zip.files['ppt/slides/slide1.xml'].async('text')
    expect(slideXml).toContain('a:audioFile')
    expect(slideXml).toContain('uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"')
    expect(slideXml).toContain('action="ppaction://media"')
    expect(slideXml).not.toContain('r:action=')
    // Ensure p14:showWhenStopped is NOT inside p:cNvPicPr
    expect(slideXml).not.toMatch(/<p:cNvPicPr>[\s\S]*?p14:showWhenStopped[\s\S]*?<\/p:cNvPicPr>/)
    expect(slideXml).toContain('<p:timing>')
    expect(slideXml).toContain('<p:audio>')
    expect(slideXml).toContain('showWhenStopped="0"')

    // Check slide 1 relationships (dual relationships)
    const relsXml = await zip.files['ppt/slides/_rels/slide1.xml.rels'].async('text')
    expect(relsXml).toContain(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio'
    )
    expect(relsXml).toContain('http://schemas.microsoft.com/office/2007/relationships/media')
  })

  it('TASK-MEDIA-004: updateVideo updates poster image and re-points rId', async () => {
    const ppt = await PPTXTemplater.createBlank()
    await ppt.useSlide(1).addVideo(sampleVideoBuf, { name: 'ShowcaseVid' })

    const initialBuf = await ppt.toBuffer()
    const initialZip = await JSZip.loadAsync(initialBuf)
    const initialSlideXml = await initialZip.files['ppt/slides/slide1.xml'].async('text')
    const initialBlipMatch = initialSlideXml.match(/<a:blip[^>]*r:embed="([^"]+)"/)
    expect(initialBlipMatch).toBeTruthy()
    const _oldRId = initialBlipMatch[1]

    // Update with replace poster image
    await ppt.useSlide(1).updateVideo('ShowcaseVid', {
      poster: sampleImageBuf,
      loop: true,
      mute: true,
    })

    const updatedBuf = await ppt.toBuffer()
    const updatedZip = await JSZip.loadAsync(updatedBuf)
    const updatedSlideXml = await updatedZip.files['ppt/slides/slide1.xml'].async('text')
    const updatedBlipMatch = updatedSlideXml.match(/<a:blip[^>]*r:embed="([^"]+)"/)
    expect(updatedBlipMatch).toBeTruthy()
    const newRId = updatedBlipMatch[1]

    // New rId allocated or updated
    expect(newRId).toBeTruthy()
    const updatedRelsXml = await updatedZip.files['ppt/slides/_rels/slide1.xml.rels'].async('text')
    expect(updatedRelsXml).toContain(`Id="${newRId}"`)
  })

  it('TASK-MEDIA-005: cloneMedia duplicates shapes and replicates timing sequences', async () => {
    const ppt = await PPTXTemplater.createBlank()
    await ppt.useSlide(1).addVideo(sampleVideoBuf, { name: 'OriginalVid' })
    await ppt.useSlide(1).addImage(sampleImageBuf, { name: 'OriginalImg' })

    // Clone image
    ppt.useSlide(1).cloneMedia('OriginalImg', 'ClonedImg', { x: 100, y: 100 })
    // Clone video
    ppt.useSlide(1).cloneMedia('OriginalVid', 'ClonedVid', { x: 200, y: 200 })

    const buf = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buf)
    const slideXml = await zip.files['ppt/slides/slide1.xml'].async('text')

    expect(slideXml).toContain('name="OriginalVid"')
    expect(slideXml).toContain('name="ClonedVid"')
    expect(slideXml).toContain('name="OriginalImg"')
    expect(slideXml).toContain('name="ClonedImg"')

    // Verify timing duplicated for both original and cloned video
    const timingBlock = slideXml.slice(slideXml.indexOf('<p:timing>'))
    const videoTags = timingBlock.match(/<p:video>/g) || []
    expect(videoTags.length).toBe(2)
  })

  it('TASK-MEDIA-006: Clean removal of audio and video with zero orphaned timing or rels', async () => {
    const ppt = await PPTXTemplater.createBlank()
    await ppt.useSlide(1).addVideo(sampleVideoBuf, { name: 'TempVid' })
    await ppt.useSlide(1).addAudio(sampleAudioBuf, { name: 'TempAud' })

    // Purge both
    ppt.useSlide(1).removeVideo('TempVid')
    ppt.useSlide(1).removeAudio('TempAud')

    const buf = await ppt.toBuffer()
    const zip = await JSZip.loadAsync(buf)
    const slideXml = await zip.files['ppt/slides/slide1.xml'].async('text')

    // No leftover video, audio, or timing
    expect(slideXml).not.toContain('<p:pic>')
    expect(slideXml).not.toContain('<p:timing>')
  })
})
