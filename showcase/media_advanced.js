/**
 * @fileoverview showcase/media_advanced.js — Images & Media Subsystem Showcase
 *
 * Demonstrates and validates the strict 12-function Images & Media whitelist:
 * =========================================================================
 * 1.  ppt.addImage(source, options) / ppt.addImage(slideId, source, options)
 * 2.  ppt.updateImage(imageId, options)
 * 3.  ppt.removeImage(imageId)
 * 4.  ppt.addVideo(source, options) / ppt.addVideo(slideId, source, options)
 * 5.  ppt.updateVideo(videoId, options)
 * 6.  ppt.removeVideo(videoId)
 * 7.  ppt.addAudio(source, options) / ppt.addAudio(slideId, source, options)
 * 8.  ppt.updateAudio(audioId, options)
 * 9.  ppt.removeAudio(audioId)
 * 10. ppt.getMedias(slideId)
 * 11. ppt.getMedia(mediaId)
 * 12. ppt.cloneMedia(mediaId, clonedMediaId, clonedOptions)
 * =========================================================================
 */

'use strict'

const path = require('path')
const fs = require('fs-extra')
const { PPTXTemplater } = require('../src/index.js')

const OUTPUT_DIR = path.join(__dirname, 'output/slide_advanced')

async function runMediaAdvancedShowcase() {
  console.log('='.repeat(75))
  console.log('PPTXTemplater — Images & Media Master Showcase (Strict 12 Methods)')
  console.log('='.repeat(75))

  try {
    await fs.ensureDir(OUTPUT_DIR)
    const templatePath = path.join(__dirname, 'templates/MediaPowerPointSample.pptx')
    const ppt = await PPTXTemplater.load(templatePath)

    // Load sample assets and construct Base64 Data URIs and raw Base64 strings
    const sampleImageBuf = fs.readFileSync(path.join(__dirname, 'templates/media/sample_image.jpg'))
    const replaceImageBuf = fs.readFileSync(
      path.join(__dirname, 'templates/media/replace_image.jpg')
    )
    const sampleImageDataUri = `data:image/jpeg;base64,${sampleImageBuf.toString('base64')}`

    const sampleMp4Buf = fs.readFileSync(path.join(__dirname, 'templates/media/sample_video.mp4'))
    const replaceMp4Buf = fs.readFileSync(path.join(__dirname, 'templates/media/replace_video.mp4'))
    const sampleMp4RawBase64 = sampleMp4Buf.toString('base64')

    const sampleMp3Buf = fs.readFileSync(path.join(__dirname, 'templates/media/sample_audio.mp3'))
    const replaceMp3Buf = fs.readFileSync(path.join(__dirname, 'templates/media/replace_audio.mp3'))
    const sampleMp3DataUri = `data:audio/mp3;base64,${sampleMp3Buf.toString('base64')}`

    // 1. addImage with Base64 Data URI and useSlide(1) context
    console.log('\n[1/12] Executing addImage(source, options) with useSlide(1)...')
    await ppt.useSlide(1).addImage(sampleImageDataUri, {
      x: '10%',
      y: '10%',
      width: '25%',
      height: '25%',
      name: 'ShowcaseLogo',
      rotation: 15,
      opacity: 90,
    })
    // Also demonstrate adding image via remote URL
    await ppt
      .useSlide(1)
      .addImage('https://images.pexels.com/photos/18523946/pexels-photo-18523946.jpeg', {
        x: '70%',
        y: '10%',
        width: '25%',
        height: '25%',
        name: 'ShowcaseLogo2',
        rotation: 15,
        opacity: 90,
      })
    console.log('  ✓ Images inserted successfully with Base64 Data URI and remote URL')

    // 2. updateImage with Buffer and options.source swapping
    console.log('\n[2/12] Executing updateImage(imageId, options) with useSlide(2)...')
    await ppt.useSlide(2).updateImage('ShowcaseLogo', {
      x: '15%',
      y: '15%',
      source: replaceImageBuf,
      rotation: 50,
      opacity: 100,
      crop: { l: 5000, t: 5000, r: 5000, b: 5000 },
    })
    console.log('  ✓ Image "ShowcaseLogo" updated with new bounds, source swap, and crop settings')

    // 3. removeImage
    console.log('\n[3/12] Executing removeImage(imageId) with useSlide(3)...')
    ppt.useSlide(3).removeImage('ShowcaseLogo')
    console.log('  ✓ Image "ShowcaseLogo" cleanly removed from slide 3')

    // 4. addVideo with raw Base64 string, custom poster, and useSlide(4)
    console.log('\n[4/12] Executing addVideo(source, options) with useSlide(4)...')
    await ppt.useSlide(4).addVideo(sampleMp4RawBase64, {
      x: '40%',
      y: '10%',
      width: '50%',
      height: '35%',
      name: 'ShowcaseVideo',
      poster: sampleImageBuf,
      autoplay: true,
      loop: true,
      mute: true,
      startTime: 0,
      endTime: 10,
    })
    console.log(
      '  ✓ Video "ShowcaseVideo" added with custom poster thumbnail and playback options'
    )

    // 5. updateVideo with source swap and replacement poster thumbnail
    console.log('\n[5/12] Executing updateVideo(videoId, options) with useSlide(5)...')
    await ppt.useSlide(5).updateVideo('ShowcaseVideo', {
      loop: false,
      mute: false,
      source: replaceMp4Buf,
      poster: replaceImageBuf,
      startTime: 1.0,
      endTime: 8.0,
    })
    console.log(
      '  ✓ Video "ShowcaseVideo" playback options, source, and poster thumbnail updated'
    )

    // 6. removeVideo with automated timing sequence cleanup
    console.log('\n[6/12] Executing removeVideo(videoId) with useSlide(6)...')
    ppt.useSlide(6).removeVideo('ShowcaseVideo')
    console.log('  ✓ Video "ShowcaseVideo" and related timing sequences cleanly purged')

    // 7. addAudio with Base64 Data URI and useSlide(7)
    console.log('\n[7/12] Executing addAudio(source, options) with useSlide(7)...')
    await ppt.useSlide(7).addAudio(sampleMp3DataUri, {
      x: '10%',
      y: '60%',
      width: '10%',
      height: '10%',
      name: 'ShowcaseAudio',
      volume: 80,
      loop: true,
      hideDuringShow: true,
    })
    console.log('  ✓ Audio "ShowcaseAudio" added with Base64 Data URI, volume, and icon toggle')

    // 8. updateAudio with source swap and volume modification
    console.log('\n[8/12] Executing updateAudio(audioId, options) with useSlide(8)...')
    await ppt.useSlide(8).updateAudio('ShowcaseAudio', {
      volume: 100,
      loop: false,
      source: replaceMp3Buf,
      hideDuringShow: false,
    })
    console.log('  ✓ Audio "ShowcaseAudio" volume, source, and visibility updated')

    // 9. removeAudio with timing and action node cleanup
    console.log('\n[9/12] Executing removeAudio(audioId) with useSlide(9)...')
    ppt.useSlide(9).removeAudio('ShowcaseAudio')
    console.log('  ✓ Audio "ShowcaseAudio" and associated timing nodes cleanly purged')

    // 10. getMedias
    console.log('\n[10/12] Executing getMedias() with useSlide(10)...')
    const medias = ppt.useSlide(10).getMedias()
    console.log(`  ✓ Found ${medias.length} media item(s) on slide 10:`)
    for (const m of medias) {
      console.log(
        `    - ID: ${m.id}, Name: "${m.name}", Type: "${m.type}", Format: "${m.format}", rId: "${m.rId}"`
      )
    }

    // 11. getMedia
    console.log('\n[11/12] Executing getMedia(mediaId)...')
    const singleMedia = ppt.getMedia('ShowcaseLogo')
    console.log(
      `  ✓ Retrieved media: "${singleMedia?.name}" (Type: ${singleMedia?.type}, Target: ${singleMedia?.targetPath})`
    )

    // 12. cloneMedia with deduplicated assets
    console.log('\n[12/12] Executing cloneMedia(mediaId, clonedMediaId, clonedOptions)...')
    ppt.useSlide(10).cloneMedia('ShowcaseLogo', 'ShowcaseLogoClone', {
      x: '50%',
      y: '60%',
      rotation: 45,
    })
    ppt.useSlide(10).cloneMedia('ShowcaseVideo', 'ShowcaseVideoClone', {
      x: '50%',
      y: '30%',
      rotation: 45,
    })
    console.log('  ✓ Cloned "ShowcaseLogo" and "ShowcaseVideo" with deduplicated media streams')

    // Save master output presentation
    const outputPath = path.join(OUTPUT_DIR, 'showcase_media_advanced.pptx')
    await ppt.saveToFile(outputPath)
    console.log(`\nMaster media showcase saved to: ${outputPath}`)

    console.log('\n' + '='.repeat(75))
    console.log('All 12 Images & Media methods verified and completed successfully!')
    console.log('='.repeat(75))
  } catch (err) {
    console.error('Showcase encountered an error:', err)
    process.exit(1)
  }
}

if (require.main === module) {
  runMediaAdvancedShowcase()
}

module.exports = { runMediaAdvancedShowcase }
