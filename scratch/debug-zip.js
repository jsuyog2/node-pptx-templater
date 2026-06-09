const { PPTXTemplater } = require('../src/index.js');
const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function main() {
  const templatePath = path.resolve(__dirname, '../templates/sample.pptx');
  const imagePath = path.resolve(__dirname, '../tests/fixtures/media__1780653961880.png'); // using a known png image
  const outputPath = path.resolve(__dirname, 'output-debug.pptx');

  if (!fs.existsSync(templatePath)) {
    console.error(`Template not found at: ${templatePath}`);
    return;
  }
  
  // Create a dummy image file if it doesn't exist
  if (!fs.existsSync(imagePath)) {
    fs.mkdirSync(path.dirname(imagePath), { recursive: true });
    // Write 100 bytes of dummy png data (with valid header)
    const dummyPng = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, ...new Array(92).fill(0)]);
    fs.writeFileSync(imagePath, dummyPng);
  }

  console.log('--- Loading Template ---');
  const ppt = await PPTXTemplater.load(templatePath);
  ppt.enableDebugZip();

  console.log('--- Replacing Image ---');
  // Just try to replace some image. Let's list images first.
  ppt.useSlide(1);
  const images = ppt.getImages();
  console.log('Images on Slide 1:', images);

  if (images.length > 0) {
    const targetImageName = images[0].name || images[0].id;
    console.log(`Replacing image "${targetImageName}" with "${imagePath}"`);
    await ppt.replaceImage(targetImageName, imagePath);
  } else {
    console.log('No images found on slide 1 to replace, trying slide 2');
    ppt.useSlide(2);
    const images2 = ppt.getImages();
    console.log('Images on Slide 2:', images2);
    if (images2.length > 0) {
      const targetImageName = images2[0].name || images2[0].id;
      console.log(`Replacing image "${targetImageName}" with "${imagePath}"`);
      await ppt.replaceImage(targetImageName, imagePath);
    }
  }

  console.log('--- Saving to File ---');
  await ppt.saveToFile(outputPath);
  console.log(`Saved output to ${outputPath}`);

  console.log('--- Auditing Saved ZIP ---');
  const zipData = fs.readFileSync(outputPath);
  const zip = await JSZip.loadAsync(zipData);

  console.log(`Total files in output ZIP: ${Object.keys(zip.files).length}`);
  
  // Let's print out details for each entry
  for (const [name, file] of Object.entries(zip.files)) {
    if (file.dir) continue;
    // JSZip v3 ZipObject metadata can be read from file properties
    const uncompressedSize = file._data ? file._data.uncompressedSize : 'unknown';
    const compressedSize = file._data ? file._data.compressedSize : 'unknown';
    const crc32 = file._data ? file._data.crc32 : 'unknown';
    const compression = file.options.compression || 'STORE';

    console.log(`Entry: ${name}`);
    console.log(`  Uncompressed Size: ${uncompressedSize}`);
    console.log(`  Compressed Size:   ${compressedSize}`);
    console.log(`  CRC32:             ${crc32}`);
    console.log(`  Compression:       ${compression}`);
    
    // Test if we can read/extract it (this will throw if uncompressed size mismatch happens)
    try {
      const buf = await file.async('nodebuffer');
      console.log(`  Successfully read: ${buf.length} bytes`);
    } catch (err) {
      console.error(`  ERROR READING ENTRY: ${err.message}`);
    }
  }
}

main().catch(err => {
  console.error('Fatal Error:', err);
});
