const { PPTXTemplater } = require('../src/index.js');
const fs = require('fs');
const path = require('path');
const JSZip = require('jszip');

async function testCompression(compressionMode) {
  const templatePath = path.resolve(__dirname, '../templates/sample.pptx');
  const imagePath = path.resolve(__dirname, '../tests/fixtures/media__1780653961880.png');
  const outputPath = path.resolve(__dirname, `output-${compressionMode.toLowerCase()}.pptx`);

  const ppt = await PPTXTemplater.load(templatePath);
  ppt.useSlide(1);
  const images = ppt.getImages();
  if (images.length > 0) {
    const targetImageName = images[0].name || images[0].id;
    await ppt.replaceImage(targetImageName, imagePath);
  }

  // Flush slide changes and content types
  await ppt.zipManager.waitForPendingWrites();
  
  // Custom generate with specified compression mode
  console.log(`Generating with compression: ${compressionMode}...`);
  const buffer = await ppt.zipManager.rawZip.generateAsync({
    type: 'nodebuffer',
    compression: compressionMode,
    compressionOptions: { level: compressionMode === 'DEFLATE' ? 6 : undefined }
  });

  fs.writeFileSync(outputPath, buffer);
  console.log(`Saved ${compressionMode} output to ${outputPath} (${buffer.length} bytes)`);

  // Validate archive
  console.log(`Validating ${compressionMode} archive...`);
  try {
    const tempPpt = await PPTXTemplater.load(outputPath);
    await tempPpt.zipManager.validateArchive();
    console.log(`✓ ${compressionMode} archive is valid and verified.`);
  } catch (err) {
    console.error(`✗ ${compressionMode} validation failed:`, err.message);
  }
}

async function main() {
  // Ensure the validateArchive method exists in PPTXTemplater
  // (We'll implement it next, but let's test the compression behavior first)
  await testCompression('STORE');
  await testCompression('DEFLATE');
}

main().catch(err => {
  console.error('Fatal Error:', err);
});
