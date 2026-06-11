const AdmZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function inspectSlides(pptxPath, label) {
  const data = fs.readFileSync(pptxPath);
  const zip = await AdmZip.loadAsync(data);
  console.log(`=== Slides in ${label} ===`);
  const files = Object.keys(zip.files).filter(name => name.startsWith('ppt/slides/slide') && name.endsWith('.xml'));
  for (const file of files) {
    const xml = await zip.file(file).async('text');
    const hasTable = xml.includes('<a:tbl>');
    console.log(`- ${file}: hasTable=${hasTable}, size=${xml.length}`);
  }
}

async function main() {
  await inspectSlides(path.resolve(__dirname, '../examples/output/basic-output.pptx'), 'basic-output.pptx');
  await inspectSlides(path.resolve(__dirname, 'basic-saved.pptx'), 'basic-saved.pptx');
}

main().catch(console.error);
