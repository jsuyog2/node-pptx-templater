const { resolve } = require('path');
const JSZip = require('jszip');
const fs = require('fs');

async function main() {
  const outputFile = resolve(__dirname, '../examples/output/basic-output.pptx');
  if (!fs.existsSync(outputFile)) {
    console.error('basic-output.pptx does not exist.');
    return;
  }
  
  const data = fs.readFileSync(outputFile);
  const zip = await JSZip.loadAsync(data);
  
  console.log('--- All Files in ZIP ---');
  const files = Object.keys(zip.files).sort();
  for (const f of files) {
    if (f.includes('notes') || f.includes('slide') || f.includes('presentation')) {
      console.log(f);
    }
  }
  
  // Let's inspect the relationships of note slides if they exist
  for (const f of files) {
    if (f.startsWith('ppt/notesSlides/_rels/')) {
      const content = await zip.file(f).async('text');
      console.log(`\n--- Rels file: ${f} ---`);
      console.log(content);
    }
  }
}

main().catch(err => console.error(err));
