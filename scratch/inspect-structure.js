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
  
  console.log('--- [Content_Types].xml ---');
  const contentTypes = await zip.file('[Content_Types].xml').async('text');
  console.log(contentTypes);
  
  console.log('\n--- ppt/presentation.xml ---');
  const presentation = await zip.file('ppt/presentation.xml').async('text');
  console.log(presentation);
  
  console.log('\n--- ppt/_rels/presentation.xml.rels ---');
  const presentationRels = await zip.file('ppt/_rels/presentation.xml.rels').async('text');
  console.log(presentationRels);

  console.log('\n--- Slide Files in ZIP ---');
  const files = Object.keys(zip.files).sort();
  for (const f of files) {
    if (f.startsWith('ppt/slides/')) {
      console.log(f);
      if (f.endsWith('.rels')) {
        const relsContent = await zip.file(f).async('text');
        console.log(`  Relationships:\n${relsContent}`);
      }
    }
  }
}

main().catch(err => console.error(err));
