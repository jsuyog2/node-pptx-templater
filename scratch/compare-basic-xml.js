const AdmZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function extractSlideXml(pptxPath, slideName) {
  const data = fs.readFileSync(pptxPath);
  const zip = await AdmZip.loadAsync(data);
  return await zip.file(slideName).async('text');
}

async function main() {
  const origXml = await extractSlideXml(path.resolve(__dirname, '../examples/output/basic-output.pptx'), 'ppt/slides/slide3.xml');
  const savedXml = await extractSlideXml(path.resolve(__dirname, 'basic-saved.pptx'), 'ppt/slides/slide2.xml');
  
  if (origXml === savedXml) {
    console.log('🎉 No changes! The generated XML matches PowerPoint saved XML exactly.');
  } else {
    console.log('⚠️ XML differs! PowerPoint modified the presentation XML.');
    const format = xml => xml
      .replace(/(<a:tr[^>]*>)/g, '\n  $1')
      .replace(/(<\/a:tr>)/g, '$1\n')
      .replace(/(<a:tc[^>]*>)/g, '\n    $1')
      .replace(/(<\/a:tc>)/g, '$1');
      
    fs.writeFileSync(path.resolve(__dirname, 'basic_original_comparison.xml'), format(origXml));
    fs.writeFileSync(path.resolve(__dirname, 'basic_saved_comparison.xml'), format(savedXml));
    console.log('Saved both files to scratch/basic_original_comparison.xml and scratch/basic_saved_comparison.xml');
  }
}

main().catch(console.error);
