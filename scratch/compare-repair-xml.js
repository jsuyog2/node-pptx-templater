const AdmZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function extractSlide3(pptxPath) {
  const data = fs.readFileSync(pptxPath);
  const zip = await AdmZip.loadAsync(data);
  return await zip.file('ppt/slides/slide3.xml').async('text');
}

async function main() {
  const origXml = await extractSlide3(path.resolve(__dirname, 'test-repair-output.pptx'));
  const savedXml = await extractSlide3(path.resolve(__dirname, 'test-repair-saved.pptx'));
  
  if (origXml === savedXml) {
    console.log('🎉 No changes! The generated XML matches PowerPoint saved XML exactly.');
  } else {
    console.log('⚠️ XML differs! PowerPoint modified the presentation XML.');
    // Let's write them both formatted so we can diff them
    const format = xml => xml
      .replace(/(<a:tr[^>]*>)/g, '\n  $1')
      .replace(/(<\/a:tr>)/g, '$1\n')
      .replace(/(<a:tc[^>]*>)/g, '\n    $1')
      .replace(/(<\/a:tc>)/g, '$1');
      
    fs.writeFileSync(path.resolve(__dirname, 'slide3_original_comparison.xml'), format(origXml));
    fs.writeFileSync(path.resolve(__dirname, 'slide3_saved_comparison.xml'), format(savedXml));
    console.log('Saved both files to scratch/slide3_original_comparison.xml and scratch/slide3_saved_comparison.xml');
  }
}

main().catch(console.error);
