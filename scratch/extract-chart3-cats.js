const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function main() {
  const samplePath = path.resolve('templates/sample.pptx');
  const zip = await JSZip.loadAsync(fs.readFileSync(samplePath));
  const xml = await zip.files['ppt/charts/chart3.xml'].async('text');

  // Let's count categories in chart3.xml
  const catMatch = /<c:cat>([\s\S]*?)<\/c:cat>/.exec(xml);
  if (catMatch) {
    const ptCountMatch = /<c:ptCount val="(\d+)"\/>/.exec(catMatch[1]);
    console.log(`Categories count: ${ptCountMatch ? ptCountMatch[1] : 'Unknown'}`);
    console.log(catMatch[1]);
  }
}

main().catch(console.error);
