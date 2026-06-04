const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function main() {
  const samplePath = path.resolve('templates/sample.pptx');
  const zip = await JSZip.loadAsync(fs.readFileSync(samplePath));
  const xml = await zip.files['ppt/charts/chart3.xml'].async('text');

  const serPattern = /<c:ser>([\s\S]*?)<\/c:ser>/g;
  let serMatch;
  let serIndex = 0;
  while ((serMatch = serPattern.exec(xml)) !== null) {
    const serContent = serMatch[1];
    
    if (serIndex === 0) {
      console.log(`\n=================== Series #0 ===================`);
      const dLblsMatch = /<c:dLbls>([\s\S]*?)<\/c:dLbls>/.exec(serContent);
      if (dLblsMatch) {
        console.log(dLblsMatch[1]);
      }
    }
    serIndex++;
  }
}

main().catch(console.error);
