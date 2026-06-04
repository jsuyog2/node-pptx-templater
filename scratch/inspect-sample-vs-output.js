const JSZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function main() {
  const samplePath = path.resolve('templates/sample.pptx');
  const outputPath = path.resolve('examples/output/basic-output.pptx');

  if (!fs.existsSync(samplePath)) {
    console.log('Sample template does not exist!');
    return;
  }
  if (!fs.existsSync(outputPath)) {
    console.log('Output does not exist!');
    return;
  }

  const sampleZip = await JSZip.loadAsync(fs.readFileSync(samplePath));
  const outputZip = await JSZip.loadAsync(fs.readFileSync(outputPath));

  console.log('=== SAMPLE.PPTX CHARTS ===');
  await printCharts(sampleZip);

  console.log('\n=== BASIC-OUTPUT.PPTX CHARTS ===');
  await printCharts(outputZip);
}

async function printCharts(zip) {
  const files = Object.keys(zip.files).filter(k => k.startsWith('ppt/charts/chart') && k.endsWith('.xml'));
  for (const f of files) {
    const xml = await zip.files[f].async('text');
    console.log(`\nChart File: ${f}`);
    
    // Find chart title
    const titleMatch = /<c:title>([\s\S]*?)<\/c:title>/.exec(xml);
    if (titleMatch) {
      const textMatch = /<a:t>([^<]*)<\/a:t>/.exec(titleMatch[1]);
      console.log(`  Title: ${textMatch ? textMatch[1] : 'Unknown'}`);
    } else {
      console.log('  Title: None');
    }

    // Let's count how many series
    const serCount = (xml.match(/<c:ser>/g) || []).length;
    console.log(`  Number of Series: ${serCount}`);

    // Print all dLbls blocks
    const dLblsPattern = /<c:dLbls>([\s\S]*?)<\/c:dLbls>/g;
    let match;
    let index = 0;
    while ((match = dLblsPattern.exec(xml)) !== null) {
      console.log(`  dLbls block #${index++}:`);
      // Let's see the dLbl children or spPr
      const dLblsContent = match[1];
      const spPrMatch = /(<c:spPr>[\s\S]*?<\/c:spPr>)/.exec(dLblsContent);
      console.log(`    Series-level spPr: ${spPrMatch ? 'Present: ' + spPrMatch[1].substring(0, 100) + '...' : 'None'}`);
      
      const dLblPattern = /<c:dLbl>([\s\S]*?)<\/c:dLbl>/g;
      let dLblMatch;
      while ((dLblMatch = dLblPattern.exec(dLblsContent)) !== null) {
        const dLblContent = dLblMatch[1];
        const idxMatch = /<c:idx val="(\d+)"\/>/.exec(dLblContent);
        const idxVal = idxMatch ? idxMatch[1] : 'unknown';
        const dLblSpPrMatch = /(<c:spPr>[\s\S]*?<\/c:spPr>)/.exec(dLblContent);
        console.log(`    dLbl idx="${idxVal}": spPr: ${dLblSpPrMatch ? 'Present: ' + dLblSpPrMatch[1] : 'None'}`);
      }
    }
  }
}

main().catch(console.error);
