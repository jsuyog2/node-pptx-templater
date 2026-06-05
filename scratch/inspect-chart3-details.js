const JSZip = require('jszip');
const fs = require('fs');

async function main() {
  const zip = await JSZip.loadAsync(fs.readFileSync('templates/sample.pptx'));
  const xml = await zip.files['ppt/charts/chart3.xml'].async('text');
  
  // Find all series names
  const serPattern = /<c:ser>([\s\S]*?)<\/c:ser>/g;
  let match;
  let idx = 0;
  while ((match = serPattern.exec(xml)) !== null) {
    const serContent = match[1];
    const nameMatch = /<c:tx>[\s\S]*?<c:v>([^<]+)<\/c:v>/.exec(serContent);
    const catMatch = /<c:cat>[\s\S]*?<c:ptCount val="(\d+)"\/>([\s\S]*?)<\/c:cat>/.exec(serContent);
    const valMatch = /<c:val>[\s\S]*?<c:ptCount val="(\d+)"\/>([\s\S]*?)<\/c:val>/.exec(serContent);
    
    console.log(`Series ${idx}:`);
    console.log(`  Name: ${nameMatch ? nameMatch[1] : 'unknown'}`);
    console.log(`  Categories count: ${catMatch ? catMatch[1] : 0}`);
    console.log(`  Values count: ${valMatch ? valMatch[1] : 0}`);
    
    // Print first few categories and values
    if (catMatch) {
      const pts = [];
      const ptPattern = /<c:pt idx="(\d+)">[\s\S]*?<c:v>([^<]+)<\/c:v>/g;
      let ptMatch;
      while ((ptMatch = ptPattern.exec(catMatch[2])) !== null) {
        pts.push(`${ptMatch[1]}:${ptMatch[2]}`);
      }
      console.log(`  Categories: ${pts.slice(0, 5).join(', ')}`);
    }
    if (valMatch) {
      const pts = [];
      const ptPattern = /<c:pt idx="(\d+)">[\s\S]*?<c:v>([^<]+)<\/c:v>/g;
      let ptMatch;
      while ((ptMatch = ptPattern.exec(valMatch[2])) !== null) {
        pts.push(`${ptMatch[1]}:${ptMatch[2]}`);
      }
      console.log(`  Values: ${pts.slice(0, 5).join(', ')}`);
    }
    idx++;
  }
}

main().catch(console.error);
