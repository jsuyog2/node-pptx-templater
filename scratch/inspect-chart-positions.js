const JSZip = require('jszip');
const fs = require('fs');

async function main() {
  const zip = await JSZip.loadAsync(fs.readFileSync('templates/sample.pptx'));
  
  // Find all slide files
  const slideFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/slides/slide') && k.endsWith('.xml'));
  for (const f of slideFiles) {
    const xml = await zip.files[f].async('text');
    console.log(`Slide file: ${f}`);
    
    // Find all graphicFrames
    const gfPattern = /<p:graphicFrame>([\s\S]*?)<\/p:graphicFrame>/g;
    let match;
    while ((match = gfPattern.exec(xml)) !== null) {
      const gfContent = match[0];
      const nameMatch = /name="([^"]+)"/.exec(gfContent);
      const idMatch = /id="([^"]+)"/.exec(gfContent);
      console.log(`  GraphicFrame Name: ${nameMatch ? nameMatch[1] : 'unknown'}, ID: ${idMatch ? idMatch[1] : 'unknown'}`);
      
      const xfrmMatch = /<p:xfrm>([\s\S]*?)<\/p:xfrm>/.exec(gfContent);
      if (xfrmMatch) {
        console.log(`    xfrm: ${xfrmMatch[1].trim()}`);
      }
      
      const relIdMatch = /r:id="([^"]+)"/.exec(gfContent);
      if (relIdMatch) {
        console.log(`    Rel ID: ${relIdMatch[1]}`);
      }
    }
  }

  // Also print info about chart XMLs
  const chartFiles = Object.keys(zip.files).filter(k => k.startsWith('ppt/charts/chart') && k.endsWith('.xml'));
  for (const f of chartFiles) {
    const xml = await zip.files[f].async('text');
    console.log(`Chart file: ${f}`);
    
    // Check chart type
    let chartType = 'unknown';
    if (xml.includes('c:barChart')) chartType = 'barChart';
    if (xml.includes('c:lineChart')) chartType = 'lineChart';
    if (xml.includes('c:pieChart')) chartType = 'pieChart';
    
    console.log(`  Type: ${chartType}`);
    
    // Check grouping (stacked, standard, etc.)
    const groupingMatch = /<c:grouping val="([^"]+)"\/>/.exec(xml);
    const dirMatch = /<c:barDir val="([^"]+)"\/>/.exec(xml);
    console.log(`  Grouping: ${groupingMatch ? groupingMatch[1] : 'none'}, Dir: ${dirMatch ? dirMatch[1] : 'none'}`);
    
    // Check plotArea layout
    const plotAreaMatch = /<c:plotArea>([\s\S]*?)<\/c:plotArea>/.exec(xml);
    if (plotAreaMatch) {
      const layoutMatch = /<c:layout>([\s\S]*?)<\/c:layout>/.exec(plotAreaMatch[1]);
      if (layoutMatch) {
        console.log(`  PlotArea Layout: ${layoutMatch[1].trim()}`);
      }
    }
  }
}

main().catch(console.error);
