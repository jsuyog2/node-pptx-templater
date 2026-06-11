const AdmZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function extractTable2Xml(pptxPath) {
  const data = fs.readFileSync(pptxPath);
  const zip = await AdmZip.loadAsync(data);
  const slideXml = await zip.file('ppt/slides/slide3.xml').async('text');
  
  // Find Table2 (graphicFrame containing name="Table2")
  // We can find <p:graphicFrame> ... name="Table2" ... </p:graphicFrame>
  const frameRegex = /<p:graphicFrame>[\s\S]*?<\/p:graphicFrame>/g;
  let match;
  while ((match = frameRegex.exec(slideXml)) !== null) {
    const frame = match[0];
    if (frame.includes('name="Table2"')) {
      const tblMatch = /<a:tbl>[\s\S]*?<\/a:tbl>/.exec(frame);
      if (tblMatch) {
        return tblMatch[0];
      }
    }
  }
  return null;
}

function formatXml(content) {
  if (!content) return [];
  return content
    .replace(/(<[^>]+>)/g, '$1\n')
    .replace(/<a:bodyPr><\/a:bodyPr>/g, '<a:bodyPr/>')
    .replace(/<a:tcPr><\/a:tcPr>/g, '<a:tcPr/>')
    .replace(/<a:lstStyle><\/a:lstStyle>/g, '<a:lstStyle/>')
    .replace(/<a:endParaRPr><\/a:endParaRPr>/g, '<a:endParaRPr/>')
    .replace(/<a:endParaRPr lang="[^"]+"><\/a:endParaRPr>/g, (m) => m.replace(/><\/a:endParaRPr>/, '/>'))
    .replace(/ \/>/g, '/>')
    .split('\n')
    .map(line => line.trim())
    .filter(line => line.length > 0);
}

async function main() {
  const origTbl = await extractTable2Xml(path.resolve(__dirname, 'test-table2-nesting-output.pptx'));
  const savedTbl = await extractTable2Xml(path.resolve(__dirname, 'test-table2-nesting-saved.pptx'));
  
  const orig = formatXml(origTbl);
  const saved = formatXml(savedTbl);
  
  let diff = [];
  const maxLines = Math.max(orig.length, saved.length);
  for (let i = 0; i < maxLines; i++) {
    const o = orig[i] || '';
    const s = saved[i] || '';
    
    if (o !== s) {
      diff.push(`Line ${i + 1}:`);
      diff.push(`- ${o}`);
      diff.push(`+ ${s}`);
      diff.push('');
    }
  }
  
  if (diff.length === 0) {
    console.log('🎉 Table2: No semantic changes! PowerPoint did NOT modify Table2 structure.');
  } else {
    console.log(`⚠️ Table2 XML differs! Saved ${diff.length / 4} differences.`);
    console.log(diff.join('\n'));
  }
}

main().catch(console.error);
