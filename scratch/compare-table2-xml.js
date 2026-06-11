const AdmZip = require('jszip');
const fs = require('fs');
const path = require('path');

async function extractSlide3(pptxPath) {
  const data = fs.readFileSync(pptxPath);
  const zip = await AdmZip.loadAsync(data);
  return await zip.file('ppt/slides/slide3.xml').async('text');
}

function formatXml(content) {
  // Add newlines after tags and normalize self-closing tags
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
  const origXml = await extractSlide3(path.resolve(__dirname, 'test-table2-nesting-output.pptx'));
  const savedXml = await extractSlide3(path.resolve(__dirname, 'test-table2-nesting-saved.pptx'));
  
  const orig = formatXml(origXml);
  const saved = formatXml(savedXml);
  
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
    console.log('🎉 No changes! PowerPoint did NOT modify the XML.');
  } else {
    console.log(`⚠️ XML differs! Saved ${diff.length / 4} semantic differences.`);
    fs.writeFileSync(path.resolve(__dirname, 'table2_diff.txt'), diff.join('\n'));
    // Log only the first 20 lines of diff
    console.log(diff.slice(0, 40).join('\n'));
  }
}

main().catch(console.error);
