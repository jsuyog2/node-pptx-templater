const { PPTXTemplater } = require('../src/index.js');
const path = require('path');
const fs = require('fs');

async function main() {
  const templatePath = path.resolve(__dirname, '../templates/sample.pptx');
  const outputPath = path.resolve(__dirname, 'test-table2-nesting-output.pptx');
  
  console.log('Loading template...');
  const ppt = await PPTXTemplater.load(templatePath);
  ppt.useSlide(3);
  
  console.log('Adding nested table row to Table2...');
  // Table2 has columns 0 and 1 horizontally merged in the template.
  // We pass 5 cells matching the grid.
  ppt.addTableRow('Table2', ['Bob', '', '     Designer', ['Product', 'Value'], '']);
  
  console.log('Saving to:', outputPath);
  await ppt.saveToFile(outputPath);
  
  const AdmZip = require('jszip');
  const data = fs.readFileSync(outputPath);
  const zip = await AdmZip.loadAsync(data);
  const slideXml = await zip.file('ppt/slides/slide3.xml').async('text');
  
  let formattedXml = slideXml
    .replace(/(<a:tr[^>]*>)/g, '\n  $1')
    .replace(/(<\/a:tr>)/g, '$1\n')
    .replace(/(<a:tc[^>]*>)/g, '\n    $1')
    .replace(/(<\/a:tc>)/g, '$1');
  
  fs.writeFileSync(path.resolve(__dirname, 'test_table2_nesting_slide3.xml'), formattedXml);
  console.log('XML saved to scratch/test_table2_nesting_slide3.xml');
  
  // Test opening it in PowerPoint
  const pptApp = NewObjectPowerPoint();
  if (pptApp) {
    console.log('Testing opening in PowerPoint via COM...');
    const resolvedPath = path.resolve(outputPath);
    try {
      const pres = pptApp.Presentations.Open(resolvedPath, 1, 0, 0);
      console.log('✅ Success: Table2 nested row opened cleanly without repair errors.');
      pres.Close();
    } catch (err) {
      console.error('❌ Error opening Table2 nested row:', err.message);
    }
    pptApp.Quit();
  }
}

function NewObjectPowerPoint() {
  try {
    const execSync = require('child_process').execSync;
    execSync('powershell -Command "New-Object -ComObject PowerPoint.Application"', { stdio: 'ignore' });
    // If it didn't throw, we can run a quick powershell script to test
    return {
      Presentations: {
        Open: (path) => {
          const execSync = require('child_process').execSync;
          execSync(`powershell -Command "$ppt = New-Object -ComObject PowerPoint.Application; $ppt.DisplayAlerts = 2; $pres = $ppt.Presentations.Open('${path}', 1, 0, 0); $pres.Close(); $ppt.Quit()"`);
        }
      },
      Quit: () => {}
    };
  } catch (e) {
    return null;
  }
}

main().catch(console.error);
