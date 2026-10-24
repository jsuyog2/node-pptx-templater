const path = require('path')
const { PPTXTemplater } = require('../src/index.js')

async function inspectSample() {
  const filePath = path.join(__dirname, 'templates/ChartPowerPointSample.pptx')
  console.log('Loading file:', filePath)

  const ppt = await PPTXTemplater.load(filePath)
  console.log('Presentation loaded successfully!')

  const slides = ppt.getSlides()
  console.log(`Total slides in deck: ${slides.length}`)

  for (let i = 0; i < slides.length; i++) {
    const s = slides[i]
    const slideIdx = i + 1
    console.log(`\n--- Slide ${slideIdx} ---`)
    console.log(`  Slide ID: ${s.slideId}, RelID: ${s.relationshipId}`)

    const layout = ppt.getSlideLayout(slideIdx)
    console.log(`  Layout: ${layout}`)

    const shapes = ppt.useSlide(slideIdx).getShapes()
    console.log(`  Total Shapes: ${shapes.length}`)
    for (const sh of shapes) {
      console.log(`    Shape ID: ${sh.shapeId || sh.id}, Name: "${sh.name}", Type: ${sh.type}`)
    }

    const charts = ppt.useSlide(slideIdx).getCharts()
    console.log(`  Total Charts: ${charts.length}`)
    for (const c of charts) {
      console.log(
        `    Chart: "${c.name}" | Type: ${c.type} | rId: ${c.rId} | zipPath: ${c.zipPath}`
      )
      console.log(
        `      Title: "${c.title}" | Series Count: ${c.series?.length} | Categories: ${c.categories?.join(', ')}`
      )
    }
  }

  // Inspect VFS for chart files and rels
  console.log('\n--- VFS Inspection ---')
  const zip = ppt.zipManager
  const chartFiles = zip.listFiles('ppt/charts/').filter(f => !f.includes('_rels'))
  console.log('Chart parts in VFS:', chartFiles)

  const embeddingFiles = zip.listFiles('ppt/embeddings/')
  console.log('Embedded workbooks in VFS:', embeddingFiles)

  const relFiles = zip.listFiles('ppt/slides/_rels/')
  console.log('Slide rels in VFS:', relFiles)
  for (const rf of relFiles) {
    const content = zip.readCachedFile(rf)
    console.log(`\n[${rf}]:\n${content}`)
  }
}

inspectSample().catch(err => {
  console.error('Inspection failed:', err)
  process.exit(1)
})
