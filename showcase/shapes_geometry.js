/**
 * @fileoverview showcase/shapes_geometry.js — Shapes & Geometry API Master Guide.
 *
 * Demonstrates and validates the strict 10-method Shapes & Geometry API boundary
 * for node-pptx-templater in a pure JavaScript runtime (zero native or Office dependencies).
 *
 * The Strict 10-Method Whitelist:
 * =========================================================================
 * 1.  ppt.addShape(slideId, type, options)
 * 2.  ppt.updateShape(targetId, options)
 * 3.  ppt.removeShape(shapeId)
 * 4.  ppt.clearShapes()
 * 5.  ppt.getShapes()
 * 6.  ppt.getShape(shapeId)
 * 7.  ppt.cloneShape(shapeId, clonedShapeId, clonedOptions)
 * 8.  ppt.getGroupShapes(groupId)
 * 9.  ppt.groupShapes(shapeIds)
 * 10. ppt.ungroupShapes(groupId)
 * =========================================================================
 */

'use strict'

const path = require('path')
const fs = require('fs-extra')
const { PPTXTemplater } = require('../src/index.js')

const SHOWCASE_DIR = __dirname
const OUTPUT_DIR = path.join(SHOWCASE_DIR, 'output/slide_advanced')
const OUTPUT_FILE = path.join(OUTPUT_DIR, 'showcase_shape_advanced.pptx')

async function runShapesShowcase() {
  console.log('='.repeat(78))
  console.log('PPTX-TEMPLATER: SHAPES & GEOMETRY 10-METHOD WHITELIST SHOWCASE')
  console.log('='.repeat(78))

  await fs.ensureDir(OUTPUT_DIR)

  // 1. Initialize a clean presentation deck
  console.log('\n[1/10] Initializing presentation...')
  // Initialize a PPTXTemplater instance with a sample presentation
  console.log('→ Step 0: Initializing PPTXTemplater with ShapePowerPointSample.pptx...')
  const templatePath = path.join(__dirname, 'templates/ShapePowerPointSample.pptx')
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1)

  // 2. addShape — Create various geometric shapes with styling and typography
  console.log('\n[2/10] ppt.addShape() — Adding DrawingML preset geometries...')
  ppt
    // Rectangle with solid fill and border
    .addShape(1, 'rectangle', {
      id: 'HeroHeaderCard',
      x: 60,
      y: '50%',
      width: 400,
      height: 100,
      fill: '#1E293B',
      border: { color: '#38BDF8', width: 2 },
      text: 'Enterprise Architecture Overview',
      textStyle: {
        fontSize: 16,
        bold: true,
        color: '#F8FAFC',
        align: 'center',
        verticalAlign: 'middle',
      },
    })
    // Rounded rectangle with percentage coordinates & custom corner radius
    .addShape(1, 'roundedRectangle', {
      id: 'MetricsPill',
      x: '55%',
      y: '50%',
      width: 320,
      height: 100,
      borderRadius: 16,
      fill: '#0F766E',
      text: 'Status: 100% Operational\n99.99% Uptime Guarantee',
      textStyle: {
        fontSize: 13,
        color: '#CCFBF1',
        align: 'center',
        lineSpacing: 1.2,
      },
    })
    // Triangle with rotation
    .addShape(1, 'triangle', {
      id: 'AccentIndicator',
      x: 60,
      y: '60%',
      width: 70,
      height: 70,
      fill: '#F59E0B',
      rotation: 90,
    })
    // Ellipse with border
    .addShape(1, 'ellipse', {
      id: 'ProcessNode1',
      x: 160,
      y: '60%',
      width: 140,
      height: 70,
      fill: '#6366F1',
      border: { color: '#4338CA', width: 1.5 },
      text: 'Ingest Queue',
      textStyle: { fontSize: 12, bold: true, color: '#FFFFFF', align: 'center' },
    })
    // Star geometry
    .addShape(1, 'star5', {
      id: 'ProcessNode2',
      x: 330,
      y: '60%',
      width: 80,
      height: 70,
      fill: '#EC4899',
    })
    // Arrow geometry
    .addShape(1, 'lineArrow', {
      id: 'chevronRightNode2',
      x: '50%',
      y: '80%',
      width: 80,
      height: 70,
      fill: '#ff0000ff',
      border: { color: '#ff0000ff', width: 2 },
    })
    // line geometry
    .addShape(1, 'line', {
      id: 'chevronRightNode2',
      x: '80%',
      y: '80%',
      width: 80,
      height: 70,
      fill: '#ff0000ff',
      border: { color: '#ff0000ff', width: 2 },
    })

  console.log(
    '  -> Added HeroHeaderCard, MetricsPill, AccentIndicator, ProcessNode1, and ProcessNode2.'
  )

  // 3. getShapes — Retrieve all shape metadata across the slide
  console.log('\n[3/10] ppt.getShapes() — Listing all shape metadata on slide 1...')

  const shapes = ppt.getShapes()

  console.log(`  -> Discovered ${shapes.length} shapes on slide 1:`)
  shapes.forEach(s => {
    console.log(
      `     - [${s.type}] "${s.name}" at (${s.x}px, ${s.y}px), size: ${s.width}x${s.height}px`
    )
  })

  // 4. getShape — Inspect an individual shape descriptor
  console.log('\n[4/10] ppt.getShape() — Querying individual shape metadata...')
  const heroCard = ppt.getShape('HeroHeaderCard')
  console.log(`  -> Queried "HeroHeaderCard":`)
  console.log(`     Geometry: ${heroCard.geometry}`)
  console.log(`     Fill: ${heroCard.fill}`)
  console.log(`     Border: ${JSON.stringify(heroCard.border)}`)
  console.log(`     Text: "${heroCard.text}"`)
  console.log(
    `     Bounds (EMU): x=${heroCard.bounds.x}, y=${heroCard.bounds.y}, w=${heroCard.bounds.width}, h=${heroCard.bounds.height}`
  )

  // 5. updateShape — In-place mutation of coordinates, dimensions, text, and styling
  console.log('\n[5/10] ppt.updateShape() — Mutating shape attributes in-place...')
  ppt.useSlide(2)
  ppt.updateShape('updateShape Circle', {
    fill: '#b813ceff',
    text: 'Status: Healthy (Live Cluster)\nLatency: 1.2ms | P99: 4.8ms',
    border: { color: '#2bff00ff', width: 2 },
  })
  ppt.updateShape('updateShape Triangle', {
    fill: '#c9ac08ff',
    text: 'Status: Healthy (Live Cluster)\nLatency: 1.2ms | P99: 4.8ms',
    border: { color: '#0004ffff', width: 2 },
  })
  ppt.updateShape('updateShape Arrow', {
    x: '50%',

    border: { color: '#ff0000ff', width: 2 },
  })
  const updatedCircle = ppt.getShape('updateShape Circle')
  console.log(
    `  -> Updated "updateShape Circle": fill is now ${updatedCircle.fill}, text:\n     "${updatedCircle.text.replace(/\n/g, ' ')}"`
  )

  // 6. removeShape & clearShapes — Deletion and evacuation
  console.log('\n[10/10] ppt.removeShape() & ppt.clearShapes() — Deletion and cleanup...')

  ppt.useSlide(3)
  ppt.removeShape('removeShape Circle')
  ppt.removeShape('removeShape Triangle')
  ppt.removeShape('removeShape Arrow')
  console.log('  -> Removed slide 3 shapes.')

  ppt.useSlide(4)
  console.log(`  -> Slide 4 created with ${ppt.getShapes().length} shapes.`)
  ppt.clearShapes()
  console.log(
    `  -> Slide 4 after clearShapes(): ${ppt.getShapes().length} shapes remaining (placeholders preserved).`
  )

  // 6. cloneShape — Duplicate an existing shape with position offsets
  console.log('\n[6/10] ppt.cloneShape() — Cloning shape with offsets...')
  ppt.useSlide(5)
  ppt.cloneShape('cloneShape Circle', 'cloneShape Circle Copy', {
    x: '75%',
    y: 210,
    text: 'Process Worker',
    fill: '#8B5CF6',
  })
  const cloned = ppt.getShape('cloneShape Circle Copy')
  console.log(
    `  -> Cloned "ProcessNode1" into "${cloned.name}" at (${cloned.x}px, ${cloned.y}px) with fill ${cloned.fill}`
  )

  // 7. groupShapes — Group multiple shapes into a composite container
  console.log(
    '\n[7/10] ppt.groupShapes() — Grouping shapes into a composite <p:grpSp> container...'
  )
  ppt.useSlide(6)
  ppt.groupShapes(['groupShapes Circle', 'groupShapes Triangle', 'groupShapes Arrow'], {
    name: 'PipelineGroup',
  })
  console.log(
    `  -> Grouped shapes ['groupShapes Circle', 'groupShapes Triangle', 'groupShapes Arrow'] into "PipelineGroup".`
  )

  // 8. getGroupShapes — Inspect child shapes inside the composite group
  console.log('\n[8/10] ppt.getGroupShapes() — Inspecting grouped child elements...')
  const groupChildren = ppt.getGroupShapes('PipelineGroup')
  console.log(`  -> "PipelineGroup" contains ${groupChildren.length} child elements:`)
  groupChildren.forEach(child => {
    console.log(`     * Child "${child.name}" [${child.type}] at (${child.x}px, ${child.y}px)`)
  })

  // 9. ungroupShapes — Dissolve composite group and reproject coordinates
  console.log(
    '\n[9/10] ppt.ungroupShapes() — Dissolving group container with coordinate re-projection...'
  )
  ppt.useSlide(7)
  ppt.ungroupShapes('ungroupShapes')
  const postUngroupShapes = ppt.getShapes()
  console.log(
    `  -> Ungroup complete. Top-level slide shape count is now ${postUngroupShapes.length}.`
  )

  try {
    await ppt.saveToFile(OUTPUT_FILE)
    console.log(`  -> Successfully exported master shapes presentation to:`)
    console.log(`     ${OUTPUT_FILE}`)
  } catch (err) {
    if (err.cause?.code === 'EBUSY' || err.message?.includes('EBUSY')) {
      const fallbackFile = path.join(OUTPUT_DIR, `showcase_shape_advanced_${Date.now()}.pptx`)
      await ppt.saveToFile(fallbackFile)
      console.log(
        `  -> Note: Target file was locked by an external application; saved to fallback:`
      )
      console.log(`     ${fallbackFile}`)
    } else {
      throw err
    }
  }
  console.log('\nAll 10 whitelisted Shapes & Geometry methods verified successfully!')
  console.log('='.repeat(78))
}

if (require.main === module) {
  runShapesShowcase().catch(err => {
    console.error('Showcase execution failed:', err)
    process.exit(1)
  })
}

module.exports = { runShapesShowcase }
