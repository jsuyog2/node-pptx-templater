/**
 * @fileoverview showcase/hyperlinks_advanced.js — Master Demonstration of the 6-Method Hyperlink API Boundary.
 *
 * Demonstrates and validates the strict 6-method Hyperlink API boundary
 * for node-pptx-templater in a pure JavaScript runtime (zero native or Office dependencies).
 *
 * The Strict 6-Method Whitelist:
 * =========================================================================
 * 1. ppt.addHyperlink(targetId, url, options)
 * 2. ppt.updateHyperlink(targetId, url, options)
 * 3. ppt.removeHyperlink(targetId)
 * 4. ppt.getHyperlinks(slideId)
 * 5. ppt.getHyperlink(targetId)
 * 6. ppt.setHyperlink(targetId, text, url, options)
 * =========================================================================
 */

'use strict'

const path = require('path')
const fs = require('fs-extra')
const { PPTXTemplater } = require('../src/index.js')

const SHOWCASE_DIR = __dirname
const OUTPUT_DIR = path.join(SHOWCASE_DIR, 'output')
const OUTPUT_FILE = path.join(OUTPUT_DIR, 'hyperlinks_advanced.pptx')

async function runHyperlinksShowcase() {
  console.log('='.repeat(78))
  console.log('PPTX-TEMPLATER: STRICT 6-METHOD HYPERLINK API WHITELIST SHOWCASE')
  console.log('='.repeat(78))

  await fs.ensureDir(OUTPUT_DIR)

  // 1. Initialize a multi-slide presentation
  console.log('\n[1/6] Initializing multi-slide presentation...')

  await fs.ensureDir(OUTPUT_DIR)
  const templatePath = path.join(__dirname, 'templates/HyperlinkPowerPointSample.pptx')
  const ppt = await PPTXTemplater.load(templatePath)

  ppt.useSlide(1)

  // 2. addHyperlink — Bind external URLs, internal jumps, and navigation actions
  console.log('\n[2/6] ppt.addHyperlink() — Binding hyperlinks to shapes and containers...')

  ppt.addHyperlink('TextHyperlink', 'https://node-pptx-templater.pages.dev/', {
    tooltip: 'Navigate to Repository Documentation',
  })
  console.log('  → Bound external URL to GitHubBtn')

  // Button 2: Internal Slide Jump (Slide 1 -> Slide 4)
  ppt.addHyperlink('ShapeHyperlink1', 4, {
    tooltip: 'Transition directly to Slide 4',
  })
  console.log('  → Bound internal slide jump (Slide 1 → Slide 4) to ShapeHyperlink1')

  // Button 3: Decoupled slideshow action preset ('nextSlide')
  ppt.addHyperlink('ShapeHyperlink2', 'last', {
    tooltip: 'Advance to next slide in show',
  })
  console.log('  → Bound decoupled action ("nextSlide") to ShapeHyperlink2')

  // Button 4: Slide number jump (Slide 1 -> Slide 3)
  ppt.addHyperlink('ShapeHyperlink3', 3, {
    tooltip: 'Transition directly to Slide 3',
  })
  console.log('  → Bound internal slide jump (Slide 1 → Slide 3) to ShapeHyperlink3')

  // Image link: Decoupled slideshow action preset ('previousSlide')
  ppt.addHyperlink('ImageHyperlink', 'end', {
    tooltip: 'Return to previous slide in show',
  })
  console.log('  → Bound decoupled action ("previousSlide") to ImageHyperlink')

  // Button 4: Temporary link for update & removal demo
  ppt.useSlide(2)
  // 3. updateHyperlink — Modify URL destination, action presets, and ScreenTip
  console.log('\n[3/6] ppt.updateHyperlink() — Updating target URL, action presets, and tooltip...')
  ppt.updateHyperlink('TextHyperlink', 'https://node-pptx-templater.pages.dev', {
    tooltip: 'Updated Production Live Environment',
  })
  ppt.updateHyperlink('ImageHyperlink', 'https://node-pptx-templater.pages.dev', {
    tooltip: 'Updated Production Live Environment',
  })
  ppt.updateHyperlink('ShapeHyperlink1', 'https://node-pptx-templater.pages.dev', {
    tooltip: 'Updated Production Live Environment',
  })
  ppt.updateHyperlink('ShapeHyperlink2', 'previous', {
    tooltip: 'Updated to previous slide action',
  })
  ppt.updateHyperlink('ShapeHyperlink3', 4, {
    tooltip: 'Updated Production Live Environment',
  })
  console.log('  → Updated TempBtn destination to Production URL')

  // 4. removeHyperlink — Detach link and purge orphaned .rels entry
  ppt.useSlide(3)

  console.log('\n[4/6] ppt.removeHyperlink() — Removing hyperlink and purging relationships...')
  ppt.removeHyperlink('TextHyperlink')
  ppt.removeHyperlink('ImageHyperlink')
  ppt.removeHyperlink('ShapeHyperlink1')
  ppt.removeHyperlink('ShapeHyperlink3')
  console.log('  → Successfully detached hyperlink from TempBtn and purged .rels entry')

  // 5. setHyperlink — Substring text run splitting with custom typography styling
  console.log('\n[5/6] ppt.setHyperlink() — Text substring run-splitting and link styling...')

  ppt.useSlide(4)

  // Substring 1: "Developer Documentation"
  ppt.setHyperlink('TextHyperlink', 'Subsystem Managers', 'https://node-pptx-templater.pages.dev', {
    color: '#38BDF8',
    underline: true,
    bold: true,
    tooltip: 'Open Developer Documentation Portal',
  })

  // Substring 2: "ValidationEngine" with direct action preset and typography styling
  ppt.setHyperlink('TextHyperlink', 'ValidationEngine', 'first', {
    action: 'first',
    style: {
      color: '#F43F5E',
      underline: true,
      italic: true,
    },
    tooltip: 'Slideshow jump to First Slide',
  })

  // Substring 3: "VirtualFileSystem" with nested options.style.color
  ppt.setHyperlink('TextHyperlink', 'VirtualFileSystem', 'https://node-pptx-templater.pages.dev', {
    style: {
      color: '#10B981',
      bold: true,
      underline: true,
    },
    tooltip: 'In-Memory Virtual File System Architecture',
  })
  console.log('  → Split text runs and applied custom colored hyperlinks to target substrings')

  ppt.useSlide(2)
  // 6. getHyperlinks & getHyperlink — Read-only introspection
  console.log('\n[6/6] ppt.getHyperlinks() & getHyperlink() — Hyperlink Introspection...')
  const slide1Links = ppt.getHyperlinks()
  console.log(`  → Total hyperlinks on Slide 1: ${slide1Links.length}`)
  for (const l of slide1Links) {
    console.log(
      `    - [${l.type}] Target: "${l.targetId}" | URL/Action: ${l.url || l.action || 'Slide ' + l.targetSlide} | Tooltip: "${l.tooltip || ''}"`
    )
  }

  const singleQuery = ppt.getHyperlink('ImageHyperlink')
  console.log(`  → getHyperlink('ImageHyperlink') result: ${singleQuery?.url}`)

  // 7. Save and export clean presentation
  console.log('\n[Export] Saving presentation to disk...')
  await ppt.saveToFile(OUTPUT_FILE)
  console.log(`✓ Showcase successfully written to: ${OUTPUT_FILE}`)
  console.log('='.repeat(78))
}

if (require.main === module) {
  runHyperlinksShowcase().catch(err => {
    console.error('Showcase execution failed:', err)
    process.exit(1)
  })
}

module.exports = { runHyperlinksShowcase }
