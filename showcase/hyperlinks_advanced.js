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
  const ppt = await PPTXTemplater.createBlank()
  ppt.addSlide() // Slide 2: Target Destination
  ppt.addSlide() // Slide 3: Summary
  console.log(`✓ Created presentation with ${ppt.slideCount} slides`)

  ppt.useSlide(1)

  // 2. addHyperlink — Bind external URLs, internal jumps, and navigation actions
  console.log('\n[2/6] ppt.addHyperlink() — Binding hyperlinks to shapes and containers...')

  // Slide 1 title shape
  await ppt.addShape('rect', {
    id: 'HeaderCard',
    name: 'HeaderCard',
    x: 80,
    y: 40,
    width: 600,
    height: 70,
    fill: '#0F172A',
    text: 'Interactive Presentationml Hyperlink Hub',
    textStyle: { color: '#38BDF8', fontSize: 20, bold: true },
  })

  // Button 1: External URL with ScreenTip
  await ppt.addShape('roundRect', {
    id: 'GitHubBtn',
    name: 'GitHubBtn',
    x: 80,
    y: 130,
    width: 250,
    height: 50,
    fill: '#2563EB',
    text: 'Visit GitHub Repository',
    textStyle: { color: '#FFFFFF', fontSize: 13, bold: true },
  })
  ppt.addHyperlink('GitHubBtn', 'https://github.com', {
    tooltip: 'Navigate to GitHub Repository',
  })
  console.log('  → Bound external URL to GitHubBtn')

  // Button 2: Internal Slide Jump (Slide 1 -> Slide 2)
  await ppt.addShape('roundRect', {
    id: 'JumpSlide2Btn',
    name: 'JumpSlide2Btn',
    x: 360,
    y: 130,
    width: 250,
    height: 50,
    fill: '#059669',
    text: 'Jump to Slide 2 (Architecture)',
    textStyle: { color: '#FFFFFF', fontSize: 13, bold: true },
  })
  ppt.addHyperlink('JumpSlide2Btn', 2, {
    tooltip: 'Transition directly to Slide 2',
  })
  console.log('  → Bound internal slide jump (Slide 1 → Slide 2) to JumpSlide2Btn')

  // Button 3: Slideshow Navigation Action (Next Slide)
  await ppt.addShape('roundRect', {
    id: 'NextActionBtn',
    name: 'NextActionBtn',
    x: 80,
    y: 200,
    width: 250,
    height: 50,
    fill: '#7C3AED',
    text: 'Next Slide Action (ppaction)',
    textStyle: { color: '#FFFFFF', fontSize: 13, bold: true },
  })
  ppt.addHyperlink('NextActionBtn', 'next', {
    action: 'next',
    tooltip: 'Advance to next slide in show',
  })
  console.log('  → Bound slideshow action ("next") to NextActionBtn')

  // Button 4: Temporary link for update & removal demo
  await ppt.addShape('roundRect', {
    id: 'TempBtn',
    name: 'TempBtn',
    x: 360,
    y: 200,
    width: 250,
    height: 50,
    fill: '#DC2626',
    text: 'Temporary External Link',
    textStyle: { color: '#FFFFFF', fontSize: 13, bold: true },
  })
  ppt.addHyperlink('TempBtn', 'https://temp-staging.example.com', {
    tooltip: 'Staging Environment',
  })

  // 3. updateHyperlink — Modify URL destination and ScreenTip
  console.log('\n[3/6] ppt.updateHyperlink() — Updating target URL and tooltip...')
  ppt.updateHyperlink('TempBtn', 'https://production.example.com/live', {
    tooltip: 'Updated Production Live Environment',
  })
  console.log('  → Updated TempBtn destination to Production URL')

  // 4. removeHyperlink — Detach link and purge orphaned .rels entry
  console.log('\n[4/6] ppt.removeHyperlink() — Removing hyperlink and purging relationships...')
  ppt.removeHyperlink('TempBtn')
  console.log('  → Successfully detached hyperlink from TempBtn and purged .rels entry')

  // 5. setHyperlink — Substring text run splitting with custom typography styling
  console.log('\n[5/6] ppt.setHyperlink() — Text substring run-splitting and link styling...')
  await ppt.addShape('rect', {
    id: 'ContentParagraph',
    name: 'ContentParagraph',
    x: 80,
    y: 280,
    width: 530,
    height: 90,
    fill: '#1E293B',
    text: 'Read our comprehensive Developer Documentation or inspect the OpenAPI Specification for endpoints.',
    textStyle: { color: '#CBD5E1', fontSize: 12 },
  })

  // Substring 1: "Developer Documentation"
  ppt.setHyperlink('ContentParagraph', 'Developer Documentation', 'https://example.com/docs', {
    color: '#38BDF8',
    underline: true,
    bold: true,
    tooltip: 'Open Developer Documentation Portal',
  })

  // Substring 2: "OpenAPI Specification"
  ppt.setHyperlink('ContentParagraph', 'OpenAPI Specification', 'https://example.com/openapi.json', {
    color: '#F43F5E',
    underline: true,
    italic: true,
    tooltip: 'View OpenAPI JSON Spec',
  })
  console.log('  → Split text runs and applied custom colored hyperlinks to target substrings')

  // Populate Slide 2 with content and return button
  ppt.useSlide(2)
  await ppt.addShape('rect', {
    id: 'Slide2Header',
    x: 80,
    y: 50,
    width: 500,
    height: 60,
    fill: '#0F172A',
    text: 'Slide 2: Architecture Deep Dive',
    textStyle: { color: '#10B981', fontSize: 20, bold: true },
  })

  await ppt.addShape('roundRect', {
    id: 'ReturnHomeBtn',
    name: 'ReturnHomeBtn',
    x: 80,
    y: 140,
    width: 220,
    height: 45,
    fill: '#334155',
    text: 'Return to Slide 1 (Home)',
    textStyle: { color: '#F8FAFC', fontSize: 12, bold: true },
  })
  ppt.addHyperlink('ReturnHomeBtn', 1, { tooltip: 'Back to Slide 1' })

  // 6. getHyperlinks & getHyperlink — Read-only introspection
  console.log('\n[6/6] ppt.getHyperlinks() & getHyperlink() — Hyperlink Introspection...')
  const slide1Links = ppt.useSlide(1).getHyperlinks(1)
  console.log(`  → Total hyperlinks on Slide 1: ${slide1Links.length}`)
  for (const l of slide1Links) {
    console.log(
      `    - [${l.type}] Target: "${l.targetId}" | URL/Action: ${l.url || l.action || 'Slide ' + l.targetSlide} | Tooltip: "${l.tooltip || ''}"`
    )
  }

  const singleQuery = ppt.getHyperlink('GitHubBtn')
  console.log(`  → getHyperlink('GitHubBtn') result: ${singleQuery?.url}`)

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
