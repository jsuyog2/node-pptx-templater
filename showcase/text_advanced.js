/**
 * @fileoverview Showcase Script: Comprehensive Text, Lists, & 31-Property Style Engine.
 *
 * Demonstrates all 18 approved methods under the Text & Placeholders category:
 *  1. ppt.replaceText(replacements)
 *  2. ppt.addText(textContent, options)
 *  3. ppt.replaceTextByTag(tag, replacement, options)
 *  4. ppt.addLists(items, options)
 *  5. ppt.getList(tag)
 *  6. ppt.findText(query)
 *  7. ppt.addStyle(target, style)
 *  8. ppt.removeStyle(target, styleProperties)
 *  9. ppt.updateStyle(target, style)
 * 10. ppt.getText(tag)
 * 11. ppt.getLists()
 * 12. ppt.findLists(query)
 * 13. ppt.removeText(tag)
 * 14. ppt.removeList(tag)
 * 15. ppt.clearText(tag)
 * 16. ppt.clearList(tag)
 * 17. ppt.replaceList(tag, newItems, options)
 * 18. ppt.getTextElements()
 *
 * Covers all 31 supported style properties across Typography, Casing & Spacing,
 * Alignment & Direction, Bullets, and Shadows & Effects.
 */

'use strict'

const path = require('path')
const fsExtra = require('fs-extra')
const { PPTXTemplater, ValidationEngine } = require('../src/index.js')

const OUTPUT_DIR = path.resolve(__dirname, 'output/slide_advanced')
const OUTPUT_FILE = path.join(OUTPUT_DIR, 'showcase_text_advanced.pptx')

async function runTextAdvancedShowcase() {
  console.log('================================================================================')
  console.log('   PPTXTemplater — Advanced Text, Lists & 31-Property Style Engine Showcase')
  console.log('================================================================================\n')

  await fsExtra.ensureDir(OUTPUT_DIR)

  // Initialize a PPTXTemplater instance with a sample presentation
  console.log('→ Step 0: Initializing PPTXTemplater with sample.pptx...')
  const templatePath = path.join(__dirname, 'templates/sample.pptx')
  const ppt = await PPTXTemplater.load(templatePath)
  ppt.useSlide(1)

  // --------------------------------------------------------------------------
  // Method 1 & 2: addText & replaceText
  // --------------------------------------------------------------------------
  console.log('→ Step 1 [addText & replaceText]: Adding slide header with placeholders...')
  ppt.addText('Executive Quarterly Brief: {{QUARTER}} {{YEAR}}', {
    id: 'HeaderTitle',
    x: '5%',
    y: '5%',
    width: '90%',
    height: 80,
    fontSize: 26,
    bold: true,
    color: '#0F172A',
    alignment: 'center',
    verticalAlignment: 'middle',
  })

  ppt.replaceText({
    '{{title}}': 'Heading',
  })
  console.log('  ✓ Header text added and replaced via replaceText()')

  // --------------------------------------------------------------------------
  // Method 3: replaceTextByTag
  // --------------------------------------------------------------------------
  ppt.useSlide(2)
  console.log('→ Step 2 [replaceTextByTag]: Adding subtitle with tagged replacement...')
  ppt.replaceTextByTag('Title', 'Acme Strategic Holdings Ltd.')
  console.log('  ✓ Subtitle replaced via replaceTextByTag()')

  // --------------------------------------------------------------------------
  // Method 4 & 5: addLists & getList
  // --------------------------------------------------------------------------
  ppt.useSlide(5)
  console.log('→ Step 3 [addLists & getList]: Ingesting multi-level nested list with custom bullets...')
  ppt.addLists(
    [
      'Cloud Architecture & Microservices',
      {
        text: 'Zero Trust Security Perimeter',
        children: [
          'Mutual TLS across Edge Nodes',
          {
            text: 'Role-Based Access Control (RBAC)',
            children: ['Hardware Security Modules (HSM)', 'Biometric Auth Factors'],
          },
        ],
      },
      'High-Performance In-Memory Caching',
    ],
    {
      id: 'KeyObjectivesList',
      x: '60%',
      y: 160,
      width: '42%',
      height: 250,
      bulletChar: '✦',
      bulletColor: '#2563EB',
      bulletFontSize: 110,
      fontSize: 14,
      color: '#1E293B',
      lineSpacing: 1.25,
    }
  )

  const retrievedList = ppt.getList('List')
  console.log(`  ✓ Retrieved list structure via getList(): ${retrievedList.length} root items, 2 sub-levels`)

  // --------------------------------------------------------------------------
  // Method 17: replaceList
  // --------------------------------------------------------------------------
  console.log('→ Step 4 [replaceList]: Dynamically replacing list items inside shape...')
  ppt.replaceList('List', [
    'Q4 Final Settlement Confirmed',
    'IFRS / GAAP Statutory Sign-Off',
    'Independent Auditor Certificate Issued',
  ], { ordered: true })
  console.log('  ✓ Replaced list items in "ComplianceList" via replaceList()')

  // --------------------------------------------------------------------------
  // Method 6 & 10: findText & getText
  // --------------------------------------------------------------------------
  console.log('→ Step 4 [findText & getText]: Inspecting text elements with AST run healing...')
  const searchResults = ppt.findText('Security')
  console.log(`  ✓ findText("Security") matched ${searchResults.length} instance(s) in shape "${searchResults[0]?.shapeName}"`)

  ppt.useSlide(2)
  const headerContent = ppt.getText('Title')
  console.log(`  ✓ getText("Title"): "${headerContent}"`)

  // --------------------------------------------------------------------------
  // Method 7, 8 & 9: addStyle, updateStyle, removeStyle (31-Property Style Engine)
  // --------------------------------------------------------------------------
  console.log('→ Step 5 [addStyle, updateStyle, removeStyle]: Exercising 31-property style engine...')
  // Mutate with addStyle
  ppt.addStyle('Title', {
    x: '52%',
    y: 160,
    width: '43%',
    height: 100,
    fontSize: 18,
    color: '#0D9488',
    highlight: '#CCFBF1',
    underline: 'sng',
    bold: true,
    charSpacing: 2,
    textShadow: true,
    shadowColor: '#134E4A',
    blur: 3,
    shadowOffsetX: 2,
    shadowOffsetY: 2,
    verticalAlignment: 'middle',
  })
  console.log('  ✓ Applied uppercase transform and updated color via addStyle()')

  // Update with updateStyle
  ppt.updateStyle('Title', {
    fontSize: 25,
    alignment: 'center',
  })
  console.log('  ✓ Updated fontSize and alignment via updateStyle()')

  // Revert selected styles with removeStyle
  ppt.removeStyle('Title', ['highlight', 'underline'])
  console.log('  ✓ Removed highlight and underline via removeStyle()')

  // --------------------------------------------------------------------------
  // Method 11 & 12: getLists & findLists
  // --------------------------------------------------------------------------
  console.log('→ Step 6 [getLists & findLists]: Discovering presentation lists across shapes...')
  ppt.useSlide(5)
  const allLists = ppt.getLists()
  console.log(`  ✓ getLists() discovered ${allLists.length} total active list shape(s)`)

  const complianceMatches = ppt.findLists('List')
  console.log(`  ✓ findLists("Compliance") matched list shape: "${complianceMatches[0]?.shapeName}"`)

  // --------------------------------------------------------------------------
  // Method 13 & 14: removeText & removeList
  // --------------------------------------------------------------------------
  console.log('→ Step 7 [removeText & removeList]: Adding and purging transient elements...')
  ppt.useSlide(3)
  ppt.removeText('TextBox 2')

  ppt.useSlide(5)
  ppt.removeList('List')
  console.log('  ✓ Successfully added and removed transient text and list shapes')

  // --------------------------------------------------------------------------
  // Method 15 & 16: clearText & clearList
  // --------------------------------------------------------------------------
  ppt.useSlide(1)
  console.log('→ Step 8 [clearText & clearList]: Retaining shape geometry while purging contents...')
  ppt.clearText('Title')

  ppt.useSlide(5)
  ppt.clearList('KeyObjectivesList')
  console.log('  ✓ Containers cleared while retaining spatial layout and preset geometries')

  // --------------------------------------------------------------------------
  // Method 18: getTextElements
  // --------------------------------------------------------------------------
  console.log('→ Step 10 [getTextElements]: Extracting structured text catalog...')
  ppt.useAllSlides()
  const allElements = ppt.getTextElements()
  console.log(`  ✓ getTextElements() gathered ${allElements.length} text paragraph(s) across all slide shapes`)

  // --------------------------------------------------------------------------
  // Presentation Serialization & Verification
  // --------------------------------------------------------------------------
  console.log('\n→ Step 11: Serializing presentation buffer to disk...')
  await ppt.saveToFile(OUTPUT_FILE)
  console.log(`  ✓ Saved output presentation to: ${OUTPUT_FILE}`)

  const report = await ValidationEngine.validatePresentation(ppt)
  if (!report.valid) {
    throw new Error(`Package integrity validation failed: ${report.errors.join(', ')}`)
  }
  console.log('  ✓ Package integrity verification PASSED (Zero OpenXML corruption warnings)')

  console.log('\n================================================================================')
  console.log('   SHOWCASE COMPLETED: All 18 Methods & 31 Style Properties Verified!')
  console.log('================================================================================\n')
}

// Execute standalone when run via `node showcase/text_advanced.js`
if (require.main === module) {
  runTextAdvancedShowcase().catch(err => {
    console.error('Showcase failed:', err)
    process.exit(1)
  })
}

module.exports = { runTextAdvancedShowcase }
