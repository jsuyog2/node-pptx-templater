const fs = require('fs');
const html = fs.readFileSync('docs/api-reference.html', 'utf8');

const expectedMethods = [
  // 1. Lifecycle & I/O
  'load', 'create', 'fromPresentationXml', 'preload', 'fromCache', 'clearCache', 'saveToFile', 'save', 'toBuffer', 'toStream', 'toBase64', 'toExpress', 'saveToFolder', 'buildPptx', 'extractPptx',
  // 2. Slides Management
  'useSlide', 'useAllSlides', 'addSlide', 'insertSlide', 'duplicateSlide', 'removeSlide', 'getSlides', 'getSlideDimensions', 'getSlideLayout', 'importSlideFrom', 'exportSlides',
  // 3. Text & Placeholders
  'replaceText', 'replaceMultiple', 'replaceTextByTag', 'updateText', 'findText', 'appendTextRun', 'prependTextRun', 'extractPlaceholders', 'getList', 'validateList',
  // 4. Tables & Cell Shapes
  'addTable', 'updateTable', 'getTableRows', 'addTableRow', 'mergeCells', 'cloneMergedRegion', 'autoFitTable', 'splitTableOnOverflow', 'getCellBounds', 'addCellShape', 'updateCellShape',
  // 5. Charts & Series Sync
  'updateChart', 'updateChartCategories', 'updateChartTitle', 'updateDataLabels', 'repairCharts', 'getCharts', 'getChartLabelPositions', 'addTextNearChartLabel',
  // 6. Shapes & Geometry
  'addShape', 'updateShape', 'updateShapeText', 'cloneShape', 'deleteShape', 'alignShapes', 'addTextAtPosition', 'getShapes', 'validateShape',
  // 7. Images & Media
  'addImage', 'replaceImage', 'getImages',
  // 8. Z-Order & Layering
  'bringToFront', 'sendToBack', 'bringForward', 'sendBackward', 'setZIndex', 'reorderObjects', 'swapObjects', 'getObjectOrder', 'normalizeZOrder',
  // 9. Hyperlinks & Actions
  'addHyperlink', 'addSlideLink', 'addImageLink', 'addShapeLink', 'addTextNavigationLink', 'addShapeNavigationLink', 'getHyperlinks',
  // 10. Validation & Diagnostics
  'validate', 'validatePresentation', 'validateTable', 'validateRelationships', 'repair', 'inspectXML', 'inspectSlide', 'inspectChart', 'debugRelationships', 'enablePerformanceProfile', 'getPerformanceMetrics', 'enableDebug'
];

console.log('Verifying 95 methods in docs/api-reference.html...');
const missingAnchors = [];
const missingSidebarLinks = [];

for (const m of expectedMethods) {
  if (!html.includes(`id="${m}"`)) missingAnchors.push(m);
  if (!html.includes(`href="#${m}"`)) missingSidebarLinks.push(m);
}

const copyBtnPattern = /<button class="copy-btn" onclick="copyCode\(this\)">Copy<\/button>/g;
const copyBtnMatches = (html.match(copyBtnPattern) || []).length;

console.log('Total methods tested:', expectedMethods.length);
console.log('Missing IDs:', missingAnchors);
console.log('Missing Sidebar/TOC Links:', missingSidebarLinks);
console.log('Total Copy buttons found:', copyBtnMatches);

if (missingAnchors.length === 0 && missingSidebarLinks.length === 0 && copyBtnMatches === 95) {
  console.log('✓ SUCCESS: All 95 methods have valid IDs, links, and copy buttons!');
} else {
  console.error('FAILED: Verification checks did not all pass.');
  process.exit(1);
}
