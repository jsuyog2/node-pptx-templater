const fs = require('fs-extra');
const { resolve } = require('path');
const path = require('path');
const { execSync } = require('child_process');

const DOCS_DIR = resolve(__dirname, '../docs');
const pkg = require('../package.json');
const VERSION = pkg.version;

// API Categories Definition
const API_CATEGORIES = {
  tables: { title: 'Tables API', desc: 'Manipulate DrawingML tables, rows, cells, auto-fitting, resizing, and cell merging.' },
  charts: { title: 'Charts API', desc: 'Sync category series, update Excel spreadsheets data caches, change titles, and manage charts.' },
  slides: { title: 'Slides API', desc: 'Duplicate, move, import, delete, and structure layout sections of slides.' },
  text: { title: 'Text API', desc: 'Execute text tag replacements, un-fragment text runs, search string tags, and apply links.' },
  images: { title: 'Images API', desc: 'Add images dynamically, swap templates, and extract file lists.' },
  shapes: { title: 'Shapes API', desc: 'Modify text within shapes, clone layout blocks, and delete shapes.' },
  layers: { title: 'Layer Stacking (Z-Order)', desc: 'Manage elements stack sorting (Bring Forward, Send to Back) within presentation slide trees.' },
  utils: { title: 'Utilities & Validation', desc: 'Load files, inspect XML elements, validation checks, and repair ZIP packages.' }
};

// Map of PPTXTemplater public method names to category keys
const METHOD_CATEGORIES = {
  // Tables
  'updateTable': 'tables',
  'addTableRow': 'tables',
  'removeTableRow': 'tables',
  'insertTableRow': 'tables',
  'cloneTableRow': 'tables',
  'updateCell': 'tables',
  'mergeCells': 'tables',
  'unmergeCells': 'tables',
  'getMergedCells': 'tables',
  'validateMergeRegion': 'tables',
  'isMergedCell': 'tables',
  'getMergeParent': 'tables',
  'getMergeRegion': 'tables',
  'splitMergedRegion': 'tables',
  'cloneMergedRegion': 'tables',
  'autoFitTable': 'tables',
  'resizeTable': 'tables',
  'getTables': 'tables',

  // Charts
  'updateChart': 'charts',
  'updateChartData': 'charts',
  'replaceChartSeries': 'charts',
  'updateChartTitle': 'charts',
  'updateChartCategories': 'charts',
  'getCharts': 'charts',
  'validateCharts': 'charts',
  'repairCharts': 'charts',

  // Slides
  'useSlide': 'slides',
  'useAllSlides': 'slides',
  'addSlide': 'slides',
  'duplicateSlide': 'slides',
  'deleteSlide': 'slides',
  'moveSlide': 'slides',
  'insertSlide': 'slides',
  'getSlides': 'slides',
  'cloneSlide': 'slides',
  'removeSlide': 'slides',
  'reorderSlides': 'slides',
  'tagSlide': 'slides',
  'exportSlides': 'slides',
  'importSlideFrom': 'slides',
  'importSlides': 'slides',

  // Text & Links
  'replaceText': 'text',
  'replaceTextByTag': 'text',
  'replaceMultiple': 'text',
  'findText': 'text',
  'getTextElements': 'text',
  'addHyperlink': 'text',
  'addSlideLink': 'text',
  'addImageLink': 'text',
  'addShapeLink': 'text',
  'addTextNavigationLink': 'text',
  'addShapeNavigationLink': 'text',

  // Images
  'replaceImage': 'images',
  'addImage': 'images',
  'removeImage': 'images',
  'getImages': 'images',

  // Shapes
  'updateShapeText': 'shapes',
  'cloneShape': 'shapes',
  'deleteShape': 'shapes',
  'getShapes': 'shapes',

  // Layer Management
  'bringForward': 'layers',
  'sendBackward': 'layers',
  'bringToFront': 'layers',
  'sendToBack': 'layers',
  'setZIndex': 'layers',
  'moveObjectBefore': 'layers',
  'moveObjectAfter': 'layers',
  'reorderObjects': 'layers',
  'getObjectOrder': 'layers',
  'applyZOrder': 'layers',
  'getTopMostObject': 'layers',
  'getBottomMostObject': 'layers',
  'swapObjects': 'layers',
  'sortObjects': 'layers',
  'normalizeZOrder': 'layers',

  // Utilities & Core
  'load': 'utils',
  'create': 'utils',
  'getInfo': 'utils',
  'validate': 'utils',
  'repair': 'utils',
  'saveToFile': 'utils',
  'toBuffer': 'utils',
  'toStream': 'utils',
  'validatePresentation': 'utils',
  'validateSlide': 'utils',
  'validateTable': 'utils',
  'validateRelationships': 'utils',
  'inspectSlide': 'utils',
  'inspectXML': 'utils',
  'inspectChart': 'utils',
  'inspectChartXML': 'utils',
  'debugRelationships': 'utils',
  'debugChartRelationships': 'utils',
  'slideCount': 'utils',
  'zipManager': 'utils',
  'xmlParser': 'utils',
  'contentTypesManager': 'utils',
  'relationshipManager': 'utils',
  'slideManager': 'utils',
  'chartManager': 'utils',
  'tableManager': 'utils',
  'shapeManager': 'utils',
  'imageManager': 'utils',
  'textManager': 'utils',
  'hyperlinkManager': 'utils',
  'mediaManager': 'utils'
};

// Rich comprehensive database of methods to fill JSDoc gaps
const API_METADATA_EXTENSIONS = {
  // --- TABLES ---
  updateTable: {
    edgeCases: 'If tableId is not found, throws TableNotFoundError. Coordinates outside structural table dimensions fail silently or align incorrectly depending on row dimensions.',
    errorHandling: 'Wrap in try/catch block to catch TableNotFoundError or SlideNotFoundError if active slide selection is empty.',
    related: ['addTableRow', 'removeTableRow', 'mergeCells'],
    xmlImpact: 'Updates `<a:t>` element text inside individual table cells under `<a:tc>`, adjusting `<a:tcPr>` border styles and alignment rules.',
    examples: {
      basic: `ppt.useSlide(1).updateTable('summary-table', [\n  ['Item', 'Value'],\n  ['Widgets', '1,200'],\n  ['Gadgets', '850']\n]);`,
      advanced: `ppt.useSlide(2).updateTable('details-table', [\n  ['Product', 'Revenue', { value: 'Growth', align: 'ctr', fill: '22c55e', bold: true }],\n  ['SaaS', '$45,000', '15.4%'],\n  ['Licensing', '$12,000', '-2.1%']\n]);`,
      production: `const reportData = await fetchReportMetrics();\nppt.useSlide(3)\n   .updateTable('sales-table', [\n     ['Q1 Metric', 'Performance', 'Target Margin'],\n     ...reportData.map(r => [\n       r.metric,\n       { value: r.perf, fill: r.perf >= r.target ? '10b981' : 'ef4444' },\n       r.target\n     ])\n   ]);`
    }
  },
  mergeCells: {
    edgeCases: 'Start Row/Col must be less than or equal to End Row/Col. Cells outside the grid boundaries throw RangeError.',
    errorHandling: 'Validates bounds before running. Best practice is to run `validateMergeRegion()` first.',
    related: ['unmergeCells', 'validateMergeRegion', 'getMergedCells'],
    xmlImpact: 'Declares `gridSpan` and `rowSpan` in the top-left origin `<a:tc>`, and applies `hMerge="1"` or `vMerge="1"` properties to all shadowed cells inside the merged region block.',
    examples: {
      basic: `ppt.useSlide(1).mergeCells('metrics-table', 1, 1, 2, 2);`,
      advanced: `// Pass coordinates config directly as an object\nppt.mergeCells({\n  tableId: 'metrics-table',\n  slide: 1,\n  startRow: 0,\n  startCol: 0,\n  endRow: 1,\n  endCol: 2\n});`,
      production: `if (ppt.validateMergeRegion('stats-table', 1, 1, 3, 2).valid) {\n  ppt.mergeCells('stats-table', 1, 1, 3, 2);\n} else {\n  console.warn('Cannot merge cells: overlapping region');\n}`
    }
  },
  unmergeCells: {
    edgeCases: 'Target must overlap an active merge region origin. Calling on an unmerged region acts as a no-op.',
    errorHandling: 'Gracefully ignores if cell coordinates do not match any active merge region.',
    related: ['mergeCells', 'splitMergedRegion'],
    xmlImpact: 'Removes `gridSpan`, `rowSpan`, `hMerge`, and `vMerge` attributes from all target cells in the region.',
    examples: {
      basic: `ppt.useSlide(1).unmergeCells('metrics-table', 1, 1, 2, 2);`,
      advanced: `// Target a specific cell inside the merge to split it\nppt.unmergeCells({\n  tableId: 'metrics-table',\n  slide: 1,\n  row: 1,\n  col: 1\n});`,
      production: `const regions = ppt.getMergedCells('data-table');\nregions.forEach(reg => {\n  if (reg.startRow === 0) {\n    ppt.unmergeCells('data-table', reg.startRow, reg.startCol, reg.endRow, reg.endCol);\n  }\n});`
    }
  },
  addTableRow: {
    edgeCases: 'Target table must exist. Row array must match table column count, otherwise cells are padded or truncated.',
    errorHandling: 'Validates columns layout before appending. Generates custom rowId hashes.',
    related: ['insertTableRow', 'removeTableRow'],
    xmlImpact: 'Appends a `<a:tr>` child block to the table, and assigns a unique collaborative ID `<a16:rowId>` to prevent file corruption.',
    examples: {
      basic: `ppt.useSlide(1).addTableRow('data-table', ['John Doe', 'Sales Manager', '$120k']);`,
      advanced: `// Add a styled row\nppt.useSlide(1).addTableRow('data-table', [\n  'Jane Smith',\n  { value: 'Director', bold: true, align: 'ctr' },\n  { value: '$180k', fill: '10b981' }\n]);`,
      production: `users.forEach(u => {\n  ppt.addTableRow('user-list-table', [u.name, u.role, u.salary]);\n});`
    }
  },
  removeTableRow: {
    edgeCases: 'Index must fall within table bounds (0 to rows.length - 1). Deleting the last row of a table leaves it structurally empty, which might alert alerts in PowerPoint.',
    errorHandling: 'Throws RangeError if rowIndex is out of bounds.',
    related: ['addTableRow', 'insertTableRow'],
    xmlImpact: 'Removes the target `<a:tr>` node entirely from the slide table XML.',
    examples: {
      basic: `ppt.useSlide(1).removeTableRow('data-table', 2);`,
      advanced: `// Delete the second row (index 1)\nppt.useSlide(1).removeTableRow('data-table', 1);`,
      production: `const tables = ppt.getTables();\nconst targetTable = tables.find(t => t.id === 'user-table');\nif (targetTable && targetTable.rows > 5) {\n  ppt.removeTableRow('user-table', targetTable.rows - 1);\n}`
    }
  },

  // --- CHARTS ---
  updateChartData: {
    edgeCases: 'Dataset categories and series coordinates must match initial chart mappings structure to avoid Excel workbook cell mismatches.',
    errorHandling: 'Throws ChartNotFoundError if chartId is missing. Verifies zip integrity before parsing.',
    related: ['updateChart', 'replaceChartSeries'],
    xmlImpact: 'Rewrites numerical categories and series data caches in chart XML, and updates cell values in the backing `ppt/embeddings/*.xlsx` workbook in the ZIP archive.',
    examples: {
      basic: `ppt.useSlide(1).updateChartData('sales-chart', {\n  categories: ['Q1', 'Q2'],\n  series: [{ name: 'Actual', values: [100, 150] }]\n});`,
      advanced: `ppt.useSlide(2).updateChartData('sales-chart', {\n  categories: ['Jan', 'Feb', 'Mar'],\n  series: [\n    { name: 'Target', values: [100, 120, 140] },\n    { name: 'Actual', values: [105, 118, 145] }\n  ]\n});`,
      production: `const salesData = await fetchSalesData();\nppt.useSlide(1).updateChartData('revenue-gauge', {\n  categories: salesData.months,\n  series: [\n    { name: 'Direct Sales', values: salesData.direct },\n    { name: 'Channel Sales', values: salesData.channel }\n  ]\n});`
    }
  },
  updateChartTitle: {
    edgeCases: 'If the template chart has no initial title layout block, adding a title might require injecting new XML blocks.',
    errorHandling: 'Safely overrides title runs if title node exists; otherwise updates layout or prints warnings.',
    related: ['updateChartData', 'updateChartCategories'],
    xmlImpact: 'Finds and replaces text run values inside `<c:title>` and child `<a:t>` nodes under chart structure.',
    examples: {
      basic: `ppt.useSlide(1).updateChartTitle('sales-chart', 'Quarterly Metrics Overview');`,
      advanced: `ppt.useSlide(2).updateChartTitle('revenue-chart', 'Global SaaS Revenue (2026)');`,
      production: `const q = getActiveQuarter();\nppt.useSlide(1).updateChartTitle('kpi-chart', \`Quarter \${q} Performance Summary\`);`
    }
  },

  // --- SLIDES ---
  duplicateSlide: {
    edgeCases: 'Index parameters are 1-based. Position index must fall between 1 and slideCount + 1.',
    errorHandling: 'Throws RangeError if indices fall outside the available collection bounds.',
    related: ['cloneSlide', 'deleteSlide', 'moveSlide'],
    xmlImpact: 'Duplicates slide layout XML file, copies all elements relationships, and appends a slide reference entry inside `ppt/presentation.xml`.',
    examples: {
      basic: `ppt.duplicateSlide(1, 2);`,
      advanced: `// Duplicate slide 1 and insert at the end\nconst count = ppt.slideCount;\nppt.duplicateSlide(1, count + 1);`,
      production: `const items = await getPortfolioItems();\nitems.forEach((item, index) => {\n  ppt.duplicateSlide(2, 3 + index);\n  ppt.useSlide(3 + index)\n     .replaceTextByTag('title', item.title)\n     .replaceTextByTag('desc', item.description);\n});`
    }
  },
  importSlideFrom: {
    edgeCases: 'Source deck must be loaded first. Deduplicates layouts, media assets, and themes to prevent PowerPoint Repair Mode prompts.',
    errorHandling: 'Validates target slide index in the source presentation. Throws error if source presentation is invalid.',
    related: ['importSlides', 'exportSlides'],
    xmlImpact: 'Remaps slide-level relationships to root presentation indices, duplicates layout links, and copies media files into the target ZIP package.',
    examples: {
      basic: `const source = await PPTXTemplater.load('template2.pptx');\nawait ppt.importSlideFrom(source, 1);`,
      advanced: `const source = await PPTXTemplater.load('slide_deck.pptx');\nawait ppt.useSlide(2).importSlideFrom(source, 'marketing-overview-slide');`,
      production: `const appendixDeck = await PPTXTemplater.load('appendix.pptx');\nfor (let i = 1; i <= appendixDeck.slideCount; i++) {\n  await ppt.importSlideFrom(appendixDeck, i);\n}`
    }
  },

  // --- TEXT ---
  replaceTextByTag: {
    edgeCases: 'Tags split across multiple text runs are healed before replacements. Standard layout handles exact matches.',
    errorHandling: 'Gracefully ignores if tag name is not present on the selected slide canvas.',
    related: ['replaceMultiple', 'replaceText'],
    xmlImpact: 'Merges fragmented `<a:r>` runs inside text paragraphs and replaces string values while preserving run font parameters.',
    examples: {
      basic: `ppt.useSlide(1).replaceTextByTag('company', 'Acme Corp');`,
      advanced: `// Search-replace with custom text configuration\nppt.useSlide(1).replaceTextByTag('year', '2026', { bold: true });`,
      production: `const profile = await getUserProfile();\nppt.useSlide(1)\n   .replaceTextByTag('firstName', profile.first)\n   .replaceTextByTag('lastName', profile.last)\n   .replaceTextByTag('email', profile.email);`
    }
  },

  // --- IMAGES ---
  replaceImage: {
    edgeCases: 'Image placeholder name/id must exist in template. Image types must match or content-type overrides must be registered.',
    errorHandling: 'Throws error if image reference is missing or if media buffer is corrupted.',
    related: ['addImage', 'removeImage'],
    xmlImpact: 'Overwrites target media target inside presentation archive or updates relationship mapping target references to point to the new image.',
    examples: {
      basic: `await ppt.useSlide(1).replaceImage('logo-img', './new-logo.png');`,
      advanced: `// Pass a binary buffer\nconst imgBuffer = fs.readFileSync('avatar.jpg');\nawait ppt.useSlide(1).replaceImage('profile-avatar', imgBuffer);`,
      production: `const users = await getTeamMembers();\nfor (let i = 0; i < users.length; i++) {\n  const user = users[i];\n  const avatarBuffer = await fetchAvatar(user.id);\n  ppt.useSlide(2 + i);\n  await ppt.replaceImage('user-avatar', avatarBuffer);\n}`
    }
  },

  // --- SHAPES ---
  cloneShape: {
    edgeCases: 'Offsets are in English Metric Units (EMU). Offset parameters specify position displacement.',
    errorHandling: 'Throws error if target shape is not found on active slide.',
    related: ['updateShapeText', 'deleteShape'],
    xmlImpact: 'Duplicates the target shape element XML node (`p:sp`) and adds positioning coordinates.',
    examples: {
      basic: `ppt.useSlide(1).cloneShape('card-bg', 'card-bg-2');`,
      advanced: `// Clone with specific displacement offsets\nppt.useSlide(1).cloneShape('card-bg', 'card-bg-2', {\n  offsetX: 360000, // 360,000 EMUs = ~1 inch\n  offsetY: 0\n});`,
      production: `const items = ['Speed', 'Stability', 'Scalability'];\nitems.forEach((item, index) => {\n  if (index > 0) {\n    ppt.cloneShape('bullet-template', \`bullet-\${index}\`, {\n      offsetX: 0,\n      offsetY: index * 400000\n    }).updateShapeText(\`bullet-\${index}\`, item);\n  } else {\n    ppt.updateShapeText('bullet-template', item);\n  }\n});`
    }
  },

  // --- LAYER MANAGEMENT ---
  bringToFront: {
    edgeCases: 'Object ID or name must match an element on slide canvas. No-op if object is already topmost.',
    errorHandling: 'Throws error if objectId is not found on targeted slide index.',
    related: ['bringForward', 'sendToBack', 'getObjectOrder'],
    xmlImpact: 'Reorders slide element nodes under `<p:spTree>`, moving the matched node to the end of the tag sequence list.',
    examples: {
      basic: `ppt.useSlide(1).bringToFront('OverlayLogo');`,
      advanced: `// Or pass as config object directly\nppt.bringToFront({ slide: 1, objectId: 'OverlayLogo' });`,
      production: `// Bring all images to the front\nconst layers = ppt.getObjectOrder(1);\nlayers.forEach(lay => {\n  if (lay.type === 'image') {\n    ppt.bringToFront(lay.id);\n  }\n});`
    }
  },

  // --- UTILITIES & CORE ---
  load: {
    edgeCases: 'Loads either string filepath or Buffer. Package must be a valid, uncorrupted OpenXML ZIP archive.',
    errorHandling: 'Throws PPTXError if package read fails or file format is invalid.',
    related: ['create', 'saveToFile'],
    xmlImpact: 'Decompresses ZIP archive, indexes content types, preloads all relationship directories, and caches XML nodes.',
    examples: {
      basic: `const ppt = await PPTXTemplater.load('./my_template.pptx');`,
      advanced: `const buffer = fs.readFileSync('template.pptx');\nconst ppt = await PPTXTemplater.load(buffer);`,
      production: `async function generateFromS3(s3Buffer) {\n  try {\n    const ppt = await PPTXTemplater.load(s3Buffer);\n    // Perform operations\n    return await ppt.toBuffer();\n  } catch (err) {\n    console.error('Error reading template from S3:', err);\n    throw err;\n  }\n}`
    }
  }
};

// Fill in other empty methods in API_METADATA_EXTENSIONS
function completeDatabaseMetadata(apis) {
  apis.forEach(api => {
    const meta = API_METADATA_EXTENSIONS[api.name] || {};
    api.category = METHOD_CATEGORIES[api.name] || 'utils';
    api.edgeCases = meta.edgeCases || 'No critical edge cases documented. Verify argument boundaries.';
    api.errorHandling = meta.errorHandling || 'Wrap inside try/catch blocks to process execution errors.';
    api.related = meta.related || [];
    api.xmlImpact = meta.xmlImpact || 'Modifies underlying OpenXML nodes to reflect updates on slide serialization.';
    api.examples = meta.examples || {
      basic: `ppt.useSlide(1).${api.name}(${api.args});`,
      advanced: `ppt.useSlide(1).${api.name}(${api.args}); // Fluent wrapper implementation`,
      production: `try {\n  ppt.useSlide(1).${api.name}(${api.args});\n} catch (err) {\n  console.error('API Error: ', err);\n}`
    };
  });
}

function getPublicMethodsFromSource() {
  const code = fs.readFileSync(resolve(__dirname, '../src/core/PPTXTemplater.js'), 'utf-8');
  // Strip out comments
  const cleanCode = code.replace(/\/\*[\s\S]*?\*\/|\/\/.*/g, '');
  
  // Match method signatures: either async or standard, e.g. name(...)
  const methodRegex = /(?:async\s+)?([#a-zA-Z0-9_]+)\s*\(([^)]*)\)\s*\{/g;
  const methods = new Set();
  const JS_KEYWORDS = new Set(['if', 'for', 'while', 'catch', 'switch', 'with', 'function']);
  let match;
  while ((match = methodRegex.exec(cleanCode)) !== null) {
    const methodName = match[1];
    if (
      !methodName.startsWith('#') &&
      !methodName.startsWith('_') &&
      methodName !== 'constructor' &&
      methodName !== 'load' && // static
      methodName !== 'create' && // static
      !JS_KEYWORDS.has(methodName)
    ) {
      methods.add(methodName);
    }
  }
  
  // Match getters too
  const getterRegex = /get\s+([a-zA-Z0-9_]+)\s*\(\)\s*\{/g;
  while ((match = getterRegex.exec(cleanCode)) !== null) {
    const getterName = match[1];
    if (!getterName.startsWith('#') && !JS_KEYWORDS.has(getterName)) {
      methods.add(getterName);
    }
  }

  return Array.from(methods);
}

/**
 * Crawls JSDocs and matches them with class declarations
 */
function extractAPIDocs() {
  const srcDir = resolve(__dirname, '../src');
  
  function walk(dir) {
    let results = [];
    const list = fs.readdirSync(dir);
    list.forEach(file => {
      const filePath = path.join(dir, file);
      const stat = fs.statSync(filePath);
      if (stat && stat.isDirectory()) {
        results = results.concat(walk(filePath));
      } else if (file.endsWith('.js')) {
        results.push(filePath);
      }
    });
    return results;
  }
  
  const files = walk(srcDir);
  const apis = [];
  
  files.forEach(file => {
    const relPath = path.relative(resolve(__dirname, '..'), file).replace(/\\/g, '/');
    const content = fs.readFileSync(file, 'utf-8');
    
    // Find class name
    const classMatch = content.match(/class\s+([a-zA-Z0-9_]+)/);
    const className = classMatch ? classMatch[1] : path.basename(file, '.js');
    
    // Pattern to match JSDoc comments followed by method signature
    const docPattern = /\/\*\*([\s\S]*?)\*\/\s*(?:async\s+|get\s+)?([a-zA-Z0-9_#]+)\s*(?:\(([^)]*)\))?/g;
    let match;
    
    while ((match = docPattern.exec(content)) !== null) {
      const jsdoc = match[1];
      const name = match[2];
      const args = match[3] || '';
      
      // Ignore private methods
      if (name.startsWith('#') || name === 'constructor') {
        continue;
      }
      
      // Only process PPTXTemplater methods and root utilities for public API
      if (className !== 'PPTXTemplater' && !relPath.includes('src/utils/xmlUtils.js') && !relPath.includes('src/utils/relationshipUtils.js')) {
        continue;
      }
      
      const lines = jsdoc.split('\n').map(l => l.replace(/^\s*\*\s?/, '').trim());
      const descLines = [];
      const params = [];
      let returns = null;
      const examples = [];
      let inExample = false;
      
      for (let line of lines) {
        if (line.startsWith('@param')) {
          inExample = false;
          const paramMatch = line.match(/@param\s+\{([^}]+)\}\s+([^\s-]+)(?:\s+-\s+(.*))?/);
          if (paramMatch) {
            params.push({
              type: paramMatch[1],
              name: paramMatch[2],
              desc: paramMatch[3] || ''
            });
          }
        } else if (line.startsWith('@returns')) {
          inExample = false;
          const returnMatch = line.match(/@returns\s+\{([^}]+)\}(?:\s+(.*))?/);
          if (returnMatch) {
            returns = {
              type: returnMatch[1],
              desc: returnMatch[2] || ''
            };
          }
        } else if (line.startsWith('@example')) {
          inExample = true;
        } else if (line.startsWith('@class') || line.startsWith('@fileoverview') || line.startsWith('@description') || line.startsWith('@private') || line.startsWith('@internal')) {
          inExample = false;
        } else {
          if (inExample) {
            examples.push(line);
          } else {
            descLines.push(line);
          }
        }
      }
      
      const description = descLines.filter(Boolean).join(' ');
      const exampleCode = examples.filter(l => l !== '').join('\n');
      
      // Deduplicate
      if (!apis.find(a => a.name === name)) {
        apis.push({
          file: relPath,
          className,
          name,
          args: args.trim(),
          description,
          params,
          returns,
          exampleCode
        });
      }
    }
  });
  
  // Fill undocumented wrapper methods of PPTXTemplater
  const sourceMethods = getPublicMethodsFromSource();
  sourceMethods.push('load', 'create');
  
  sourceMethods.forEach(method => {
    if (!apis.find(a => a.name === method)) {
      // Find implementation inside manager JS JSDocs if exists
      let desc = 'Delegates core actions to slide element sub-managers.';
      let args = '()';
      let params = [];
      let returns = { type: 'PPTXTemplater', desc: 'The fluent engine instance.' };
      
      apis.push({
        file: 'src/core/PPTXTemplater.js',
        className: 'PPTXTemplater',
        name: method,
        args,
        description: desc,
        params,
        returns,
        exampleCode: ''
      });
    }
  });
  
  completeDatabaseMetadata(apis);
  return apis;
}

function validateAPIBuild(apis) {
  const sourceMethods = getPublicMethodsFromSource();
  sourceMethods.push('load', 'create');
  
  const unmapped = [];
  for (const method of sourceMethods) {
    if (!METHOD_CATEGORIES[method]) {
      unmapped.push(method);
    }
  }
  
  if (unmapped.length > 0) {
    console.error('Validation failed: The following public methods are not categorized in METHOD_CATEGORIES:', unmapped);
    process.exit(1);
  }
  
  console.log(`Validation passed: All ${sourceMethods.length} methods correctly mapped and documented!`);
}

function escapeHtml(str) {
  if (!str) return '';
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function renderSidebarAPIs(apis) {
  let html = '';
  for (const [catId, catInfo] of Object.entries(API_CATEGORIES)) {
    const catApis = apis.filter(a => a.category === catId);
    
    html += `
      <div class="space-y-1 mt-4 class-sidebar-group" id="group-api-${catId}">
        <h3 class="text-xs uppercase text-gray-500 font-semibold tracking-wider px-3 flex items-center justify-between cursor-pointer hover:text-white transition-colors" onclick="toggleSidebarGroup('api-${catId}')">
          <span>${escapeHtml(catInfo.title)}</span>
          <span class="text-[9px] transform transition-transform duration-200" id="arrow-api-${catId}" style="transform: rotate(-90deg);">▼</span>
        </h3>
        <ul id="sidebar-list-api-${catId}" class="space-y-1 pl-2 transition-all duration-300 hidden">
          <li><a href="#api-${catId}" class="nav-item block px-3 py-1.5 text-xs text-gray-400 hover:text-white rounded-md transition-colors font-medium">Overview</a></li>
          ${catApis.map(m => `
            <li><a href="#method-${m.name}" class="nav-item block px-3 py-1 text-[11px] text-gray-500 hover:text-white rounded-md transition-colors pl-4 font-mono">${escapeHtml(m.name)}()</a></li>
          `).join('')}
        </ul>
      </div>
    `;
  }
  return html;
}

function renderContentAPIs(apis) {
  let html = '';
  for (const [catId, catInfo] of Object.entries(API_CATEGORIES)) {
    const catApis = apis.filter(a => a.category === catId);
    
    html += `
      <!-- Category ${catId} Section -->
      <section id="api-${catId}" class="doc-section space-y-6 hidden">
        <div class="flex items-center gap-3 border-b border-white/5 pb-4">
          <span class="px-2 py-0.5 rounded bg-brand-500/10 text-brand-400 font-mono text-[10px] font-bold uppercase tracking-wider">category</span>
          <h1 class="font-title text-3xl font-extrabold text-white">${escapeHtml(catInfo.title)}</h1>
        </div>
        
        <p class="text-sm text-gray-400 leading-relaxed">${escapeHtml(catInfo.desc)}</p>
        
        <div class="grid grid-cols-1 gap-6 mt-6">
    `;
    
    catApis.forEach(m => {
      const signature = `${m.name}(${m.args})`;
      
      let paramsTable = '';
      if (m.params.length > 0) {
        paramsTable = `
          <div class="overflow-x-auto mt-4 border border-white/5 rounded-xl bg-slate-900/10">
            <table class="w-full text-xs text-gray-400">
              <thead class="bg-white/5 text-gray-200">
                <tr>
                  <th class="px-4 py-2.5 text-left font-semibold">Parameter</th>
                  <th class="px-4 py-2.5 text-left font-semibold">Type</th>
                  <th class="px-4 py-2.5 text-left font-semibold">Description</th>
                </tr>
              </thead>
              <tbody class="divide-y divide-white/5">
                ${m.params.map(p => `
                  <tr class="hover:bg-white/[0.02]">
                    <td class="px-4 py-2.5 font-semibold text-indigo-300 font-mono">${escapeHtml(p.name)}</td>
                    <td class="px-4 py-2.5 font-mono text-pink-400">${escapeHtml(p.type)}</td>
                    <td class="px-4 py-2.5 text-gray-300">${escapeHtml(p.desc)}</td>
                  </tr>
                `).join('')}
              </tbody>
            </table>
          </div>
        `;
      }
      
      let returnsBlock = '';
      if (m.returns) {
        returnsBlock = `
          <div class="mt-4 p-3 bg-white/[0.02] border border-white/5 rounded-lg text-xs text-gray-400 flex items-center gap-2">
            <span class="font-semibold text-gray-200 uppercase tracking-wider text-[10px] bg-emerald-500/10 text-emerald-400 px-2 py-0.5 rounded">returns</span>
            <span class="font-mono text-emerald-400 font-semibold">${escapeHtml(m.returns.type)}</span>
            <span class="text-gray-300">${escapeHtml(m.returns.desc)}</span>
          </div>
        `;
      }
      
      let exampleBlock = '';
      if (m.examples) {
        exampleBlock = `
          <div class="mt-4 space-y-4">
            <div>
              <span class="text-xxs font-semibold text-gray-500 block mb-1 font-title uppercase tracking-wider">Basic Usage</span>
              <pre class="relative group"><code class="language-javascript block p-4 bg-[#05070c] border border-white/5 rounded-xl text-indigo-300 font-mono text-xs overflow-x-auto">${escapeHtml(m.examples.basic)}</code><button class="copy-btn absolute top-3.5 right-3.5 text-[10px] bg-white/5 hover:bg-white/10 text-gray-400 hover:text-white px-2.5 py-1 rounded-lg border border-white/10 transition-all">Copy</button></pre>
            </div>
            <div>
              <span class="text-xxs font-semibold text-gray-500 block mb-1 font-title uppercase tracking-wider">Advanced Usage</span>
              <pre class="relative group"><code class="language-javascript block p-4 bg-[#05070c] border border-white/5 rounded-xl text-indigo-300 font-mono text-xs overflow-x-auto">${escapeHtml(m.examples.advanced)}</code><button class="copy-btn absolute top-3.5 right-3.5 text-[10px] bg-white/5 hover:bg-white/10 text-gray-400 hover:text-white px-2.5 py-1 rounded-lg border border-white/10 transition-all">Copy</button></pre>
            </div>
            <div>
              <span class="text-xxs font-semibold text-gray-500 block mb-1 font-title uppercase tracking-wider">Production Setup</span>
              <pre class="relative group"><code class="language-javascript block p-4 bg-[#05070c] border border-white/5 rounded-xl text-indigo-300 font-mono text-xs overflow-x-auto">${escapeHtml(m.examples.production)}</code><button class="copy-btn absolute top-3.5 right-3.5 text-[10px] bg-white/5 hover:bg-white/10 text-gray-400 hover:text-white px-2.5 py-1 rounded-lg border border-white/10 transition-all">Copy</button></pre>
            </div>
          </div>
        `;
      }
      
      html += `
        <!-- Method ${m.name} Card -->
        <div id="method-${m.name}" class="glass-card method-card p-6 bg-slate-900/30 border border-white/5 rounded-2xl space-y-4 relative overflow-hidden transition-all duration-300 hover:border-brand-500/30">
          <div class="flex flex-wrap items-center justify-between gap-3 border-b border-white/5 pb-3">
            <div class="flex items-center gap-2">
              <span class="px-2 py-0.5 rounded bg-brand-500/10 text-brand-400 font-mono text-[9px] font-bold uppercase tracking-wider">method</span>
              <h3 class="font-title text-xl font-bold text-white hover:text-brand-400 transition-colors">${escapeHtml(m.name)}</h3>
            </div>
            <div class="font-mono text-xs text-indigo-400/90 bg-indigo-500/5 border border-indigo-500/10 px-3 py-1 rounded-full">
              ${escapeHtml(signature)}
            </div>
          </div>
          
          <p class="text-sm text-gray-300 leading-relaxed">${escapeHtml(m.description || 'No description available.')}</p>
          
          <div class="grid grid-cols-1 md:grid-cols-2 gap-4 text-xs text-gray-400 bg-white/[0.01] border border-white/5 rounded-xl p-4">
            <div>
              <strong class="text-gray-300 block mb-1">🛡️ Edge Cases & Warnings</strong>
              <p>${escapeHtml(m.edgeCases)}</p>
            </div>
            <div>
              <strong class="text-gray-300 block mb-1">🏗️ OpenXML & ZIP Impact</strong>
              <p>${escapeHtml(m.xmlImpact)}</p>
            </div>
          </div>
          
          ${paramsTable}
          ${returnsBlock}
          ${exampleBlock}
        </div>
      `;
    });
    
    html += `
        </div>
      </section>
    `;
  }
  
  return html;
}

function generateReadmeAPI(apis) {
  let md = '';
  for (const [catId, catInfo] of Object.entries(API_CATEGORIES)) {
    const catApis = apis.filter(a => a.category === catId);
    md += `\n### ${catInfo.title}\n\n`;
    
    catApis.forEach(m => {
      md += `#### \`${m.name}(${m.args})\`\n`;
      md += `${m.description}\n\n`;
      if (m.params.length > 0) {
        md += `* **Arguments**:\n`;
        m.params.forEach(p => {
          md += `  * \`${p.name}\` (\`${p.type}\`): ${p.desc}\n`;
        });
      }
      if (m.returns) {
        md += `* **Returns**: \`${m.returns.type}\` - ${m.returns.desc}\n`;
      }
      if (m.examples && m.examples.basic) {
        md += `\n\`\`\`javascript\n${m.examples.basic}\n\`\`\`\n\n`;
      }
    });
    md += `---\n`;
  }
  return md.trim();
}

async function syncReadme(apis) {
  const readmePath = resolve(__dirname, '../README.md');
  if (!fs.existsSync(readmePath)) return;
  
  let readme = await fs.readFile(readmePath, 'utf-8');
  const startTag = '<!-- API_REFERENCE_START -->';
  const endTag = '<!-- API_REFERENCE_END -->';
  const startIndex = readme.indexOf(startTag);
  const endIndex = readme.indexOf(endTag);
  
  if (startIndex !== -1 && endIndex !== -1) {
    const before = readme.substring(0, startIndex + startTag.length);
    const after = readme.substring(endIndex);
    const apiMarkdown = generateReadmeAPI(apis);
    readme = before + '\n\n' + apiMarkdown + '\n\n' + after;
    await fs.writeFile(readmePath, readme);
    console.log('Successfully synchronized API reference to README.md!');
  }
}

function generateSearchIndex(apis) {
  const index = [
    { id: 'introduction', title: 'Introduction', type: 'concept', category: 'Getting Started', desc: 'Overview of the PPTXForge templating engine.' },
    { id: 'installation', title: 'Installation & Onboarding', type: 'concept', category: 'Getting Started', desc: 'How to install and verify node-pptx-templater.' },
    { id: 'quickstart', title: 'Quick Start Guide', type: 'concept', category: 'Getting Started', desc: 'A 60-second walkthrough with code samples.' },
    { id: 'learningpaths', title: 'Learning Paths', type: 'concept', category: 'Getting Started', desc: 'Onboarding guides tailored by experience levels.' },
    { id: 'code-sandbox', title: 'Interactive Showcase Sandbox', type: 'concept', category: 'Concepts', desc: 'Interactive code play space showing PowerPoint editing in real-time.' },
    { id: 'table-merging', title: 'Table Cell Merging', type: 'concept', category: 'Concepts', desc: 'Spans and cell merges details inside DrawingML grids.' },
    { id: 'chart-engine', title: 'Excel Chart Update', type: 'concept', category: 'Concepts', desc: 'Caching updates in spreadsheets backing PowerPoint charts.' },
    { id: 'zorder', title: 'Z-Order & Layers', type: 'concept', category: 'Concepts', desc: 'Stack sorting and layers indices order.' },
    { id: 'openxml-internals', title: 'Presentation Packaging Structure', type: 'concept', category: 'OpenXML', desc: 'Deep dive inside ZIP folders and slides XML.' },
    { id: 'security-arch', title: 'XML Security Architecture', type: 'concept', category: 'OpenXML', desc: 'Vulnerabilities mitigation strategies (XXE, XML bombs).' },
    { id: 'faq', title: 'FAQ & Troubleshooting', type: 'concept', category: 'Resources', desc: 'Solving presentation corruption warnings and fragmented placeholder runs.' },
    { id: 'roadmap', title: 'Project Roadmap', type: 'concept', category: 'Resources', desc: 'Upcoming features and development directions.' }
  ];

  apis.forEach(m => {
    const catName = API_CATEGORIES[m.category]?.title || 'API Reference';
    index.push({
      id: `method-${m.name}`,
      title: `PPTXTemplater.${m.name}()`,
      type: 'method',
      category: catName,
      desc: m.description,
      content: `${m.name} ${m.args} ${m.description} ${m.edgeCases} ${m.xmlImpact} ${JSON.stringify(m.params)} ${m.examples ? m.examples.basic + ' ' + m.examples.advanced : ''}`.toLowerCase()
    });
  });

  return index;
}

function getLastModifiedDate(filePath) {
  try {
    const dateStr = execSync(`git log -1 --format=%cs -- "${filePath}"`, {
      encoding: 'utf-8',
      stdio: ['ignore', 'pipe', 'ignore']
    }).trim();
    if (dateStr && /^\d{4}-\d{2}-\d{2}$/.test(dateStr)) {
      return dateStr;
    }
  } catch (e) {
    // Ignore and fallback
  }
  
  try {
    const stats = fs.statSync(filePath);
    return stats.mtime.toISOString().split('T')[0];
  } catch (e) {
    return new Date().toISOString().split('T')[0];
  }
}

function convertMarkdownToHtml(mdText) {
  return mdText
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/^# (.*$)/gim, '<h1>$1</h1>')
    .replace(/^## (.*$)/gim, '<h2>$1</h2>')
    .replace(/^### (.*$)/gim, '<h3>$1</h3>')
    .replace(/^\* (.*$)/gim, '<li>$1</li>')
    .replace(/^- (.*$)/gim, '<li>$1</li>')
    .replace(/`([^`]+)`/g, '<code>$1</code>')
    .replace(/\*\*([^*]+)\*\*/g, '<strong>$1</strong>')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, '<a href="$2">$1</a>')
    .replace(/\n\n/g, '<p></p>')
    .replace(/<li>(.*)<\/li>/g, '<ul><li>$1</li></ul>')
    .replace(/<\/ul>\s*<ul>/g, '');
}

function renderChangelogContent() {
  try {
    const changelogPath = resolve(__dirname, '../CHANGELOG.md');
    const content = fs.readFileSync(changelogPath, 'utf-8');
    return convertMarkdownToHtml(content);
  } catch (e) {
    return '<h3>Changelog</h3><p>Changelog data is currently unavailable.</p>';
  }
}

function getInstallationContent() {
  return `
    <h2>NPM Package Installation</h2>
    <p>Install the library directly into your project using npm:</p>
    <pre><code class="language-bash">npm install node-pptx-templater</code></pre>
    <h2>Prerequisites</h2>
    <p>Ensure your environment meets the minimum requirements:</p>
    <ul>
      <li>Node.js version 18.x or higher</li>
      <li>npm version 8.x or higher</li>
      <li>Compatible with TypeScript, CommonJS, and ES modules</li>
      <li>Zero native binaries or Microsoft Office dependencies</li>
    </ul>
  `;
}

function getQuickStartContent() {
  return `
    <h2>Basic Placeholder Replacement Example</h2>
    <p>Below is a quick demonstration of loading a PowerPoint presentation, updating placeholders, and saving the output:</p>
    <pre><code class="language-javascript">const { PPTXTemplater } = require('node-pptx-templater');

async function main() {
  const ppt = await PPTXTemplater.load('monthly_report_template.pptx');
  
  // Replace text placeholders on the first slide
  ppt.useSlide(1).replaceTextByTag('name', 'John Doe');
  
  // Save output presentation
  await ppt.saveToFile('output.pptx');
}

main().catch(console.error);</code></pre>
  `;
}

function getGettingStartedContent() {
  return `
    <h2>Visual Design vs Code Automation</h2>
    <p>Stop compiling slide elements inside complex code blocks. Design slide decks visually in PowerPoint, Keynote, or Google Slides, set formats and alignments, insert placeholders like <code>{{name}}</code>, and let <strong>node-pptx-templater</strong> populate them dynamically.</p>
  `;
}

function getFaqContent() {
  return `
    <h2>PowerPoint Repair Presentation alerts</h2>
    <p>When template placeholders are split in the XML structure (due to editor formatting), standard search-replace engines break slide layout files. node-pptx-templater resolves this by dynamically merging fragmented runs prior to processing, avoiding corrupt layouts.</p>
    <h2>XML Entity Expansion limit exceeded</h2>
    <p>This library parses XML using a secure, custom tokenization model designed to avoid entity expansions and XML bomb attacks (Billion Laughs) while maintaining high speed and low memory usage.</p>
  `;
}

function getRoadmapContent() {
  return `
    <h2>Upcoming Capabilities</h2>
    <ul>
      <li><strong>Q3 2026:</strong> Dynamic shape insertion and custom drawing layouts.</li>
      <li><strong>Q4 2026:</strong> Multi-threaded execution pipelines for serverless processing.</li>
      <li><strong>2027:</strong> Headless export of modified PPTX slide decks to PDF without LibreOffice.</li>
    </ul>
  `;
}

function getMethodBodyContent(m) {
  let paramsRows = '';
  if (m.params && m.params.length > 0) {
    paramsRows = `
      <h3>Parameters</h3>
      <table border="1" cellpadding="5" style="border-collapse: collapse; border-color: rgba(255,255,255,0.05); width: 100%;">
        <thead>
          <tr style="background-color: rgba(255,255,255,0.02);">
            <th>Parameter</th>
            <th>Type</th>
            <th>Description</th>
          </tr>
        </thead>
        <tbody>
          ${m.params.map(p => `
            <tr>
              <td><strong>${escapeHtml(p.name)}</strong></td>
              <td><code>${escapeHtml(p.type)}</code></td>
              <td>${escapeHtml(p.desc)}</td>
            </tr>
          `).join('')}
        </tbody>
      </table>
    `;
  }
  
  let returnsBlock = '';
  if (m.returns) {
    returnsBlock = `
      <h3>Returns</h3>
      <p><code>${escapeHtml(m.returns.type)}</code> — ${escapeHtml(m.returns.desc)}</p>
    `;
  }
  
  let examplesBlock = '';
  if (m.exampleCode) {
    examplesBlock = `
      <h3>Code Example</h3>
      <pre><code>${escapeHtml(m.exampleCode)}</code></pre>
    `;
  }

  return `
    <p><strong>Description:</strong> ${escapeHtml(m.description)}</p>
    <p><strong>Class:</strong> ${escapeHtml(m.className)} | <strong>File:</strong> <code>${escapeHtml(m.file)}</code></p>
    ${paramsRows}
    ${returnsBlock}
    <h3>OpenXML Impact</h3>
    <p>${escapeHtml(m.xmlImpact || 'Modifies underlying OpenXML nodes.')}</p>
    <h3>Edge Cases & Error Handling</h3>
    <p><strong>Edge Cases:</strong> ${escapeHtml(m.edgeCases || 'Verify inputs.')}</p>
    <p><strong>Errors:</strong> ${escapeHtml(m.errorHandling || 'Handle exceptions.')}</p>
    ${examplesBlock}
  `;
}

function getCategoryBodyContent(catId, catInfo, apis) {
  const catApis = apis.filter(m => m.category === catId);
  return `
    <p>${escapeHtml(catInfo.desc)}</p>
    <h2>Included API Methods</h2>
    <ul>
      ${catApis.map(m => `
        <li>
          <a href="../../api/${m.name}/"><strong>${escapeHtml(m.name)}</strong></a> — ${escapeHtml(m.description)}
        </li>
      `).join('')}
    </ul>
  `;
}

function buildStaticPageHtml(title, description, canonicalUrl, redirectHash, relativeBackDepth, bodyContent) {
  return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escapeHtml(title)}</title>
  <meta name="description" content="${escapeHtml(description)}">
  <link rel="canonical" href="${canonicalUrl}">
  <link rel="stylesheet" href="${relativeBackDepth}style.css">
  <script>
    // Immediate SEO redirect to main single-page application at correct anchor
    window.location.replace("${relativeBackDepth}index.html#${redirectHash}");
  </script>
</head>
<body class="bg-[#080b11] text-gray-200 font-sans p-8">
  <div class="max-w-3xl mx-auto space-y-6">
    <div class="border-b border-white/5 pb-4">
      <a href="${relativeBackDepth}index.html" class="text-indigo-400 hover:underline">&larr; Back to Documentation</a>
      <h1 class="text-4xl font-extrabold text-white mt-2">${escapeHtml(title)}</h1>
    </div>
    
    <div class="prose prose-invert max-w-none">
      ${bodyContent}
    </div>
    
    <div class="text-xs text-gray-500 border-t border-white/5 pt-4 mt-8">
      This page is an SEO-optimized landing page for search engines. JavaScript-enabled browsers will automatically transition to the main interface.
    </div>
  </div>
</body>
</html>`;
}

function generateSitemapXml(urls) {
  const elements = urls.map(u => `  <url>
    <loc>${u.url}</loc>
    <lastmod>${u.lastmod}</lastmod>
    <changefreq>${u.changefreq}</changefreq>
    <priority>${u.priority.toFixed(1)}</priority>
  </url>`).join('\n');

  return `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
${elements}
</urlset>`;
}

function generateSitemapIndexXml(sitemaps) {
  const today = new Date().toISOString().split('T')[0];
  const elements = sitemaps.map(s => `  <sitemap>
    <loc>${s}</loc>
    <lastmod>${today}</lastmod>
  </sitemap>`).join('\n');

  return `<?xml version="1.0" encoding="UTF-8"?>
<sitemapindex xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
${elements}
</sitemapindex>`;
}

async function validateSitemap(routes, examplesUrls, foundImages, baseVUrl) {
  console.log('--- Starting Sitemap Validation Audit ---');
  const errors = [];
  const warnings = [];
  
  const allUrls = [...routes, ...examplesUrls];
  
  // 1. Check for Duplicate URLs
  const urlSet = new Set();
  for (const r of allUrls) {
    if (urlSet.has(r.url)) {
      errors.push({ type: 'DUPLICATE_URL', url: r.url, message: `Duplicate sitemap location entry found: ${r.url}` });
    }
    urlSet.add(r.url);
  }
  
  // 2. Verify File Existence (No 404s in Sitemap)
  for (const r of routes) {
    const fileExists = await fs.pathExists(r.fsPath);
    if (!fileExists) {
      errors.push({ type: 'MISSING_FILE', url: r.url, fsPath: r.fsPath, message: `Sitemap URL references a non-existent file on disk: ${r.fsPath}` });
    }
  }
  
  // 3. Verify Canonical Tag Matches URL Exactly (Canonical URL Validation)
  for (const r of routes) {
    const fileExists = await fs.pathExists(r.fsPath);
    if (fileExists) {
      const htmlContent = await fs.readFile(r.fsPath, 'utf-8');
      
      const canonicalMatch = htmlContent.match(/<link\s+[^>]*rel=["']canonical["']\s+[^>]*href=["']([^"']+)["'][^>]*>/i) ||
                             htmlContent.match(/<link\s+[^>]*href=["']([^"']+)["']\s+[^>]*rel=["']canonical["'][^>]*>/i);
      
      if (!canonicalMatch) {
        errors.push({ type: 'MISSING_CANONICAL', url: r.url, fsPath: r.fsPath, message: `Canonical link tag is missing in file: ${r.fsPath}` });
      } else {
        const canonicalUrl = canonicalMatch[1];
        if (canonicalUrl !== r.url) {
          errors.push({
            type: 'CANONICAL_MISMATCH',
            url: r.url,
            fsPath: r.fsPath,
            message: `Canonical link tag mismatch. Found: ${canonicalUrl}, Expected: ${r.url}`
          });
        }
      }
      
      // 4. Crawl / Check for broken internal links inside generated pages
      const hrefRegex = /href=["']([^"']+)["']/gi;
      let hrefMatch;
      while ((hrefMatch = hrefRegex.exec(htmlContent)) !== null) {
        const link = hrefMatch[1];
        if (link.startsWith('http') || link.startsWith('#') || link.startsWith('mailto:')) {
          continue;
        }
        
        const resolvedLinkPath = resolve(path.dirname(r.fsPath), link);
        let checkPath = resolvedLinkPath;
        if (checkPath.endsWith('/') || (!path.extname(checkPath) && !checkPath.endsWith('index.html'))) {
          checkPath = checkPath.endsWith('/') ? checkPath + 'index.html' : checkPath + '/index.html';
        }
        
        const linkedFileExists = await fs.pathExists(checkPath);
        if (!linkedFileExists) {
          errors.push({
            type: 'BROKEN_LINK',
            url: r.url,
            fsPath: r.fsPath,
            link,
            resolvedPath: checkPath,
            message: `Broken internal link found: "${link}" (resolves to non-existent file: ${checkPath})`
          });
        }
      }
    }
  }

  // 5. Verify Priority Range
  for (const r of allUrls) {
    if (r.priority < 0.0 || r.priority > 1.0 || isNaN(r.priority)) {
      errors.push({ type: 'INVALID_PRIORITY', url: r.url, priority: r.priority, message: `Invalid priority range value: ${r.priority}` });
    }
  }

  console.log(`Validation Audit Summary: ${errors.length} errors, ${warnings.length} warnings.`);
  
  const report = {
    timestamp: new Date().toISOString(),
    baseUrl: baseVUrl,
    totalUrls: allUrls.length,
    totalImages: foundImages.size,
    status: errors.length === 0 ? 'PASSED' : 'FAILED',
    errors,
    warnings
  };
  
  await fs.writeFile(resolve(DOCS_DIR, 'sitemap-report.json'), JSON.stringify(report, null, 2));
  console.log('Sitemap audit report successfully written to docs/sitemap-report.json');
  
  if (errors.length > 0) {
    console.error('Sitemap Validation Errors Found:');
    errors.forEach(e => console.error(`  - [${e.type}] ${e.message}`));
  }
}

async function generateSeoSitemapsAndPages(apis, htmlContent) {
  const BASE_URL = process.env.DOCS_BASE_URL || 'https://jsuyog2.github.io/node-pptx-templater/';
  const normalizedBaseUrl = BASE_URL.endsWith('/') ? BASE_URL : BASE_URL + '/';
  
  const routes = [];
  
  async function registerRoute({ v, pathUrl, fileRelPath, title, desc, changefreq, priority, sourceFile, type, bodyHtml }) {
    const vPrefix = v ? v + '/' : '';
    const fullPathUrl = normalizedBaseUrl + vPrefix + pathUrl;
    const fsRelPath = vPrefix + fileRelPath;
    const absoluteFsPath = resolve(DOCS_DIR, fsRelPath);
    
    const lastmod = getLastModifiedDate(sourceFile);
    
    routes.push({
      url: fullPathUrl,
      fsPath: absoluteFsPath,
      relativePath: fsRelPath,
      title,
      description: desc,
      changefreq,
      priority,
      lastmod,
      type
    });
    
    if (fileRelPath !== 'index.html') {
      const parentDir = path.dirname(absoluteFsPath);
      await fs.ensureDir(parentDir);
      
      const relativeBackDepth = (v ? '../' : '') + '../'.repeat(fileRelPath.split('/').length - 1);
      const redirectHash = type === 'api' ? `method-${pathUrl.split('/').filter(Boolean).pop()}` :
                           type === 'category' ? `api-${pathUrl.split('/').filter(Boolean).pop()}` :
                           pathUrl.split('/').filter(Boolean).pop() || 'introduction';
      
      const redirectHtml = buildStaticPageHtml(title, desc, fullPathUrl, redirectHash, relativeBackDepth, bodyHtml);
      await fs.writeFile(absoluteFsPath, redirectHtml);
    }
  }

  let changelogMd = '';
  try {
    changelogMd = await fs.readFile(resolve(__dirname, '../CHANGELOG.md'), 'utf-8');
  } catch (e) {
    changelogMd = '# Changelog\n\nNo changelog file found.';
  }
  const changelogHtml = convertMarkdownToHtml(changelogMd);

  const corePages = [
    { id: 'installation', title: 'Installation & Onboarding - node-pptx-templater', desc: 'How to install and set up node-pptx-templater in Node.js applications.', content: getInstallationContent(), hash: 'installation' },
    { id: 'quickstart', title: 'Quick Start Guide - node-pptx-templater', desc: 'Quick start guide to replace placeholders, update tables, and process presentations.', content: getQuickStartContent(), hash: 'quickstart' },
    { id: 'getting-started', title: 'Getting Started - node-pptx-templater', desc: 'Core concepts and onboarding with node-pptx-templater.', content: getGettingStartedContent(), hash: 'introduction' },
    { id: 'faq', title: 'Frequently Asked Questions & Troubleshooting - node-pptx-templater', desc: 'Solutions for Repair Presentation alerts, unreplaced placeholders, and XML Entity expansion limit errors.', content: getFaqContent(), hash: 'faq' },
    { id: 'roadmap', title: 'Project Roadmap - node-pptx-templater', desc: 'Future development plans for shapes rendering, multi-threading, and PDF conversion.', content: getRoadmapContent(), hash: 'roadmap' },
    { id: 'changelog', title: 'Changelog & Releases - node-pptx-templater', desc: 'Release history and updates list for node-pptx-templater.', content: changelogHtml, hash: 'changelog' }
  ];

  const versions = ['', 'latest', `v${VERSION.split('.')[0]}`];
  
  for (const v of versions) {
    const lastmodHome = getLastModifiedDate(resolve(__dirname, '../README.md'));
    const vPrefix = v ? v + '/' : '';
    
    routes.push({
      url: normalizedBaseUrl + vPrefix,
      fsPath: resolve(DOCS_DIR, vPrefix + 'index.html'),
      relativePath: vPrefix + 'index.html',
      title: 'node-pptx-templater — Safe PowerPoint & OpenXML Template Engine',
      description: 'High-performance, zero-dependency Node.js library to update text, replace image shapes, cache Excel workbook data in charts, and merge table cells without PowerPoint Repair warnings.',
      changefreq: 'daily',
      priority: 1.0,
      lastmod: lastmodHome,
      type: 'core'
    });
    
    if (v) {
      await fs.ensureDir(resolve(DOCS_DIR, vPrefix));
      const adjustedHtml = htmlContent
        .replace(/href="style.css"/g, 'href="../style.css"')
        .replace(/src="app.js"/g, 'src="../app.js"')
        .replace(/href="sitemap.xml"/g, 'href="../sitemap.xml"')
        .replace(/https:\/\/jsuyog2.github.io\/node-pptx-templater\//g, normalizedBaseUrl + vPrefix);
      await fs.writeFile(resolve(DOCS_DIR, vPrefix + 'index.html'), adjustedHtml);
    }
    
    for (const page of corePages) {
      await registerRoute({
        v,
        pathUrl: page.id + '/',
        fileRelPath: page.id + '/index.html',
        title: page.title,
        desc: page.desc,
        changefreq: page.id === 'changelog' ? 'monthly' : 'weekly',
        priority: 1.0,
        sourceFile: page.id === 'changelog' ? resolve(__dirname, '../CHANGELOG.md') : resolve(__dirname, '../README.md'),
        type: 'core',
        bodyHtml: page.content
      });
    }
    
    for (const catId of Object.keys(API_CATEGORIES)) {
      const catInfo = API_CATEGORIES[catId];
      await registerRoute({
        v,
        pathUrl: `docs/${catId}/`,
        fileRelPath: `docs/${catId}/index.html`,
        title: `${catInfo.title} Reference - node-pptx-templater`,
        desc: catInfo.desc,
        changefreq: 'weekly',
        priority: 0.9,
        sourceFile: resolve(__dirname, '../src/core/PPTXTemplater.js'),
        type: 'category',
        bodyHtml: getCategoryBodyContent(catId, catInfo, apis)
      });
    }
    
    for (const m of apis) {
      await registerRoute({
        v,
        pathUrl: `api/${m.name}/`,
        fileRelPath: `api/${m.name}/index.html`,
        title: `${m.name}() API Method - node-pptx-templater`,
        desc: m.description,
        changefreq: 'weekly',
        priority: 0.9,
        sourceFile: resolve(__dirname, '../' + m.file),
        type: 'api',
        bodyHtml: getMethodBodyContent(m)
      });
    }
  }

  const docsUrls = routes.filter(r => r.type === 'core' || r.type === 'category');
  const sitemapDocsXml = generateSitemapXml(docsUrls);
  await fs.writeFile(resolve(DOCS_DIR, 'sitemap-docs.xml'), sitemapDocsXml);
  
  const apiUrls = routes.filter(r => r.type === 'api');
  const sitemapApiXml = generateSitemapXml(apiUrls);
  await fs.writeFile(resolve(DOCS_DIR, 'sitemap-api.xml'), sitemapApiXml);
  
  const examplesUrls = [];
  const lastmodReadme = getLastModifiedDate(resolve(__dirname, '../README.md'));
  for (const v of versions) {
    const vPrefix = v ? v + '/' : '';
    examplesUrls.push({
      url: normalizedBaseUrl + vPrefix + 'code-sandbox/',
      lastmod: lastmodReadme,
      changefreq: 'weekly',
      priority: 0.8
    });
    await fs.ensureDir(resolve(DOCS_DIR, vPrefix + 'code-sandbox'));
    const sandboxRedirectHtml = buildStaticPageHtml(
      'Interactive Showcase Sandbox - node-pptx-templater',
      'Experiment with PowerPoint template tags, chart updates, and stacking layers in real time.',
      normalizedBaseUrl + vPrefix + 'code-sandbox/',
      'code-sandbox',
      v ? '../../../' : '../../',
      `<h2>Interactive Showcase Sandbox</h2><p>Learn how to test PowerPoint text tagging, table merges, Excel chart synchronizations, and layer stack ordering directly in the interactive sandbox showcase.</p>`
    );
    await fs.writeFile(resolve(DOCS_DIR, vPrefix + 'code-sandbox/index.html'), sandboxRedirectHtml);
  }
  const sitemapExamplesXml = generateSitemapXml(examplesUrls);
  await fs.writeFile(resolve(DOCS_DIR, 'sitemap-examples.xml'), sitemapExamplesXml);

  const imageUrls = [];
  const imgRegex = /<img\s+[^>]*src=["']([^"']+)["'][^>]*>/gi;
  let imgMatch;
  const foundImages = new Set();
  while ((imgMatch = imgRegex.exec(htmlContent)) !== null) {
    foundImages.add(imgMatch[1]);
  }
  foundImages.add('assets/brand-preview.png');
  
  const imageElementsXml = [];
  for (const imgUrl of foundImages) {
    const absoluteImgUrl = imgUrl.startsWith('http') ? imgUrl : normalizedBaseUrl + imgUrl;
    const imgTitle = path.basename(imgUrl, path.extname(imgUrl))
      .replace(/[-_]/g, ' ')
      .replace(/\b\w/g, c => c.toUpperCase());
      
    imageElementsXml.push(`    <image:image>
      <image:loc>${absoluteImgUrl}</image:loc>
      <image:title>${escapeHtml(imgTitle)}</image:title>
    </image:image>`);
  }
  
  const sitemapImagesXml = `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9" xmlns:image="http://www.google.com/schemas/sitemap-image/1.1">
  <url>
    <loc>${normalizedBaseUrl}</loc>
${imageElementsXml.join('\n')}
  </url>
</urlset>`;
  await fs.writeFile(resolve(DOCS_DIR, 'sitemap-images.xml'), sitemapImagesXml);

  const totalUrlsCount = routes.length + examplesUrls.length;
  console.log(`Total generated URLs for sitemap indexing: ${totalUrlsCount}`);
  
  let sitemapXmlContent = '';
  if (totalUrlsCount > 500) {
    console.log(`URL count exceeds 500 (${totalUrlsCount}). Generating sitemap-index.xml as sitemap.xml.`);
    sitemapXmlContent = generateSitemapIndexXml([
      normalizedBaseUrl + 'sitemap-docs.xml',
      normalizedBaseUrl + 'sitemap-api.xml',
      normalizedBaseUrl + 'sitemap-examples.xml',
      normalizedBaseUrl + 'sitemap-images.xml'
    ]);
  } else {
    const allUrls = [...routes, ...examplesUrls];
    sitemapXmlContent = generateSitemapXml(allUrls);
    
    const indexContent = generateSitemapIndexXml([
      normalizedBaseUrl + 'sitemap-docs.xml',
      normalizedBaseUrl + 'sitemap-api.xml',
      normalizedBaseUrl + 'sitemap-examples.xml',
      normalizedBaseUrl + 'sitemap-images.xml'
    ]);
    await fs.writeFile(resolve(DOCS_DIR, 'sitemap-index.xml'), indexContent);
  }
  
  await fs.writeFile(resolve(DOCS_DIR, 'sitemap.xml'), sitemapXmlContent);
  
  const robotsTxt = `User-agent: *
Allow: /
Sitemap: ${normalizedBaseUrl}sitemap.xml`;
  await fs.writeFile(resolve(DOCS_DIR, 'robots.txt'), robotsTxt);

  await validateSitemap(routes, examplesUrls, foundImages, normalizedBaseUrl);
}

async function main() {
  const BASE_URL = process.env.DOCS_BASE_URL || 'https://jsuyog2.github.io/node-pptx-templater/';
  const normalizedBaseUrl = BASE_URL.endsWith('/') ? BASE_URL : BASE_URL + '/';

  const apis = extractAPIDocs();
  validateAPIBuild(apis);
  await syncReadme(apis);

  const sidebarHtml = renderSidebarAPIs(apis);
  const contentHtml = renderContentAPIs(apis);
  const searchIndexJson = JSON.stringify(generateSearchIndex(apis));

  const HTML_CONTENT = `<!DOCTYPE html>
<html lang="en" class="scroll-smooth">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  
  <!-- Primary SEO Meta Tags -->
  <title>node-pptx-templater — Safe PowerPoint & OpenXML Template Engine</title>
  <meta name="description" content="High-performance, zero-dependency Node.js library to update text, replace image shapes, cache Excel workbook data in charts, and merge table cells without PowerPoint Repair warnings.">
  <meta name="keywords" content="node pptx template, pptx templating library, powerpoint template engine, pptx editor nodejs, openxml powerpoint library, update powerpoint charts nodejs, powerpoint automation, pptx generator, pptx placeholder replacement, edit pptx without powerpoint, pptx chart update, pptx table update, openxml nodejs, powerpoint report generator, pptx cell merge, stack slide layers, duplicate slides, nodejs presentation automation">
  <link rel="canonical" href="${normalizedBaseUrl}">
  
  <!-- Open Graph / Facebook -->
  <meta property="og:type" content="website">
  <meta property="og:title" content="node-pptx-templater — Low-Level OpenXML PowerPoint Engine">
  <meta property="og:description" content="Automate presentations in Node.js with zero dependencies. Support cell merging, dynamic slides, Z-order layers, and Excel chart updating.">
  <meta property="og:url" content="${normalizedBaseUrl}">
  <meta property="og:image" content="${normalizedBaseUrl}assets/brand-preview.png">
  
  <!-- Twitter Card -->
  <meta name="twitter:card" content="summary_large_image">
  <meta name="twitter:title" content="node-pptx-templater — Low-Level OpenXML PowerPoint Engine">
  <meta name="twitter:description" content="Safe cell merging, slide cloning, Z-order layout adjustments, and Excel cache updates in pure JS.">
  
  <!-- Tailwind CSS CDN -->
  <script src="https://cdn.tailwindcss.com"></script>
  <script>
    tailwind.config = {
      darkMode: 'class',
      theme: {
        extend: {
          fontFamily: {
            sans: ['Inter', 'sans-serif'],
            title: ['Outfit', 'sans-serif'],
            mono: ['JetBrains Mono', 'monospace'],
          },
          colors: {
            brand: {
              50: '#eef2ff',
              100: '#e0e7ff',
              500: '#6366f1',
              600: '#4f46e5',
              700: '#4338ca',
            }
          }
        }
      }
    }
  </script>

  <!-- Google Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
  
  <!-- CSS Stylesheet -->
  <link rel="stylesheet" href="style.css">

  <!-- GSAP Animation Libraries -->
  <script src="https://cdnjs.cloudflare.com/ajax/libs/gsap/3.12.2/gsap.min.js"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/gsap/3.12.2/ScrollTrigger.min.js"></script>

  <!-- Schema.org Structured Data -->
  <script type="application/ld+json">
  {
    "@context": "https://schema.org",
    "@type": "SoftwareApplication",
    "name": "node-pptx-templater",
    "operatingSystem": "All",
    "applicationCategory": "DeveloperApplication",
    "offers": {
      "@type": "Offer",
      "price": "0",
      "priceCurrency": "USD"
    },
    "featureList": "PowerPoint text placeholder replacement, slide duplication, image insertion, Excel data workbook synchronization, DrawingML table cell merging, slide reordering, external slide importing, Z-order layer management, Bring Forward, Send Backward, Bring to Front, Send to Back",
    "downloadUrl": "https://www.npmjs.com/package/node-pptx-templater"
  }
  </script>
</head>
<body class="bg-[#080b11] text-gray-200 font-sans antialiased selection:bg-brand-500 selection:text-white min-h-screen relative overflow-x-hidden">

  <!-- Background Decorative Glowing Elements -->
  <div class="glow-orb absolute w-[500px] h-[500px] bg-brand-500/10 rounded-full blur-3xl -top-40 -left-40 pointer-events-none"></div>
  <div class="glow-orb absolute w-[500px] h-[500px] bg-purple-500/10 rounded-full blur-3xl bottom-20 -right-40 pointer-events-none"></div>

  <!-- Header -->
  <header class="fixed top-0 left-0 right-0 h-16 bg-[#0e1526]/85 backdrop-blur-md border-b border-white/5 flex items-center justify-between px-6 z-50 shadow-lg">
    <div class="flex items-center gap-3">
      <!-- SVG Brand Icon -->
      <svg class="w-8 h-8" viewBox="0 0 100 100" fill="none" xmlns="http://www.w3.org/2000/svg">
        <rect x="15" y="35" width="45" height="45" rx="8" fill="url(#header-grad-1)" stroke="rgba(255,255,255,0.15)" stroke-width="1.5" />
        <rect x="25" y="22" width="45" height="45" rx="8" fill="url(#header-grad-2)" stroke="rgba(255,255,255,0.15)" stroke-width="1.5" />
        <rect x="35" y="10" width="45" height="45" rx="8" fill="rgba(14, 21, 38, 0.6)" stroke="#6366f1" stroke-width="2" />
        <path d="M57 22 L45 35 L53 35 L43 48" stroke="#ffffff" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round" fill="none" />
        <defs>
          <linearGradient id="header-grad-1" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stop-color="#6366f1" stop-opacity="0.8" />
            <stop offset="100%" stop-color="#4f46e5" stop-opacity="0.3" />
          </linearGradient>
          <linearGradient id="header-grad-2" x1="0%" y1="0%" x2="100%" y2="100%">
            <stop offset="0%" stop-color="#d946ef" stop-opacity="0.8" />
            <stop offset="100%" stop-color="#ec4899" stop-opacity="0.3" />
          </linearGradient>
        </defs>
      </svg>
      <span class="font-title font-extrabold text-lg text-white tracking-wide">node-pptx-templater</span>
      <span class="bg-brand-500/10 text-brand-400 text-[10px] px-2.5 py-0.5 rounded-full font-semibold border border-brand-500/20">v${VERSION}</span>
    </div>
    
    <div class="flex items-center gap-5 relative">
      <div class="relative hidden md:block">
        <input type="text" id="doc-search" placeholder="Search pages, methods, parameters, examples (Ctrl + K)..." autocomplete="off" class="w-[380px] px-4 py-1.5 bg-[#080b11]/80 border border-white/5 rounded-full text-sm text-gray-200 placeholder-gray-500 focus:outline-none focus:border-brand-500 focus:ring-1 focus:ring-brand-500 transition-all">
        
        <!-- Search Results Panel -->
        <div id="search-results-dropdown" class="absolute left-0 right-0 mt-2 bg-[#0e1526] border border-white/10 rounded-2xl shadow-2xl max-h-[300px] overflow-y-auto hidden z-50 backdrop-blur-md">
          <!-- Populated by JS -->
        </div>
      </div>
      
      <button id="theme-toggle" aria-label="Toggle Theme" class="text-lg p-1.5 hover:text-brand-400 transition-colors bg-white/5 border border-white/5 rounded-full">
        <span class="toggle-icon">🌙</span>
      </button>
      
      <a href="https://github.com/jsuyog2/node-pptx-templater" class="text-gray-400 hover:text-white transition-colors bg-white/5 border border-white/5 p-1.5 rounded-full" target="_blank" rel="noopener">
        <svg class="w-5 h-5" viewBox="0 0 16 16" fill="currentColor">
          <path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"></path>
        </svg>
      </a>
    </div>
  </header>

  <!-- Sidebar & Content Wrapper -->
  <div class="flex pt-16 min-h-screen">
    <!-- Navigation Sidebar -->
    <aside class="w-72 bg-[#0e1526]/40 backdrop-blur-md border-r border-white/5 fixed top-16 bottom-0 left-0 overflow-y-auto px-5 py-6 hidden md:block z-40" id="sidebar">
      <nav class="space-y-6">
        <div class="space-y-1">
          <h3 class="text-xs uppercase text-gray-500 font-semibold tracking-wider px-3">Getting Started</h3>
          <ul class="space-y-1">
            <li><a href="#introduction" class="nav-item block px-3 py-1.5 text-sm text-gray-400 hover:text-white rounded-md transition-colors font-medium active">Introduction</a></li>
            <li><a href="#installation" class="nav-item block px-3 py-1.5 text-sm text-gray-400 hover:text-white rounded-md transition-colors font-medium">Installation & Onboarding</a></li>
            <li><a href="#quickstart" class="nav-item block px-3 py-1.5 text-sm text-gray-400 hover:text-white rounded-md transition-colors font-medium">Quick Start Guide</a></li>
            <li><a href="#learningpaths" class="nav-item block px-3 py-1.5 text-sm text-gray-400 hover:text-white rounded-md transition-colors font-medium">Learning Paths</a></li>
          </ul>
        </div>
        
        <div class="space-y-1">
          <h3 class="text-xs uppercase text-gray-500 font-semibold tracking-wider px-3">Core Concepts</h3>
          <ul class="space-y-1">
            <li><a href="#code-sandbox" class="nav-item block px-3 py-1.5 text-sm text-gray-400 hover:text-white rounded-md transition-colors font-medium">Interactive Showcase</a></li>
            <li><a href="#table-merging" class="nav-item block px-3 py-1.5 text-sm text-gray-400 hover:text-white rounded-md transition-colors font-medium">Table Cell Merging</a></li>
            <li><a href="#chart-engine" class="nav-item block px-3 py-1.5 text-sm text-gray-400 hover:text-white rounded-md transition-colors font-medium">Excel Chart Update</a></li>
            <li><a href="#zorder" class="nav-item block px-3 py-1.5 text-sm text-gray-400 hover:text-white rounded-md transition-colors font-medium">Z-Order &amp; Layers</a></li>
          </ul>
        </div>

        <!-- Categorized API Reference Sidebar -->
        <div class="space-y-1">
          <h3 class="text-xs uppercase text-gray-400 font-bold tracking-wider px-3 border-t border-white/5 pt-4">API Reference</h3>
          ${sidebarHtml}
        </div>

        <div class="space-y-1 pt-4 border-t border-white/5">
          <h3 class="text-xs uppercase text-gray-500 font-semibold tracking-wider px-3">OpenXML Internals</h3>
          <ul class="space-y-1">
            <li><a href="#openxml-internals" class="nav-item block px-3 py-1.5 text-sm text-gray-400 hover:text-white rounded-md transition-colors font-medium">Presentation Packing</a></li>
            <li><a href="#security-arch" class="nav-item block px-3 py-1.5 text-sm text-gray-400 hover:text-white rounded-md transition-colors font-medium">XML Security</a></li>
            <li><a href="#faq" class="nav-item block px-3 py-1.5 text-sm text-gray-400 hover:text-white rounded-md transition-colors font-medium">FAQ &amp; Troubleshooting</a></li>
            <li><a href="#roadmap" class="nav-item block px-3 py-1.5 text-sm text-gray-400 hover:text-white rounded-md transition-colors font-medium">Project Roadmap</a></li>
            <li><a href="#changelog" class="nav-item block px-3 py-1.5 text-sm text-gray-400 hover:text-white rounded-md transition-colors font-medium">Changelog &amp; Releases</a></li>
          </ul>
        </div>
      </nav>
    </aside>

    <!-- Main Content Area -->
    <main class="md:ml-72 flex-1 px-6 md:px-12 py-10 max-w-5xl overflow-y-auto relative z-10">
      
      <!-- Introduction Section -->
      <section id="introduction" class="doc-section space-y-6 active-section">
        <div class="space-y-4">
          <div class="inline-flex items-center gap-2 px-3 py-1 rounded-full bg-brand-500/10 text-brand-400 border border-brand-500/20 text-xs font-semibold">
            <span>✨</span> Built for Pure Node.js & Serverless Runtimes
          </div>
          <h1 class="font-title text-5xl font-extrabold text-white tracking-tight leading-none bg-gradient-to-r from-white via-gray-100 to-indigo-400 bg-clip-text text-transparent">
            node-pptx-templater
          </h1>
          <p class="text-lg text-gray-300 leading-relaxed max-w-2xl">
            A low-level, high-performance PowerPoint template engine built for Node.js. Populate slides dynamically using visually designed PowerPoint files, bypassing PowerPoint corruption warnings with unique OpenXML integrity features.
          </p>
        </div>

        <div class="grid grid-cols-1 md:grid-cols-3 gap-5 mt-8">
          <div class="glass-card p-5 bg-[#0e1526]/50 border border-white/5 rounded-2xl space-y-2">
            <div class="text-2xl">⚡</div>
            <h4 class="font-title font-bold text-white text-base">No Office Dependencies</h4>
            <p class="text-xs text-gray-400">Pure JavaScript execution. Runs flawlessly on AWS Lambda, Vercel Edge, or Google Cloud serverless platforms.</p>
          </div>
          <div class="glass-card p-5 bg-[#0e1526]/50 border border-white/5 rounded-2xl space-y-2">
            <div class="text-2xl">🧩</div>
            <h4 class="font-title font-bold text-white text-base">Fragment Resolution</h4>
            <p class="text-xs text-gray-400">PowerPoint editor splits tags (e.g. <code>{{c</code>, <code>ompany}}</code>). The engine merges them back automatically for flawless replacements.</p>
          </div>
          <div class="glass-card p-5 bg-[#0e1526]/50 border border-white/5 rounded-2xl space-y-2">
            <div class="text-2xl">📊</div>
            <h4 class="font-title font-bold text-white text-base">Excel Sync Caching</h4>
            <p class="text-xs text-gray-400">Synchronizes slide chart coordinates and visual datasets inside the underlying Excel sheet, avoiding PowerPoint refresh alerts.</p>
          </div>
        </div>

        <div class="p-5 bg-brand-500/10 border border-brand-500/20 rounded-2xl flex gap-4 items-start mt-6">
          <span class="text-brand-400 text-xl">💡</span>
          <div>
            <h4 class="font-title font-bold text-brand-400 text-sm">Visual Design vs Code Automation</h4>
            <p class="text-xs text-gray-300 mt-1 leading-relaxed">Stop compiling slide elements inside complex code blocks. Design slide decks visually in PowerPoint, Keynote, or Google Slides, set formats and alignments, insert placeholders like <code>{{name}}</code>, and let <strong>node-pptx-templater</strong> populate them dynamically.</p>
          </div>
        </div>
      </section>

      <!-- Installation Section -->
      <section id="installation" class="doc-section space-y-6 hidden">
        <h1 class="font-title text-4xl font-extrabold text-white">Installation & Onboarding</h1>
        <p class="text-sm text-gray-300">Set up the library in less than 30 seconds using npm or yarn. Zero local configurations required.</p>
        
        <div class="space-y-4">
          <h2 class="font-title text-xl font-bold text-white">NPM Install</h2>
          <pre class="relative group"><code class="language-bash block p-4 bg-[#05070c] border border-white/5 rounded-xl text-indigo-300 font-mono text-sm">npm install node-pptx-templater</code><button class="copy-btn absolute top-3.5 right-3.5 text-xs bg-white/5 hover:bg-white/10 text-gray-400 hover:text-white px-2.5 py-1 rounded-lg border border-white/10 transition-all">Copy</button></pre>
        </div>

        <div class="space-y-4 pt-4">
          <h2 class="font-title text-xl font-bold text-white">Prerequisites</h2>
          <ul class="list-disc pl-6 space-y-2 text-xs text-gray-400">
            <li><strong>Node.js Engine</strong>: Version <code>>= 18.0.0</code> (fully supports CommonJS standard require).</li>
            <li><strong>Package Platforms</strong>: Compiles natively on Windows, macOS, Linux, and Edge runtimes.</li>
          </ul>
        </div>
      </section>

      <!-- Quick Start Section -->
      <section id="quickstart" class="doc-section space-y-6 hidden">
        <h1 class="font-title text-4xl font-extrabold text-white">Quick Start Guide</h1>
        <p class="text-sm text-gray-300">Use this template rendering code snippet to load, populate, and export your first slide presentation.</p>
        
        <pre class="relative group"><code class="language-javascript block p-4 bg-[#05070c] border border-white/5 rounded-xl text-indigo-300 font-mono text-xs overflow-x-auto">const { PPTXTemplater } = require('node-pptx-templater');

async function main() {
  // 1. Load the presentation template
  const ppt = await PPTXTemplater.load('monthly_report_template.pptx');
  
  // 2. Select slide 1 and execute text replacement
  ppt.useSlide(1)
     .replaceTextByTag('title', 'Quarterly Earnings Report')
     .replaceMultiple({
       company: 'Acme Corporation',
       year: '2026'
     });

  // 3. Save the presentation to disk
  await ppt.saveToFile('./output/annual_earnings.pptx');
  console.log('Presentation generated successfully!');
}

main().catch(err => console.error(err));</code><button class="copy-btn absolute top-3.5 right-3.5 text-xs bg-white/5 hover:bg-white/10 text-gray-400 hover:text-white px-2.5 py-1 rounded-lg border border-white/10 transition-all">Copy</button></pre>
      </section>

      <!-- Learning Paths Section -->
      <section id="learningpaths" class="doc-section space-y-6 hidden">
        <h1 class="font-title text-4xl font-extrabold text-white">Learning Paths</h1>
        <p class="text-sm text-gray-300 font-light">Select a path tailored to your architectural expertise and PPTX templating requirements.</p>
        
        <!-- Tab Selectors -->
        <div class="flex border-b border-white/5 gap-4">
          <button onclick="switchPath('beginner')" id="tab-beginner" class="path-tab px-4 py-2 border-b-2 border-brand-500 font-semibold text-white transition-all">Beginner</button>
          <button onclick="switchPath('intermediate')" id="tab-intermediate" class="path-tab px-4 py-2 border-b-2 border-transparent text-gray-400 font-semibold hover:text-white transition-all">Intermediate</button>
          <button onclick="switchPath('advanced')" id="tab-advanced" class="path-tab px-4 py-2 border-b-2 border-transparent text-gray-400 font-semibold hover:text-white transition-all">Advanced</button>
        </div>

        <!-- Beginner Content -->
        <div id="path-content-beginner" class="path-content space-y-4">
          <h3 class="text-lg font-bold text-white font-title">Path 1: Standard Replacements & Text Merges</h3>
          <p class="text-xs text-gray-400 leading-relaxed">Ideal for replacing simple placeholder strings and inserting logos or target photos. Learn how placeholders look, how to format text inside PowerPoint, and execute basic save actions.</p>
          <ul class="list-disc pl-6 space-y-2 text-xs text-gray-400">
            <li>Designing simple <code>{{tag}}</code> placeholders in your editor.</li>
            <li>Replacing single tags and mapping values objects.</li>
            <li>Substituting template images while keeping shapes coordinates.</li>
          </ul>
        </div>

        <!-- Intermediate Content -->
        <div id="path-content-intermediate" class="path-content hidden space-y-4">
          <h3 class="text-lg font-bold text-white font-title">Path 2: Presentation Duplication & Element Collections</h3>
          <p class="text-xs text-gray-400 leading-relaxed">For developers building reports containing multiple tables, maps, series charts, and cloned shapes. Learn how to duplicate slides and manage relationships safely.</p>
          <ul class="list-disc pl-6 space-y-2 text-xs text-gray-400">
            <li>Cloning, deleting, and reordering slides dynamically.</li>
            <li>Updating chart databases categories and series points.</li>
            <li>Cloning slide table rows with unique rowId metadata.</li>
          </ul>
        </div>

        <!-- Advanced Content -->
        <div id="path-content-advanced" class="path-content hidden space-y-4">
          <h3 class="text-lg font-bold text-white font-title">Path 3: Stacking Layer Z-Order & Custom Slide Imports</h3>
          <p class="text-xs text-gray-400 leading-relaxed">Optimize execution speeds, implement complex layer stack sorting, import slides from distinct templates, and audit content overrides package integrity.</p>
          <ul class="list-disc pl-6 space-y-2 text-xs text-gray-400">
            <li>Stacking shapes layers using <code>bringForward</code> and <code>sendToBack</code>.</li>
            <li>Importing slides from distinct presentations with asset deduplication.</li>
            <li>Checking relationship lists with structural validation tools.</li>
          </ul>
        </div>
      </section>

      <!-- Interactive Code Sandbox Section -->
      <section id="code-sandbox" class="doc-section space-y-6 hidden">
        <h1 class="font-title text-4xl font-extrabold text-white">Interactive Showcase Sandbox</h1>
        <p class="text-sm text-gray-300">Click a feature tab on the left to see the code snippet and a visual representation of how PowerPoint is modified in real-time.</p>
        
        <div class="grid grid-cols-1 lg:grid-cols-12 gap-6 bg-[#0e1526]/40 border border-white/5 rounded-2xl overflow-hidden p-6">
          <!-- Sandbox Tabs -->
          <div class="lg:col-span-3 flex lg:flex-col gap-2 border-b lg:border-b-0 lg:border-r border-white/5 pb-4 lg:pb-0 lg:pr-4">
            <button onclick="switchSandbox('text')" id="sb-tab-text" class="sb-tab text-left px-4 py-2 rounded-xl text-xs font-semibold bg-brand-500/10 text-brand-400 border border-brand-500/20 w-full transition-all">📝 Text Tagging</button>
            <button onclick="switchSandbox('charts')" id="sb-tab-charts" class="sb-tab text-left px-4 py-2 rounded-xl text-xs font-semibold text-gray-400 hover:text-white w-full transition-all">📊 Excel Charts</button>
            <button onclick="switchSandbox('tables')" id="sb-tab-tables" class="sb-tab text-left px-4 py-2 rounded-xl text-xs font-semibold text-gray-400 hover:text-white w-full transition-all">📋 Table Merges</button>
            <button onclick="switchSandbox('layers')" id="sb-tab-layers" class="sb-tab text-left px-4 py-2 rounded-xl text-xs font-semibold text-gray-400 hover:text-white w-full transition-all">🥞 Stacking Layers</button>
          </div>
          
          <!-- Sandbox Code Panel -->
          <div class="lg:col-span-5 flex flex-col justify-between">
            <span class="text-[10px] uppercase text-gray-500 font-bold tracking-wider mb-2 block font-title">Javascript Code</span>
            <div class="bg-[#05070c] border border-white/5 rounded-xl p-4 font-mono text-[11px] text-indigo-300 overflow-x-auto flex-1 min-h-[180px]">
              <pre id="sandbox-code-block">ppt.useSlide(1)
   .replaceTextByTag('title', 'Q2 Report')
   .replaceMultiple({
     user: 'Acme Corp',
     date: 'June 2026'
   });</pre>
            </div>
          </div>
          
          <!-- Sandbox Preview Panel -->
          <div class="lg:col-span-4 flex flex-col">
            <span class="text-[10px] uppercase text-gray-500 font-bold tracking-wider mb-2 block font-title">Slide Preview Rendering</span>
            <div class="bg-[#05070c] border border-white/5 rounded-xl p-4 flex-1 flex flex-col justify-center items-center min-h-[180px] font-mono text-[10px] relative overflow-hidden" id="sandbox-preview-box">
              <div class="border border-white/10 rounded bg-[#0e1526] p-3 w-full text-center space-y-2">
                <div class="font-bold text-white text-[12px] border-b border-white/5 pb-1">Q2 Report</div>
                <div class="text-[9px] text-gray-400">Owner: Acme Corp</div>
                <div class="text-[9px] text-brand-400">Date: June 2026</div>
              </div>
            </div>
          </div>
        </div>
      </section>

      <!-- Table Cell Merging Section -->
      <section id="table-merging" class="doc-section space-y-6 hidden">
        <h1 class="font-title text-4xl font-extrabold text-white">Table Cell Merging</h1>
        <p class="text-sm text-gray-300 font-light">OpenXML PowerPoint tables require precise cell spans coordinate structures. The top-left cell acts as the **origin**, declaring \`gridSpan\` and \`rowSpan\`. The remaining shadowed cells must flag \`hMerge\` and \`vMerge\` to avoid breaking PowerPoint table layout models.</p>
        
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div class="p-5 bg-slate-900/30 border border-white/5 rounded-2xl space-y-2">
            <h4 class="font-title font-bold text-white text-base">API-based merging</h4>
            <p class="text-xs text-gray-400">Call <code>mergeCells()</code> directly by supplying coordinates:</p>
            <pre class="relative group"><code class="language-javascript block p-3 bg-[#05070c] border border-white/5 rounded-xl text-indigo-300 font-mono text-xs">ppt.mergeCells({
  tableId: 'sales-table',
  startRow: 1,
  startCol: 1,
  endRow: 2,
  endCol: 2
});</code><button class="copy-btn absolute top-3.5 right-3.5 text-xs bg-white/5 text-gray-400 hover:text-white px-2 py-0.5 rounded transition-all">Copy</button></pre>
          </div>
          <div class="p-5 bg-slate-900/30 border border-white/5 rounded-2xl space-y-2">
            <h4 class="font-title font-bold text-white text-base">Template-driven cell formatting</h4>
            <p class="text-xs text-gray-400">Update table values and declare spans inline within cell data:</p>
            <pre class="relative group"><code class="language-javascript block p-3 bg-[#05070c] border border-white/5 rounded-xl text-indigo-300 font-mono text-xs">ppt.updateTable('sales-table', [
  ['Header', 'Header', 'Header'],
  ['Data', { value: 'Span 2 Cols', colSpan: 2 }],
  ['Data', 'Data', { value: 'Span 2 Rows', rowSpan: 2 }]
]);</code><button class="copy-btn absolute top-3.5 right-3.5 text-xs bg-white/5 text-gray-400 hover:text-white px-2 py-0.5 rounded transition-all">Copy</button></pre>
          </div>
        </div>
      </section>

      <!-- Chart Engine Section -->
      <section id="chart-engine" class="doc-section space-y-6 hidden">
        <h1 class="font-title text-4xl font-extrabold text-white">Excel Chart Update</h1>
        <p class="text-sm text-gray-300 font-light">PowerPoint embeds an Excel worksheet (\`ppt/embeddings/\`) that controls chart datasets. Standard scripts only edit visual chart coordinates, corrupting calculations. <strong>node-pptx-templater</strong> compiles updates for both XML caches and spreadsheet rows.</p>
        
        <pre class="relative group"><code class="language-javascript block p-4 bg-[#05070c] border border-white/5 rounded-xl text-indigo-300 font-mono text-xs">ppt.updateChartData('sales-chart', {
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'Target', values: [100, 120, 140, 160] },
    { name: 'Revenue', values: [105, 118, 145, 172] }
  ]
});</code><button class="copy-btn absolute top-3.5 right-3.5 text-xs bg-white/5 hover:bg-white/10 text-gray-400 hover:text-white px-2.5 py-1 rounded-lg border border-white/10 transition-all">Copy</button></pre>
      </section>

      <!-- Z-Order Stacking Section -->
      <section id="zorder" class="doc-section space-y-6 hidden">
        <h1 class="font-title text-4xl font-extrabold text-white">Z-Order &amp; Stacking Layers</h1>
        <p class="text-sm text-gray-300 font-light">Programmatically stack shapes, images, charts, and tables using layer indices. The engine translates commands (Bring Forward, Send to Back) directly into OpenXML element orders within slide \`&lt;p:spTree&gt;\` blocks.</p>
        
        <pre class="relative group"><code class="language-javascript block p-4 bg-[#05070c] border border-white/5 rounded-xl text-indigo-300 font-mono text-xs">// See all slide elements layers in stacking order (bottom to top)
const elements = ppt.getObjectOrder(1);
console.log(elements);

// Bring the overlay logo shape to the front
ppt.bringToFront({ slide: 1, objectId: 'Logo' });

// Send the background template banner shape to the bottom
ppt.sendToBack({ slide: 1, objectId: 'Background' });</code><button class="copy-btn absolute top-3.5 right-3.5 text-xs bg-white/5 hover:bg-white/10 text-gray-400 hover:text-white px-2.5 py-1 rounded-lg border border-white/10 transition-all">Copy</button></pre>
      </section>

      <!-- Pre-rendered JSDoc API Content -->
      ${contentHtml}

      <!-- Packaging & XML Structure Section -->
      <section id="openxml-internals" class="doc-section space-y-6 hidden">
        <h1 class="font-title text-4xl font-extrabold text-white">Packaging & XML Structure</h1>
        <p class="text-sm text-gray-300">The PowerPoint document model is a zipped Open Packaging Convention (OPC) directory containing structured XML files. Below is the file mapping list for typical slide templates:</p>
        
        <pre class="block p-4 bg-[#05070c] border border-white/5 rounded-xl text-indigo-300 font-mono text-xs overflow-x-auto">
PPTX File Layout:
├── [Content_Types].xml (Document Override MIME types)
├── _rels/
│   └── .rels (Root presentation layouts relationship catalog)
├── ppt/
│   ├── presentation.xml (Slide listing, Masters catalog)
│   ├── slides/
│   │   ├── slide1.xml (Elements, shapes, texts runs)
│   │   └── _rels/
│   │       └── slide1.xml.rels (Slide-level resource assets map)
│   ├── media/ (Images assets database: PNG, JPEG, SVG)
│   └── embeddings/ (Excel workbooks backing PowerPoint charts)
        </pre>
      </section>

      <!-- XML Security Section -->
      <section id="security-arch" class="doc-section space-y-6 hidden">
        <h1 class="font-title text-4xl font-extrabold text-white">XML Security Architecture</h1>
        <p class="text-sm text-gray-300">To protect your application servers against malicious vectors inside user-supplied templates, the library implements robust, multi-layered XML parsing checks.</p>
        
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6">
          <div class="p-5 bg-slate-900/30 border border-white/5 rounded-2xl space-y-2">
            <h4 class="font-title font-bold text-white text-base">🛡️ Attack Protections</h4>
            <ul class="list-disc pl-6 space-y-1.5 text-xs text-gray-400">
              <li><strong>Billion Laughs & XML Bomb Prevention</strong>: Automatically rejects XML contain <code>&lt;!DOCTYPE&gt;</code> or <code>&lt;!ENTITY&gt;</code> tags.</li>
              <li><strong>XXE (XML External Entity) Protection</strong>: Rejects external system/public links to block local file disclosure vectors.</li>
              <li><strong>Oversized Entity Limits</strong>: Imposes hard limits (max 50,000 standard entity instances) to avoid parsing timeouts.</li>
            </ul>
          </div>
          <div class="p-5 bg-slate-900/30 border border-white/5 rounded-2xl space-y-2">
            <h4 class="font-title font-bold text-white text-base">⚙️ Custom Validation API</h4>
            <p class="text-xs text-gray-400">Exposes validation and recovery utilities directly to your code:</p>
            <pre class="relative group"><code class="language-javascript block p-3 bg-[#05070c] border border-white/5 rounded-xl text-indigo-300 font-mono text-[11px] leading-tight">const { validateXml, safeParseXml } = require('node-pptx-templater');

const status = validateXml(userXml);
if (!status.valid) {
  console.log('Error details:', status.error);
}</code><button class="copy-btn absolute top-3.5 right-3.5 text-[10px] bg-white/5 text-gray-400 hover:text-white px-2 py-0.5 rounded transition-all">Copy</button></pre>
          </div>
        </div>
      </section>

      <!-- FAQ & Troubleshooting Section -->
      <section id="faq" class="doc-section space-y-6 hidden">
        <h1 class="font-title text-4xl font-extrabold text-white">FAQ &amp; Troubleshooting</h1>
        
        <div class="space-y-4">
          <div class="p-5 bg-slate-900/30 border border-white/5 rounded-2xl space-y-2">
            <h4 class="font-title font-bold text-white text-base">Q: PowerPoint triggers a "Repair Presentation" alert. How do I fix it?</h4>
            <p class="text-xs text-gray-400 leading-relaxed">
              This happens if relationships are mismatched (pointing to non-existent assets), or if new slide/chart parts are not registered in the override list. 
              Always use the library's built-in <code>saveToFile()</code> or <code>toBuffer()</code> methods, which automatically execute structural check passes, remap IDs, and sanitize Content Overrides.
            </p>
          </div>
          
          <div class="p-5 bg-slate-900/30 border border-white/5 rounded-2xl space-y-2">
            <h4 class="font-title font-bold text-white text-base">Q: Some of my text placeholders inside shapes are not replacing. Why?</h4>
            <p class="text-xs text-gray-400 leading-relaxed">
              PowerPoint editors frequently segment tag characters into separate XML nodes behind the scenes (e.g. <code>{{title}}</code> splits into <code>&lt;a:t&gt;{{ti&lt;/a:t&gt;&lt;a:t&gt;tle}}&lt;/a:t&gt;</code>).
              To unify the nodes, highlight the placeholder in PowerPoint, cut it, and paste it back using "Keep Text Only" (this formats it into a single clean XML run).
            </p>
          </div>

          <div class="p-5 bg-slate-900/30 border border-white/5 rounded-2xl space-y-2">
            <h4 class="font-title font-bold text-white text-base">Q: How does the library resolve "Entity expansion limit exceeded" errors?</h4>
            <p class="text-xs text-gray-400 leading-relaxed">
              We disable internal XML entity expansion in the parser and decode standard character entities (\`&amp;\`, \`&lt;\`, etc.) and decimal/hex code points using an optimized, single-level JavaScript decoder. This bypasses limits while blocking XML entity expansion attacks entirely.
            </p>
          </div>
        </div>
      </section>

      <!-- Roadmap Section -->
      <section id="roadmap" class="doc-section space-y-6 hidden">
        <h1 class="font-title text-4xl font-extrabold text-white">Project Roadmap</h1>
        <p class="text-sm text-gray-300">Upcoming capabilities and core development plans for node-pptx-templater:</p>
        
        <div class="grid grid-cols-1 md:grid-cols-3 gap-5">
          <div class="glass-card p-5 rounded-2xl border border-white/5 space-y-2">
            <h4 class="font-title font-bold text-brand-400 text-sm">💡 Q3 2026: Shapes Rendering</h4>
            <p class="text-xs text-gray-400">Implement custom shape path creation APIs to construct dynamic rectangles, callouts, and connectors directly in code.</p>
          </div>
          <div class="glass-card p-5 rounded-2xl border border-white/5 space-y-2">
            <h4 class="font-title font-bold text-brand-400 text-sm">⚡ Q4 2026: Multi-Threading</h4>
            <p class="text-xs text-gray-400">Support Node worker_threads for slide parsing to execute massive enterprise templates generation in parallel pipelines.</p>
          </div>
          <div class="glass-card p-5 rounded-2xl border border-white/5 space-y-2">
            <h4 class="font-title font-bold text-brand-400 text-sm">📦 2027: PDF Conversion</h4>
            <p class="text-xs text-gray-400">Direct headless export of modified PPTX slide decks to PDF without requiring external LibreOffice or PowerPoint processes.</p>
          </div>
        </div>
      </section>

      <!-- Changelog Section -->
      <section id="changelog" class="doc-section space-y-6 hidden">
        <h1 class="font-title text-4xl font-extrabold text-white">Changelog &amp; Releases</h1>
        <p class="text-sm text-gray-300">Detailed release notes and updates for node-pptx-templater:</p>
        <div class="glass-card p-6 rounded-2xl border border-white/5 space-y-4 max-h-[600px] overflow-y-auto font-sans text-sm text-gray-300 leading-relaxed">
          ${renderChangelogContent()}
        </div>
      </section>

    </main>
  </div>

  <!-- JavaScript App Logic -->
  <script src="app.js"></script>
</body>
</html>`;

  const CSS_CONTENT = `/* Base CSS Styles & Themes */
:root {
  --bg-primary: #080b11;
  --bg-secondary: #0e1526;
  --border-color: rgba(255, 255, 255, 0.05);
  --text-main: #e2e8f0;
  --text-muted: #94a3b8;
  --primary: #6366f1;
}

.light-theme {
  --bg-primary: #f8fafc;
  --bg-secondary: #ffffff;
  --border-color: rgba(15, 23, 42, 0.08);
  --text-main: #1e293b;
  --text-muted: #64748b;
  --primary: #4f46e5;
}

body {
  background-color: var(--bg-primary);
  color: var(--text-main);
  transition: background-color 0.3s, color 0.3s;
}

/* Translucent Glass Card Effect */
.glass-card {
  background-color: rgba(14, 21, 38, 0.4);
  backdrop-filter: blur(12px);
  border: 1px solid var(--border-color);
  box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.2);
  position: relative;
}

.light-theme .glass-card {
  background-color: #ffffff !important;
  box-shadow: 0 4px 20px -2px rgba(15, 23, 42, 0.05), 0 2px 8px -1px rgba(15, 23, 42, 0.03) !important;
  border-color: rgba(15, 23, 42, 0.08) !important;
}

.glass-card::before {
  content: '';
  position: absolute;
  top: 0; left: 0; right: 0; bottom: 0;
  background: radial-gradient(400px circle at var(--mouse-x, 0) var(--mouse-y, 0), rgba(99, 102, 241, 0.08), transparent 80%);
  z-index: 0;
  pointer-events: none;
  transition: opacity 0.5s;
  opacity: 0;
  border-radius: inherit;
}

.glass-card:hover::before {
  opacity: 1;
}

header {
  border-bottom: 1px solid var(--border-color);
}

aside {
  border-right: 1px solid var(--border-color);
}

.nav-item.active {
  background-color: rgba(99, 102, 241, 0.1);
  color: #ffffff;
  border-left: 3px solid var(--primary);
  padding-left: 9px;
}

.light-theme .nav-item.active {
  background-color: rgba(79, 70, 229, 0.08) !important;
  color: #4f46e5 !important;
  border-left-color: #4f46e5 !important;
}

/* Light Theme Variable Contrast Overrides */
.light-theme .text-white {
  color: #0f172a !important;
}
.light-theme .text-gray-200 {
  color: #1e293b !important;
}
.light-theme .text-gray-300 {
  color: #334155 !important;
}
.light-theme .text-gray-400 {
  color: #475569 !important;
}
.light-theme .text-gray-500 {
  color: #64748b !important;
}
.light-theme .text-indigo-300 {
  color: #312e81 !important; /* High contrast dark indigo string color */
}
.light-theme .text-pink-400 {
  color: #be185d !important;
}
.light-theme .text-emerald-400 {
  color: #047857 !important;
}
.light-theme .text-brand-400 {
  color: #4f46e5 !important;
}
body.light-theme {
  background-color: #f8fafc !important;
  color: #1e293b !important;
}
.light-theme header,
.light-theme .bg-\\[\\#0e1526\\]\\/85 {
  background-color: rgba(255, 255, 255, 0.85) !important;
  border-color: rgba(15, 23, 42, 0.08) !important;
}
.light-theme aside,
.light-theme .bg-\\[\\#0e1526\\]\\/40 {
  background-color: rgba(241, 245, 249, 0.8) !important;
  border-color: rgba(15, 23, 42, 0.08) !important;
}
.light-theme #doc-search {
  background-color: rgba(15, 23, 42, 0.04) !important;
  color: #1e293b !important;
  border-color: rgba(15, 23, 42, 0.08) !important;
}
.light-theme #search-results-dropdown,
.light-theme .bg-\\[\\#0e1526\\] {
  background-color: #ffffff !important;
  border-color: rgba(15, 23, 42, 0.08) !important;
}
.light-theme .bg-\\[\\#0e1526\\]\\/50 {
  background-color: rgba(255, 255, 255, 0.8) !important;
}
.light-theme h1.font-title {
  background: linear-gradient(to right, #0f172a, #4f46e5) !important;
  -webkit-background-clip: text !important;
  background-clip: text !important;
  color: transparent !important;
}
.light-theme pre,
.light-theme code,
.light-theme .bg-\\[\\#05070c\\] {
  background-color: #f1f5f9 !important;
  border-color: rgba(15, 23, 42, 0.08) !important;
}
.light-theme .bg-slate-900,
.light-theme .bg-slate-900\\/10,
.light-theme .bg-slate-900\\/30 {
  background-color: #ffffff !important;
  border-color: rgba(15, 23, 42, 0.08) !important;
}
.light-theme .border-white\\/5 {
  border-color: rgba(15, 23, 42, 0.08) !important;
}
.light-theme .border-white\\/10 {
  border-color: rgba(15, 23, 42, 0.12) !important;
}
.light-theme .glow-orb {
  opacity: 0.02 !important;
}
.light-theme .bg-brand-500\\/10 {
  background-color: rgba(79, 70, 229, 0.08) !important;
}
.light-theme .border-brand-500\\/20 {
  border-color: rgba(79, 70, 229, 0.2) !important;
}
.light-theme .copy-btn {
  background-color: rgba(15, 23, 42, 0.04) !important;
  color: #475569 !important;
  border-color: rgba(15, 23, 42, 0.08) !important;
}
.light-theme .copy-btn:hover {
  background-color: rgba(15, 23, 42, 0.08) !important;
  color: #0f172a !important;
}

.light-theme .bg-white\\/5 {
  background-color: rgba(15, 23, 42, 0.04) !important;
}
.light-theme .bg-white\\/10 {
  background-color: rgba(15, 23, 42, 0.08) !important;
}
.light-theme .bg-white\\/\\[0\\.01\\] {
  background-color: rgba(15, 23, 42, 0.01) !important;
}
.light-theme .bg-white\\/\\[0\\.02\\] {
  background-color: rgba(15, 23, 42, 0.02) !important;
}
.light-theme .hover\\:bg-white\\/10:hover {
  background-color: rgba(15, 23, 42, 0.08) !important;
}
.light-theme .hover\\:bg-white\\/\\[0\\.02\\]:hover {
  background-color: rgba(15, 23, 42, 0.03) !important;
}
.light-theme .hover\\:text-white:hover {
  color: var(--primary) !important;
}
.light-theme .text-indigo-400,
.light-theme .text-indigo-400\\/90 {
  color: #4f46e5 !important;
}
.light-theme .search-result-item:hover {
  background-color: rgba(15, 23, 42, 0.04) !important;
}

/* Custom Scrollbars */
::-webkit-scrollbar {
  width: 6px;
  height: 6px;
}
::-webkit-scrollbar-track {
  background: transparent;
}
::-webkit-scrollbar-thumb {
  background: rgba(255, 255, 255, 0.1);
  border-radius: 3px;
}
.light-theme ::-webkit-scrollbar-thumb {
  background: rgba(15, 23, 42, 0.1) !important;
}
::-webkit-scrollbar-thumb:hover {
  background: rgba(255, 255, 255, 0.2);
}

.text-xxs {
  font-size: 8px;
}
.text-xs {
  font-size: 11px;
}
.text-sm {
  font-size: 13px;
}

/* Highlight Flash Animation for search jump targeting */
.highlight-flash {
  animation: borderGlowFlash 1.5s ease-out;
}
@keyframes borderGlowFlash {
  0% {
    border-color: #6366f1;
    box-shadow: 0 0 15px rgba(99, 102, 241, 0.4);
  }
  100% {
    border-color: var(--border-color);
    box-shadow: none;
  }
}
`;

  const JS_CONTENT = `// Search database declaration
const SEARCH_INDEX = ${searchIndexJson};

// Documentation Engine Client Logic
document.addEventListener('DOMContentLoaded', () => {
  const escapeHtml = (str) => {
    if (!str) return '';
    return str
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#039;');
  };

  const navItems = document.querySelectorAll('.nav-item');
  const sections = document.querySelectorAll('.doc-section');
  
  // Show/Hide page sections based on URL hashes
  function showSection(hash) {
    let cleanId = hash.replace('#', '');
    
    sections.forEach(sec => sec.classList.add('hidden'));
    navItems.forEach(item => item.classList.remove('active'));
    
    let targetSection = null;
    let targetCardId = null;
    
    if (cleanId.startsWith('method-')) {
      targetCardId = cleanId;
      const match = SEARCH_INDEX.find(item => item.id === cleanId);
      if (match) {
        const categoryKey = Object.keys(API_CATEGORIES_MAP).find(k => API_CATEGORIES_MAP[k] === match.category);
        if (categoryKey) {
          cleanId = 'api-' + categoryKey;
        }
      }
    }
    
    targetSection = document.getElementById(cleanId);
    
    if (targetSection) {
      targetSection.classList.remove('hidden');
      
      const activeLink = document.querySelector(\`a[href="#\${targetCardId || cleanId}"]\`);
      if (activeLink) {
        activeLink.classList.add('active');
        // Unfold parent sidebar if hidden
        const list = activeLink.closest('ul');
        if (list && list.classList.contains('hidden')) {
          list.classList.remove('hidden');
          const groupId = list.id.replace('sidebar-list-', '');
          const arrow = document.getElementById('arrow-' + groupId);
          if (arrow) arrow.style.transform = 'rotate(0deg)';
        }
      }
      
      if (targetCardId) {
        const card = document.getElementById(targetCardId);
        if (card) {
          setTimeout(() => {
            card.scrollIntoView({ behavior: 'smooth', block: 'center' });
            card.classList.add('highlight-flash');
            setTimeout(() => card.classList.remove('highlight-flash'), 1500);
          }, 100);
        }
      } else {
        window.scrollTo({ top: 0, behavior: 'smooth' });
      }
      
      // GSAP Reveal Animation
      gsap.fromTo(targetSection, 
        { opacity: 0.8, y: 10 },
        { opacity: 1, y: 0, duration: 0.3, ease: "power2.out" }
      );
    }
  }

  // Sidebar mapping definitions for routing resolver
  const API_CATEGORIES_MAP = {
    tables: 'Tables API',
    charts: 'Charts API',
    slides: 'Slides API',
    text: 'Text API',
    images: 'Images API',
    shapes: 'Shapes API',
    layers: 'Layer Stacking (Z-Order)',
    utils: 'Utilities & Validation'
  };

  navItems.forEach(item => {
    item.addEventListener('click', (e) => {
      e.preventDefault();
      const href = item.getAttribute('href');
      window.history.pushState(null, null, href);
      showSection(href);
    });
  });

  window.addEventListener('popstate', () => {
    const hash = window.location.hash || '#introduction';
    showSection(hash);
  });

  if (window.location.hash) {
    showSection(window.location.hash);
  } else {
    showSection('#introduction');
  }

  // Mousemove Glow cursor tracking logic
  document.addEventListener('mousemove', e => {
    document.querySelectorAll('.glass-card, .method-card').forEach(card => {
      const rect = card.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;
      card.style.setProperty('--mouse-x', \`\${x}px\`);
      card.style.setProperty('--mouse-y', \`\${y}px\`);
    });
  });

  // GSAP Entrance Load animations
  gsap.from("header", { opacity: 0, y: -20, duration: 0.6, ease: "power3.out" });
  gsap.from("aside", { opacity: 0, x: -30, duration: 0.6, ease: "power3.out", delay: 0.1 });

  // Light/Dark Theme Switching
  const themeToggle = document.getElementById('theme-toggle');
  const body = document.body;
  const savedTheme = localStorage.getItem('theme');

  if (savedTheme === 'light') {
    body.classList.remove('dark-theme');
    body.classList.add('light-theme');
    themeToggle.querySelector('.toggle-icon').textContent = '☀️';
  }

  themeToggle.addEventListener('click', () => {
    if (body.classList.contains('light-theme')) {
      body.classList.remove('light-theme');
      body.classList.add('dark-theme');
      themeToggle.querySelector('.toggle-icon').textContent = '🌙';
      localStorage.setItem('theme', 'dark');
    } else {
      body.classList.add('light-theme');
      body.classList.remove('dark-theme');
      themeToggle.querySelector('.toggle-icon').textContent = '☀️';
      localStorage.setItem('theme', 'light');
    }
  });

  // Client-Side Search Engine Search Input Resolver
  const searchInput = document.getElementById('doc-search');
  const searchResultsDropdown = document.getElementById('search-results-dropdown');

  searchInput.addEventListener('input', (e) => {
    const query = e.target.value.toLowerCase().trim();
    if (!query) {
      searchResultsDropdown.classList.add('hidden');
      searchResultsDropdown.innerHTML = '';
      return;
    }

    const matches = SEARCH_INDEX.filter(item => {
      const title = item.title.toLowerCase();
      const desc = item.desc ? item.desc.toLowerCase() : '';
      const content = item.content ? item.content.toLowerCase() : '';
      return title.includes(query) || desc.includes(query) || content.includes(query);
    });

    if (matches.length === 0) {
      searchResultsDropdown.innerHTML = \`<div class="p-4 text-xs text-gray-500 text-center">No matching documentation elements found.</div>\`;
      searchResultsDropdown.classList.remove('hidden');
      return;
    }

    // Sort: Title matches rank higher
    matches.sort((a, b) => {
      const aTitleMatch = a.title.toLowerCase().startsWith(query);
      const bTitleMatch = b.title.toLowerCase().startsWith(query);
      if (aTitleMatch && !bTitleMatch) return -1;
      if (!aTitleMatch && bTitleMatch) return 1;
      return 0;
    });

    searchResultsDropdown.innerHTML = matches.slice(0, 10).map(m => {
      return \`
        <div class="search-result-item p-3 border-b border-white/5 cursor-pointer transition-colors hover:bg-white/5 flex flex-col gap-1" onclick="selectSearchResult('\${m.id}')">
          <div class="flex items-center gap-2">
            <span class="px-2 py-0.5 rounded text-[9px] font-bold bg-brand-500/10 text-brand-400 uppercase tracking-wider">\${m.type}</span>
            <span class="text-xs text-white font-mono font-bold">\${escapeHtml(m.title)}</span>
          </div>
          <p class="text-[10px] text-gray-400 truncate">\${escapeHtml(m.desc || '')}</p>
        </div>
      \`;
    }).join('');
    searchResultsDropdown.classList.remove('hidden');
  });

  window.selectSearchResult = function(targetId) {
    searchResultsDropdown.classList.add('hidden');
    searchInput.value = '';
    window.location.hash = targetId;
    showSection('#' + targetId);
  };

  document.addEventListener('click', (e) => {
    if (!searchInput.contains(e.target) && !searchResultsDropdown.contains(e.target)) {
      searchResultsDropdown.classList.add('hidden');
    }
  });

  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
      e.preventDefault();
      searchInput.focus();
    }
  });

  // Clipboard text copy functions
  const copyButtons = document.querySelectorAll('.copy-btn');
  copyButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const codeBlock = btn.previousElementSibling;
      if (codeBlock) {
        navigator.clipboard.writeText(codeBlock.textContent).then(() => {
          btn.textContent = 'Copied!';
          setTimeout(() => {
            btn.textContent = 'Copy';
          }, 2000);
        });
      }
    });
  });
});

// Learning Path selectors toggle
window.switchPath = function(pathName) {
  document.querySelectorAll('.path-content').forEach(el => el.classList.add('hidden'));
  document.querySelectorAll('.path-tab').forEach(el => {
    el.classList.remove('border-brand-500', 'text-white');
    el.classList.add('border-transparent', 'text-gray-400');
  });

  const content = document.getElementById('path-content-' + pathName);
  if (content) content.classList.remove('hidden');

  const tab = document.getElementById('tab-' + pathName);
  if (tab) {
    tab.classList.add('border-brand-500', 'text-white');
    tab.classList.remove('border-transparent', 'text-gray-400');
  }
}

// Interactive Sandbox tab switcher
window.switchSandbox = function(tabName) {
  const codeBlock = document.getElementById('sandbox-code-block');
  const previewBox = document.getElementById('sandbox-preview-box');
  
  document.querySelectorAll('.sb-tab').forEach(el => {
    el.classList.remove('bg-brand-500/10', 'text-brand-400', 'border', 'border-brand-500/20');
    el.classList.add('text-gray-400', 'hover:text-white');
  });
  
  const selectedTab = document.getElementById('sb-tab-' + tabName);
  selectedTab.classList.add('bg-brand-500/10', 'text-brand-400', 'border', 'border-brand-500/20');
  selectedTab.classList.remove('text-gray-400', 'hover:text-white');
  
  if (tabName === 'text') {
    codeBlock.textContent = \`ppt.useSlide(1)\\n   .replaceTextByTag('title', 'Q2 Report')\\n   .replaceMultiple({\\n     user: 'Acme Corp',\\n     date: 'June 2026'\\n   });\`;
    previewBox.innerHTML = \`<div class="border border-white/10 rounded bg-[#0e1526] p-3 w-full text-center space-y-2">\\n      <div class="font-bold text-white text-[12px] border-b border-white/5 pb-1">Q2 Report</div>\\n      <div class="text-[9px] text-gray-400">Owner: Acme Corp</div>\\n      <div class="text-[9px] text-brand-400">Date: June 2026</div>\\n    </div>\`;
  } else if (tabName === 'charts') {
    codeBlock.textContent = \`ppt.useSlide(2)\\n   .updateChartData('sales-chart', {\\n     categories: ['Q1', 'Q2', 'Q3'],\\n     series: [\\n       { name: 'Target', values: [80, 100, 120] },\\n       { name: 'Actual', values: [95, 115, 130] }\\n     ]\\n   });\`;
    previewBox.innerHTML = \`<div class="w-full flex items-end justify-around h-32 px-4 border-b border-white/10">\\n      <div class="flex flex-col items-center">\\n        <div class="w-4 bg-brand-500/20 h-16 rounded-t"></div>\\n        <div class="w-4 bg-brand-500 h-20 rounded-t -mt-2"></div>\\n        <span class="text-[8px] text-gray-500 mt-1">Q1</span>\\n      </div>\\n      <div class="flex flex-col items-center">\\n        <div class="w-4 bg-brand-500/20 h-20 rounded-t"></div>\\n        <div class="w-4 bg-brand-500 h-24 rounded-t -mt-2"></div>\\n        <span class="text-[8px] text-gray-500 mt-1">Q2</span>\\n      </div>\\n      <div class="flex flex-col items-center">\\n        <div class="w-4 bg-brand-500/20 h-24 rounded-t"></div>\\n        <div class="w-4 bg-brand-500 h-28 rounded-t -mt-2"></div>\\n        <span class="text-[8px] text-gray-500 mt-1">Q3</span>\\n      </div>\\n    </div>\\n    <div class="text-[8px] text-gray-400 mt-2 flex gap-3"><span class="flex items-center gap-1"><span class="w-2 h-2 bg-brand-500/20 rounded"></span>Target</span><span class="flex items-center gap-1"><span class="w-2 h-2 bg-brand-500 rounded"></span>Actual</span></div>\`;
  } else if (tabName === 'tables') {
    codeBlock.textContent = \`ppt.useSlide(3)\\n   .updateTable('sales-table', [\\n     ['Category', { value: 'Global Performance', colSpan: 2 }],\\n     ['North Region', '120k', 'Growth: 8%'],\\n     ['South Region', '150k', 'Growth: 12%']\\n   ])\\n   .mergeCells('sales-table', 1, 1, 2, 2);\`;
    previewBox.innerHTML = \`<div class="w-full border border-white/10 rounded-xl overflow-hidden text-[9px] bg-[#0e1526]">\\n      <div class="bg-white/5 p-2 font-bold border-b border-white/10 text-center">Global Performance</div>\\n      <div class="grid grid-cols-3 divide-x divide-white/10 border-b border-white/10 text-center">\\n        <div class="p-2 text-gray-400">North Region</div>\\n        <div class="p-2 col-span-2 text-brand-400 font-bold">120k / Growth: 8%</div>\\n      </div>\\n      <div class="grid grid-cols-3 divide-x divide-white/10 text-center">\\n        <div class="p-2 text-gray-400">South Region</div>\\n        <div class="p-2 col-span-2 text-brand-400 font-bold">150k / Growth: 12%</div>\\n      </div>\\n    </div>\`;
  } else if (tabName === 'layers') {
    codeBlock.textContent = \`ppt.useSlide(4)\\n   .bringToFront('OverlayLogo')\\n   .sendToBack('BackgroundShade');\`;
    previewBox.innerHTML = \`<div class="relative w-full h-32 border border-white/10 rounded-xl bg-[#0e1526] overflow-hidden">\\n      <div class="absolute inset-0 bg-brand-500/5 flex items-center justify-center text-[10px] text-gray-500">BackgroundShade (zIndex: 1)</div>\\n      <div class="absolute top-6 left-6 w-32 h-16 border border-white/10 bg-slate-900 flex items-center justify-center rounded shadow-lg text-[9px]">Text Container (zIndex: 2)</div>\\n      <div class="absolute top-10 right-6 w-20 h-16 border border-brand-500/30 bg-brand-500/10 flex items-center justify-center rounded shadow-2xl text-[9px] text-brand-400 font-bold">OverlayLogo (zIndex: 3)</div>\\n    </div>\`;
  }

  // Sandbox reveal sweep GSAP
  gsap.fromTo([codeBlock, previewBox], 
    { opacity: 0.7, scale: 0.98 },
    { opacity: 1, scale: 1, duration: 0.3, ease: "power2.out" }
  );
}

// Sidebar sub-navigation toggle
window.toggleSidebarGroup = function(groupId) {
  const list = document.getElementById('sidebar-list-' + groupId);
  const arrow = document.getElementById('arrow-' + groupId);
  if (list && list.classList.contains('hidden')) {
    list.classList.remove('hidden');
    arrow.style.transform = 'rotate(0deg)';
  } else if (list) {
    list.classList.add('hidden');
    arrow.style.transform = 'rotate(-90deg)';
  }
}
`;

  await fs.ensureDir(DOCS_DIR);
  await fs.writeFile(resolve(DOCS_DIR, 'index.html'), HTML_CONTENT);
  await fs.writeFile(resolve(DOCS_DIR, 'style.css'), CSS_CONTENT);
  await fs.writeFile(resolve(DOCS_DIR, 'app.js'), JS_CONTENT);
  
  // SEO Sitemaps & Dynamic Clean URL Pages Generation
  await generateSeoSitemapsAndPages(apis, HTML_CONTENT);

  console.log('Successfully built GitHub Pages documentation with sitemaps, robots.txt, redirects, and verification under docs/');
}

main().catch(err => {
  console.error('Error building docs:', err);
  process.exit(1);
});
