/**
 * @fileoverview Script to export all decoupled documentation data into structured JSON files.
 * Generates:
 *   - docs-app/src/data/apiReference.json
 *   - docs-app/src/data/guides.json
 *   - docs-app/src/data/navigation.json
 *   - docs-app/src/data/showcase.json
 *   - docs-app/src/data/siteMeta.json
 */

const fs = require('fs');
const path = require('path');

const ROOT_DIR = path.resolve(__dirname, '..');
const DOCS_APP_DATA = path.join(ROOT_DIR, 'docs-app', 'src', 'data');
const DOCS_DATA = path.join(ROOT_DIR, 'docs', 'data');

// Ensure output directories exist
fs.mkdirSync(DOCS_APP_DATA, { recursive: true });
fs.mkdirSync(DOCS_DATA, { recursive: true });

// Load existing MODULES from scripts/build-docs.js
const buildDocsSource = fs.readFileSync(path.join(ROOT_DIR, 'scripts', 'build-docs.js'), 'utf8');
const match = buildDocsSource.match(/const MODULES = (\[[\s\S]*?\n\];)/);
if (!match) {
  throw new Error('Could not find MODULES in scripts/build-docs.js');
}

// Evaluate MODULES safely
const rawModules = eval(match[1]);

// 1. Process apiReference.json
// The required categories:
// Lifecycle & I/O, Slides, Text, Tables, Charts, Shapes, Cell Shapes, Images, Z-Order, Hyperlinks, Validation
const CATEGORY_MAP = {
  'lifecycle-io': 'Lifecycle & I/O',
  'slides-management': 'Slides',
  'text-placeholders': 'Text',
  'charts-series-sync': 'Charts',
  'shapes-geometry': 'Shapes',
  'images-media': 'Images',
  'z-order-layering': 'Z-Order',
  'hyperlinks-actions': 'Hyperlinks',
  'validation-diagnostics': 'Validation'
};

const CELL_SHAPE_METHODS = new Set([
  'getCellBounds',
  'addCellShape',
  'updateCellShape',
  'removeCellShape',
  'getCellShape',
  'getCellPosition'
]);

const apiMethods = [];

rawModules.forEach(mod => {
  mod.methods.forEach(method => {
    let category = CATEGORY_MAP[mod.id] || 'General';
    if (mod.id === 'tables-cell-shapes') {
      if (CELL_SHAPE_METHODS.has(method.name)) {
        category = 'Cell Shapes';
      } else {
        category = 'Tables';
      }
    }

    // Format parameters: array with name, type, required, default, description
    const formattedParams = (method.params || []).map(p => ({
      name: p.name,
      type: p.type,
      required: Boolean(p.required),
      default: p.defaultVal !== undefined ? p.defaultVal : '—',
      description: p.description
    }));

    apiMethods.push({
      name: method.name,
      signature: method.signature || `ppt.${method.name}()`,
      category: category,
      description: (method.description || '').replace(/<[^>]+>/g, ''), // Plain text / markdown clean
      rawDescription: method.description || '',
      isStatic: Boolean(method.isStatic),
      isAsync: Boolean(method.isAsync),
      openXml: method.openXml ? method.openXml.replace(/<[^>]+>/g, '') : '',
      parameters: formattedParams,
      returns: method.returns || { type: 'void', description: 'No return value' },
      exampleCode: (method.example || '').trim()
    });
  });
});

// Add extra high-frequency utility methods that are present in PPTXTemplater.js
const EXTRA_METHODS = [
  {
    name: 'toArrayBuffer',
    signature: 'ppt.toArrayBuffer()',
    category: 'Lifecycle & I/O',
    description: 'Serializes presentation into a raw in-memory ArrayBuffer, ideal for Cloudflare Workers, Vercel Edge, and Fetch API streaming.',
    isStatic: false,
    isAsync: true,
    openXml: 'Flushes VirtualFileSystem into an uncompressed or compressed PKZip ArrayBuffer payload.',
    parameters: [],
    returns: { type: 'Promise<ArrayBuffer>', description: 'Resolves to binary ArrayBuffer representing the PPTX archive.' },
    exampleCode: `const arrayBuffer = await ppt.toArrayBuffer();\nreturn new Response(arrayBuffer, { headers: { 'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation' } });`
  },
  {
    name: 'toExpress',
    signature: 'ppt.toExpress(res, filename = "presentation.pptx")',
    category: 'Lifecycle & I/O',
    description: 'Pipes the generated presentation stream directly into an Express.js HTTP response with appropriate attachment headers and MIME types.',
    isStatic: false,
    isAsync: true,
    openXml: 'Streams ZIP archive entries directly over the HTTP response socket.',
    parameters: [
      { name: 'res', type: 'Response', required: true, default: '—', description: 'Express HTTP response object.' },
      { name: 'filename', type: 'string', required: false, default: '"presentation.pptx"', description: 'Filename for Content-Disposition header.' }
    ],
    returns: { type: 'Promise<void>', description: 'Resolves when stream flushing is complete.' },
    exampleCode: `app.get('/download', async (req, res) => {\n  const ppt = await PPTXTemplater.load('template.pptx');\n  await ppt.toExpress(res, 'Report.pptx');\n});`
  },
  {
    name: 'deleteSlide',
    signature: 'ppt.deleteSlide(slideIndex)',
    category: 'Slides',
    description: 'Removes a slide from the presentation by 1-based index and purges its relationships and layout mappings from internal manifests.',
    isStatic: false,
    isAsync: false,
    openXml: 'Purges slide ID from ppt/presentation.xml and removes ppt/slides/slideN.xml and .rels.',
    parameters: [
      { name: 'slideIndex', type: 'number', required: true, default: '—', description: '1-based index of slide to remove.' }
    ],
    returns: { type: 'PPTXTemplater', description: 'Chainable instance.' },
    exampleCode: `ppt.deleteSlide(3); // Deletes the 3rd slide`
  },
  {
    name: 'getBounds',
    signature: 'ppt.getBounds(objectIdOrName)',
    category: 'Shapes',
    description: 'Calculates the pixel bounding box (x, y, width, height) of an element on the currently active slide.',
    isStatic: false,
    isAsync: false,
    openXml: 'Queries a:off (offset) and a:ext (extent) coordinates and converts EMUs to integer pixels.',
    parameters: [
      { name: 'objectIdOrName', type: 'string', required: true, default: '—', description: 'Element name or spId.' }
    ],
    returns: { type: '{ x: number, y: number, width: number, height: number }', description: 'Pixel bounding box coordinates.' },
    exampleCode: `ppt.useSlide(1);\nconst bounds = ppt.getBounds('HeaderTitle');\nconsole.log(bounds.width, bounds.height);`
  },
  {
    name: 'setBounds',
    signature: 'ppt.setBounds(objectIdOrName, options)',
    category: 'Shapes',
    description: 'Updates position and dimensions of an element using integer pixels or string units (e.g. "2in", "100px", "5cm").',
    isStatic: false,
    isAsync: false,
    openXml: 'Translates units into English Metric Units (EMU) and mutates a:off / a:ext XML nodes.',
    parameters: [
      { name: 'objectIdOrName', type: 'string', required: true, default: '—', description: 'Element identifier.' },
      { name: 'options', type: 'Object', required: true, default: '—', description: 'Coordinates { x, y, width, height }.' }
    ],
    returns: { type: 'PPTXTemplater', description: 'Chainable instance.' },
    exampleCode: `ppt.setBounds('HeaderBox', { width: '8in', height: '1.5in' });`
  },
  {
    name: 'stepForward',
    signature: 'ppt.stepForward(objectIdOrName)',
    category: 'Z-Order',
    description: 'Moves an element one step forward in the visual stacking order (Z-index) within the slide shape tree.',
    isStatic: false,
    isAsync: false,
    openXml: 'Permutes the child nodes inside p:spTree forward by one position.',
    parameters: [
      { name: 'objectIdOrName', type: 'string', required: true, default: '—', description: 'Target element identifier.' }
    ],
    returns: { type: 'PPTXTemplater', description: 'Chainable instance.' },
    exampleCode: `ppt.useSlide(1).stepForward('WatermarkShape');`
  },
  {
    name: 'stepBackward',
    signature: 'ppt.stepBackward(objectIdOrName)',
    category: 'Z-Order',
    description: 'Moves an element one step backward in the visual stacking order (Z-index) within the slide shape tree.',
    isStatic: false,
    isAsync: false,
    openXml: 'Permutes the child nodes inside p:spTree backward by one position.',
    parameters: [
      { name: 'objectIdOrName', type: 'string', required: true, default: '—', description: 'Target element identifier.' }
    ],
    returns: { type: 'PPTXTemplater', description: 'Chainable instance.' },
    exampleCode: `ppt.useSlide(1).stepBackward('OverlayBox');`
  },
  {
    name: 'removeCellShape',
    signature: 'ppt.removeCellShape(tableId, row, col)',
    category: 'Cell Shapes',
    description: 'Removes an overlaid cell shape or badge from a specific table cell coordinate.',
    isStatic: false,
    isAsync: false,
    openXml: 'Deletes associated geometric shape child in p:spTree mapped to table cell bounds.',
    parameters: [
      { name: 'tableId', type: 'string', required: true, default: '—', description: 'Table identifier.' },
      { name: 'row', type: 'number', required: true, default: '—', description: '0-based row index.' },
      { name: 'col', type: 'number', required: true, default: '—', description: '0-based column index.' }
    ],
    returns: { type: 'PPTXTemplater', description: 'Chainable instance.' },
    exampleCode: `ppt.removeCellShape('Financials', 1, 3);`
  }
];

EXTRA_METHODS.forEach(m => {
  if (!apiMethods.some(existing => existing.name === m.name)) {
    apiMethods.push(m);
  }
});

// Group methods by Category
const CATEGORY_LIST = [
  'Lifecycle & I/O',
  'Slides',
  'Text',
  'Tables',
  'Charts',
  'Shapes',
  'Cell Shapes',
  'Images',
  'Z-Order',
  'Hyperlinks',
  'Validation'
];

const apiReferenceData = {
  version: '3.0.0-alpha.0',
  totalMethods: apiMethods.length,
  categories: CATEGORY_LIST,
  methodsByCategory: {},
  methods: apiMethods
};

CATEGORY_LIST.forEach(cat => {
  apiReferenceData.methodsByCategory[cat] = apiMethods.filter(m => m.category === cat);
});

// 2. Process guides.json
const guidesData = [
  {
    id: 'getting-started',
    title: 'Getting Started & System Requirements',
    slug: 'getting-started',
    badge: 'Core',
    readingTime: '4 min read',
    summary: 'Install node-pptx-templater, configure Node.js runtimes, and understand the zero-native-dependency architecture.',
    sections: [
      {
        heading: 'Architecture Overview',
        text: 'node-pptx-templater is a 100% pure JavaScript OpenXML presentation engine built specifically for Node.js (>=20.12.0) and Serverless environments. Unlike legacy solutions that require spawning LibreOffice, Java runtimes (POI), or headless Chromium browsers, node-pptx-templater directly reads and mutates the internal OpenXML ZIP container in memory with zero C++ addons or binary dependencies.'
      },
      {
        heading: 'Installation',
        text: 'Install via your package manager of choice:',
        code: `# Using npm\nnpm install node-pptx-templater\n\n# Using pnpm\npnpm add node-pptx-templater\n\n# Using yarn\nyarn add node-pptx-templater`,
        language: 'bash'
      },
      {
        heading: 'Dual ESM and CommonJS Support',
        text: 'The library is packaged with first-class TypeScript declarations (.d.ts) and supports both CommonJS require() and modern ESM import syntax out of the box without bundler configuration.',
        code: `// CommonJS\nconst { PPTXTemplater } = require('node-pptx-templater');\n\n// ESM / TypeScript\nimport { PPTXTemplater } from 'node-pptx-templater';`,
        language: 'javascript'
      }
    ]
  },
  {
    id: 'quick-start',
    title: '5-Minute Quick Start',
    slug: 'quick-start',
    badge: 'Quickstart',
    readingTime: '5 min read',
    summary: 'Load a PPTX template, perform text tag replacements, update native charts with Excel caching, and save.',
    sections: [
      {
        heading: 'Basic Presentation Automation',
        text: 'Automate a complete quarterly report in under 15 lines of code. The engine manages layout relationships, run fragmentation, and XML namespaces automatically.',
        code: `import { PPTXTemplater } from 'node-pptx-templater';\n\nasync function run() {\n  // 1. Load template from file path or Buffer\n  const ppt = await PPTXTemplater.load('./templates/quarterly-review.pptx');\n\n  // 2. Target slide 1 and replace metadata\n  ppt.useSlide(1).replaceText({\n    '{{COMPANY}}': 'Acme Global Technologies',\n    '{{QUARTER}}': 'Q3 2026',\n    '{{PRESENTER}}': 'Operations Leadership'\n  });\n\n  // 3. Update chart series on slide 2 with Excel cache sync\n  ppt.useSlide(2);\n  await ppt.updateChart('RevenueChart', {\n    categories: ['Q1', 'Q2', 'Q3', 'Q4 (Proj)'],\n    series: [\n      { name: 'Cloud ARR', values: [18.4, 22.1, 26.8, 31.5] },\n      { name: 'AI Services', values: [8.2, 12.6, 17.4, 23.0] }\n    ]\n  });\n\n  // 4. Save to disk\n  await ppt.saveToFile('./output/Final_Q3_Review.pptx');\n  console.log('Presentation generated successfully without Office corruption!');\n}\n\nrun();`,
        language: 'javascript'
      }
    ]
  },
  {
    id: 'serverless',
    title: 'Serverless Runtimes (Lambda, Cloudflare, Vercel Edge)',
    slug: 'serverless',
    badge: 'Serverless',
    readingTime: '7 min read',
    summary: 'Deploying high-throughput PPTX generation on AWS Lambda, Cloudflare Workers, and Vercel Edge with zero filesystem dependencies.',
    sections: [
      {
        heading: 'Why Traditional Tools Fail on Serverless',
        text: 'Traditional PowerPoint automation tools rely on LibreOffice CLI (which requires a 500MB+ Docker image, takes 15+ seconds to boot, and consumes 1GB+ RAM) or Java-based Apache POI. node-pptx-templater has a pure JavaScript footprint under 2MB, boots in <15ms, and runs entirely in memory without writing temporary files to /tmp.'
      },
      {
        heading: 'AWS Lambda (Buffer & S3 Streaming)',
        text: 'Load templates from S3 and export directly as an in-memory Node.js Buffer:',
        code: `import { PPTXTemplater } from 'node-pptx-templater';\nimport { S3Client, GetObjectCommand, PutObjectCommand } from '@aws-sdk/client-s3';\n\nconst s3 = new S3Client({ region: 'us-east-1' });\n\nexport const handler = async (event) => {\n  const s3Data = await s3.send(new GetObjectCommand({\n    Bucket: 'my-templates-bucket',\n    Key: 'executive-summary.pptx'\n  }));\n\n  const templateBuffer = Buffer.from(await s3Data.Body.transformToByteArray());\n  const ppt = await PPTXTemplater.load(templateBuffer);\n\n  ppt.useSlide('all').replaceText(event.payload.replacements);\n\n  // Export buffer directly into memory\n  const outputBuffer = await ppt.toBuffer();\n\n  await s3.send(new PutObjectCommand({\n    Bucket: 'my-output-bucket',\n    Key: \`reports/\${event.payload.reportId}.pptx\`,\n    Body: outputBuffer,\n    ContentType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation'\n  }));\n\n  return { statusCode: 200, body: JSON.stringify({ success: true }) };\n};`,
        language: 'javascript'
      },
      {
        heading: 'Cloudflare Workers (Pure JS V8 Isolates)',
        text: 'Cloudflare Workers do not provide Node.js fs APIs. node-pptx-templater provides toArrayBuffer() to stream PPTX files directly into standard HTTP Response objects:',
        code: `import { PPTXTemplater } from 'node-pptx-templater';\n\nexport default {\n  async fetch(request, env) {\n    const templateUrl = 'https://assets.example.com/template.pptx';\n    const res = await fetch(templateUrl);\n    const arrayBuffer = await res.arrayBuffer();\n\n    const ppt = await PPTXTemplater.load(arrayBuffer);\n    ppt.useSlide(1).replaceText({ '{{USER_NAME}}': 'Edge Subscriber' });\n\n    const outputArrayBuffer = await ppt.toArrayBuffer();\n\n    return new Response(outputArrayBuffer, {\n      headers: {\n        'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',\n        'Content-Disposition': 'attachment; filename="GeneratedDeck.pptx"'\n      }\n    });\n  }\n};`,
        language: 'javascript'
      },
      {
        heading: 'Vercel Edge & Next.js API Routes',
        text: 'Serve real-time PowerPoint downloads directly from Next.js App Router route handlers:',
        code: `// app/api/generate-pptx/route.js\nimport { PPTXTemplater } from 'node-pptx-templater';\nimport fs from 'fs/promises';\nimport path from 'path';\n\nexport async function POST(req) {\n  const data = await req.json();\n  const templatePath = path.join(process.cwd(), 'templates', 'invoice.pptx');\n  const buffer = await fs.readFile(templatePath);\n\n  const ppt = await PPTXTemplater.load(buffer);\n  ppt.useSlide(1).replaceText(data.replacements);\n  const outputBuffer = await ppt.toBuffer();\n\n  return new Response(outputBuffer, {\n    headers: {\n      'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',\n      'Content-Disposition': 'attachment; filename="invoice.pptx"'\n    }\n  });\n}`,
        language: 'javascript'
      }
    ]
  },
  {
    id: 'run-fragment-healing',
    title: 'Run Fragment Healing & Anti-Corruption Engine',
    slug: 'run-fragment-healing',
    badge: 'OpenXML Engine',
    readingTime: '6 min read',
    summary: 'Overcoming PowerPoint split <a:r> text run fragmentation at the AST level without regex bugs.',
    sections: [
      {
        heading: 'The Fragmented Run Problem Explained',
        text: 'When a human types "{{COMPANY_NAME}}" into PowerPoint, spellcheckers, auto-correct, or grammar tools frequently break that single placeholder across multiple disjoint XML nodes. In the raw PresentationML, it becomes:\n\n<a:p>\n  <a:r><a:t>{{COMP</a:t></a:r>\n  <a:r><a:rPr b="1"/><a:t>ANY_</a:t></a:r>\n  <a:r><a:t>NAME}}</a:t></a:r>\n</a:p>\n\nNaive search-and-replace regex engines fail completely because "{{COMPANY_NAME}}" does not exist in any single text node.'
      },
      {
        heading: 'How node-pptx-templater Heals Runs',
        text: 'node-pptx-templater implements an AST-level paragraph normalizer. It analyzes all <a:r> child elements within each <a:p> paragraph, constructs a virtual coordinate map of the text, detects placeholder boundaries spanning multiple runs, merges the text while preserving typography styles (bold, font size, color, language), and prunes the emptied sibling nodes cleanly. Result: 100% replacement accuracy with zero layout deformation.'
      },
      {
        heading: 'Example in Action',
        text: 'You never have to manually clean or re-type templates in PowerPoint:',
        code: `const ppt = await PPTXTemplater.load('human-edited-template.pptx');\n\n// Even if {{DISCOUNT_RATE}} was typed with typos and edits,\n// the Run Normalizer seamlessly resolves and substitutes it:\nppt.useSlide('all').replaceText({\n  '{{DISCOUNT_RATE}}': '24.5%',\n  '{{EXPIRATION_DATE}}': 'December 31, 2026'\n});`,
        language: 'javascript'
      }
    ]
  },
  {
    id: 'excel-dual-sync',
    title: 'Excel Dual-Sync: Preventing Repair Presentation Warnings',
    slug: 'excel-dual-sync',
    badge: 'Charts',
    readingTime: '6 min read',
    summary: 'Synchronizing DrawingML <c:chartSpace> and embedded .xlsx workbook data caches.',
    sections: [
      {
        heading: 'Why Other Libraries Cause "Repair Presentation" Prompts',
        text: 'In Microsoft PowerPoint, native charts are dual-layered:\n1. The chart visual representation is defined in XML: ppt/charts/chart1.xml (<c:chartSpace>).\n2. The underlying data source is stored as an embedded binary Excel workbook: ppt/embeddings/Microsoft_Excel_Worksheet1.xlsx.\n\nWhen standard open-source libraries only update the XML <c:strCache> or <c:numCache>, Microsoft Office detects a desynchronization between the visual cache and the binary Excel spreadsheet when a user clicks "Edit Data in Excel". PowerPoint then throws the infamous error:\n"PowerPoint found a problem with content in presentation.pptx. PowerPoint can attempt to repair the presentation."'
      },
      {
        heading: 'The Dual-Sync Solution',
        text: 'node-pptx-templater synchronizes BOTH layers simultaneously:\n- It mutates the DrawingML <c:ser>, <c:cat>, and <c:val> series cache trees.\n- It parses and updates the embedded .xlsx sheet XML (sheet1.xml) and shared strings table within the internal ZIP package.\nWhen the recipient opens the deck in PowerPoint or clicks "Edit Data in Excel", the numbers, series, and formulas match 100% with zero repair prompts.'
      },
      {
        heading: 'Usage Example',
        text: 'Single API call to update series and sync workbooks:',
        code: `await ppt.useSlide(2).updateChart('FinancialPerformanceChart', {\n  categories: ['FY23', 'FY24', 'FY25', 'FY26'],\n  series: [\n    { name: 'Gross Revenue', values: [42.1, 56.4, 78.9, 104.2] },\n    { name: 'Operating Income', values: [12.4, 18.2, 27.5, 39.8] }\n  ]\n});`,
        language: 'javascript'
      }
    ]
  },
  {
    id: 'table-merging',
    title: 'Dynamic Tables, Row Merging & Auto-Pagination',
    slug: 'table-merging',
    badge: 'Tables',
    readingTime: '8 min read',
    summary: 'Handling complex rowspan/gridspan cell merges, row striping, and auto-pagination across slide overflows.',
    sections: [
      {
        heading: 'Row Merging & Nested Data',
        text: 'Standard OpenXML tables can be easily corrupted if vertical merge (vMerge) or horizontal merge (hMerge) attributes are mismatched. node-pptx-templater automatically manages grid coordinates, span calculations, and cell hiding rules.'
      },
      {
        heading: 'Auto-Pagination with splitTableOnOverflow',
        text: 'When dynamically generating tables with 50+ rows from database records, tables frequently spill off the bottom of the slide. splitTableOnOverflow calculates row heights, detects boundary collisions, deep-clones the parent slide, repeats header rows, and paginates rows cleanly across subsequent slides.',
        code: `ppt.useSlide(3);\n\n// Add 60 rows from SQL query\nfor (const record of queryResults) {\n  ppt.addTableRow('InventoryTable', [\n    record.sku,\n    record.name,\n    record.warehouse,\n    record.quantity\n  ]);\n}\n\n// Automatically split across new slides if it exceeds slide height\nawait ppt.splitTableOnOverflow('InventoryTable', {\n  bottomMargin: '1in',\n  repeatHeaders: 1\n});`,
        language: 'javascript'
      }
    ]
  },
  {
    id: 'z-order-stacking',
    title: 'Z-Order Stacking & Spatial Geometry Engine',
    slug: 'z-order-stacking',
    badge: 'Geometry',
    readingTime: '5 min read',
    summary: 'Precise layer stacking, bounding box transformations, and English Metric Unit (EMU) alignment.',
    sections: [
      {
        heading: 'Understanding OpenXML Coordinates (EMU)',
        text: 'In PowerPoint, coordinates are stored in English Metric Units (1 inch = 914,400 EMUs; 1 cm = 360,000 EMUs; 1 pt = 12,700 EMUs). node-pptx-templater includes a high-precision spatial unit engine to prevent floating-point drift.'
      },
      {
        heading: 'Z-Order Layer Management',
        text: 'Control visual depth without recreating elements or destroying attached animations and relationships:',
        code: `ppt.useSlide(1);\n\n// Bring badge or logo to top\nppt.bringToFront('CompanyLogo');\n\n// Push background graphic to back\nppt.sendToBack('GradientBackground');\n\n// Step-wise nudging\nppt.stepForward('Watermark');\nppt.stepBackward('OverlayCard');`,
        language: 'javascript'
      }
    ]
  }
];

// 3. Process navigation.json
const navigationData = {
  header: [
    { title: 'Home', href: '/' },
    { title: 'Guides', href: '/guides' },
    { title: 'API Reference', href: '/api-reference' },
    { title: 'Showcase', href: '/showcase' },
    { title: 'Feature Inventory', href: '/features' }
  ],
  sidebar: [
    {
      group: 'Getting Started',
      items: [
        { title: 'Introduction & Setup', href: '/guides/getting-started' },
        { title: '5-Minute Quick Start', href: '/guides/quick-start' },
        { title: 'Serverless Runtimes', href: '/guides/serverless' }
      ]
    },
    {
      group: 'Core Architecture',
      items: [
        { title: 'Run Fragment Healing', href: '/guides/run-fragment-healing' },
        { title: 'Excel Dual-Sync', href: '/guides/excel-dual-sync' },
        { title: 'Tables & Overflow Pagination', href: '/guides/table-merging' },
        { title: 'Z-Order & Spatial Geometry', href: '/guides/z-order-stacking' }
      ]
    },
    {
      group: 'API Reference by Category',
      items: CATEGORY_LIST.map(cat => ({
        title: cat,
        href: `/api-reference?category=${encodeURIComponent(cat)}`,
        category: cat,
        count: apiReferenceData.methodsByCategory[cat].length
      }))
    },
    {
      group: 'Examples & Showcase',
      items: [
        { title: 'Unified Showcase Demos', href: '/showcase' },
        { title: 'Verified Feature Inventory', href: '/features' }
      ]
    }
  ],
  social: {
    github: 'https://github.com/jsuyog2/node-pptx-templater',
    npm: 'https://www.npmjs.com/package/node-pptx-templater',
    license: 'https://opensource.org/licenses/MIT',
    issues: 'https://github.com/jsuyog2/node-pptx-templater/issues'
  }
};

// 4. Process showcase.json
const showcaseData = [
  {
    id: 'demo-01',
    number: '01',
    title: 'Basic Report & 16:9 Presentation',
    category: 'Layout & Typography',
    description: 'Load a blank or template deck, set widescreen 16:9 dimensions, and substitute executive summary metadata.',
    code: `const ppt = await PPTXTemplater.load('./showcase/templates/sample.pptx');\nppt.useSlide(1).setSlideSize('16:9');\nppt.replaceMultiple({\n  '{{title}}': 'Executive Business Review',\n  '{{subtitle}}': 'Q3 2026 Strategic Performance Update',\n  '{{author}}': 'Operations Team'\n});\nawait ppt.save('./showcase/output/01_basic_report.pptx');`
  },
  {
    id: 'demo-02',
    number: '02',
    title: 'Chart & Excel Dual-Sync',
    category: 'Charts',
    description: 'Update clustered column and line charts while keeping embedded Excel workbook cache in complete synchronization.',
    code: `ppt.useSlide(1);\nawait ppt.updateChart('chart1', {\n  categories: ['Q1 2026', 'Q2 2026', 'Q3 2026', 'Q4 2026 (Proj)'],\n  series: [\n    { name: 'Enterprise Cloud', values: [18.4, 22.1, 26.8, 31.5] },\n    { name: 'AI Solutions', values: [8.2, 12.6, 17.4, 23.0] },\n    { name: 'Professional Services', values: [5.1, 5.4, 5.8, 6.2] }\n  ]\n});`
  },
  {
    id: 'demo-03',
    number: '03',
    title: 'Tables, Striping & Overflow Pagination',
    category: 'Tables',
    description: 'Populate 50+ row tabular datasets, alternate row colors, and automatically split overflow rows across new slides.',
    code: `ppt.useSlide(2);\nfor (const item of workstreams) {\n  ppt.addTableRow('projectStatusTable', item);\n}\nppt.stripeRows('projectStatusTable', '#F8FAFC', '#FFFFFF');\nawait ppt.splitTableOnOverflow('projectStatusTable', { bottomMargin: '1in', repeatHeaders: 1 });`
  },
  {
    id: 'demo-04',
    number: '04',
    title: 'Nested Tables & Merged Cells (Rowspan)',
    category: 'Tables',
    description: 'Render multi-level hierarchical records with vertical cell merging (rowspan) without XML table corruption.',
    code: `ppt.useSlide(2);\n// Insert department row with vertical rowspan over 3 child teams\nppt.addTableRow('teamTable', ['Engineering', 'Platform Architecture', '24', 'Alex Mercer'], { mergeStrategy: 'rowspan' });\nppt.addTableRow('teamTable', ['', 'Data Infrastructure', '18', 'Priya Sharma']);\nppt.addTableRow('teamTable', ['', 'Frontend Experience', '16', 'Liam O\\’Connor']);`
  },
  {
    id: 'demo-05',
    number: '05',
    title: 'Shapes, Z-Order & Spatial Alignment',
    category: 'Shapes',
    description: 'Inject modern UI callouts, cards, and banners, coordinate English Metric Units, and bring active elements to front.',
    code: `await ppt.addShape('roundRect', {\n  x: '5%', y: '15%', width: '90%', height: '22%',\n  fill: '1E293B', text: 'Critical Milestone Passed'\n});\nppt.bringToFront('CompanyBadge');\nppt.alignShapes(['Card1', 'Card2', 'Card3'], 'distributeHorizontally');`
  },
  {
    id: 'demo-06',
    number: '06',
    title: 'Cell Shapes & Dynamic Badges',
    category: 'Cell Shapes',
    description: 'Position status pill badges (Green circle, Amber triangle, Red flag) directly inside table cells.',
    code: `ppt.addTableRow('kpiTable', [\n  'Service Availability',\n  '99.98%',\n  { type: 'circle', fill: '10B981', label: 'Operational' }\n]);`
  },
  {
    id: 'demo-07',
    number: '07',
    title: 'Images: PNG, WebP & Base64 Data URIs',
    category: 'Images',
    description: 'Embed high-resolution images, corporate logos, and base64 avatars with automatic MIME and Content-Type registration.',
    code: `await ppt.replaceImage('SpeakerAvatar', './assets/avatar.png');\nawait ppt.addImage({\n  source: 'data:image/png;base64,iVBORw0KGgo...', \n  x: '75%', y: '5%', width: '20%', height: '10%'\n});`
  },
  {
    id: 'demo-08',
    number: '08',
    title: 'Slide Cloning & Relationship Remapping',
    category: 'Slides',
    description: 'Clone template master slides repeatedly, reorder slides dynamically, and import slides across presentations.',
    code: `await ppt.duplicateSlide(2, 3); // Duplicate slide 2 into position 3\nppt.reorderSlides([1, 3, 2]);\nawait ppt.importSlideFrom('./library.pptx', 1);`
  },
  {
    id: 'demo-09',
    number: '09',
    title: 'Hyperlinks & Slide Navigation Actions',
    category: 'Hyperlinks',
    description: 'Attach external URL web links and internal jump-to-slide navigation actions to shapes and text runs.',
    code: `ppt.addHyperlink('DocumentationButton', 'https://node-pptx-templater.pages.dev');\nppt.addSlideLink('TableOfContentsItem', 4); // Jump to slide 4`
  },
  {
    id: 'demo-10',
    number: '10',
    title: 'Typography & Text Runs',
    category: 'Text',
    description: 'Format complex paragraphs with mixed fonts, bold weights, hex colors, and inline hyperlinks inside a single paragraph.',
    code: `ppt.useSlide(1);\nppt.appendTextRun('SummaryPara', 'Approved by leadership: ',\n  { bold: false, color: '64748B' }\n);\nppt.appendTextRun('SummaryPara', 'COMPLETED',\n  { bold: true, color: '10B981' }\n);`
  },
  {
    id: 'demo-11',
    number: '11',
    title: 'OpenXML Folder Workflow',
    category: 'Lifecycle & I/O',
    description: 'Unpack presentations into readable XML folder structures for Git auditing, diffing, and rebuild back to PPTX.',
    code: `await PPTXTemplater.extractPptx('./template.pptx', './unpacked-git-repo');\n// Inspect or edit presentation.xml in Git\nawait PPTXTemplater.buildPptx('./unpacked-git-repo', './rebuilt.pptx');`
  },
  {
    id: 'demo-12',
    number: '12',
    title: 'Serverless Buffer Export',
    category: 'Lifecycle & I/O',
    description: 'Produce in-memory binary archives for high-speed delivery in AWS Lambda, Cloudflare Workers, and Docker microservices.',
    code: `const buffer = await ppt.toBuffer();\nconsole.log(\`Generated \${buffer.length} bytes in memory without touching disk.\`);`
  }
];

// 5. Process siteMeta.json
const siteMetaData = {
  title: 'node-pptx-templater — High-Performance PowerPoint OpenXML Template Engine',
  tagline: 'Populate PowerPoint templates dynamically in Node.js & Serverless runtimes without Microsoft Office dependencies.',
  description: 'Production-ready, pure JavaScript PowerPoint (PPTX) OpenXML template engine for Node.js and Serverless runtimes. Dynamically replace text, update charts with Excel cache sync, format tables, and merge slides without PowerPoint corruption.',
  author: 'node-pptx-templater contributors',
  version: '3.0.0-alpha.0',
  license: 'MIT',
  baseUrl: 'https://node-pptx-templater.pages.dev',
  repositoryUrl: 'https://github.com/jsuyog2/node-pptx-templater',
  npmUrl: 'https://www.npmjs.com/package/node-pptx-templater',
  issuesUrl: 'https://github.com/jsuyog2/node-pptx-templater/issues',
  minNodeVersion: '>=20.12.0',
  social: {
    twitterHandle: '@jsuyog2',
    ogBannerUrl: 'https://node-pptx-templater.pages.dev/og-banner.png',
    twitterCardUrl: 'https://node-pptx-templater.pages.dev/twitter-card.png'
  },
  keywordsCatalog: {
    core: [
      'node-pptx-templater',
      'node.js pptx template engine',
      'openxml powerpoint generator',
      'pure js pptx',
      'powerpoint automation nodejs'
    ],
    serverless: [
      'serverless powerpoint lambda',
      'cloudflare workers pptx',
      'vercel edge powerpoint',
      'aws lambda pptx generation',
      'in-memory pptx buffer'
    ],
    features: [
      'powerpoint chart excel sync',
      'run fragment healing',
      'table merging pptx',
      'table cell merge pptx',
      'split table on overflow',
      'z-order shapes pptx',
      'slide cloning openxml'
    ],
    compliance: [
      'non-corrupting pptx',
      'no office repair prompt',
      'openxml standard compliant',
      'drawingml presentationml'
    ]
  }
};

// Write files to both docs-app/src/data and docs/data
const targets = [DOCS_APP_DATA, DOCS_DATA];

targets.forEach(dir => {
  fs.writeFileSync(path.join(dir, 'apiReference.json'), JSON.stringify(apiReferenceData, null, 2), 'utf8');
  fs.writeFileSync(path.join(dir, 'guides.json'), JSON.stringify(guidesData, null, 2), 'utf8');
  fs.writeFileSync(path.join(dir, 'navigation.json'), JSON.stringify(navigationData, null, 2), 'utf8');
  fs.writeFileSync(path.join(dir, 'showcase.json'), JSON.stringify(showcaseData, null, 2), 'utf8');
  fs.writeFileSync(path.join(dir, 'siteMeta.json'), JSON.stringify(siteMetaData, null, 2), 'utf8');
});

console.log('✓ Successfully decoupled and generated all documentation JSON data schemas:');
console.log(`  - apiReference.json (${apiReferenceData.totalMethods} methods across ${CATEGORY_LIST.length} categories)`);
console.log(`  - guides.json (${guidesData.length} comprehensive guides)`);
console.log(`  - navigation.json (sidebar, header, and social links)`);
console.log(`  - showcase.json (${showcaseData.length} production demonstrations)`);
console.log(`  - siteMeta.json (global metadata and SEO keyword catalog)`);
