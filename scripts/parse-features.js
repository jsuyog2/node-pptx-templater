const fs = require('fs');
const path = require('path');

const ROOT_DIR = path.resolve(__dirname, '..');
const FEATURES_MD_PATH = path.join(ROOT_DIR, 'project', 'Features.md');
const DATA_DIR = path.join(ROOT_DIR, 'docs-site', 'src', 'data');

if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

// 1. Parse project/Features.md
const markdown = fs.readFileSync(FEATURES_MD_PATH, 'utf8');
const lines = markdown.split('\n');

const categories = [];
let currentCategory = null;

for (let i = 0; i < lines.length; i++) {
  const line = lines[i].trim();
  
  // Category heading: ## 1. Input & Ingestion Flexibility
  const catMatch = line.match(/^##\s+(\d+)\.\s+(.+)$/);
  if (catMatch) {
    if (currentCategory) {
      categories.push(currentCategory);
    }
    const catNum = parseInt(catMatch[1], 10);
    const catTitle = catMatch[2].replace(/\(Flagship\)/g, '').trim();
    const isFlagship = catMatch[2].includes('(Flagship)');
    
    currentCategory = {
      id: catTitle.toLowerCase().replace(/[^a-z0-9]+/g, '-'),
      number: catNum,
      title: catTitle,
      isFlagship: isFlagship,
      items: []
    };
    continue;
  }

  // Feature item: 1. **Direct PPTX File Ingestion**: Load standard `.pptx` documents...
  const itemMatch = line.match(/^(\d+)\.\s+\*\*(.+?)\*\*:\s*(.+)$/);
  if (itemMatch && currentCategory) {
    const itemNum = parseInt(itemMatch[1], 10);
    const itemTitle = itemMatch[2].trim();
    const itemDesc = itemMatch[3].trim();
    currentCategory.items.push({
      id: itemNum,
      title: itemTitle,
      description: itemDesc
    });
  }
}

if (currentCategory) {
  categories.push(currentCategory);
}

const featuresData = {
  version: '2.0 / 3.0',
  totalFeatures: categories.reduce((sum, c) => sum + c.items.length, 0),
  totalCategories: categories.length,
  categories: categories.map(c => ({
    ...c,
    count: c.items.length
  }))
};

fs.writeFileSync(
  path.join(DATA_DIR, 'features.json'),
  JSON.stringify(featuresData, null, 2),
  'utf8'
);

console.log(`✓ Generated features.json with ${featuresData.totalFeatures} features across ${featuresData.totalCategories} categories.`);

// 2. Generate docsNavigation.json
const docsNavigation = {
  header: [
    { title: "Home", href: "/" },
    { title: "Docs", href: "/docs" },
    { title: "Features", href: "/features" },
    { title: "API Reference", href: "/docs/api" }
  ],
  sidebar: [
    {
      title: "Getting Started",
      items: [
        { title: "Overview", href: "/docs", icon: "BookOpen" },
        { title: "Installation Guides", href: "/docs/installation", icon: "Download" },
        { title: "Compatibility Matrix", href: "/docs/compatibility", icon: "Cpu" }
      ]
    },
    {
      title: "API Reference",
      href: "/docs/api",
      items: [
        { title: "Lifecycle & I/O", href: "/docs/api#Lifecycle%20%26%20I%2FO", category: "Lifecycle & I/O", count: 16 },
        { title: "Slides", href: "/docs/api#Slides", category: "Slides", count: 12 },
        { title: "Text & Placeholders", href: "/docs/api#Text%20%26%20Placeholders", category: "Text & Placeholders", count: 10 },
        { title: "Tables & Cell Shapes", href: "/docs/api#Tables%20%26%20Cell%20Shapes", category: "Tables & Cell Shapes", count: 11 },
        { title: "Charts & Series Sync", href: "/docs/api#Charts%20%26%20Series%20Sync", category: "Charts & Series Sync", count: 8 },
        { title: "Shapes & Geometry", href: "/docs/api#Shapes%20%26%20Geometry", category: "Shapes & Geometry", count: 11 },
        { title: "Images & Media", href: "/docs/api#Images%20%26%20Media", category: "Images & Media", count: 3 },
        { title: "Z-Order & Layering", href: "/docs/api#Z-Order%20%26%20Layering", category: "Z-Order & Layering", count: 11 },
        { title: "Hyperlinks", href: "/docs/api#Hyperlinks", category: "Hyperlinks", count: 7 },
        { title: "Diagnostics & Validation", href: "/docs/api#Diagnostics%20%26%20Validation", category: "Diagnostics & Validation", count: 12 }
      ]
    }
  ],
  external: {
    github: "https://github.com/jsuyog2/node-pptx-templater",
    npm: "https://www.npmjs.com/package/node-pptx-templater",
    issues: "https://github.com/jsuyog2/node-pptx-templater/issues"
  }
};

fs.writeFileSync(
  path.join(DATA_DIR, 'docsNavigation.json'),
  JSON.stringify(docsNavigation, null, 2),
  'utf8'
);
console.log('✓ Generated docsNavigation.json');

// 3. Generate compatibility.json
const compatibilityData = {
  title: "Runtime Compatibility & Serverless Guarantees",
  philosophy: "Pure JavaScript OpenXML parsing with zero native binary dependencies (no Microsoft Office, no LibreOffice, no Java VM, no Electron, no node-gyp).",
  runtimes: [
    {
      name: "Node.js (>= 18, 20.x, 22.x, 24.x)",
      tier: "Tier 1 - Primary Support",
      status: "Fully Supported",
      notes: "Full support for CJS (require) and ESM (import) with native TypeScript definitions (.d.ts). High-performance in-memory Virtual File System.",
      bootTime: "< 15ms",
      bundleSize: "< 2MB"
    },
    {
      name: "AWS Lambda (Node.js 18 / 20 / 22)",
      tier: "Tier 1 - Serverless",
      status: "Fully Supported",
      notes: "Instant cold-start execution without disk writes. Export directly via ppt.toBuffer() or stream to S3 buckets.",
      bootTime: "< 25ms",
      bundleSize: "< 2MB"
    },
    {
      name: "Cloudflare Workers (V8 Isolates)",
      tier: "Tier 1 - Edge Computing",
      status: "Fully Supported",
      notes: "Zero Node.js fs dependencies. Ingest from web streams or ArrayBuffer, export via ppt.toArrayBuffer() directly to HTTP responses.",
      bootTime: "< 5ms",
      bundleSize: "< 2MB"
    },
    {
      name: "Vercel Edge & Serverless Functions",
      tier: "Tier 1 - Edge / Serverless",
      status: "Fully Supported",
      notes: "Compatible with Next.js App Router route handlers (app/api/generate/route.js). Zero temporary file disk pollution.",
      bootTime: "< 20ms",
      bundleSize: "< 2MB"
    },
    {
      name: "Bun (>= 1.0)",
      tier: "Tier 1 - Alternate Runtime",
      status: "Fully Supported",
      notes: "High-throughput execution leveraging Bun fast ZIP and Buffer APIs.",
      bootTime: "< 10ms",
      bundleSize: "< 2MB"
    },
    {
      name: "Deno (>= 1.38)",
      tier: "Tier 2 - Compatible",
      status: "Supported via npm: specifier",
      notes: "Usable via import { PPTXTemplater } from 'npm:node-pptx-templater'.",
      bootTime: "< 25ms",
      bundleSize: "< 2MB"
    }
  ],
  guarantees: [
    {
      title: "Zero Office / Electron Daemon",
      description: "Unlike unoconv, LibreOffice, or Puppeteer, node-pptx-templater does not launch background daemons or require 500MB+ Docker containers."
    },
    {
      title: "In-Memory Virtual File System",
      description: "Operates 100% in RAM via Map<string, Uint8Array>. Eliminates /tmp write contention, disk IOPS costs, and concurrency race conditions."
    },
    {
      title: "XXE & Security Guardrails",
      description: "Built-in protection against XML External Entity (XXE) and Billion Laughs exponential entity expansion attacks via strict parser configuration."
    },
    {
      title: "Office Repair Warning Prevention",
      description: "Synchronous DrawingML and embedded Excel workbook cache updating guarantees slides open cleanly without Microsoft PowerPoint repair prompts."
    }
  ]
};

fs.writeFileSync(
  path.join(DATA_DIR, 'compatibility.json'),
  JSON.stringify(compatibilityData, null, 2),
  'utf8'
);
console.log('✓ Generated compatibility.json');
