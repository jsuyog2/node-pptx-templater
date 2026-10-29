'use client';

import apiReferenceData from '../data/apiReference.json';

const STATIC_ROUTES = [
  {
    id: 'route-home',
    title: 'Home Presentation & Overview',
    subtitle: 'High-performance pure JavaScript PowerPoint template engine',
    category: 'Routes',
    type: 'route',
    href: '/',
    keywords: ['home', 'start', 'landing', 'powerpoint', 'template', 'engine', 'pure js', 'overview'],
  },
  {
    id: 'route-docs',
    title: 'Documentation Overview',
    subtitle: 'Paradigm comparison, AST run healing, chart dual-sync, VFS',
    category: 'Documentation',
    type: 'guide',
    href: '/docs',
    keywords: ['docs', 'architecture', 'vfs', 'ast', 'healing', 'virtual file system', 'security'],
  },
  {
    id: 'route-installation',
    title: 'Installation & Setup Guide',
    subtitle: 'npm, pnpm, yarn, bun, ESM/CJS, TypeScript definitions, serverless',
    category: 'Documentation',
    type: 'guide',
    href: '/docs/installation',
    keywords: ['install', 'npm', 'pnpm', 'yarn', 'bun', 'typescript', 'esm', 'cjs', 'setup'],
  },
  {
    id: 'route-capabilities',
    title: 'Capabilities & Compatibility Matrix',
    subtitle: 'Runtime matrix across Node.js, AWS Lambda, Cloudflare Workers, Vercel Edge, Bun, Deno',
    category: 'Documentation',
    type: 'guide',
    href: '/docs/capabilities',
    keywords: ['capabilities', 'matrix', 'compatibility', 'lambda', 'workers', 'edge', 'runtimes', 'serverless'],
  },
  {
    id: 'route-api',
    title: 'Complete API Reference Catalog',
    subtitle: 'All 8 standardized categories with 135 methods and signatures',
    category: 'API Reference',
    type: 'route',
    href: '/docs/api-reference',
    keywords: ['api', 'reference', 'methods', 'functions', 'catalog', 'signatures'],
  },
];

const GUIDE_SLIDES = [
  {
    id: 'slide-zero-dep',
    title: 'Zero Office Dependencies & Pure JS Runtime',
    subtitle: 'Serverless, Edge & Node.js Native Execution without Office or LibreOffice',
    category: 'Overview',
    type: 'guide',
    href: '/docs#slide-1',
    keywords: ['zero', 'office', 'libreoffice', 'pure js', 'edge', 'lambda', 'workers'],
  },
  {
    id: 'slide-quickstart',
    title: 'Under 15 Lines of Code Quickstart',
    subtitle: 'End-to-end ingest, mutation, and export pipeline',
    category: 'Overview',
    type: 'guide',
    href: '/docs#slide-2',
    keywords: ['quickstart', 'example', 'pipeline', 'load', 'replaceText', 'updateChart', 'saveToFile'],
  },
  {
    id: 'slide-vfs',
    title: 'In-Memory Virtual File System (VFS)',
    subtitle: 'Sub-15ms cold start without disk I/O',
    category: 'Architecture',
    type: 'guide',
    href: '/docs#slide-3',
    keywords: ['vfs', 'virtual file system', 'memory', 'cold start', 'lambda', 'performance'],
  },
  {
    id: 'slide-ast-healing',
    title: 'AST Run Fragment Healing Engine',
    subtitle: 'Automatic coalescing of split template tokens across <a:r> tags',
    category: 'Text & Runs',
    type: 'guide',
    href: '/docs#slide-4',
    keywords: ['ast', 'run healing', 'fragmentation', 'replaceText', 'tokens', 'placeholders'],
  },
  {
    id: 'slide-chart-dual-sync',
    title: 'Chart Excel Dual-Sync Architecture',
    subtitle: 'Simultaneous DrawingML and embedded XLSX synchronization',
    category: 'Charts',
    type: 'guide',
    href: '/docs#slide-5',
    keywords: ['chart', 'excel', 'dual-sync', 'drawingml', 'xlsx', 'repair', 'graphs'],
  },
  {
    id: 'slide-table-pagination',
    title: 'Dynamic Multi-Slide Table Auto-Pagination',
    subtitle: 'Automatic row height EMU calculation and header repetition',
    category: 'Tables',
    type: 'guide',
    href: '/docs#slide-6',
    keywords: ['table', 'pagination', 'split', 'overflow', 'emu', 'headers', 'rows'],
  },
  {
    id: 'slide-security',
    title: 'Hardened Package & XML Security',
    subtitle: 'XXE, Billion Laughs, and Zip-Slip resistance',
    category: 'Security',
    type: 'guide',
    href: '/docs#slide-7',
    keywords: ['security', 'xxe', 'billion laughs', 'zip slip', 'hardening', 'safeParseXml'],
  },
];

let cachedIndex = null;

export function getSearchIndex() {
  if (cachedIndex) return cachedIndex;

  const items = [...STATIC_ROUTES, ...GUIDE_SLIDES];

  // Ingest all 135 API Methods across the 8 categories
  if (apiReferenceData && apiReferenceData.categories) {
    apiReferenceData.categories.forEach((cat) => {
      const methods = apiReferenceData.methodsByCategory?.[cat] || [];
      methods.forEach((m) => {
        const paramNames = (m.parameters || []).map((p) => p.name).join(' ');
        items.push({
          id: `api-${cat}-${m.name}`,
          title: `${m.name}()`,
          subtitle: m.description || m.signature || '',
          signature: m.signature || `ppt.${m.name}()`,
          category: cat,
          type: 'api',
          href: `/docs/api-reference?category=${encodeURIComponent(cat)}`,
          keywords: [
            m.name.toLowerCase(),
            cat.toLowerCase(),
            'method',
            'api',
            paramNames.toLowerCase(),
            m.isAsync ? 'async' : 'sync',
          ],
        });
      });
    });
  }

  cachedIndex = items;
  return items;
}

export function searchItems(query = '', limit = 12) {
  const q = query.trim().toLowerCase();
  if (!q) return [];

  const index = getSearchIndex();

  const results = [];
  for (const item of index) {
    let score = 0;
    const titleLower = item.title.toLowerCase();
    const subtitleLower = (item.subtitle || '').toLowerCase();
    const categoryLower = (item.category || '').toLowerCase();
    const sigLower = (item.signature || '').toLowerCase();

    // Exact name match
    if (titleLower === q || titleLower === `${q}()`) {
      score += 100;
    } else if (titleLower.startsWith(q)) {
      score += 50;
    } else if (titleLower.includes(q)) {
      score += 30;
    }

    // Category match
    if (categoryLower.includes(q)) {
      score += 20;
    }

    // Signature match
    if (sigLower.includes(q)) {
      score += 15;
    }

    // Subtitle / description match
    if (subtitleLower.includes(q)) {
      score += 10;
    }

    // Keyword array match
    if (item.keywords && item.keywords.some((kw) => kw.includes(q))) {
      score += 10;
    }

    if (score > 0) {
      results.push({ item, score });
    }
  }

  // Sort by score descending
  results.sort((a, b) => b.score - a.score);

  return results.slice(0, limit).map((r) => r.item);
}
