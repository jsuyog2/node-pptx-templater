const fs = require('fs-extra');
const { resolve } = require('path');
const path = require('path');

const DOCS_DIR = resolve(__dirname, '../docs');
const pkg = require('../package.json');
const VERSION = pkg.version;

/**
 * Dynamically crawls Node.js source files under src/ and extracts classes, 
 * methods, descriptions, arguments, parameters, returns, and examples from JSDoc.
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
    const docPattern = /\/\*\*([\s\S]*?)\*\/\s*(?:async\s+)?([a-zA-Z0-9_#]+)\s*\(([^)]*)\)/g;
    let match;
    
    while ((match = docPattern.exec(content)) !== null) {
      const jsdoc = match[1];
      const name = match[2];
      const args = match[3];
      
      // Ignore private helper methods and constructor
      if (name.startsWith('#') || name === 'constructor') {
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
  });
  
  return apis;
}

function renderSidebarAPIs(apis) {
  const classes = {};
  apis.forEach(api => {
    if (!classes[api.className]) {
      classes[api.className] = [];
    }
    classes[api.className].push(api);
  });
  
  let html = '';
  for (const className of Object.keys(classes)) {
    html += `
      <div class="space-y-1 mt-4">
        <h3 class="text-xs uppercase text-gray-500 font-semibold tracking-wider px-3 flex items-center justify-between cursor-pointer hover:text-white transition-colors" onclick="toggleSidebarGroup('class-${className.toLowerCase()}')">
          <span>${className} API</span>
          <span class="text-[9px] transform transition-transform duration-200" id="arrow-class-${className.toLowerCase()}">▼</span>
        </h3>
        <ul id="sidebar-list-class-${className.toLowerCase()}" class="space-y-1 pl-2 transition-all duration-300">
          <li><a href="#class-${className.toLowerCase()}" class="nav-item block px-3 py-1.5 text-xs text-gray-400 hover:text-white rounded-md transition-colors font-medium">Overview</a></li>
          ${classes[className].map(m => `
            <li><a href="#method-${className.toLowerCase()}-${m.name.toLowerCase()}" class="nav-item block px-3 py-1 text-[11px] text-gray-500 hover:text-white rounded-md transition-colors pl-4 font-mono">${m.name}()</a></li>
          `).join('')}
        </ul>
      </div>
    `;
  }
  
  return html;
}

function renderContentAPIs(apis) {
  const classes = {};
  apis.forEach(api => {
    if (!classes[api.className]) {
      classes[api.className] = [];
    }
    classes[api.className].push(api);
  });
  
  let html = '';
  
  for (const [className, methods] of Object.entries(classes)) {
    html += `
      <!-- Class ${className} Section -->
      <section id="class-${className.toLowerCase()}" class="doc-section space-y-6 hidden">
        <div class="flex items-center gap-3 border-b border-white/5 pb-4">
          <span class="px-2 py-0.5 rounded bg-brand-500/10 text-brand-400 font-mono text-[10px] font-bold uppercase tracking-wider">class</span>
          <h1 class="font-title text-3xl font-extrabold text-white">${className}</h1>
        </div>
        
        <p class="text-sm text-gray-400">Located in: <code class="text-indigo-400 font-mono text-xs bg-indigo-500/5 px-2 py-1 rounded border border-indigo-500/10">${methods[0].file}</code></p>
        
        <div class="grid grid-cols-1 gap-6 mt-6">
    `;
    
    methods.forEach(m => {
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
                    <td class="px-4 py-2.5 font-semibold text-indigo-300 font-mono">${p.name}</td>
                    <td class="px-4 py-2.5 font-mono text-pink-400">${p.type}</td>
                    <td class="px-4 py-2.5 text-gray-300">${p.desc}</td>
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
            <span class="font-mono text-emerald-400 font-semibold">${m.returns.type}</span>
            <span class="text-gray-300">${m.returns.desc}</span>
          </div>
        `;
      }
      
      let exampleBlock = '';
      if (m.exampleCode) {
        exampleBlock = `
          <div class="mt-4">
            <span class="text-xs font-semibold text-gray-500 block mb-1.5 font-title">USAGE EXAMPLE</span>
            <pre class="relative group"><code class="language-javascript block p-4 bg-[#05070c] border border-white/5 rounded-xl text-indigo-300 font-mono text-xs overflow-x-auto">${m.exampleCode}</code><button class="copy-btn absolute top-3.5 right-3.5 text-[10px] bg-white/5 hover:bg-white/10 text-gray-400 hover:text-white px-2.5 py-1 rounded-lg border border-white/10 transition-all">Copy</button></pre>
          </div>
        `;
      }
      
      html += `
        <!-- Method ${m.name} Card -->
        <div id="method-${className.toLowerCase()}-${m.name.toLowerCase()}" class="glass-card p-6 bg-slate-900/30 border border-white/5 rounded-2xl space-y-4 relative overflow-hidden transition-all duration-300 hover:border-brand-500/30">
          <div class="flex flex-wrap items-center justify-between gap-3 border-b border-white/5 pb-3">
            <div class="flex items-center gap-2">
              <span class="px-2 py-0.5 rounded bg-brand-500/10 text-brand-400 font-mono text-[9px] font-bold uppercase tracking-wider">method</span>
              <h3 class="font-title text-xl font-bold text-white hover:text-brand-400 transition-colors">${m.name}</h3>
            </div>
            <div class="font-mono text-xs text-indigo-400/90 bg-indigo-500/5 border border-indigo-500/10 px-3 py-1 rounded-full">
              ${signature}
            </div>
          </div>
          <p class="text-sm text-gray-300 leading-relaxed">${m.description || 'No description available.'}</p>
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

async function main() {
  const apis = extractAPIDocs();
  const sidebarHtml = renderSidebarAPIs(apis);
  const contentHtml = renderContentAPIs(apis);

  const HTML_CONTENT = `<!DOCTYPE html>
<html lang="en" class="scroll-smooth">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  
  <!-- Primary SEO Meta Tags -->
  <title>node-pptx-templater — Safe PowerPoint & OpenXML Template Engine</title>
  <meta name="description" content="High-performance, zero-dependency Node.js library to update text, replace image shapes, cache Excel workbook data in charts, and merge table cells without PowerPoint Repairwarnings.">
  <meta name="keywords" content="node pptx template, pptx templating library, powerpoint template engine, pptx editor nodejs, openxml powerpoint library, update powerpoint charts nodejs, powerpoint automation, pptx generator, pptx placeholder replacement, edit pptx without powerpoint, pptx chart update, pptx table update, openxml nodejs, powerpoint report generator, pptx cell merge, stack slide layers, duplicate slides, nodejs presentation automation">
  <link rel="canonical" href="https://jsuyog2.github.io/node-pptx-templater/">
  
  <!-- Open Graph / Facebook -->
  <meta property="og:type" content="website">
  <meta property="og:title" content="node-pptx-templater — Low-Level OpenXML PowerPoint Engine">
  <meta property="og:description" content="Automate presentations in Node.js with zero dependencies. Support cell merging, dynamic slides, Z-order layers, and Excel chart updating.">
  <meta property="og:url" content="https://jsuyog2.github.io/node-pptx-templater/">
  <meta property="og:image" content="https://jsuyog2.github.io/node-pptx-templater/assets/brand-preview.png">
  
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
    
    <div class="flex items-center gap-5">
      <div class="relative hidden md:block">
        <input type="text" id="doc-search" placeholder="Search API methods (Ctrl + K)..." autocomplete="off" class="w-80 px-4 py-1.5 bg-[#080b11]/80 border border-white/5 rounded-full text-sm text-gray-200 placeholder-gray-500 focus:outline-none focus:border-brand-500 focus:ring-1 focus:ring-brand-500 transition-all">
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

        <!-- Dynamic API References Sidebar List -->
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
        <p class="text-sm text-gray-300">OpenXML PowerPoint tables require precise cell spans coordinate structures. The top-left cell acts as the **origin**, declaring \`gridSpan\` and \`rowSpan\`. The remaining shadowed cells must flag \`hMerge\` and \`vMerge\` to avoid breaking PowerPoint table layout models.</p>
        
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
        <p class="text-sm text-gray-300">PowerPoint embeds an Excel worksheet (\`ppt/embeddings/\`) that controls chart datasets. Standard scripts only edit visual chart coordinates, corrupting calculations. <strong>node-pptx-templater</strong> compiles updates for both XML caches and spreadsheet rows.</p>
        
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
        <p class="text-sm text-gray-300">Programmatically stack shapes, images, charts, and tables using layer indices. The engine translates commands (Bring Forward, Send to Back) directly into OpenXML element orders within slide \`&lt;p:spTree&gt;\` blocks.</p>
        
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
  --border-color: rgba(0, 0, 0, 0.06);
  --text-main: #0f172a;
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
  background-color: rgba(255, 255, 255, 0.6);
  box-shadow: 0 8px 32px 0 rgba(0, 0, 0, 0.05);
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
  background-color: rgba(79, 70, 229, 0.08);
  color: var(--primary);
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
::-webkit-scrollbar-thumb:hover {
  background: rgba(255, 255, 255, 0.2);
}

/* Custom layout sizes */
.text-xxs {
  font-size: 8px;
}
.text-xs {
  font-size: 11px;
}
.text-sm {
  font-size: 13px;
}
`;

  const JS_CONTENT = `// Documentation Engine Client Logic
document.addEventListener('DOMContentLoaded', () => {
  const navItems = document.querySelectorAll('.nav-item');
  const sections = document.querySelectorAll('.doc-section, .class-docs-section');
  
  // Show/Hide page sections based on URL hashes
  function showSection(hash) {
    const cleanId = hash.replace('#', '');
    
    sections.forEach(sec => sec.classList.add('hidden'));
    navItems.forEach(item => item.classList.remove('active'));
    
    const targetSection = document.getElementById(cleanId);
    if (targetSection) {
      targetSection.classList.remove('hidden');
      window.scrollTo({ top: 0, behavior: 'smooth' });
      
      // GSAP Liquid Glass Reveal Animation
      gsap.fromTo(targetSection, 
        { opacity: 0, y: 15 },
        { opacity: 1, y: 0, duration: 0.4, ease: "power2.out" }
      );
    }
    
    const activeLink = document.querySelector(\`a[href="#\${cleanId}"]\`);
    if (activeLink) {
      activeLink.classList.add('active');
    }
  }

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

  // Liquid Glass Glow Cursor Tracker
  document.addEventListener('mousemove', e => {
    document.querySelectorAll('.glass-card, .method-card').forEach(card => {
      const rect = card.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;
      card.style.setProperty('--mouse-x', \`\${x}px\`);
      card.style.setProperty('--mouse-y', \`\${y}px\`);
    });
  });

  // GSAP Entrance Load Animations
  gsap.from("header", { opacity: 0, y: -20, duration: 0.6, ease: "power3.out" });
  gsap.from("aside", { opacity: 0, x: -30, duration: 0.6, ease: "power3.out", delay: 0.1 });
  gsap.from(".glow-orb", { opacity: 0, scale: 0.8, duration: 1.5, ease: "power2.out" });

  // Light/Dark Theme Switcher
  const themeToggle = document.getElementById('theme-toggle');
  const body = document.body;
  const savedTheme = localStorage.getItem('theme');

  if (savedTheme === 'light') {
    body.classList.remove('dark-theme');
    body.classList.add('light-theme', 'bg-white', 'text-slate-800');
    body.classList.remove('bg-[#080b11]', 'text-gray-200');
    themeToggle.querySelector('.toggle-icon').textContent = '☀️';
  }

  themeToggle.addEventListener('click', () => {
    if (body.classList.contains('light-theme')) {
      body.classList.remove('light-theme', 'bg-white', 'text-slate-800');
      body.classList.add('bg-[#080b11]', 'text-gray-200');
      themeToggle.querySelector('.toggle-icon').textContent = '🌙';
      localStorage.setItem('theme', 'dark');
    } else {
      body.classList.add('light-theme', 'bg-white', 'text-slate-800');
      body.classList.remove('bg-[#080b11]', 'text-gray-200');
      themeToggle.querySelector('.toggle-icon').textContent = '☀️';
      localStorage.setItem('theme', 'light');
    }
  });

  // Client Sidebar Navigation Filter Search
  const searchInput = document.getElementById('doc-search');
  searchInput.addEventListener('input', (e) => {
    const query = e.target.value.toLowerCase().trim();
    if (!query) {
      document.querySelectorAll('.nav-item').forEach(item => item.parentElement.style.display = 'block');
      document.querySelectorAll('.class-sidebar-group').forEach(group => group.style.display = 'block');
      return;
    }

    document.querySelectorAll('.nav-item').forEach(item => {
      const text = item.textContent.toLowerCase();
      const parent = item.parentElement;
      if (text.includes(query)) {
        parent.style.display = 'block';
      } else {
        parent.style.display = 'none';
      }
    });
  });

  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
      e.preventDefault();
      searchInput.focus();
    }
  });

  // Clipboard Text Copier
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
  
  // SEO Sitemap.xml
  const sitemap = `<?xml version="1.0" encoding="UTF-8"?>
<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">
  <url>
    <loc>https://jsuyog2.github.io/node-pptx-templater/</loc>
    <lastmod>${new Date().toISOString().split('T')[0]}</lastmod>
    <changefreq>monthly</changefreq>
    <priority>1.0</priority>
  </url>
</urlset>`;
  await fs.writeFile(resolve(DOCS_DIR, 'sitemap.xml'), sitemap.trim());

  // SEO Robots.txt
  const robots = `User-agent: *
Allow: /
Sitemap: https://jsuyog2.github.io/node-pptx-templater/sitemap.xml`;
  await fs.writeFile(resolve(DOCS_DIR, 'robots.txt'), robots.trim());

  console.log('Successfully built GitHub Pages documentation with sitemap and robots.txt under docs/');
}

main().catch(err => {
  console.error('Error building docs:', err);
  process.exit(1);
});
