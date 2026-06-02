const fsExtra = require('fs-extra');
const { resolve } = require('path');

const DOCS_DIR = resolve(__dirname, '../docs');
const pkg = require('../package.json');
const VERSION = pkg.version;

const HTML_CONTENT = `<!DOCTYPE html>
<html lang="en" class="scroll-smooth">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>node-pptx-templater — Low-Level OpenXML Engine Docs</title>
  <meta name="description" content="High-performance, low-level PowerPoint OpenXML template engine for Node.js. Dynamically replace text, images, charts, and tables without PowerPoint Repair errors.">
  
  <!-- Search Engine Meta Tags -->
  <meta name="keywords" content="PowerPoint template, PPTX template engine, PowerPoint automation, PPTX generator, Node.js PowerPoint, OpenXML PowerPoint, merge cells, dynamic slides">
  <link rel="canonical" href="https://jsuyog2.github.io/node-pptx-templater/">
  
  <!-- Open Graph / Twitter Card -->
  <meta property="og:type" content="website">
  <meta property="og:title" content="node-pptx-templater — Low-Level OpenXML PowerPoint Engine">
  <meta property="og:description" content="PowerPoint template automation in Node.js with zero dependencies. Safe cell merging, Excel cache sync, and slide imports.">
  <meta property="og:url" content="https://jsuyog2.github.io/node-pptx-templater/">
  <meta name="twitter:card" content="summary_large_image">
  
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
  
  <link rel="stylesheet" href="style.css">

  <!-- Schema.org JSON-LD Structured Data -->
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
    "featureList": "PowerPoint text placeholder replacement, slide duplication, image insertion, Excel data workbook synchronization, DrawingML table cell merging, slide reordering, external slide importing",
    "downloadUrl": "https://www.npmjs.com/package/node-pptx-templater"
  }
  </script>
</head>
<body class="bg-[#0b0f19] text-gray-200 font-sans antialiased selection:bg-brand-500 selection:text-white transition-colors duration-200">
  <!-- Top Navigation Header -->
  <header class="fixed top-0 left-0 right-0 h-16 bg-[#131b2e]/80 backdrop-blur-md border-b border-gray-800 flex items-center justify-between px-6 z-50">
    <div class="flex items-center gap-3">
      <span class="text-2xl">📊</span>
      <span class="font-title font-bold text-xl bg-gradient-to-r from-brand-500 to-purple-500 bg-clip-text text-transparent">node-pptx-templater</span>
      <span class="bg-brand-500/10 text-brand-500 text-xs px-2.5 py-0.5 rounded-full font-semibold border border-brand-500/30">v${VERSION}</span>
    </div>
    
    <div class="flex items-center gap-4">
      <div class="relative hidden md:block">
        <input type="text" id="doc-search" placeholder="Search docs (Ctrl + K)..." autocomplete="off" class="w-80 px-4 py-1.5 bg-[#0b0f19] border border-gray-800 rounded-full text-sm text-gray-200 placeholder-gray-500 focus:outline-none focus:border-brand-500 focus:ring-1 focus:ring-brand-500">
      </div>
      
      <button id="theme-toggle" aria-label="Toggle Theme" class="text-xl p-1 hover:text-brand-500 transition-colors">
        <span class="toggle-icon">🌙</span>
      </button>
      
      <a href="https://github.com/jsuyog2/node-pptx-templater" class="text-gray-400 hover:text-white transition-colors" target="_blank" rel="noopener">
        <svg height="24" viewBox="0 0 16 16" width="24" fill="currentColor">
          <path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"></path>
        </svg>
      </a>
    </div>
  </header>

  <!-- Sidebar & Main Content Wrapper -->
  <div class="flex pt-16 min-h-screen">
    <!-- Left Navigation Sidebar -->
    <aside class="w-72 bg-[#131b2e] border-r border-gray-800 fixed top-16 bottom-0 left-0 overflow-y-auto px-6 py-8 hidden md:block" id="sidebar">
      <nav class="space-y-8">
        <div class="space-y-2">
          <h3 class="text-xs uppercase text-gray-500 font-semibold tracking-wider">Getting Started</h3>
          <ul class="space-y-1">
            <li><a href="#introduction" class="nav-item block px-3 py-2 text-sm text-gray-400 hover:text-white rounded-md transition-colors active">Introduction</a></li>
            <li><a href="#installation" class="nav-item block px-3 py-2 text-sm text-gray-400 hover:text-white rounded-md transition-colors">Installation & Onboarding</a></li>
            <li><a href="#quickstart" class="nav-item block px-3 py-2 text-sm text-gray-400 hover:text-white rounded-md transition-colors">Quick Start Guide</a></li>
            <li><a href="#learningpaths" class="nav-item block px-3 py-2 text-sm text-gray-400 hover:text-white rounded-md transition-colors">Learning Paths</a></li>
          </ul>
        </div>
        
        <div class="space-y-2">
          <h3 class="text-xs uppercase text-gray-500 font-semibold tracking-wider">Core Features</h3>
          <ul class="space-y-1">
            <li><a href="#text-replacement" class="nav-item block px-3 py-2 text-sm text-gray-400 hover:text-white rounded-md transition-colors">Text Tag Replacement</a></li>
            <li><a href="#image-replacement" class="nav-item block px-3 py-2 text-sm text-gray-400 hover:text-white rounded-md transition-colors">Image Substitution</a></li>
            <li><a href="#slide-management" class="nav-item block px-3 py-2 text-sm text-gray-400 hover:text-white rounded-md transition-colors">Slide Duplication & CRUD</a></li>
            <li><a href="#slide-import" class="nav-item block px-3 py-2 text-sm text-gray-400 hover:text-white rounded-md transition-colors">Cross-Template Imports</a></li>
          </ul>
        </div>

        <div class="space-y-2">
          <h3 class="text-xs uppercase text-gray-500 font-semibold tracking-wider">Advanced Elements</h3>
          <ul class="space-y-1">
            <li><a href="#table-merging" class="nav-item block px-3 py-2 text-sm text-gray-400 hover:text-white rounded-md transition-colors">Table Cell Merging</a></li>
            <li><a href="#chart-engine" class="nav-item block px-3 py-2 text-sm text-gray-400 hover:text-white rounded-md transition-colors">Interactive Chart Updates</a></li>
            <li><a href="#navigation-links" class="nav-item block px-3 py-2 text-sm text-gray-400 hover:text-white rounded-md transition-colors">Action Hyperlinks</a></li>
          </ul>
        </div>

        <div class="space-y-2">
          <h3 class="text-xs uppercase text-gray-500 font-semibold tracking-wider">OpenXML Architecture</h3>
          <ul class="space-y-1">
            <li><a href="#openxml-internals" class="nav-item block px-3 py-2 text-sm text-gray-400 hover:text-white rounded-md transition-colors">Packaging & XML Structure</a></li>
            <li><a href="#manifest-managers" class="nav-item block px-3 py-2 text-sm text-gray-400 hover:text-white rounded-md transition-colors">Managers & Core Validation</a></li>
            <li><a href="#troubleshooting" class="nav-item block px-3 py-2 text-sm text-gray-400 hover:text-white rounded-md transition-colors">Troubleshooting & FAQs</a></li>
          </ul>
        </div>
      </nav>
    </aside>

    <!-- Main Content Area -->
    <main class="md:ml-72 flex-1 px-6 md:px-12 py-10 max-w-5xl overflow-y-auto">
      
      <!-- Introduction Section -->
      <section id="introduction" class="doc-section space-y-6 active-section">
        <h1 class="font-title text-4xl font-extrabold text-white">Introduction</h1>
        <p class="text-lg text-gray-300">
          Welcome to the documentation for <strong>node-pptx-templater</strong>—a low-level, zero-dependency, high-performance PowerPoint template engine built for Node.js. It modifies PPTX files by applying structured, schema-validated XML edits directly to the package, solving standard text fragmentation and table corruption errors once and for all.
        </p>
        
        <div class="p-4 bg-brand-500/10 border-l-4 border-brand-500 rounded-r-xl">
          <span class="font-semibold text-brand-500 block">⚡ Visual Design vs Code Automation</span>
          With <code>node-pptx-templater</code>, you separate design from code. Create slide layouts, animations, transitions, and headers inside standard PowerPoint editors (Microsoft Office, Google Slides, or Keynote) and bind data fields using simple placeholders (<code>{{company}}</code>, <code>{{sales-table}}</code>).
        </div>

        <h2 class="font-title text-2xl font-bold text-white mt-8">Key Advantages</h2>
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mt-4">
          <div class="p-5 bg-[#131b2e] border border-gray-800 rounded-xl">
            <h4 class="font-semibold text-white mb-2">🚀 Pure JavaScript & Zip Compression</h4>
            <p class="text-sm text-gray-400">Zero dependencies on local installations of Microsoft Office, LibreOffice, or Java. Easily deploys to serverless runtimes (AWS Lambda, Vercel, Google Cloud Functions).</p>
          </div>
          <div class="p-5 bg-[#131b2e] border border-gray-800 rounded-xl">
            <h4 class="font-semibold text-white mb-2">🔀 Run Fragment Resolution</h4>
            <p class="text-sm text-gray-400">PowerPoint splits placeholders across separate XML tags (e.g. <code>{{c</code>, <code>ompany}}</code>). Our parser dynamically merges text runs to ensure 100% replacement success.</p>
          </div>
          <div class="p-5 bg-[#131b2e] border border-gray-800 rounded-xl">
            <h4 class="font-semibold text-white mb-2">📊 Caching Workbook Updater</h4>
            <p class="text-sm text-gray-400">Updates underlying chart data Excel spreadsheets and synchronizes visual data cache matrices to avoid refresh alerts in PowerPoint Online.</p>
          </div>
          <div class="p-5 bg-[#131b2e] border border-gray-800 rounded-xl">
            <h4 class="font-semibold text-white mb-2">🛡️ Table rowId Sanitizer</h4>
            <p class="text-sm text-gray-400">Dynamically allocates unique 32-bit row identification hashes whenever a row is cloned or inserted, eliminating PowerPoint Repair prompts.</p>
          </div>
        </div>
      </section>

      <!-- Installation & Onboarding Section -->
      <section id="installation" class="doc-section space-y-6">
        <h1 class="font-title text-4xl font-extrabold text-white">Installation & Onboarding</h1>
        <p class="text-gray-300">Set up the library in less than 5 minutes. Learn how it integrates with your existing Node.js codebases.</p>
        
        <h2 class="font-title text-2xl font-bold text-white mt-6">NPM Setup</h2>
        <pre class="relative group"><code class="language-bash block p-4 bg-[#090c15] border border-gray-800 rounded-lg text-indigo-300 font-mono text-sm">npm install node-pptx-templater</code><button class="copy-btn absolute top-3 right-3 text-xs bg-gray-800 text-gray-300 hover:text-white px-2 py-1 rounded transition-colors">Copy</button></pre>
        
        <h2 class="font-title text-2xl font-bold text-white mt-6">Prerequisites</h2>
        <ul class="list-disc pl-6 space-y-2 text-gray-400 text-sm">
          <li><strong>Node.js Engine</strong>: >= 18.0.0 (Supports CommonJS out of the box).</li>
          <li><strong>Module Support</strong>: Designed to run flawlessly across standard Node environments, Vercel edge routes, and AWS Lambdas.</li>
        </ul>
      </section>

      <!-- Quick Start Section -->
      <section id="quickstart" class="doc-section space-y-6">
        <h1 class="font-title text-4xl font-extrabold text-white">Quick Start Guide</h1>
        <p class="text-gray-300">Generate your first presentation file in under 10 minutes by copying the following snippet.</p>
        
        <pre class="relative group"><code class="language-javascript block p-4 bg-[#090c15] border border-gray-800 rounded-lg text-indigo-300 font-mono text-sm">const { PPTXTemplater } = require('node-pptx-templater');

async function buildPresentation() {
  // Load the design template
  const ppt = await PPTXTemplater.load('marketing_template.pptx');
  
  // Replace text on slide 1
  ppt.useSlide(1)
     .replaceTextByTag('header', 'Acme Q2 Roadmap')
     .replaceMultiple({
       date: 'June 2026',
       speaker: 'John Doe'
     });

  // Export non-corrupted PPTX
  await ppt.saveToFile('./output/q2_roadmap.pptx');
  console.log('Presentation generated successfully!');
}

buildPresentation();</code><button class="copy-btn absolute top-3 right-3 text-xs bg-gray-800 text-gray-300 hover:text-white px-2 py-1 rounded transition-colors">Copy</button></pre>
      </section>

      <!-- Learning Paths Section -->
      <section id="learningpaths" class="doc-section space-y-6">
        <h1 class="font-title text-4xl font-extrabold text-white">Learning Paths</h1>
        <p class="text-gray-300">Choose a path tailored to your technical requirements and familiarity with OpenXML PowerPoint architectures.</p>
        
        <!-- Tab Selectors -->
        <div class="flex border-b border-gray-800 gap-4">
          <button onclick="switchPath('beginner')" id="tab-beginner" class="path-tab px-4 py-2 border-b-2 border-brand-500 font-semibold text-white">Beginner Path</button>
          <button onclick="switchPath('intermediate')" id="tab-intermediate" class="path-tab px-4 py-2 border-b-2 border-transparent text-gray-400 font-semibold hover:text-white">Intermediate Path</button>
          <button onclick="switchPath('advanced')" id="tab-advanced" class="path-tab px-4 py-2 border-b-2 border-transparent text-gray-400 font-semibold hover:text-white">Advanced Path</button>
        </div>

        <!-- Beginner Content -->
        <div id="path-content-beginner" class="path-content space-y-4">
          <h3 class="text-lg font-bold text-white">Path 1: Beginner Level</h3>
          <p class="text-sm text-gray-400">Focuses on template concepts, setting up placeholders in your presentation editor, and replacing basic text/image fields.</p>
          <ul class="list-decimal pl-6 space-y-2 text-sm text-gray-400">
            <li>What is a PPTX Template? (Just standard PPTX files with text tags like <code>{{name}}</code>).</li>
            <li>How to format shapes and placeholders cleanly in PowerPoint.</li>
            <li>Performing simple mail-merge replacement operations.</li>
          </ul>
        </div>

        <!-- Intermediate Content -->
        <div id="path-content-intermediate" class="path-content hidden space-y-4">
          <h3 class="text-lg font-bold text-white">Path 2: Intermediate Level</h3>
          <p class="text-sm text-gray-400">Covers table structure modifications, chart workbook calculations, and duplicating templates dynamically.</p>
          <ul class="list-decimal pl-6 space-y-2 text-sm text-gray-400">
            <li>Handling slide duplicates to build reports from single layouts.</li>
            <li>Adding, deleting, and updating slide table rows safely.</li>
            <li>Replacing series and category lists inside existing charts.</li>
          </ul>
        </div>

        <!-- Advanced Content -->
        <div id="path-content-advanced" class="path-content hidden space-y-4">
          <h3 class="text-lg font-bold text-white">Path 3: Advanced Level</h3>
          <p class="text-sm text-gray-400">For developers seeking to build complex presentation workflows, custom extensions, and understand OpenXML packaging rules.</p>
          <ul class="list-decimal pl-6 space-y-2 text-sm text-gray-400">
            <li>OpenXML elements, relationships (<code>.rels</code>), and content overrides.</li>
            <li>Writing extensions to interact with slide geometry paths.</li>
            <li>Optimizing execution speed and garbage collection when generating files.</li>
          </ul>
        </div>
      </section>

      <!-- Text Replacement Section -->
      <section id="text-replacement" class="doc-section space-y-6">
        <h1 class="font-title text-4xl font-extrabold text-white">Text Tag Replacement</h1>
        
        <div>
          <span class="bg-indigo-500/10 text-indigo-400 text-xs px-2.5 py-1 rounded font-semibold border border-indigo-500/30">1. Beginner Explanation</span>
          <p class="text-sm text-gray-400 mt-2">Replace simple text tags such as <code>{{title}}</code> or <code>{{date}}</code> with dynamic variables. The library keeps the original font, color, bold/italic, alignment, and size styles exactly as designed in PowerPoint.</p>
        </div>

        <div>
          <span class="bg-indigo-500/10 text-indigo-400 text-xs px-2.5 py-1 rounded font-semibold border border-indigo-500/30">2. Step-by-Step Guide</span>
          <ol class="list-decimal pl-6 space-y-2 text-sm text-gray-400 mt-2">
            <li>Draw a Text Box shape in PowerPoint or Google Slides.</li>
            <li>Write your placeholder surrounded by braces (e.g. <code>{{client}}</code>).</li>
            <li>Call <code>ppt.replaceText({ client: "John Doe" })</code>.</li>
          </ol>
        </div>

        <div>
          <span class="bg-indigo-500/10 text-indigo-400 text-xs px-2.5 py-1 rounded font-semibold border border-indigo-500/30">3. Code Example</span>
          <pre class="relative group"><code class="language-javascript block p-4 bg-[#090c15] border border-gray-800 rounded-lg text-indigo-300 font-mono text-sm">ppt.useSlide(1);
ppt.replaceTextByTag('client', 'John Doe');

// Replace multiple at once
ppt.replaceMultiple({
  'date': '2026-06-02',
  'revenue': '$1.2M'
});</code><button class="copy-btn absolute top-3 right-3 text-xs bg-gray-800 text-gray-300 hover:text-white px-2 py-1 rounded transition-colors">Copy</button></pre>
        </div>

        <div>
          <span class="bg-indigo-500/10 text-indigo-400 text-xs px-2.5 py-1 rounded font-semibold border border-indigo-500/30">4. Common Mistakes & Troubleshooting</span>
          <p class="text-sm text-gray-400 mt-2">
            <strong>Placeholder not replacing:</strong> PowerPoint often splits text runs behind the scenes due to spelling checks or manual typing pauses. In slide XML, this splits <code>{{client}}</code> into:
            <code class="block bg-gray-950 p-2 my-2 text-xs text-red-400">&lt;a:t&gt;{{cl&lt;/a:t&gt;&lt;a:t&gt;ient}}&lt;/a:t&gt;</code>
            <strong>How to resolve:</strong> Highlight the placeholder in PowerPoint, cut it, and paste it back as "Keep Text Only" to merge the XML tags. Alternatively, set <code>PPTX_LOG_LEVEL=debug</code> to identify fragmented runs.
          </p>
        </div>
      </section>

      <!-- Image Replacement Section -->
      <section id="image-replacement" class="doc-section space-y-6">
        <h1 class="font-title text-4xl font-extrabold text-white">Image Substitution</h1>
        
        <div>
          <span class="bg-indigo-500/10 text-indigo-400 text-xs px-2.5 py-1 rounded font-semibold border border-indigo-500/30">1. Beginner Explanation</span>
          <p class="text-sm text-gray-400 mt-2">Substitute picture shapes in your templates (e.g. placeholder logos or headshots) with dynamic image files. The library updates the underlying image data binary in the ZIP archive, keeping the exact dimensions, rotation, coordinates, and styling limits.</p>
        </div>

        <div>
          <span class="bg-indigo-500/10 text-indigo-400 text-xs px-2.5 py-1 rounded font-semibold border border-indigo-500/30">2. Code Example</span>
          <pre class="relative group"><code class="language-javascript block p-4 bg-[#090c15] border border-gray-800 rounded-lg text-indigo-300 font-mono text-sm">// Find slide image shapes
const images = ppt.getImages();
console.log(images); 
// Output: [ { id: 'Picture 1', name: 'logo' } ]

// Substitute image binary
await ppt.replaceImage('logo', 'path/to/new_logo.png');

// Alternatively, insert a brand new image
await ppt.addImage('path/to/badge.png', {
  x: 1.5,       // Position X (inches)
  y: 2.0,       // Position Y (inches)
  width: 2.5,   // Width (inches)
  height: 2.5   // Height (inches)
});</code><button class="copy-btn absolute top-3 right-3 text-xs bg-gray-800 text-gray-300 hover:text-white px-2 py-1 rounded transition-colors">Copy</button></pre>
        </div>
      </section>

      <!-- Slide Management Section -->
      <section id="slide-management" class="doc-section space-y-6">
        <h1 class="font-title text-4xl font-extrabold text-white">Slide Duplication & CRUD</h1>
        <p class="text-gray-300">Assemble presentations dynamically by copying template layout slides and reordering them as needed.</p>
        
        <pre class="relative group"><code class="language-javascript block p-4 bg-[#090c15] border border-gray-800 rounded-lg text-indigo-300 font-mono text-sm">// Duplicate slide 1 to build a second slide at index 2
ppt.duplicateSlide(1, 2);

// Reorder slide 3 to the front of the presentation
ppt.moveSlide(3, 1);

// Delete temporary slide
ppt.deleteSlide(4);</code><button class="copy-btn absolute top-3 right-3 text-xs bg-gray-800 text-gray-300 hover:text-white px-2 py-1 rounded transition-colors">Copy</button></pre>
      </section>

      <!-- Slide Import Section -->
      <section id="slide-import" class="doc-section space-y-6">
        <h1 class="font-title text-4xl font-extrabold text-white">Cross-Template Imports</h1>
        <p class="text-gray-300">Deep-copy a slide from an external presentation template into your active target presentation.</p>
        
        <div class="p-4 bg-brand-500/10 border-l-4 border-brand-500 rounded-r-xl text-sm text-gray-300">
          <strong>📦 Automated Media & Layout Remapping:</strong> Direct file copying corrupts slides because layouts, masters, images, and relationships are not linked in the destination namespace. <code>importSlideFrom</code> automatically resolves, copies, and maps these slide assets, deduplicating image files in the target archive to save disk space.
        </div>

        <pre class="relative group"><code class="language-javascript block p-4 bg-[#090c15] border border-gray-800 rounded-lg text-indigo-300 font-mono text-sm">const targetPresentation = await PPTXTemplater.load('corporate_base.pptx');
const marketingSource = await PPTXTemplater.load('marketing_campaign.pptx');

// Deep import slide 3 of marketing deck into the base presentation
await targetPresentation.importSlideFrom(marketingSource, 3);

await targetPresentation.saveToFile('merged_corporate_presentation.pptx');</code><button class="copy-btn absolute top-3 right-3 text-xs bg-gray-800 text-gray-300 hover:text-white px-2 py-1 rounded transition-colors">Copy</button></pre>
      </section>

      <!-- Table Merging Section -->
      <section id="table-merging" class="doc-section space-y-6">
        <h1 class="font-title text-4xl font-extrabold text-white">Table Cell Merging</h1>
        
        <div>
          <span class="bg-indigo-500/10 text-indigo-400 text-xs px-2.5 py-1 rounded font-semibold border border-indigo-500/30">1. Beginner Explanation</span>
          <p class="text-sm text-gray-400 mt-2">
            In PowerPoint, merging cells requires managing strict table grids. 
            The top-left cell is designated as the <strong>origin cell</strong> and holds <code>gridSpan</code> (columns spanned) and <code>rowSpan</code> (rows spanned) attributes.
            The remaining cells in the merged region are **shadowed cells** and must contain the attributes <code>hMerge="1"</code> and <code>vMerge="1"</code>. 
            Our Table engine manages these attributes automatically, moving merged text cleanly into the origin cell.
          </p>
        </div>

        <div>
          <span class="bg-indigo-500/10 text-indigo-400 text-xs px-2.5 py-1 rounded font-semibold border border-indigo-500/30">2. Visual Example</span>
          <div class="p-4 bg-[#131b2e] border border-gray-800 rounded-lg font-mono text-xs text-indigo-300 mt-2">
            <pre>
Before Merging:
┌───────────┬───────────┐
│ Row 1 C1  │ Row 1 C2  │
├───────────┼───────────┤
│ Row 2 C1  │ Row 2 C2  │
└───────────┴───────────┘

After Merging (Row 1-2, Col 1-2):
┌───────────────────────┐
│ Row 1 C1 \n Row 1 C2   │
│ Row 2 C1 \n Row 2 C2   │ (Single Merged Cell)
└───────────────────────┘
            </pre>
          </div>
        </div>

        <div>
          <span class="bg-indigo-500/10 text-indigo-400 text-xs px-2.5 py-1 rounded font-semibold border border-indigo-500/30">3. Code Example</span>
          <pre class="relative group"><code class="language-javascript block p-4 bg-[#090c15] border border-gray-800 rounded-lg text-indigo-300 font-mono text-sm">// Populate dynamic table data
ppt.updateTable('sales-table', [
  ['Category', 'Actual', 'Growth', 'Rating'],
  ['Engineering', '140k', '12%', { value: 'Good', align: 'ctr', fill: '10b981' }],
  ['Marketing', '110k', '8%', { value: 'Review', align: 'ctr', fill: 'f59e0b' }]
]);

// Merge Row 1 Column 1 through Row 2 Column 2
ppt.mergeCells({
  tableId: 'sales-table',
  startRow: 1,
  startCol: 1,
  endRow: 2,
  endCol: 2
});

// Unmerge cells back to normal
ppt.unmergeCells({
  tableId: 'sales-table',
  row: 1,
  col: 1
});</code><button class="copy-btn absolute top-3 right-3 text-xs bg-gray-800 text-gray-300 hover:text-white px-2 py-1 rounded transition-colors">Copy</button></pre>
        </div>

        <div>
          <span class="bg-indigo-500/10 text-indigo-400 text-xs px-2.5 py-1 rounded font-semibold border border-indigo-500/30">4. Template-driven Cell Merging</span>
          <p class="text-sm text-gray-400 mt-2">
            You can configure merges inline inside the <code>updateTable</code> parameters by supplying rowSpan / colSpan objects or using a config array:
          </p>
          <pre class="relative group"><code class="language-javascript block p-4 bg-[#090c15] border border-gray-800 rounded-lg text-indigo-300 font-mono text-sm">ppt.updateTable('sales-table', {
  rows: [
    ['Header 1', 'Header 2', 'Header 3'],
    ['Normal Cell', { value: 'Column Merge Cell', colSpan: 2 }],
    ['Normal Cell', 'Normal Cell', { value: 'Row Merge Cell', rowSpan: 2 }]
  ],
  merge: [
    { startRow: 0, startCol: 0, endRow: 0, endCol: 2 }
  ]
});</code><button class="copy-btn absolute top-3 right-3 text-xs bg-gray-800 text-gray-300 hover:text-white px-2 py-1 rounded transition-colors">Copy</button></pre>
        </div>
      </section>

      <!-- Chart Engine Section -->
      <section id="chart-engine" class="doc-section space-y-6">
        <h1 class="font-title text-4xl font-extrabold text-white">Interactive Chart Updates</h1>
        
        <div>
          <span class="bg-indigo-500/10 text-indigo-400 text-xs px-2.5 py-1 rounded font-semibold border border-indigo-500/30">1. Beginner Explanation</span>
          <p class="text-sm text-gray-400 mt-2">
            PowerPoint charts are backed by an XML description and an embedded binary Excel sheet. 
            If you only update the slide XML, PowerPoint will render a cached image until clicked. 
            Our Chart Engine updates both the slide XML data points (<code>strCache</code>, <code>numCache</code>) and cell records in the Excel spreadsheet binary.
            This ensures your charts refresh instantly without alerts.
          </p>
        </div>

        <div>
          <span class="bg-indigo-500/10 text-indigo-400 text-xs px-2.5 py-1 rounded font-semibold border border-indigo-500/30">2. Code Example</span>
          <pre class="relative group"><code class="language-javascript block p-4 bg-[#090c15] border border-gray-800 rounded-lg text-indigo-300 font-mono text-sm">ppt.updateChartData('revenue-chart', {
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'Target Revenue', values: [100, 120, 140, 160] },
    { name: 'Actual Revenue', values: [105, 118, 145, 172] }
  ]
});</code><button class="copy-btn absolute top-3 right-3 text-xs bg-gray-800 text-gray-300 hover:text-white px-2 py-1 rounded transition-colors">Copy</button></pre>
        </div>
      </section>

      <!-- Action Hyperlinks Section -->
      <section id="navigation-links" class="doc-section space-y-6">
        <h1 class="font-title text-4xl font-extrabold text-white">Action Hyperlinks</h1>
        <p class="text-gray-300">Set slide action jumps, URL links, or custom paths on shapes and text blocks in your deck.</p>
        
        <pre class="relative group"><code class="language-javascript block p-4 bg-[#090c15] border border-gray-800 rounded-lg text-indigo-300 font-mono text-sm">// Make matching text link to the next slide
ppt.addTextNavigationLink({
  element: 'Proceed to summary',
  action: 'next' // 'next', 'previous', 'first', 'last'
});

// Bind navigation target to a vector button shape
ppt.addShapeNavigationLink({
  shapeId: 'CloseButton',
  action: 'last'
});</code><button class="copy-btn absolute top-3 right-3 text-xs bg-gray-800 text-gray-300 hover:text-white px-2 py-1 rounded transition-colors">Copy</button></pre>
      </section>

      <!-- Packaging & XML Structure Section -->
      <section id="openxml-internals" class="doc-section space-y-6">
        <h1 class="font-title text-4xl font-extrabold text-white">Packaging & XML Structure</h1>
        <p class="text-gray-300">A detailed lookup of standard OPC PPTX directories and where components are located.</p>
        
        <div class="architecture-diagram p-4 bg-[#131b2e] border border-gray-800 rounded-lg font-mono text-xs text-indigo-300">
          <pre>
PPTX Archive structure:
├── [Content_Types].xml (MIME type override registry)
├── _rels/
│   └── .rels (Root relations maps)
├── ppt/
│   ├── presentation.xml (Slide inventory records)
│   ├── _rels/
│   │   └── presentation.xml.rels (Global assets layout links)
│   ├── slides/
│   │   ├── slide1.xml (Slide 1 canvas elements)
│   │   └── _rels/
│   │       └── slide1.xml.rels (Slide-level assets references)
│   ├── slideLayouts/ (Slide theme layout components)
│   ├── slideMasters/ (Slide presentation master masters)
│   ├── media/ (Embedded image assets: PNG, JPEG, SVG)
│   └── embeddings/ (Chart worksheets: Excel documents)
          </pre>
        </div>
      </section>

      <!-- Managers & Core Validation Section -->
      <section id="manifest-managers" class="doc-section space-y-6">
        <h1 class="font-title text-4xl font-extrabold text-white">Managers & Validation</h1>
        <p class="text-gray-300">How the templating engine protects presentation validity using internal validation managers.</p>
        
        <div class="grid grid-cols-1 md:grid-cols-2 gap-6 mt-4">
          <div class="p-5 bg-[#131b2e] border border-gray-800 rounded-xl">
            <h4 class="font-semibold text-white mb-2">🛡️ Integrity Validation Engine</h4>
            <p class="text-sm text-gray-400">Verifies XML structure, checks table grids, and remaps assets relationship keys before writing output directories.</p>
          </div>
          <div class="p-5 bg-[#131b2e] border border-gray-800 rounded-xl">
            <h4 class="font-semibold text-white mb-2">📁 ContentTypesManager</h4>
            <p class="text-sm text-gray-400">Maintains <code>[Content_Types].xml</code> Override nodes, making sure new assets have valid XML mime headers.</p>
          </div>
          <div class="p-5 bg-[#131b2e] border border-gray-800 rounded-xl">
            <h4 class="font-semibold text-white mb-2">🔗 RelationshipManager</h4>
            <p class="text-sm text-gray-400">Manages relative slide links and re-maps relationship indexes (<code>.rels</code> files) dynamically.</p>
          </div>
          <div class="p-5 bg-[#131b2e] border border-gray-800 rounded-xl">
            <h4 class="font-semibold text-white mb-2">📝 SlideManager</h4>
            <p class="text-sm text-gray-400">Caches slide markup buffers for fast reads/writes, updating the master slides catalog.</p>
          </div>
        </div>
      </section>

      <!-- Troubleshooting Section -->
      <section id="troubleshooting" class="doc-section space-y-6">
        <h1 class="font-title text-4xl font-extrabold text-white">Troubleshooting & FAQs</h1>
        
        <div class="space-y-4">
          <div class="p-5 bg-[#131b2e] border border-gray-800 rounded-xl">
            <h4 class="font-semibold text-white mb-2">Q: Why does Microsoft PowerPoint display "Repair content" on open?</h4>
            <p class="text-sm text-gray-400">
              This typically happens if slide assets relationship indexes (e.g. mapping of a chart or image) are missing, duplicate rowIds exist, or new XML content is missing from <code>[Content_Types].xml</code>. 
              Always use the library's built-in <code>saveToFile()</code> or <code>toBuffer()</code> methods, which run validation passes and automatically resolve these issues.
            </p>
          </div>

          <div class="p-5 bg-[#131b2e] border border-gray-800 rounded-xl">
            <h4 class="font-semibold text-white mb-2">Q: My text placeholders are not getting replaced. What's wrong?</h4>
            <p class="text-sm text-gray-400">
              PowerPoint often splits run nodes (e.g. <code>{{name}}</code> is split in slide XML into separate tags like <code>{{na</code> and <code>me}}</code>). 
              To fix this, cut the placeholder text in PowerPoint, and paste it back using "Keep Text Only" to merge the XML runs.
            </p>
          </div>
        </div>
      </section>

    </main>
  </div>

  <script src="app.js"></script>
</body>
</html>
`;

const CSS_CONTENT = `/* Variables & Custom Themes */
:root {
  --bg-primary: #0b0f19;
  --bg-secondary: #131b2e;
  --border-color: rgba(255, 255, 255, 0.08);
  --text-main: #f3f4f6;
  --text-muted: #9ca3af;
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

/* Glassmorphism sidebar & header */
#sidebar {
  background-color: rgba(19, 27, 46, 0.6);
  backdrop-filter: blur(16px);
}
.light-theme #sidebar {
  background-color: rgba(255, 255, 255, 0.6);
}

/* Layout transition styles */
body {
  background-color: var(--bg-primary);
  color: var(--text-main);
}

.nav-item.active {
  color: var(--text-main) !important;
  background-color: rgba(99, 102, 241, 0.15) !important;
  border-left: 3px solid var(--primary);
  font-weight: 500;
}

.doc-section {
  display: none;
}
.doc-section.active-section {
  display: block;
  animation: fadeIn 0.3s ease-out;
}

@keyframes fadeIn {
  from { opacity: 0; transform: translateY(8px); }
  to { opacity: 1; transform: translateY(0); }
}
`;

const JS_CONTENT = `document.addEventListener('DOMContentLoaded', () => {
  // Navigation Router
  const navItems = document.querySelectorAll('.nav-item');
  const sections = document.querySelectorAll('.doc-section');

  function showSection(targetId) {
    let cleanId = targetId.replace('#', '');
    sections.forEach(s => s.classList.remove('active-section'));
    navItems.forEach(n => n.classList.remove('active'));

    const targetSection = document.getElementById(cleanId);
    if (targetSection) {
      targetSection.classList.add('active-section');
      window.scrollTo({ top: 0, behavior: 'smooth' });
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
  }

  // Theme Toggler
  const themeToggle = document.getElementById('theme-toggle');
  const body = document.body;
  const savedTheme = localStorage.getItem('theme');

  if (savedTheme === 'light') {
    body.classList.remove('dark-theme');
    body.classList.add('light-theme', 'bg-white', 'text-slate-800');
    body.classList.remove('bg-[#0b0f19]', 'text-gray-200');
    themeToggle.querySelector('.toggle-icon').textContent = '☀️';
  }

  themeToggle.addEventListener('click', () => {
    if (body.classList.contains('light-theme')) {
      body.classList.remove('light-theme', 'bg-white', 'text-slate-800');
      body.classList.add('bg-[#0b0f19]', 'text-gray-200');
      themeToggle.querySelector('.toggle-icon').textContent = '🌙';
      localStorage.setItem('theme', 'dark');
    } else {
      body.classList.add('light-theme', 'bg-white', 'text-slate-800');
      body.classList.remove('bg-[#0b0f19]', 'text-gray-200');
      themeToggle.querySelector('.toggle-icon').textContent = '☀️';
      localStorage.setItem('theme', 'light');
    }
  });

  // Client Search Indexer
  const searchInput = document.getElementById('doc-search');
  searchInput.addEventListener('input', (e) => {
    const query = e.target.value.toLowerCase().trim();
    if (!query) {
      navItems.forEach(item => item.parentElement.style.display = 'block');
      return;
    }

    navItems.forEach(item => {
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

  // Clipboard Copier
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

// Learning Path selector
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
`;

async function main() {
  await fsExtra.ensureDir(DOCS_DIR);
  await fsExtra.writeFile(resolve(DOCS_DIR, 'index.html'), HTML_CONTENT);
  await fsExtra.writeFile(resolve(DOCS_DIR, 'style.css'), CSS_CONTENT);
  await fsExtra.writeFile(resolve(DOCS_DIR, 'app.js'), JS_CONTENT);
  
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
  await fsExtra.writeFile(resolve(DOCS_DIR, 'sitemap.xml'), sitemap.trim());

  // SEO Robots.txt
  const robots = `User-agent: *
Allow: /
Sitemap: https://jsuyog2.github.io/node-pptx-templater/sitemap.xml`;
  await fsExtra.writeFile(resolve(DOCS_DIR, 'robots.txt'), robots.trim());

  console.log('Successfully built GitHub Pages documentation with sitemap and robots.txt under docs/');
}

main().catch(err => {
  console.error('Error building docs:', err);
  process.exit(1);
});
