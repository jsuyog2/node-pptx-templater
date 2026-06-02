const fsExtra = require('fs-extra');
const { resolve } = require('path');

const DOCS_DIR = resolve(__dirname, '../docs');

const HTML_CONTENT = `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>node-pptx-templater — Low-Level OpenXML Engine</title>
  <!-- Google Fonts -->
  <link rel="preconnect" href="https://fonts.googleapis.com">
  <link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
  <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=Outfit:wght@400;500;600;700;800&family=JetBrains+Mono:wght@400;500&display=swap" rel="stylesheet">
  <link rel="stylesheet" href="style.css">
</head>
<body class="dark-theme">
  <!-- Top Navigation Header -->
  <header class="top-nav">
    <div class="logo-container">
      <span class="logo-icon">📊</span>
      <span class="logo-text">node-pptx-templater</span>
      <span class="logo-badge">v1.0.0</span>
    </div>
    <div class="search-container">
      <input type="text" id="doc-search" placeholder="Search docs (Ctrl + K)..." autocomplete="off">
    </div>
    <div class="header-actions">
      <button id="theme-toggle" aria-label="Toggle Theme">
        <span class="toggle-icon">🌙</span>
      </button>
      <a href="https://github.com/jsuyog2/node-pptx-templater" class="github-link" target="_blank" rel="noopener">
        <svg height="24" viewBox="0 0 16 16" width="24" fill="currentColor">
          <path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38 0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13-.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66.07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15-.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0 1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82 1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01 1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"></path>
        </svg>
      </a>
    </div>
  </header>

  <!-- Sidebar & Main Content Wrapper -->
  <div class="container">
    <!-- Left Navigation Sidebar -->
    <aside class="sidebar" id="sidebar">
      <nav class="sidebar-nav">
        <div class="nav-section">
          <h3>Getting Started</h3>
          <ul>
            <li><a href="#introduction" class="nav-item active">Introduction</a></li>
            <li><a href="#installation" class="nav-item">Installation</a></li>
            <li><a href="#quickstart" class="nav-item">Quick Start</a></li>
          </ul>
        </div>
        <div class="nav-section">
          <h3>Core Modules</h3>
          <ul>
            <li><a href="#slide-import" class="nav-item">Slide Import Engine</a></li>
            <li><a href="#navigation-hyperlinks" class="nav-item">Navigation Hyperlinks</a></li>
            <li><a href="#content-types-manager" class="nav-item">ContentTypesManager</a></li>
            <li><a href="#relationship-manager" class="nav-item">RelationshipManager</a></li>
          </ul>
        </div>
        <div class="nav-section">
          <h3>API Abstractions</h3>
          <ul>
            <li><a href="#pptx-templater-api" class="nav-item">PPTXTemplater Class</a></li>
            <li><a href="#chart-engine" class="nav-item">Chart Engine</a></li>
            <li><a href="#table-manager" class="nav-item">Table Engine</a></li>
          </ul>
        </div>
        <div class="nav-section">
          <h3>Guides</h3>
          <ul>
            <li><a href="#openxml-architecture" class="nav-item">OpenXML Architecture</a></li>
            <li><a href="#troubleshooting" class="nav-item">Troubleshooting</a></li>
          </ul>
        </div>
      </nav>
    </aside>

    <!-- Main Content Area -->
    <main class="content-area">
      <!-- Introduction Section -->
      <section id="introduction" class="doc-section active-section">
        <h1>Introduction</h1>
        <p class="lead">A low-level, high-performance PowerPoint OpenXML templating engine for Node.js. It generates and edits PPTX files directly through structured XML manipulation, avoiding fragile regex replacements and eliminating standard "Repair Mode" issues.</p>
        
        <div class="alert note">
          <strong>⚡ Why node-pptx-templater?</strong> It runs on pure Javascript/ESM with zero native PPTX generator dependencies, making it ultra-lightweight and ideal for high-throughput server environments or serverless runtimes.
        </div>

        <h2>Features At A Glance</h2>
        <div class="features-grid">
          <div class="feature-card">
            <h4>📦 Structured XML Parsing</h4>
            <p>Every file modification uses a fast, reliable DOM-like abstraction instead of raw text replacement.</p>
          </div>
          <div class="feature-card">
            <h4>🔄 In-Memory Manifests</h4>
            <p>Guards packaging integrity via atomic serialization of relationships and content types.</p>
          </div>
          <div class="feature-card">
            <h4>📋 Deep Slide Imports</h4>
            <p>Clones slides across templates, preserving all associated layouts, charts, links, and styling.</p>
          </div>
          <div class="feature-card">
            <h4>📊 Accurate Chart Updates</h4>
            <p>Synchronizes chart XML and underlying data caches directly inside embedded Excel sheets.</p>
          </div>
        </div>
      </section>

      <!-- Installation Section -->
      <section id="installation" class="doc-section">
        <h1>Installation</h1>
        <p>Install the library via npm. Make sure you meet the requirement of having Node.js version 18.0.0 or higher.</p>
        <pre><code class="language-bash">npm install node-pptx-templater</code></pre>
        <h3>System Requirements</h3>
        <ul>
          <li><strong>Node.js:</strong> >= 18.0.0</li>
          <li><strong>Module system:</strong> ES Modules (<code>"type": "module"</code> in package.json)</li>
        </ul>
      </section>

      <!-- Quick Start Section -->
      <section id="quickstart" class="doc-section">
        <h1>Quick Start</h1>
        <p>Get up and running in under 60 seconds with this simple template rendering example.</p>
        <pre><code class="language-javascript">import { PPTXTemplater } from 'node-pptx-templater';

// Load the source template
const ppt = await PPTXTemplater.load('template.pptx');

// Select slide 1 and execute operations
ppt.useSlide(1);

// Safely replace text placeholders
ppt.replaceText({
  '{{title}}': 'Quarterly Earnings Report',
  '{{date}}': 'May 2026'
});

// Update an interactive bar chart
ppt.updateChart('sales-chart', {
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'Revenue ($M)', values: [120, 145, 170, 210] }
  ]
});

// Save non-corrupted PPTX to disk
await ppt.saveToFile('./output/annual_report.pptx');</code></pre>
      </section>

      <!-- Slide Import Section -->
      <section id="slide-import" class="doc-section">
        <h1>Slide Import Engine</h1>
        <p>The Slide Import engine allows robust deep copying of individual slides from one template into another.</p>
        <p>Unlike basic copy implementations, <code>importSlideFrom</code> performs an exhaustive search and deep-copy of all associated relationships, slide layouts, embedded media, and chart workbooks, remapping them into the new presentation's namespace.</p>
        
        <h2>Example: Merging Slide Decks</h2>
        <pre><code class="language-javascript">import { PPTXTemplater } from 'node-pptx-templater';

const master = await PPTXTemplater.load('master_template.pptx');
const source = await PPTXTemplater.load('department_results.pptx');

// Deep-copy slide 2 from department_results into master template
await master.importSlideFrom(source, 2);

// Save the combined deck
await master.saveToFile('merged_presentation.pptx');</code></pre>

        <div class="alert tip">
          <strong>Pro-tip:</strong> When importing a slide, media deduplication automatically checks if the image is already present in the destination deck, preventing asset bloating!
        </div>
      </section>

      <!-- Navigation Hyperlinks Section -->
      <section id="navigation-hyperlinks" class="doc-section">
        <h1>Navigation Hyperlinks</h1>
        <p>Link slides internally or attach special navigation actions directly to text runs or vector shapes using our clean API.</p>
        
        <h2>Adding Text Navigation Links</h2>
        <p>Make matching text clickable to jump directly to standard navigation targets (<code>next</code>, <code>previous</code>, <code>first</code>, or <code>last</code> slide).</p>
        <pre><code class="language-javascript">ppt.addTextNavigationLink({
  slide: 1,
  element: 'Go to Next Slide',
  action: 'next'
});</code></pre>

        <h2>Adding Shape Navigation Links</h2>
        <p>Bind custom navigation actions to vector shapes or button graphics inside your template.</p>
        <pre><code class="language-javascript">ppt.addShapeNavigationLink({
  slide: 1,
  shapeId: 'Back Button',
  action: 'previous'
});</code></pre>
      </section>

      <!-- ContentTypesManager Section -->
      <section id="content-types-manager" class="doc-section">
        <h1>ContentTypesManager</h1>
        <p>Every single file inside an OpenXML ZIP package must have its type registered inside the centralized manifest <code>[Content_Types].xml</code>.</p>
        <p>Our centralized <code>ContentTypesManager</code> replaces standard fragile string replacements with a parsed in-memory node-tree representation of the manifest. Changes are flushed only during the final package writing step to completely eliminate race conditions.</p>
        
        <h2>API Reference</h2>
        <pre><code class="language-javascript">// Inside PPTXTemplater lifecycle:
const contentTypes = ppt.contentTypesManager;

// Override registration for new slide XML
contentTypes.addOverride('/ppt/slides/slide3.xml', 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml');

// Default MIME registration
contentTypes.addDefault('webp', 'image/webp');</code></pre>
      </section>

      <!-- RelationshipManager Section -->
      <section id="relationship-manager" class="doc-section">
        <h1>RelationshipManager</h1>
        <p>Relationships define how slides link to layouts, masters, charts, and media assets. Standard regex tools corrupt relationship chains if IDs are not fully unique.</p>
        <p>The <code>RelationshipManager</code> provides bulletproof remapping and creation of relative relationship paths inside standard <code>.rels</code> files.</p>
        
        <h2>Usage Example</h2>
        <pre><code class="language-javascript">const rels = ppt.relationshipManager;

// Add a safe relationship to slides/slide1.xml
const rId = rels.addRelationship(
  'ppt/slides/slide1.xml',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  '../media/image10.png'
);
// Returns unique ID (e.g. 'rId5')</code></pre>
      </section>

      <!-- PPTXTemplater API Section -->
      <section id="pptx-templater-api" class="doc-section">
        <h1>PPTXTemplater</h1>
        <p>The main class exposed by the library to load, select, modify, and save presentation templates.</p>
        
        <h2>Properties & Getters</h2>
        <table class="api-table">
          <thead>
            <tr>
              <th>Property</th>
              <th>Type</th>
              <th>Description</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><code>slideCount</code></td>
              <td><code>number</code></td>
              <td>Returns the total count of active slides in the loaded deck.</td>
            </tr>
            <tr>
              <td><code>relationshipManager</code></td>
              <td><code>RelationshipManager</code></td>
              <td>Access direct relationship operations.</td>
            </tr>
            <tr>
              <td><code>contentTypesManager</code></td>
              <td><code>ContentTypesManager</code></td>
              <td>Access manifest definitions.</td>
            </tr>
          </tbody>
        </table>

        <h2>Main Methods</h2>
        <table class="api-table">
          <thead>
            <tr>
              <th>Method</th>
              <th>Arguments</th>
              <th>Returns</th>
              <th>Description</th>
            </tr>
          </thead>
          <tbody>
            <tr>
              <td><code>useSlide</code></td>
              <td><code>(slideNumberOrTag)</code></td>
              <td><code>PPTXTemplater</code></td>
              <td>Restricts text, table, and chart updates to only target specific slides.</td>
            </tr>
            <tr>
              <td><code>replaceText</code></td>
              <td><code>(replacements)</code></td>
              <td><code>PPTXTemplater</code></td>
              <td>Performs structured substitution on text, handles run fragmentation cleanly.</td>
            </tr>
            <tr>
              <td><code>saveToFile</code></td>
              <td><code>(filePath)</code></td>
              <td><code>Promise&lt;void&gt;</code></td>
              <td>Saves all dirty cached XML parts and updates metadata counts before writing.</td>
            </tr>
          </tbody>
        </table>
      </section>

      <!-- Chart Engine Section -->
      <section id="chart-engine" class="doc-section">
        <h1>Chart Engine</h1>
        <p>PowerPoint charts are structured via complex pairings of a Chart XML part and an embedded Excel Worksheet binary containing matching row/column data cells.</p>
        <p>Our Chart Engine automatically updates the underlying series formulas, Category axis values, and regenerates both <code>strCache</code> and <code>numCache</code> so Google Slides and LibreOffice display charts instantly without requiring user-refresh triggers.</p>
        
        <h2>Updating Chart Data</h2>
        <pre><code class="language-javascript">ppt.updateChart('revenue-chart', {
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [
    { name: 'Target', values: [100, 120, 140, 160] },
    { name: 'Actual', values: [105, 118, 145, 172] }
  ]
});</code></pre>
      </section>

      <!-- Table Manager Section -->
      <section id="table-manager" class="doc-section">
        <h1>Table Engine</h1>
        <p>Dynamically scale template tables. The Table Engine lets you replace placeholder rows while fully preserving the styles, fonts, backgrounds, and cell heights of the original template row.</p>

        <h2>Cell Merging & Unmerging</h2>
        <p>PowerPoint DrawingML tables require strict grid adherence. For cell merging, the top-left cell acts as the <strong>origin</strong> (carrying <code>gridSpan</code> and <code>rowSpan</code> attributes) while other cells in the merge region are <strong>shadowed</strong> (carrying <code>hMerge</code> and <code>vMerge</code> flags). <code>node-pptx-templater</code> provides full compatibility with PowerPoint, Google Slides, and LibreOffice.</p>

        <h3>1. Direct Cell Merging</h3>
        <p>Merge horizontal columns, vertical rows, or a rectangular block using a simple configuration object:</p>
        <pre><code class="language-javascript">// Merge rows 1-3 and columns 1-3 in 'sales-table'
ppt.mergeCells({
  slide: 1,
  tableId: 'sales-table',
  startRow: 1,
  startCol: 1,
  endRow: 3,
  endCol: 3
});</code></pre>

        <h3>2. Unmerging / Splitting Cells</h3>
        <p>Split a merged region back to its base components. You can target the region containing a specific cell coordinate:</p>
        <pre><code class="language-javascript">// Split the merged block containing cell (2, 2)
ppt.unmergeCells({
  slide: 1,
  tableId: 'sales-table',
  row: 2,
  col: 2
});</code></pre>

        <h3>3. Template-driven merges</h3>
        <p>You can define cell merging dynamically during <code>updateTable</code> by passing objects with <code>colSpan</code> / <code>rowSpan</code> properties or providing a <code>merge</code> array configuration:</p>
        <pre><code class="language-javascript">ppt.updateTable('sales-table', {
  rows: [
    ['Header 1', 'Header 2', 'Header 3'],
    ['Normal Cell', { value: 'Spanned Cell', colSpan: 2 }],
    ['Normal Cell', 'Normal Cell', { value: 'Spanned Row', rowSpan: 2 }]
  ],
  merge: [
    { startRow: 0, startCol: 0, endRow: 0, endCol: 2 }
  ]
});</code></pre>
      </section>

      <!-- OpenXML Architecture Section -->
      <section id="openxml-architecture" class="doc-section">
        <h1>OpenXML Architecture</h1>
        <p>An inside look at how files are packaged and managed within the PPTX OpenXML structure.</p>
        <div class="architecture-diagram">
          <pre>
            ┌───────────────────────────────────────────────┐
            │                 PPTXTemplater                 │
            │                  (Public API)                 │
            └────────┬──────────────────────────────┬───────┘
                     │                              │
            ┌────────▼────────┐            ┌────────▼────────┐
            │   SlideManager  │            │   ChartManager  │
            │  (discovery &amp;   │            │   (Excel &amp; cache│
            │   slide CRUD)   │            │    synchronizer)│
            └────────┬────────┘            └────────┬────────┘
                     │                              │
            ┌────────▼──────────────────────────────▼────────┐
            │                 XMLParser                     │
            │         (fast-xml-parser parsing tree)         │
            └────────┬──────────────────────────────┬───────┘
                     │                              │
            ┌────────▼────────┐            ┌────────▼────────┐
            │ ContentTypesMgr │            │ RelationshipMgr │
            │ (Overrides list)│            │ (.rels mapper)  │
            └────────┬────────┘            └────────┬────────┘
                     │                              │
            ┌────────▼──────────────────────────────▼────────┐
            │                 ZipManager                    │
            │            (JSZip raw serialization)          │
            └───────────────────────────────────────────────┘
          </pre>
        </div>
      </section>

      <!-- Troubleshooting Section -->
      <section id="troubleshooting" class="doc-section">
        <h1>Troubleshooting</h1>
        <h3>My presentation shows "PowerPoint found a problem with content..."</h3>
        <p>This is commonly caused by out-of-sync slide count metadata in <code>docProps/app.xml</code> or an override registration missing from <code>[Content_Types].xml</code>.</p>
        <div class="alert note">
          <strong>The Solution:</strong> Always make sure to use <code>ppt.saveToFile()</code> or <code>ppt.toBuffer()</code>, which automatically run metadata-synchronization passes and flush all manifest changes.
        </div>

        <h3>My placeholders aren't replacing!</h3>
        <p>PowerPoint split-runs break placeholders like <code>{{placeholder}}</code> into multiple separate <code>&lt;a:t&gt;</code> tags. Enable logging output to identify where fragmentation occurs:</p>
        <pre><code class="language-bash">PPTX_LOG_LEVEL=debug node script.js</code></pre>
      </section>
    </main>
  </div>

  <script src="app.js"></script>
</body>
</html>
`;

const CSS_CONTENT = `/* Variables & Themes */
:root {
  --font-title: 'Outfit', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  --font-body: 'Inter', -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
  --font-mono: 'JetBrains Mono', Courier, monospace;

  /* Premium Dark Palette */
  --bg-primary: #0b0f19;
  --bg-secondary: #131b2e;
  --bg-card: #1c2742;
  --border-color: rgba(255, 255, 255, 0.08);
  --text-main: #f3f4f6;
  --text-muted: #9ca3af;
  --primary: #6366f1;
  --primary-glow: rgba(99, 102, 241, 0.15);
  --accent: #a855f7;
  --accent-glow: rgba(168, 85, 247, 0.15);
  --success: #10b981;
  --success-glow: rgba(16, 185, 129, 0.1);
  --code-bg: #090c15;
}

.light-theme {
  /* Premium Light Palette */
  --bg-primary: #f8fafc;
  --bg-secondary: #ffffff;
  --bg-card: #f1f5f9;
  --border-color: rgba(0, 0, 0, 0.06);
  --text-main: #0f172a;
  --text-muted: #64748b;
  --primary: #4f46e5;
  --primary-glow: rgba(79, 70, 229, 0.08);
  --accent: #8b5cf6;
  --accent-glow: rgba(139, 92, 246, 0.08);
  --success: #059669;
  --success-glow: rgba(5, 150, 105, 0.08);
  --code-bg: #0f172a;
}

/* Global Reset */
* {
  box-sizing: border-box;
  margin: 0;
  padding: 0;
  transition: background-color 0.3s ease, border-color 0.3s ease, color 0.2s ease;
}

body {
  background-color: var(--bg-primary);
  color: var(--text-main);
  font-family: var(--font-body);
  line-height: 1.6;
  -webkit-font-smoothing: antialiased;
}

/* Header Styling */
.top-nav {
  position: fixed;
  top: 0;
  left: 0;
  right: 0;
  height: 70px;
  background-color: rgba(19, 27, 46, 0.8);
  backdrop-filter: blur(12px);
  border-bottom: 1px solid var(--border-color);
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 0 30px;
  z-index: 100;
}

.light-theme .top-nav {
  background-color: rgba(255, 255, 255, 0.85);
}

.logo-container {
  display: flex;
  align-items: center;
  gap: 10px;
}

.logo-icon {
  font-size: 24px;
}

.logo-text {
  font-family: var(--font-title);
  font-weight: 700;
  font-size: 20px;
  background: linear-gradient(135deg, var(--primary), var(--accent));
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
}

.logo-badge {
  background: var(--primary-glow);
  color: var(--primary);
  font-size: 11px;
  padding: 2px 8px;
  border-radius: 20px;
  font-weight: 600;
  border: 1px solid rgba(99, 102, 241, 0.3);
}

.search-container input {
  width: 320px;
  padding: 10px 16px;
  background-color: var(--bg-primary);
  border: 1px solid var(--border-color);
  border-radius: 30px;
  color: var(--text-main);
  font-size: 14px;
  outline: none;
}

.search-container input:focus {
  border-color: var(--primary);
  box-shadow: 0 0 10px var(--primary-glow);
}

.header-actions {
  display: flex;
  align-items: center;
  gap: 20px;
}

#theme-toggle {
  background: none;
  border: none;
  cursor: pointer;
  font-size: 20px;
  color: var(--text-main);
}

.github-link {
  color: var(--text-main);
  opacity: 0.8;
}

.github-link:hover {
  opacity: 1;
}

/* Sidebar Styling */
.container {
  display: flex;
  margin-top: 70px;
  min-height: calc(100vh - 70px);
}

.sidebar {
  width: 280px;
  background-color: var(--bg-secondary);
  border-right: 1px solid var(--border-color);
  position: fixed;
  top: 70px;
  bottom: 0;
  left: 0;
  overflow-y: auto;
  padding: 30px 20px;
  z-index: 90;
}

.sidebar-nav {
  display: flex;
  flex-direction: column;
  gap: 25px;
}

.nav-section h3 {
  font-family: var(--font-title);
  font-size: 12px;
  text-transform: uppercase;
  color: var(--text-muted);
  letter-spacing: 0.1em;
  margin-bottom: 12px;
}

.nav-section ul {
  list-style: none;
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.nav-item {
  display: block;
  padding: 8px 12px;
  color: var(--text-muted);
  text-decoration: none;
  font-size: 14px;
  border-radius: 8px;
}

.nav-item:hover {
  color: var(--text-main);
  background-color: var(--border-color);
}

.nav-item.active {
  color: var(--text-main);
  background-color: var(--primary-glow);
  font-weight: 500;
  border-left: 3px solid var(--primary);
}

/* Content Area */
.content-area {
  margin-left: 280px;
  flex: 1;
  padding: 50px 80px;
  max-width: 1000px;
  overflow-y: auto;
}

.doc-section {
  display: none;
}

.doc-section.active-section {
  display: block;
  animation: fadeIn 0.4s ease-out;
}

/* Headings */
h1 {
  font-family: var(--font-title);
  font-size: 40px;
  font-weight: 800;
  margin-bottom: 20px;
  background: linear-gradient(135deg, var(--text-main), var(--text-muted));
  -webkit-background-clip: text;
  -webkit-text-fill-color: transparent;
}

h2 {
  font-family: var(--font-title);
  font-size: 26px;
  font-weight: 700;
  margin-top: 40px;
  margin-bottom: 15px;
}

h3 {
  font-family: var(--font-title);
  font-size: 18px;
  font-weight: 600;
  margin-top: 25px;
  margin-bottom: 10px;
}

p {
  color: var(--text-muted);
  margin-bottom: 20px;
  font-size: 16px;
}

p.lead {
  font-size: 18px;
  color: var(--text-main);
}

/* Code Blocks */
pre {
  background-color: var(--code-bg);
  border-radius: 12px;
  padding: 20px;
  overflow-x: auto;
  border: 1px solid var(--border-color);
  margin-bottom: 25px;
}

code {
  font-family: var(--font-mono);
  font-size: 14px;
  color: #a5b4fc;
}

/* Alerts */
.alert {
  padding: 16px 20px;
  border-radius: 12px;
  margin: 25px 0;
  border-left: 4px solid var(--primary);
  background-color: var(--primary-glow);
}

.alert.tip {
  border-left-color: var(--accent);
  background-color: var(--accent-glow);
}

.alert.note {
  border-left-color: var(--success);
  background-color: var(--success-glow);
}

/* Features Grid */
.features-grid {
  display: grid;
  grid-template-columns: repeat(auto-fit, minmax(220px, 1fr));
  gap: 20px;
  margin-top: 30px;
}

.feature-card {
  background-color: var(--bg-secondary);
  border: 1px solid var(--border-color);
  border-radius: 12px;
  padding: 20px;
}

.feature-card h4 {
  font-family: var(--font-title);
  font-size: 16px;
  margin-bottom: 10px;
  color: var(--text-main);
}

.feature-card p {
  font-size: 13px;
  margin-bottom: 0;
}

/* Tables */
.api-table {
  width: 100%;
  border-collapse: collapse;
  margin: 30px 0;
}

.api-table th, .api-table td {
  border-bottom: 1px solid var(--border-color);
  padding: 12px 16px;
  text-align: left;
  font-size: 14px;
}

.api-table th {
  font-family: var(--font-title);
  color: var(--text-main);
  background-color: var(--bg-secondary);
  font-weight: 600;
}

.api-table td code {
  color: var(--accent);
}

/* Diagrams */
.architecture-diagram pre {
  background-color: var(--bg-secondary);
  color: var(--text-main);
  border-color: var(--border-color);
  font-family: var(--font-mono);
  font-size: 13px;
  line-height: 1.3;
}

@keyframes fadeIn {
  from { opacity: 0; transform: translateY(10px); }
  to { opacity: 1; transform: translateY(0); }
}

/* Responsive design */
@media(max-width: 900px) {
  .sidebar { width: 0; display: none; }
  .content-area { margin-left: 0; padding: 30px; }
  .search-container { display: none; }
}
`;

const JS_CONTENT = `document.addEventListener('DOMContentLoaded', () => {
  // Navigation Routing System
  const navItems = document.querySelectorAll('.nav-item');
  const sections = document.querySelectorAll('.doc-section');

  function showSection(targetId) {
    let cleanId = targetId.replace('#', '');
    
    // Deactivate all
    sections.forEach(s => s.classList.remove('active-section'));
    navItems.forEach(n => n.classList.remove('active'));

    // Activate selected
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

  // Bind sidebar nav links
  navItems.forEach(item => {
    item.addEventListener('click', (e) => {
      e.preventDefault();
      const href = item.getAttribute('href');
      window.history.pushState(null, null, href);
      showSection(href);
    });
  });

  // Handle hash changes on back/forward buttons
  window.addEventListener('popstate', () => {
    const hash = window.location.hash || '#introduction';
    showSection(hash);
  });

  // Load initial page hash if exists
  if (window.location.hash) {
    showSection(window.location.hash);
  }

  // Theme Toggle Mechanism
  const themeToggle = document.getElementById('theme-toggle');
  const body = document.body;
  
  // Set default theme from localStorage
  const savedTheme = localStorage.getItem('theme');
  if (savedTheme === 'light') {
    body.classList.remove('dark-theme');
    body.classList.add('light-theme');
    themeToggle.querySelector('.toggle-icon').textContent = '☀️';
  }

  themeToggle.addEventListener('click', () => {
    if (body.classList.contains('dark-theme')) {
      body.classList.remove('dark-theme');
      body.classList.add('light-theme');
      themeToggle.querySelector('.toggle-icon').textContent = '☀️';
      localStorage.setItem('theme', 'light');
    } else {
      body.classList.remove('light-theme');
      body.classList.add('dark-theme');
      themeToggle.querySelector('.toggle-icon').textContent = '🌙';
      localStorage.setItem('theme', 'dark');
    }
  });

  // Search Filter Mechanism
  const searchInput = document.getElementById('doc-search');
  searchInput.addEventListener('input', (e) => {
    const query = e.target.value.toLowerCase().trim();
    if (!query) {
      // Restore default sidebar list visibility
      document.querySelectorAll('.nav-section li').forEach(li => li.style.display = 'block');
      return;
    }

    // Filter items based on navigation tags
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

  // Global Ctrl + K search hotkey
  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
      e.preventDefault();
      searchInput.focus();
    }
  });
});
`;

async function main() {
  await fsExtra.ensureDir(DOCS_DIR);
  await fsExtra.writeFile(resolve(DOCS_DIR, 'index.html'), HTML_CONTENT);
  await fsExtra.writeFile(resolve(DOCS_DIR, 'style.css'), CSS_CONTENT);
  await fsExtra.writeFile(resolve(DOCS_DIR, 'app.js'), JS_CONTENT);
  console.log('Successfully built GitHub Pages documentation under docs/');
}

main().catch(err => {
  console.error('Error building docs:', err);
  process.exit(1);
});
