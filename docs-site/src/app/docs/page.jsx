'use client';

import React from 'react';
import Link from 'next/link';
import { ArrowRight, Sparkles, Zap, FileSpreadsheet, Layers, ShieldCheck, Cpu, Terminal } from 'lucide-react';
import PowerPointSlideDeck from '../../components/PowerPointSlideDeck';
import CodeBlock from '../../components/CodeBlock';

const OVERVIEW_SLIDES = [
  {
    id: 'slide-1',
    category: 'Zero Dependencies',
    title: 'Zero Office Dependencies & Pure JavaScript Runtime',
    subtitle: 'Serverless, Edge & Node.js Native Execution',
    signature: 'Runtime Architecture & Compatibility',
    content: (
      <div className="space-y-4">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)] leading-relaxed">
          Runs entirely in memory without Microsoft Office, LibreOffice, Java, or Electron. It requires no heavy third-party generator packages like Puppeteer or PptxGenJS, enabling lightweight execution across Node.js, serverless runtimes, and edge workers.
        </p>

        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-1">
          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)] text-center">
            <Zap className="w-5 h-5 text-[#D24726] mx-auto mb-1" />
            <div className="text-base font-bold text-[var(--ppt-text-primary)]">0 Dependencies</div>
            <div className="text-[10px] text-[var(--ppt-text-secondary)]">No Office / LibreOffice / Java</div>
          </div>
          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)] text-center">
            <Cpu className="w-5 h-5 text-[#D24726] mx-auto mb-1" />
            <div className="text-base font-bold text-[var(--ppt-text-primary)]">Pure JS</div>
            <div className="text-[10px] text-[var(--ppt-text-secondary)]">No C++ Addons or Canvas</div>
          </div>
          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)] text-center">
            <ShieldCheck className="w-5 h-5 text-emerald-500 mx-auto mb-1" />
            <div className="text-base font-bold text-[var(--ppt-text-primary)]">Edge Ready</div>
            <div className="text-[10px] text-[var(--ppt-text-secondary)]">Lambda, Vercel &amp; Workers</div>
          </div>
        </div>

        <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
          <div className="text-xs font-bold text-[#D24726] mb-1">Lightweight Execution Anywhere</div>
          <p className="text-[11px] text-[var(--ppt-text-secondary)] leading-relaxed">
            Eliminates complex container builds and native binary compilation (`node-gyp`). Operates directly within standard V8 isolates with zero temporary disk writes to <code className="font-mono text-[#D24726]">/tmp</code>.
          </p>
        </div>
      </div>
    ),
  },
  {
    id: 'slide-2',
    category: 'Quickstart',
    title: 'Under 15 Lines of Code',
    subtitle: 'End-to-End Ingest, Mutation, and Export Pipeline',
    signature: 'PPTXTemplater Pipeline',
    content: (
      <div className="space-y-3">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)]">
          Load an existing template, substitute text variables with AST run-fragment healing, update native charts with Excel sheet dual-sync, and export cleanly:
        </p>
        <CodeBlock
          code={`import { PPTXTemplater } from 'node-pptx-templater';

// 1. Ingest existing template into memory (Zero disk writes)
const ppt = await PPTXTemplater.load('./templates/quarterly-report.pptx');

// 2. Perform text replacement
ppt.useSlide(1).replaceText({
  '{{COMPANY}}': 'Acme Global AI',
  '{{REVENUE}}': '$48.5M ARR'
});

// 3. Update native charts
ppt.useSlide(2);
await ppt.updateChart('ARRChart', {
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [{ name: 'ARR ($M)', values: [24.5, 31.2, 39.8, 48.5] }]
});

// 4. Export cleanly
await ppt.saveToFile('./output/Quarterly_Report_2026.pptx');`}
          language="javascript"
          filename="quickstart.mjs"
        />
      </div>
    ),
  },
  {
    id: 'slide-3',
    category: 'Input & Multi-Target Export',
    title: 'Versatile Input & Multi-Target Export Pipeline',
    subtitle: 'Strict 18-Method Flexible I/O Ecosystem',
    signature: 'PPTXTemplater I/O System',
    content: (
      <div className="space-y-3">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)] leading-relaxed">
          Ingests presentations from files, Buffers, URLs, streams, directories, and Flat OPC XML. Exports directly to disk files, Node.js Buffers, Streams, ArrayBuffers, and Base64 strings.
        </p>

        <CodeBlock
          code={`// 1. Ingest from any source (File, Buffer, Stream, URL, Flat OPC XML)
const ppt = await PPTXTemplater.load('./templates/report.pptx');

// 2. Perform slide mutations and content replacements
ppt.useSlide(1).replaceText({ '{{COMPANY}}': 'Acme Global' });

// 3. Export to your target runtime destination
await ppt.saveToFile('./output/report.pptx');     // File on disk
const buffer = await ppt.toBuffer();              // S3 / Blob buffer
return await ppt.toResponse({ filename: 'deck.pptx' }); // Web API Response`}
          language="javascript"
          filename="io-pipeline.mjs"
        />

        <p className="text-[11px] text-[var(--ppt-text-secondary)]">
          Also features dedicated streaming pipe adapters directly for web frameworks: <code className="font-mono text-[#D24726]">ppt.toExpress(res)</code> and <code className="font-mono text-[#D24726]">ppt.toFastify(reply)</code>.
        </p>
      </div>
    ),
  },
  {
    id: 'slide-4',
    category: 'Create Blank Presentations',
    title: 'Create Blank Presentations From Scratch',
    subtitle: 'In-Memory Synthesis Without External Template Files',
    signature: 'PPTXTemplater.createBlank(options)',
    content: (
      <div className="space-y-3">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)] leading-relaxed">
          Initialize minimal, compliant PowerPoint decks entirely in memory without loading external template files. Supports custom aspect ratio presets such as 16:9, 4:3, or explicit EMU dimensions.
        </p>

        <CodeBlock
          code={`import { PPTXTemplater } from 'node-pptx-templater';

// Initialize a clean 16:9 widescreen presentation in memory
const ppt = await PPTXTemplater.createBlank({ aspectRatio: '16:9' });

// Immediately target slide 1 and construct layouts via fluent chaining
ppt.useSlide(1);
await ppt.addShape('rect', {
  x: '10%', y: '15%', width: '80%', height: '30%',
  text: 'Generated Completely From Scratch',
  style: { fill: '#1E293B', color: '#FFFFFF' }
});

await ppt.saveToFile('./output/from-scratch.pptx');`}
          language="javascript"
          filename="create-blank.js"
        />
      </div>
    ),
  },
  {
    id: 'slide-5',
    category: 'Slide & Element',
    title: 'Complete Slide & Element Lifecycle Management',
    subtitle: 'Fine-Grained Manipulation Across All Presentation Objects',
    signature: 'Lifecycle & Stacking Order',
    content: (
      <div className="space-y-3">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)] leading-relaxed">
          Add, update, clone, move, style, reorder, or delete slides and elements with full layout control. Handles rich text, media, vector shapes, tables, and charts with absolute spatial positioning.
        </p>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
            <span className="text-xs font-bold text-[#D24726]">Slide Lifecycle</span>
            <p className="text-[11px] text-[var(--ppt-text-secondary)] mt-1 leading-relaxed">
              Target (<code className="font-mono text-[var(--ppt-text-primary)]">useSlide</code>), clone (<code className="font-mono text-[var(--ppt-text-primary)]">cloneSlide</code>), move, reorder, tag, partition into sections (<code className="font-mono text-[var(--ppt-text-primary)]">addSection</code>), or import across presentations with deduplication.
            </p>
          </div>
          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
            <span className="text-xs font-bold text-[#D24726]">Z-Index &amp; Element Stacking</span>
            <p className="text-[11px] text-[var(--ppt-text-secondary)] mt-1 leading-relaxed">
              Easily bring elements forward, send them backward, or set exact Z-index stacking orders (<code className="font-mono text-[var(--ppt-text-primary)]">bringToFront</code>, <code className="font-mono text-[var(--ppt-text-primary)]">sendToBack</code>) while preserving OpenXML tree structure.
            </p>
          </div>
        </div>

        <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)] text-xs font-mono space-y-1">
          <div className="text-[var(--ppt-text-secondary)]">// Fluent lifecycle chaining example:</div>
          <div className="text-[var(--ppt-text-primary)]">
            ppt.useSlide(1).cloneSlide(1).moveSlide(2, 4).bringToFront(&apos;chart1&apos;);
          </div>
        </div>
      </div>
    ),
  },
  {
    id: 'slide-6',
    category: 'Corporate Support',
    title: 'Native Corporate Presentation & Classifier Support',
    subtitle: 'Preserves Enterprise Slide Masters, Themes & Markings',
    signature: 'Corporate Compliance & Metadata',
    content: (
      <div className="space-y-4">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)] leading-relaxed">
          Preserves corporate slide master templates, custom XML metadata, and enterprise security classifiers without file corruption. Avoids schema-breaking alterations to ensure enterprise-compliant presentations.
        </p>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
            <div className="text-xs font-bold text-[#D24726] mb-1 flex items-center gap-1.5">
              <span className="inline-block w-1.5 h-1.5 rounded-full bg-[#D24726]"></span>
              Sensitivity &amp; Security Classifiers
            </div>
            <p className="text-xs text-[var(--ppt-text-secondary)] leading-relaxed">
              Maintains organization-mandated classification markings, sensitivity labels, corporate headers, and footers during automated data injection.
            </p>
          </div>

          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
            <div className="text-xs font-bold text-[#D24726] mb-1 flex items-center gap-1.5">
              <span className="inline-block w-1.5 h-1.5 rounded-full bg-[#D24726]"></span>
              Custom XML &amp; Slide Masters
            </div>
            <p className="text-xs text-[var(--ppt-text-secondary)] leading-relaxed">
              Custom XML parts, proprietary add-in metadata, and branded corporate color palettes remain completely untouched and valid across updates.
            </p>
          </div>
        </div>
      </div>
    ),
  },
  {
    id: 'slide-7',
    category: 'Charts & Tables',
    title: 'All 17 Native Chart Types & Matrix Table Engine',
    subtitle: 'Dual Excel Workbook Synchronization & Robust Cell Manipulation',
    signature: 'Charts & Tables Management',
    content: (
      <div className="space-y-3">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)] leading-relaxed">
          Supports all 17 PowerPoint chart types with dual Excel sync, plus a full table management engine to add, style, merge, and remove tables, rows, or columns.
        </p>

        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
            <span className="text-xs font-bold text-[#D24726]">17 Native Chart Types</span>
            <p className="text-[11px] text-[var(--ppt-text-secondary)] mt-1 leading-relaxed">
              Column, Line, Pie, Bar, Area, Scatter, Map, Stock, Surface, Radar, Treemap, Sunburst, Histogram, Box &amp; Whisker, Waterfall, Funnel, and Combo. Updates both DrawingML XML and embedded Excel spreadsheets simultaneously.
            </p>
          </div>
          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
            <span className="text-xs font-bold text-[#D24726]">Matrix Table Engine</span>
            <p className="text-[11px] text-[var(--ppt-text-secondary)] mt-1 leading-relaxed">
              Add, modify, and delete rows/columns with unique 32-bit <code className="font-mono text-[var(--ppt-text-primary)]">&lt;a16:rowId&gt;</code> tags, rectangular cell merges (<code className="font-mono text-[var(--ppt-text-primary)]">gridSpan</code>, <code className="font-mono text-[var(--ppt-text-primary)]">rowSpan</code>), and cell shape overlays without layout reflow.
            </p>
          </div>
        </div>
      </div>
    ),
  },
];;

export default function DocsPage() {
  return (
    <PowerPointSlideDeck
      slides={OVERVIEW_SLIDES}
      categoryTitle="Overview"
    />
  );
}
