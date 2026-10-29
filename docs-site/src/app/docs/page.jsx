'use client';

import React from 'react';
import Link from 'next/link';
import { ArrowRight, Sparkles, Zap, FileSpreadsheet, Layers, ShieldCheck, Cpu, Terminal } from 'lucide-react';
import PowerPointSlideDeck from '../../components/PowerPointSlideDeck';
import CodeBlock from '../../components/CodeBlock';

const OVERVIEW_SLIDES = [
  {
    id: 'slide-1',
    category: 'Overview',
    title: 'The Two PowerPoint Paradigms',
    subtitle: 'Canvas From Scratch vs Template-Driven Mutation',
    signature: 'Paradigm Comparison',
    content: (
      <div className="space-y-4">
        <p className="text-sm text-[var(--ppt-text-secondary)]">
          Most presentation libraries make developers construct slides from blank canvases in code. `node-pptx-templater` flips this paradigm by treating PowerPoint files created by designers as data-bindable templates.
        </p>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-4 pt-2">
          <div className="ppt-card-surface p-4 rounded-lg border border-[var(--ppt-card-border)]">
            <div className="text-xs font-bold text-rose-500 mb-1 flex items-center gap-1">
              <span>✕</span> Legacy Canvas-From-Scratch
            </div>
            <p className="text-xs text-[var(--ppt-text-secondary)] leading-relaxed">
              Coding every box coordinate, color hex, font hierarchy, and margin manually in JavaScript (<code className="text-[#D24726]">pptx.addShape()</code>). Designers cannot review decks without git PRs, and developers spend days adjusting pixel coordinates.
            </p>
          </div>

          <div className="ppt-card-surface p-4 rounded-lg border-2 border-[#D24726]">
            <div className="text-xs font-bold text-[#D24726] mb-1 flex items-center gap-1">
              <span>✓</span> Template-Driven Mutation (Our Approach)
            </div>
            <p className="text-xs text-[var(--ppt-text-secondary)] leading-relaxed">
              Designers create gorgeous templates in Microsoft PowerPoint with corporate slide masters, fonts, themes, and layouts. Developers simply inject dynamic data into placeholders, charts, and tables in memory.
            </p>
          </div>
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

// 2. Perform safe, AST run-fragment healed text replacement
ppt.useSlide(1).replaceText({
  '{{COMPANY}}': 'Acme Global AI',
  '{{REVENUE}}': '$48.5M ARR'
});

// 3. Update native charts with automated Excel workbook dual-sync
ppt.useSlide(2);
await ppt.updateChart('ARRChart', {
  categories: ['Q1', 'Q2', 'Q3', 'Q4'],
  series: [{ name: 'ARR ($M)', values: [24.5, 31.2, 39.8, 48.5] }]
});

// 4. Export cleanly without Microsoft PowerPoint repair warnings
await ppt.saveToFile('./output/Quarterly_Report_2026.pptx');`}
          language="javascript"
          filename="quickstart.mjs"
        />
      </div>
    ),
  },
  {
    id: 'slide-3',
    category: 'Architecture',
    title: 'In-Memory Virtual File System (VFS)',
    subtitle: 'Sub-15ms Cold-Start for AWS Lambda & Edge Isolates',
    signature: 'Virtual File System',
    content: (
      <div className="space-y-4">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)] leading-relaxed">
          The engine never writes temporary files to <code className="font-mono text-[#D24726]">/tmp</code>. It ingests OPC PKZip archives directly into an in-memory Virtual File System backed by <code className="font-mono text-[var(--ppt-text-primary)]">Map&lt;string, Uint8Array&gt;</code>.
        </p>

        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-1">
          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)] text-center">
            <Zap className="w-5 h-5 text-[#D24726] mx-auto mb-1" />
            <div className="text-base font-bold text-[var(--ppt-text-primary)]">&lt; 15 ms</div>
            <div className="text-[10px] text-[var(--ppt-text-secondary)]">Serverless Cold-Start</div>
          </div>
          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)] text-center">
            <Cpu className="w-5 h-5 text-[#D24726] mx-auto mb-1" />
            <div className="text-base font-bold text-[var(--ppt-text-primary)]">0 Disk I/O</div>
            <div className="text-[10px] text-[var(--ppt-text-secondary)]">Pure RAM Buffer Ops</div>
          </div>
          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)] text-center">
            <ShieldCheck className="w-5 h-5 text-emerald-500 mx-auto mb-1" />
            <div className="text-base font-bold text-[var(--ppt-text-primary)]">&lt; 2 MB</div>
            <div className="text-[10px] text-[var(--ppt-text-secondary)]">Zero-Addon Footprint</div>
          </div>
        </div>
      </div>
    ),
  },
  {
    id: 'slide-4',
    category: 'Text & Runs',
    title: 'AST Run Fragment Healing Engine',
    subtitle: 'Eliminating the Dreaded Split-Token Dilemma',
    signature: 'AST Run Normalizer',
    content: (
      <div className="space-y-3">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)] leading-relaxed">
          When humans type tokens like <code className="font-mono text-[#D24726]">{'{{COMPANY}}'}</code> in Microsoft PowerPoint, spell-checkers and font adjustments split them across multiple XML run elements (<code className="font-mono text-[var(--ppt-text-primary)]">&lt;a:r&gt;</code>).
        </p>
        <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)] text-xs font-mono space-y-1">
          <div className="text-rose-500 font-bold">// What PowerPoint actually stores in slide1.xml:</div>
          <div className="text-[var(--ppt-text-secondary)]">&lt;a:r&gt;&lt;a:t&gt;{'{{'}&lt;/a:t&gt;&lt;/a:r&gt;&lt;a:r&gt;&lt;a:t&gt;COMP&lt;/a:t&gt;&lt;/a:r&gt;&lt;a:r&gt;&lt;a:t&gt;ANY{'}}'}&lt;/a:t&gt;&lt;/a:r&gt;</div>
        </div>
        <p className="text-xs text-[var(--ppt-text-secondary)]">
          Naive regex find-and-replace fails completely. Our AST scanner coalesces text runs transparently, performs the replacement, and preserves formatting without losing styles.
        </p>
      </div>
    ),
  },
  {
    id: 'slide-5',
    category: 'Charts',
    title: 'Chart Excel Dual-Sync Architecture',
    subtitle: 'Preventing PowerPoint Repair Content Warnings',
    signature: 'Chart Dual-Sync',
    content: (
      <div className="space-y-3">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)] leading-relaxed">
          PowerPoint charts store data in two separate locations: DrawingML XML (<code className="font-mono text-[var(--ppt-text-primary)]">&lt;c:chartSpace&gt;</code>) and an embedded binary Excel workbook (<code className="font-mono text-[var(--ppt-text-primary)]">ppt/embeddings/Microsoft_Excel_Worksheet.xlsx</code>).
        </p>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 pt-1">
          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
            <span className="text-xs font-bold text-[#D24726]">1. DrawingML XML</span>
            <p className="text-[11px] text-[var(--ppt-text-secondary)] mt-1">
              Controls what is visually painted when the presentation is opened.
            </p>
          </div>
          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
            <span className="text-xs font-bold text-[#D24726]">2. Embedded XLSX</span>
            <p className="text-[11px] text-[var(--ppt-text-secondary)] mt-1">
              Controls what users see when they right-click &quot;Edit Data in Excel&quot;.
            </p>
          </div>
        </div>
        <p className="text-xs text-emerald-500 font-medium">
          ✓ node-pptx-templater synchronizes both simultaneously in memory. Zero repair dialogs.
        </p>
      </div>
    ),
  },
  {
    id: 'slide-6',
    category: 'Tables',
    title: 'Dynamic Multi-Slide Table Auto-Pagination',
    subtitle: 'Automatic Row Height EMU Calculation & Header Repetition',
    signature: 'Table Pagination',
    content: (
      <div className="space-y-3">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)] leading-relaxed">
          When reporting data with 50+ rows, a single slide table would overflow the bottom canvas. The engine dynamically calculates row heights in English Metric Units (EMUs) and paginates rows across cloned slides:
        </p>
        <CodeBlock
          code={`await ppt.updateTable('BillingTable', {
  data: transactionsList, // 100+ rows
  options: {
    paginate: true,
    maxSlideHeight: 4800000, // EMU boundary
    repeatHeader: true,      // Copies table headers on slide 2, 3, etc.
    preserveSummaryRow: true // Anchors summary row on final slide
  }
});`}
          language="javascript"
          filename="table-pagination.mjs"
        />
      </div>
    ),
  },
  {
    id: 'slide-7',
    category: 'Security',
    title: 'Hardened Package & XML Security',
    subtitle: 'Enterprise-Grade XXE, Entity Bomb & Zip-Slip Resistance',
    signature: 'Security Hardening',
    content: (
      <div className="space-y-4">
        <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)]">
          Built from the ground up for zero-trust enterprise multi-tenant cloud environments:
        </p>

        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
            <ShieldCheck className="w-4 h-4 text-[#D24726] mb-1" />
            <h4 className="text-xs font-bold text-[var(--ppt-text-primary)]">XXE Blocking</h4>
            <p className="text-[10px] text-[var(--ppt-text-secondary)] mt-1">
              Entity resolution disabled in custom XML parser. No external DTD resolution.
            </p>
          </div>

          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
            <ShieldCheck className="w-4 h-4 text-[#D24726] mb-1" />
            <h4 className="text-xs font-bold text-[var(--ppt-text-primary)]">Billion Laughs Guard</h4>
            <p className="text-[10px] text-[var(--ppt-text-secondary)] mt-1">
              Limits recursion depth and entity expansion size to prevent memory starvation attacks.
            </p>
          </div>

          <div className="ppt-card-surface p-3 rounded-lg border border-[var(--ppt-card-border)]">
            <ShieldCheck className="w-4 h-4 text-[#D24726] mb-1" />
            <h4 className="text-xs font-bold text-[var(--ppt-text-primary)]">Zip-Slip Defense</h4>
            <p className="text-[10px] text-[var(--ppt-text-secondary)] mt-1">
              Strict path sanitization prevents directory traversal attacks during ZIP extraction.
            </p>
          </div>
        </div>

        <div className="pt-2 flex justify-end">
          <Link
            href="/docs/installation"
            className="inline-flex items-center gap-1.5 px-4 py-2 rounded text-xs font-bold bg-[#D24726] text-white hover:bg-[#b83b1e] transition-colors"
          >
            <span>Next: Installation Guide</span>
            <ArrowRight className="w-3.5 h-3.5" />
          </Link>
        </div>
      </div>
    ),
  },
];

export default function DocsPage() {
  return (
    <PowerPointSlideDeck
      slides={OVERVIEW_SLIDES}
      categoryTitle="Overview"
    />
  );
}
