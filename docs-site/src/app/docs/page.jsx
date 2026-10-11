import React from 'react';
import Link from 'next/link';
import { ArrowRight, Sparkles, Layers, Cpu, FileSpreadsheet, ShieldCheck, Zap, Download } from 'lucide-react';
import CodeBlock from '../../components/CodeBlock';

export const metadata = {
  title: 'Overview & Architecture',
  description: 'Understand the architectural philosophy and templating paradigm of node-pptx-templater.',
};

const SAMPLE_OVERVIEW_CODE = `import { PPTXTemplater } from 'node-pptx-templater';

// 1. Load an existing PowerPoint template with design, layout & masters
const ppt = await PPTXTemplater.load('./templates/quarterly-report.pptx');

// 2. Safely substitute text variables with AST run-fragment healing
ppt.useSlide(1).replaceText({
  '{{COMPANY_NAME}}': 'Acme Global AI',
  '{{REPORT_DATE}}': 'October 2026',
  '{{REVENUE_TOTAL}}': '$48.5M'
});

// 3. Update native charts with automated Excel workbook dual-sync
ppt.useSlide(2);
await ppt.updateChart('ARRChart', {
  categories: ['Q1', 'Q2', 'Q3', 'Q4 (Proj)'],
  series: [{ name: 'ARR ($M)', values: [24.5, 31.2, 39.8, 48.5] }]
});

// 4. Export directly without Microsoft PowerPoint repair warnings
await ppt.saveToFile('./output/Quarterly_Report_2026.pptx');`;

export default function DocsOverviewPage() {
  return (
    <article className="space-y-12">
      {/* Page Header */}
      <div>
        <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full glass-pill text-xs font-semibold mb-4">
          <Sparkles className="w-3.5 h-3.5" />
          <span>Documentation Overview</span>
        </div>
        <h1 className="text-3xl sm:text-4xl font-extrabold tracking-tight text-slate-900 dark:text-white">
          Architectural Overview &amp; Templating Paradigm
        </h1>
        <p className="mt-3 text-base sm:text-lg text-slate-600 dark:text-slate-300 leading-relaxed">
          `node-pptx-templater` bridges the gap between high-fidelity visual design in Microsoft PowerPoint and programmatic serverless report automation in pure JavaScript.
        </p>
      </div>

      {/* The Two Paradigms: Visual Design vs Code Canvas */}
      <section className="space-y-4">
        <h2 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900 dark:text-white">
          The Two PowerPoint Paradigms
        </h2>
        
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6 pt-2">
          {/* Paradigm A */}
          <div className="glass-card p-6 rounded-2xl space-y-3">
            <div className="w-8 h-8 rounded-xl bg-rose-500/10 text-rose-500 flex items-center justify-center font-bold text-sm">
              ✕
            </div>
            <h3 className="text-base font-bold text-slate-900 dark:text-white">
              Canvas From Scratch (Legacy Approach)
            </h3>
            <p className="text-xs text-slate-600 dark:text-slate-400 leading-relaxed">
              Coding every box coordinate, color hex, font hierarchy, and margin manually in JavaScript code (`pptx.addShape()`, `pptx.addText()`). Designers cannot review decks without git pull requests and developers spend days adjusting pixel coordinates.
            </p>
          </div>

          {/* Paradigm B: node-pptx-templater */}
          <div className="glass-card p-6 rounded-2xl space-y-3 border-brand-500/30 bg-brand-50/20 dark:bg-brand-950/20">
            <div className="w-8 h-8 rounded-xl bg-brand-500/15 text-brand-600 dark:text-brand-400 flex items-center justify-center font-bold text-sm">
              ✓
            </div>
            <h3 className="text-base font-bold text-brand-700 dark:text-brand-300">
              Template-Driven Mutation (Our Paradigm)
            </h3>
            <p className="text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
              Designers create gorgeous templates directly inside Microsoft PowerPoint with corporate slide masters, fonts, themes, and layouts. Developers simply inject dynamic data into placeholders, charts, and tables in memory.
            </p>
          </div>
        </div>
      </section>

      {/* Quick Example */}
      <section className="space-y-4">
        <h2 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900 dark:text-white">
          Under 15 Lines of Code
        </h2>
        <p className="text-xs sm:text-sm text-slate-600 dark:text-slate-400">
          Load a presentation, substitute text across fragmented XML runs, update native charts with Excel sheet dual-sync, and export cleanly:
        </p>
        <CodeBlock code={SAMPLE_OVERVIEW_CODE} language="javascript" filename="report-generator.mjs" />
      </section>

      {/* Core Architectural Subsystems */}
      <section className="space-y-4">
        <h2 className="text-xl sm:text-2xl font-bold tracking-tight text-slate-900 dark:text-white">
          Core Engine Subsystems
        </h2>
        
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 pt-1">
          <div className="glass-card p-5 rounded-2xl space-y-2">
            <div className="flex items-center gap-2 text-brand-600 dark:text-brand-400 font-semibold text-sm">
              <Zap className="w-4 h-4" />
              <h4>In-Memory Virtual File System</h4>
            </div>
            <p className="text-xs text-slate-600 dark:text-slate-400 leading-relaxed">
              Extracts OPC PKZip archives directly into an in-memory `Map&lt;string, Uint8Array&gt;`. Zero temporary disk file writes, sub-15ms cold boot, perfect for AWS Lambda and Cloudflare Workers.
            </p>
          </div>

          <div className="glass-card p-5 rounded-2xl space-y-2">
            <div className="flex items-center gap-2 text-emerald-600 dark:text-emerald-400 font-semibold text-sm">
              <Sparkles className="w-4 h-4" />
              <h4>Run Fragment Healing Engine</h4>
            </div>
            <p className="text-xs text-slate-600 dark:text-slate-400 leading-relaxed">
              When humans type tokens like <code className="font-mono text-brand-600 dark:text-brand-300">{'{{NAME}}'}</code>, PowerPoint splits them into disjoint XML run nodes. Our AST scanner coalesces text runs transparently.
            </p>
          </div>

          <div className="glass-card p-5 rounded-2xl space-y-2">
            <div className="flex items-center gap-2 text-indigo-600 dark:text-indigo-400 font-semibold text-sm">
              <FileSpreadsheet className="w-4 h-4" />
              <h4>Chart Excel Dual-Sync</h4>
            </div>
            <p className="text-xs text-slate-600 dark:text-slate-400 leading-relaxed">
              PowerPoint charts store data in both DrawingML and embedded Excel workbooks. Updating both parts synchronously eliminates the infamous &quot;PowerPoint found a problem with content&quot; repair prompt.
            </p>
          </div>

          <div className="glass-card p-5 rounded-2xl space-y-2">
            <div className="flex items-center gap-2 text-purple-600 dark:text-purple-400 font-semibold text-sm">
              <Layers className="w-4 h-4" />
              <h4>Dynamic Table Auto-Pagination</h4>
            </div>
            <p className="text-xs text-slate-600 dark:text-slate-400 leading-relaxed">
              Recursively slices overflowing rows across cloned slides against calculated vertical margins while repeating headers and preserving summary rows on the final slide.
            </p>
          </div>
        </div>
      </section>

      {/* Navigation Footers */}
      <div className="pt-8 border-t border-slate-200 dark:border-slate-800 flex flex-wrap items-center justify-between gap-4">
        <div>
          <span className="text-xs text-slate-500">Ready to install?</span>
          <p className="text-sm font-semibold text-slate-900 dark:text-white">Jump to Installation Guides</p>
        </div>
        <Link
          href="/docs/installation"
          className="flex items-center gap-2 px-5 py-2.5 rounded-xl text-xs font-semibold bg-brand-600 text-white hover:bg-brand-500 shadow-md shadow-brand-500/20 transition-all"
        >
          <span>Installation Guide</span>
          <ArrowRight className="w-3.5 h-3.5" />
        </Link>
      </div>
    </article>
  );
}
