'use client';

import React, { useRef, useEffect } from 'react';
import Link from 'next/link';
import {
  ArrowRight,
  Zap,
  ShieldCheck,
  Server,
  Layers,
  FileSpreadsheet,
  Split,
  Code2,
  Cpu,
  BarChart3,
  Table,
  Image as ImageIcon,
  CheckCircle2,
  XCircle,
  Sparkles,
  BookOpen
} from 'lucide-react';
import CodeBlock from '../components/CodeBlock';
import siteMeta from '../data/siteMeta.json';
import { animateHero } from '../animations/gsapTransitions';

const QUICKSTART_CODE = `import { PPTXTemplater } from 'node-pptx-templater';

// 1. Ingest existing presentation template (.pptx) or Buffer
const ppt = await PPTXTemplater.load('./templates/executive-review.pptx');

// 2. Perform safe, run-fragment-healed text replacements
ppt.useSlide(1).replaceText({
  '{{COMPANY}}': 'Acme Global AI',
  '{{QUARTER}}': 'Q3 2026',
  '{{REVENUE}}': '$42.8M'
});

// 3. Update native PowerPoint charts with cached Excel sheet dual-sync
ppt.useSlide(2);
await ppt.updateChart('ARRPerformanceChart', {
  categories: ['Q1', 'Q2', 'Q3', 'Q4 (Proj)'],
  series: [
    { name: 'Cloud ARR', values: [18.4, 22.1, 26.8, 31.5] },
    { name: 'AI Services', values: [8.2, 12.6, 17.4, 23.0] }
  ]
});

// 4. Save cleanly without Microsoft PowerPoint repair warnings
await ppt.saveToFile('./output/Quarterly_Executive_Review.pptx');
console.log('✓ PowerPoint presentation generated in 42ms!');`;

const COMPARISON_DATA = [
  {
    feature: 'Architecture & Engine',
    pptxTemplater: 'Pure JavaScript AST & OpenXML Parser',
    libreOffice: 'Heavy C++ headless daemon (Docker)',
    poi: 'Heavy Java VM & classpath',
    pythonPptx: 'Python CPython runtime'
  },
  {
    feature: 'Cold-Start Boot Time',
    pptxTemplater: '< 15 ms',
    libreOffice: '8,000 – 15,000 ms',
    poi: '1,500 – 4,000 ms',
    pythonPptx: '300 – 900 ms'
  },
  {
    feature: 'Binary Container Size',
    pptxTemplater: '< 2 MB (npm install)',
    libreOffice: '500 MB – 1.2 GB',
    poi: '45 MB – 90 MB',
    pythonPptx: '30 MB – 80 MB'
  },
  {
    feature: 'Serverless Compatible',
    pptxTemplater: 'AWS Lambda, Cloudflare Workers, Vercel Edge',
    libreOffice: 'Lambda container only (high memory)',
    poi: 'Lambda container only',
    pythonPptx: 'AWS Lambda (Python runtime)'
  },
  {
    feature: 'Run Fragment Healing',
    pptxTemplater: 'Full AST-level coalescing & repair',
    libreOffice: 'None (naive string replacement)',
    poi: 'Manual paragraph run manipulation',
    pythonPptx: 'Manual run-level loops'
  },
  {
    feature: 'Embedded Excel Dual-Sync',
    pptxTemplater: 'Automated DrawingML + XLSX sync',
    libreOffice: 'Requires full recalculation pass',
    poi: 'Manual POI-HSSF / XSSF hookup',
    pythonPptx: 'Partial (frequently triggers Repair)'
  }
];

const DOMAINS = [
  {
    title: 'Lifecycle & I/O',
    count: 16,
    icon: Server,
    desc: 'Load from path, stream, Buffer, or Base64. Export to disk, HTTP response, or in-memory ArrayBuffer.'
  },
  {
    title: 'Slides Management',
    count: 12,
    icon: Layers,
    desc: 'Add, duplicate, delete, reorder, clone across decks, and inspect slide layouts and relationships.'
  },
  {
    title: 'Text & Placeholders',
    count: 10,
    icon: Sparkles,
    desc: 'Run-fragment healing engine ensures {{tags}} split across multiple XML nodes are safely replaced.'
  },
  {
    title: 'Tables & Cell Shapes',
    count: 11,
    icon: Table,
    desc: 'Dynamic rows, rectangular cell span merging, auto-fitting, and cell-anchored shape placements.'
  },
  {
    title: 'Charts & Series Sync',
    count: 8,
    icon: BarChart3,
    desc: 'Update bar, pie, and line charts while seamlessly caching data in embedded Excel workbooks.'
  },
  {
    title: 'Shapes & Geometry',
    count: 11,
    icon: Cpu,
    desc: 'Inject geometric shapes, format fill/strokes, position coordinates, and manage rotation.'
  },
  {
    title: 'Images & Media',
    count: 3,
    icon: ImageIcon,
    desc: 'Embed PNG, JPEG, SVG, and WebP with auto-calculated aspect ratios and relationship remapping.'
  },
  {
    title: 'Z-Order & Layering',
    count: 11,
    icon: Split,
    desc: 'Bring to front, send to back, shift layers, and normalize object stacking orders in PresentationML.'
  },
  {
    title: 'Diagnostics & Validation',
    count: 12,
    icon: ShieldCheck,
    desc: 'Inspect XML AST, validate relationship integrity, and verify schema compliance before export.'
  }
];

export default function HomePage() {
  const heroRef = useRef(null);

  useEffect(() => {
    if (heroRef.current) {
      return animateHero(heroRef.current);
    }
  }, []);

  return (
    <div className="space-y-24 pb-16">
      {/* Hero Section */}
      <section
        ref={heroRef}
        className="relative pt-12 md:pt-20 px-4 sm:px-6 lg:px-8 max-w-7xl mx-auto text-center"
      >
        <div className="hero-badge inline-flex items-center gap-2 px-3 py-1.5 rounded-full glass-pill text-xs font-semibold mb-6">
          <Zap className="w-3.5 h-3.5 text-brand-500" />
          <span>Next-Gen Pure JavaScript PowerPoint Engine</span>
          <span className="w-1.5 h-1.5 rounded-full bg-brand-500"></span>
          <span className="font-mono">{siteMeta.version}</span>
        </div>

        <h1 className="hero-title text-4xl sm:text-6xl lg:text-7xl font-extrabold text-slate-900 dark:text-white tracking-tight leading-[1.1] max-w-5xl mx-auto">
          High-Performance{' '}
          <span className="text-transparent bg-clip-text bg-gradient-to-r from-brand-600 via-sky-500 to-indigo-600 dark:from-brand-400 dark:via-sky-300 dark:to-indigo-400">
            OpenXML PowerPoint
          </span>{' '}
          Template Engine
        </h1>

        <p className="hero-description mt-6 text-lg sm:text-xl text-slate-600 dark:text-slate-300 max-w-3xl mx-auto leading-relaxed">
          {siteMeta.description}
        </p>

        {/* CTA Buttons */}
        <div className="mt-8 flex flex-wrap items-center justify-center gap-4">
          <Link
            href="/docs"
            className="hero-cta px-6 py-3.5 rounded-xl font-semibold text-sm bg-gradient-to-r from-brand-600 to-indigo-600 hover:from-brand-500 hover:to-indigo-500 text-white shadow-lg shadow-brand-500/25 flex items-center gap-2 transition-all hover:scale-[1.02]"
          >
            <BookOpen className="w-4 h-4" />
            <span>Read the Docs</span>
            <ArrowRight className="w-4 h-4" />
          </Link>
          <Link
            href="/features"
            className="hero-cta px-6 py-3.5 rounded-xl font-semibold text-sm glass-card hover:bg-slate-100 dark:hover:bg-slate-800 transition-all flex items-center gap-2 hover:scale-[1.02]"
          >
            <Sparkles className="w-4 h-4 text-brand-500" />
            <span>All Features</span>
          </Link>
          <Link
            href="/docs/api"
            className="hero-cta px-6 py-3.5 rounded-xl font-semibold text-sm glass-card hover:bg-slate-100 dark:hover:bg-slate-800 transition-all flex items-center gap-2 hover:scale-[1.02]"
          >
            <Code2 className="w-4 h-4 text-brand-500" />
            <span>Browse APIs</span>
          </Link>
        </div>

        {/* Quick Install Bar */}
        <div className="hero-install mt-8 inline-flex items-center gap-3 px-4 py-2 rounded-xl glass-card text-xs font-mono text-slate-700 dark:text-slate-300 shadow-inner">
          <span className="text-slate-400 dark:text-slate-500">$</span>
          <span>npm install node-pptx-templater</span>
        </div>
      </section>

      {/* Interactive Quickstart Section */}
      <section className="max-w-5xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="text-center mb-6">
          <h2 className="text-2xl sm:text-3xl font-bold text-slate-900 dark:text-white tracking-tight">
            Generate Presentations in Under 15 Lines of Code
          </h2>
          <p className="text-sm text-slate-600 dark:text-slate-400 mt-2">
            No Microsoft Office, no Java runtime, and zero external binary dependencies.
          </p>
        </div>
        <CodeBlock code={QUICKSTART_CODE} language="javascript" filename="quickstart.mjs" />
      </section>

      {/* Feature Highlights Grid */}
      <section className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="text-center mb-12">
          <h2 className="text-3xl font-bold text-slate-900 dark:text-white tracking-tight">
            Engineered for Production Reliability
          </h2>
          <p className="text-sm text-slate-600 dark:text-slate-400 mt-2">
            Eliminates the common failure modes of legacy template engines.
          </p>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          <div className="glass-card p-6 rounded-2xl">
            <div className="w-10 h-10 rounded-xl bg-brand-500/10 border border-brand-500/30 flex items-center justify-center mb-4">
              <Sparkles className="w-5 h-5 text-brand-500" />
            </div>
            <h3 className="text-lg font-bold text-slate-900 dark:text-white mb-2">Run Fragment Healing</h3>
            <p className="text-xs text-slate-600 dark:text-slate-400 leading-relaxed">
              When humans edit PowerPoint templates, tags like <code className="text-brand-600 dark:text-brand-300">{'{{NAME}}'}</code> are fractured across disjoint <code className="text-brand-600 dark:text-brand-300">&lt;a:r&gt;</code> run nodes. Our AST scanner coalesces text runs transparently without breaking formatting.
            </p>
          </div>

          <div className="glass-card p-6 rounded-2xl">
            <div className="w-10 h-10 rounded-xl bg-emerald-500/10 border border-emerald-500/30 flex items-center justify-center mb-4">
              <FileSpreadsheet className="w-5 h-5 text-emerald-500" />
            </div>
            <h3 className="text-lg font-bold text-slate-900 dark:text-white mb-2">Chart Excel Dual-Sync</h3>
            <p className="text-xs text-slate-600 dark:text-slate-400 leading-relaxed">
              Modern PowerPoint charts store series data in two places: DrawingML chart definitions and embedded Excel workbooks. Updating both simultaneously prevents Microsoft PowerPoint &quot;Repair Presentation&quot; warnings.
            </p>
          </div>

          <div className="glass-card p-6 rounded-2xl">
            <div className="w-10 h-10 rounded-xl bg-purple-500/10 border border-purple-500/30 flex items-center justify-center mb-4">
              <Server className="w-5 h-5 text-purple-500" />
            </div>
            <h3 className="text-lg font-bold text-slate-900 dark:text-white mb-2">Pure V8 Isolates</h3>
            <p className="text-xs text-slate-600 dark:text-slate-400 leading-relaxed">
              Runs in 100% memory without temporary files. Native compatibility with Cloudflare Workers, AWS Lambda (sub-15ms cold start), Vercel Edge, Bun, and Node.js.
            </p>
          </div>
        </div>
      </section>

      {/* Architectural Comparison Matrix */}
      <section className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="text-center mb-10">
          <h2 className="text-3xl font-bold text-slate-900 dark:text-white tracking-tight">
            How It Compares to Legacy Solutions
          </h2>
          <p className="text-sm text-slate-600 dark:text-slate-400 mt-2">
            Why modern engineering teams choose node-pptx-templater for mission-critical reporting.
          </p>
        </div>

        <div className="overflow-x-auto rounded-2xl glass-card">
          <table className="w-full text-left text-xs">
            <thead className="bg-slate-100/90 dark:bg-slate-950/80 uppercase font-mono text-[11px] text-slate-500 dark:text-slate-400 border-b border-slate-200 dark:border-slate-800">
              <tr>
                <th className="px-5 py-3.5">Capability</th>
                <th className="px-5 py-3.5 text-brand-600 dark:text-brand-400 font-bold bg-brand-500/10">node-pptx-templater</th>
                <th className="px-5 py-3.5">LibreOffice CLI</th>
                <th className="px-5 py-3.5">Apache POI</th>
                <th className="px-5 py-3.5">python-pptx</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-200 dark:divide-slate-800/60 font-sans">
              {COMPARISON_DATA.map((row, idx) => (
                <tr key={idx} className="hover:bg-slate-50/50 dark:hover:bg-slate-900/40">
                  <td className="px-5 py-3.5 font-semibold text-slate-900 dark:text-white whitespace-nowrap">
                    {row.feature}
                  </td>
                  <td className="px-5 py-3.5 font-semibold text-brand-700 dark:text-brand-300 bg-brand-500/5">
                    <span className="flex items-center gap-1.5">
                      <CheckCircle2 className="w-4 h-4 text-brand-500 shrink-0" />
                      {row.pptxTemplater}
                    </span>
                  </td>
                  <td className="px-5 py-3.5 text-slate-500 dark:text-slate-400">
                    <span className="flex items-center gap-1.5">
                      <XCircle className="w-4 h-4 text-rose-500 shrink-0" />
                      {row.libreOffice}
                    </span>
                  </td>
                  <td className="px-5 py-3.5 text-slate-500 dark:text-slate-400">
                    {row.poi}
                  </td>
                  <td className="px-5 py-3.5 text-slate-500 dark:text-slate-400">
                    {row.pythonPptx}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>

      {/* 10 Domain Areas Grid */}
      <section className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
        <div className="flex flex-col sm:flex-row items-start sm:items-end justify-between mb-8 gap-4">
          <div>
            <h2 className="text-3xl font-bold text-slate-900 dark:text-white tracking-tight">
              Specialized Domain APIs
            </h2>
            <p className="text-sm text-slate-600 dark:text-slate-400 mt-1">
              Comprehensive control over the complete OpenXML PresentationML schema.
            </p>
          </div>
          <Link
            href="/docs/api"
            className="text-xs font-semibold text-brand-600 dark:text-brand-400 hover:text-brand-500 flex items-center gap-1.5"
          >
            <span>View all methods</span>
            <ArrowRight className="w-3.5 h-3.5" />
          </Link>
        </div>

        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-5">
          {DOMAINS.map((domain) => {
            const Icon = domain.icon;
            return (
              <Link
                key={domain.title}
                href={`/docs/api?category=${encodeURIComponent(domain.title)}`}
                className="group glass-card p-5 rounded-2xl hover:border-brand-500/40 hover:-translate-y-1 transition-all duration-200"
              >
                <div className="flex items-center justify-between mb-3">
                  <div className="w-9 h-9 rounded-xl bg-slate-100 dark:bg-slate-800 group-hover:bg-brand-500/20 flex items-center justify-center transition-colors">
                    <Icon className="w-4 h-4 text-brand-600 dark:text-brand-400" />
                  </div>
                  <span className="text-xs font-mono px-2 py-0.5 rounded-full bg-slate-100 dark:bg-slate-800 text-slate-700 dark:text-slate-300 group-hover:text-brand-600 dark:group-hover:text-brand-300">
                    {domain.count} APIs
                  </span>
                </div>
                <h3 className="text-base font-bold text-slate-900 dark:text-white group-hover:text-brand-600 dark:group-hover:text-brand-300 transition-colors">
                  {domain.title}
                </h3>
                <p className="text-xs text-slate-600 dark:text-slate-400 mt-2 leading-relaxed">
                  {domain.desc}
                </p>
              </Link>
            );
          })}
        </div>
      </section>
    </div>
  );
}
