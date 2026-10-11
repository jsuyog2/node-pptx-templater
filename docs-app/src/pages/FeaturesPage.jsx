import React, { useState } from 'react';
import { CheckCircle2, AlertCircle, Clock, ShieldCheck, Search } from 'lucide-react';
import SEOHead from '../components/SEOHead';
import apiData from '../data/apiReference.json';

export default function FeaturesPage() {
  const [search, setSearch] = useState('');

  const subsystems = [
    { name: 'Core & Lifecycle', status: 'Supported', count: 16, desc: 'PackageLoader, ZIP inflation, VFS in-memory cache, and stream pipelines.' },
    { name: 'Slide Management', status: 'Supported', count: 11, desc: 'Slide targeting, cloning, reordering, deletion, and cross-deck import.' },
    { name: 'Text & Run Healing', status: 'Supported', count: 10, desc: 'AST run normalizer, split tag repair, font & typography styling.' },
    { name: 'Native Charts & Excel Sync', status: 'Supported', count: 8, desc: 'Clustered column, line, bar, pie, and embedded XLSX workbook caching.' },
    { name: 'Tables & Merging', status: 'Supported', count: 8, desc: 'Dynamic rows, alternating row striping, rowspan/gridspan, and splitTableOnOverflow.' },
    { name: 'Cell Shapes & Badges', status: 'Supported', count: 3, desc: 'Status badges and vector shapes anchored inside table cells.' },
    { name: 'Shapes & Geometry', status: 'Supported', count: 9, desc: 'Vector shapes, text boxes, coordinates, bounding box queries, and EMU conversions.' },
    { name: 'Images & Media', status: 'Supported', count: 3, desc: 'PNG, JPEG, WebP, Base64 data URIs, and TIFF image injection.' },
    { name: 'Z-Order Stacking', status: 'Supported', count: 9, desc: 'Visual layer reordering (bringToFront, sendToBack, stepForward, stepBackward).' },
    { name: 'Hyperlinks & Actions', status: 'Supported', count: 7, desc: 'Web URLs and internal slide-to-slide jump action links.' },
    { name: 'Validation & Diagnostics', status: 'Supported', count: 12, desc: 'OpenXML schema compliance validators, chart inspection, and XML repair.' }
  ];

  const filteredSubsystems = subsystems.filter((s) =>
    s.name.toLowerCase().includes(search.toLowerCase()) ||
    s.desc.toLowerCase().includes(search.toLowerCase())
  );

  return (
    <div>
      <SEOHead
        title="Verified Feature Catalog — node-pptx-templater"
        description="Audit-verified feature catalog and subsystem capabilities of node-pptx-templater v3.0.0."
        canonicalUrl="https://node-pptx-templater.pages.dev/features"
      />

      {/* Header Banner */}
      <div className="mb-8 pb-6 border-b border-slate-200/80 dark:border-slate-800/80">
        <div className="flex flex-wrap items-center gap-2 mb-2">
          <span className="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border border-emerald-500/20">
            Feature Inventory
          </span>
          <span className="text-xs text-slate-500">100% Audit Verified</span>
        </div>
        <h1 className="text-2xl sm:text-4xl font-extrabold tracking-tight text-slate-900 dark:text-slate-100 mb-3">
          Verified Feature Catalog
        </h1>
        <p className="text-xs sm:text-sm text-slate-600 dark:text-slate-300 max-w-3xl leading-relaxed">
          Every method, property, and subsystem is verified against actual test suites (285 passing Vitest tests) and production OpenXML validation routines.
        </p>

        {/* Search Input */}
        <div className="relative mt-6 max-w-md">
          <input
            type="text"
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            placeholder="Search subsystems & capabilities..."
            className="w-full px-4 py-2 text-xs sm:text-sm rounded-xl bg-slate-100 dark:bg-slate-900 border border-slate-200 dark:border-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500/50 text-slate-900 dark:text-slate-100 placeholder-slate-400"
          />
        </div>
      </div>

      {/* Subsystem Cards Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
        {filteredSubsystems.map((sub, idx) => (
          <div
            key={idx}
            className="glass-card rounded-2xl p-6 border border-slate-200/80 dark:border-slate-800/80"
          >
            <div className="flex items-center justify-between gap-3 mb-2">
              <h3 className="text-base font-bold text-slate-900 dark:text-slate-100">
                {sub.name}
              </h3>
              <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full text-xs font-semibold bg-emerald-500/10 text-emerald-600 dark:text-emerald-400">
                <CheckCircle2 className="w-3.5 h-3.5" /> {sub.status}
              </span>
            </div>
            <p className="text-xs sm:text-sm text-slate-600 dark:text-slate-400 leading-relaxed mb-4">
              {sub.desc}
            </p>
            <div className="text-xs font-mono text-blue-600 dark:text-blue-400 font-medium">
              {sub.count} verified public APIs
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
