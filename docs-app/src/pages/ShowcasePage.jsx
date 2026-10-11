import React, { useState } from 'react';
import { Sparkles, Terminal, ArrowRight, Play, Copy, Check } from 'lucide-react';
import CodeBlock from '../components/CodeBlock';
import SEOHead from '../components/SEOHead';
import showcaseData from '../data/showcase.json';

export default function ShowcasePage() {
  const [selectedCategory, setSelectedCategory] = useState('All');
  const [copiedCommand, setCopiedCommand] = useState(false);

  const categories = ['All', ...new Set(showcaseData.map((d) => d.category))];

  const filteredDemos =
    selectedCategory === 'All'
      ? showcaseData
      : showcaseData.filter((d) => d.category === selectedCategory);

  const handleCopyCommand = () => {
    navigator.clipboard.writeText('npm run examples');
    setCopiedCommand(true);
    setTimeout(() => setCopiedCommand(false), 2000);
  };

  return (
    <div>
      <SEOHead
        title="Showcase & Production Scenarios — node-pptx-templater"
        description="Explore 12 production-ready showcase scenarios for node-pptx-templater: charts, overflow tables, shapes, badges, slide cloning, serverless buffers, and more."
        canonicalUrl="https://node-pptx-templater.pages.dev/showcase"
      />

      {/* Header Banner */}
      <div className="mb-8 pb-6 border-b border-slate-200/80 dark:border-slate-800/80">
        <div className="flex flex-wrap items-center gap-2 mb-2">
          <span className="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-emerald-500/10 text-emerald-600 dark:text-emerald-400 border border-emerald-500/20">
            Showcase Demos
          </span>
          <span className="text-xs text-slate-500">12 Production Demonstrations</span>
        </div>
        <h1 className="text-2xl sm:text-4xl font-extrabold tracking-tight text-slate-900 dark:text-slate-100 mb-3">
          Unified Showcase Suite
        </h1>
        <p className="text-xs sm:text-sm text-slate-600 dark:text-slate-300 max-w-3xl leading-relaxed mb-6">
          Every scenario is runnable in a single self-contained script (<code>showcase/examples.js</code>) with embedded JSON data. Run all 12 demonstrations locally with one command:
        </p>

        {/* Command Runner Box */}
        <div className="flex items-center justify-between max-w-md px-4 py-2.5 rounded-xl bg-slate-900 text-slate-100 border border-slate-800 font-mono text-xs sm:text-sm shadow-md">
          <div className="flex items-center gap-2">
            <span className="text-emerald-400 select-none">$</span>
            <span>npm run examples</span>
          </div>
          <button
            onClick={handleCopyCommand}
            className="flex items-center gap-1 px-2.5 py-1 rounded-md text-xs font-sans text-slate-400 hover:text-white bg-slate-800 hover:bg-slate-700 transition-colors"
          >
            {copiedCommand ? (
              <>
                <Check className="w-3.5 h-3.5 text-emerald-400" />
                <span className="text-emerald-400 font-medium">Copied!</span>
              </>
            ) : (
              <>
                <Copy className="w-3.5 h-3.5" />
                <span>Copy</span>
              </>
            )}
          </button>
        </div>

        {/* Filter Pills */}
        <div className="flex flex-wrap gap-2 mt-6">
          {categories.map((cat) => (
            <button
              key={cat}
              onClick={() => setSelectedCategory(cat)}
              className={`px-3 py-1.5 rounded-xl text-xs font-medium transition-all ${
                selectedCategory === cat
                  ? 'bg-blue-600 text-white shadow-sm'
                  : 'bg-slate-100 dark:bg-slate-800/80 text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-700'
              }`}
            >
              {cat}
            </button>
          ))}
        </div>
      </div>

      {/* Showcase Cards Grid */}
      <div className="space-y-8">
        {filteredDemos.map((demo) => (
          <article
            key={demo.id}
            id={demo.id}
            className="glass-card rounded-2xl p-6 sm:p-8 border border-slate-200/80 dark:border-slate-800/80"
          >
            <div className="flex flex-wrap items-center justify-between gap-3 mb-3">
              <div className="flex items-center gap-2.5">
                <span className="w-7 h-7 rounded-lg bg-blue-600/10 text-blue-600 dark:text-blue-400 font-mono font-bold text-xs flex items-center justify-center">
                  {demo.number}
                </span>
                <h3 className="text-base sm:text-xl font-bold text-slate-900 dark:text-slate-100">
                  {demo.title}
                </h3>
              </div>
              <span className="px-2.5 py-0.5 rounded-full text-xs font-medium bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300">
                {demo.category}
              </span>
            </div>

            <p className="text-xs sm:text-sm text-slate-600 dark:text-slate-300 leading-relaxed mb-4">
              {demo.description}
            </p>

            <CodeBlock
              code={demo.code}
              language="javascript"
              title={`Demo ${demo.number}: ${demo.title}`}
            />
          </article>
        ))}
      </div>
    </div>
  );
}
