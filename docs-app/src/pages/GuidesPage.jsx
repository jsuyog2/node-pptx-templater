import React, { useState, useEffect } from 'react';
import { BookOpen, Clock, ChevronRight, Hash, ArrowLeft } from 'lucide-react';
import CodeBlock from '../components/CodeBlock';
import SEOHead from '../components/SEOHead';
import guidesData from '../data/guides.json';

export default function GuidesPage({ initialGuideSlug }) {
  const [selectedSlug, setSelectedSlug] = useState(initialGuideSlug || 'getting-started');

  useEffect(() => {
    if (initialGuideSlug) {
      setSelectedSlug(initialGuideSlug);
    }
  }, [initialGuideSlug]);

  const activeGuide = guidesData.find((g) => g.slug === selectedSlug) || guidesData[0];

  return (
    <div>
      <SEOHead
        title={`${activeGuide.title} — node-pptx-templater Guides`}
        description={activeGuide.summary}
        canonicalUrl={`https://node-pptx-templater.pages.dev/guides/${activeGuide.slug}`}
        articleData={{
          title: activeGuide.title,
          summary: activeGuide.summary
        }}
      />

      {/* Guide Navigation Header */}
      <div className="mb-8 pb-6 border-b border-slate-200/80 dark:border-slate-800/80">
        <div className="flex flex-wrap items-center gap-2 mb-2">
          <span className="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-purple-500/10 text-purple-600 dark:text-purple-400 border border-purple-500/20">
            Technical Guides
          </span>
          <span className="text-xs text-slate-500">Architecture & Best Practices</span>
        </div>
        <h1 className="text-2xl sm:text-4xl font-extrabold tracking-tight text-slate-900 dark:text-slate-100 mb-3">
          Deep-Dive Architectural Guides
        </h1>
        <p className="text-xs sm:text-sm text-slate-600 dark:text-slate-300 max-w-3xl leading-relaxed">
          Master the internal mechanics of OpenXML PowerPoint manipulation: zero-dependency serverless deployments, AST paragraph normalizers, binary Excel dual-sync caching, and automated table overflow pagination.
        </p>

        {/* Quick Guide Picker Pills */}
        <div className="flex flex-wrap gap-2 mt-6">
          {guidesData.map((guide) => (
            <button
              key={guide.slug}
              onClick={() => setSelectedSlug(guide.slug)}
              className={`px-3 py-1.5 rounded-xl text-xs font-medium transition-all ${
                selectedSlug === guide.slug
                  ? 'bg-blue-600 text-white shadow-sm'
                  : 'bg-slate-100 dark:bg-slate-800/80 text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-700'
              }`}
            >
              {guide.title.split('&')[0].split('(')[0].trim()}
            </button>
          ))}
        </div>
      </div>

      {/* Active Guide Content */}
      <article className="glass-card rounded-3xl p-6 sm:p-10 border border-slate-200/80 dark:border-slate-800/80">
        <div className="flex flex-wrap items-center justify-between gap-3 pb-6 border-b border-slate-200/80 dark:border-slate-800/80">
          <div>
            <div className="flex items-center gap-2 mb-2">
              <span className="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-500/10 text-blue-600 dark:text-blue-400">
                {activeGuide.badge}
              </span>
              <span className="flex items-center gap-1 text-xs text-slate-400">
                <Clock className="w-3.5 h-3.5" /> {activeGuide.readingTime}
              </span>
            </div>
            <h2 className="text-xl sm:text-3xl font-bold text-slate-900 dark:text-slate-100 tracking-tight">
              {activeGuide.title}
            </h2>
          </div>
        </div>

        <p className="text-sm sm:text-base text-slate-600 dark:text-slate-300 leading-relaxed my-6 font-medium">
          {activeGuide.summary}
        </p>

        {/* Guide Sections */}
        <div className="space-y-8 mt-6">
          {activeGuide.sections.map((sec, idx) => (
            <div key={idx} className="scroll-mt-24">
              <h3 className="text-lg sm:text-xl font-bold text-slate-800 dark:text-slate-200 mb-3 flex items-center gap-2">
                <span className="w-1.5 h-5 rounded-full bg-blue-600 inline-block"></span>
                {sec.heading}
              </h3>
              <p className="text-xs sm:text-sm text-slate-600 dark:text-slate-300 leading-relaxed whitespace-pre-line mb-4">
                {sec.text}
              </p>
              {sec.code && (
                <CodeBlock
                  code={sec.code}
                  language={sec.language || 'javascript'}
                  title={sec.heading}
                />
              )}
            </div>
          ))}
        </div>
      </article>
    </div>
  );
}
