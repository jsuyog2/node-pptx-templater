import React from 'react';
import Link from 'next/link';
import { notFound } from 'next/navigation';
import { BookOpen, Clock, ChevronRight, ArrowLeft, ArrowRight, Layers, Sparkles } from 'lucide-react';
import guidesData from '../../../data/guides.json';
import siteMeta from '../../../data/siteMeta.json';
import CodeBlock from '../../../components/CodeBlock';

// Static generation params for Next.js output: 'export'
export async function generateStaticParams() {
  return guidesData.map((guide) => ({
    slug: guide.slug,
  }));
}

// Generate dynamic metadata for SEO
export async function generateMetadata({ params }) {
  const guide = guidesData.find((g) => g.slug === params.slug);
  if (!guide) {
    return {
      title: 'Guide Not Found',
    };
  }

  return {
    title: guide.title,
    description: guide.summary,
    alternates: {
      canonical: `/guides/${guide.slug}/`,
    },
    openGraph: {
      title: `${guide.title} | node-pptx-templater Guide`,
      description: guide.summary,
      url: `${siteMeta.domainUrl}/guides/${guide.slug}/`,
      images: [
        {
          url: '/og-banner.png',
          width: 1200,
          height: 630,
          alt: guide.title,
        },
      ],
    },
    twitter: {
      card: 'summary_large_image',
      title: `${guide.title} | node-pptx-templater Guide`,
      description: guide.summary,
      images: ['/twitter-card.png'],
    },
  };
}

export default function GuidePage({ params }) {
  const guideIndex = guidesData.findIndex((g) => g.slug === params.slug);
  if (guideIndex === -1) {
    notFound();
  }

  const guide = guidesData[guideIndex];
  const prevGuide = guideIndex > 0 ? guidesData[guideIndex - 1] : null;
  const nextGuide = guideIndex < guidesData.length - 1 ? guidesData[guideIndex + 1] : null;

  // JSON-LD TechArticle schema
  const articleSchema = {
    '@context': 'https://schema.org',
    '@type': 'TechArticle',
    headline: guide.title,
    description: guide.summary,
    author: {
      '@type': 'Person',
      name: siteMeta.author,
    },
    publisher: {
      '@type': 'Organization',
      name: 'node-pptx-templater',
      logo: {
        '@type': 'ImageObject',
        url: `${siteMeta.domainUrl}/logo.png`,
      },
    },
    mainEntityOfPage: `${siteMeta.domainUrl}/guides/${guide.slug}/`,
    datePublished: '2026-10-11',
  };

  return (
    <article className="max-w-4xl mx-auto px-4 sm:px-6 lg:px-8 py-10 sm:py-14">
      {/* Schema.org TechArticle Script */}
      <script
        type="application/ld+json"
        dangerouslySetInnerHTML={{ __html: JSON.stringify(articleSchema) }}
      />

      {/* Breadcrumb Navigation */}
      <nav className="flex items-center gap-2 text-xs text-slate-400 mb-6">
        <Link href="/" className="hover:text-white transition-colors">
          Docs
        </Link>
        <ChevronRight className="w-3 h-3 text-slate-600" />
        <span className="text-slate-300">Guides</span>
        <ChevronRight className="w-3 h-3 text-slate-600" />
        <span className="text-brand-400 font-medium truncate">{guide.title}</span>
      </nav>

      {/* Header Info */}
      <header className="mb-10 pb-8 border-b border-slate-800">
        <div className="flex items-center gap-3 mb-4">
          <span className="px-3 py-1 rounded-full text-xs font-semibold bg-brand-500/15 text-brand-400 border border-brand-500/30">
            {guide.badge || 'Guide'}
          </span>
          <span className="flex items-center gap-1.5 text-xs text-slate-400">
            <Clock className="w-3.5 h-3.5" />
            {guide.readingTime || '5 min read'}
          </span>
        </div>

        <h1 className="text-3xl sm:text-4xl lg:text-5xl font-extrabold text-white tracking-tight leading-tight">
          {guide.title}
        </h1>

        <p className="mt-4 text-base sm:text-lg text-slate-300 leading-relaxed">
          {guide.summary}
        </p>
      </header>

      {/* Guide Content Sections */}
      <div className="space-y-12">
        {guide.sections.map((section, idx) => (
          <section key={idx} className="space-y-4">
            <h2 className="text-2xl font-bold text-white tracking-tight flex items-center gap-2.5">
              <span className="w-7 h-7 rounded-lg bg-slate-800 border border-slate-700 text-xs font-mono text-brand-400 flex items-center justify-center">
                {idx + 1}
              </span>
              <span>{section.heading}</span>
            </h2>

            <p className="text-sm sm:text-base text-slate-300 leading-relaxed whitespace-pre-line">
              {section.text}
            </p>

            {section.code && (
              <CodeBlock
                code={section.code}
                language={section.language || 'javascript'}
              />
            )}
          </section>
        ))}
      </div>

      {/* Guide Pagination (Prev / Next) */}
      <nav className="mt-16 pt-8 border-t border-slate-800 grid grid-cols-1 sm:grid-cols-2 gap-4">
        {prevGuide ? (
          <Link
            href={`/guides/${prevGuide.slug}/`}
            className="p-4 rounded-xl border border-slate-800 hover:border-slate-700 bg-slate-900/40 hover:bg-slate-900/80 transition-all flex flex-col group text-left"
          >
            <span className="text-[11px] font-medium text-slate-500 flex items-center gap-1 mb-1">
              <ArrowLeft className="w-3 h-3 group-hover:-translate-x-0.5 transition-transform" />
              Previous Guide
            </span>
            <span className="text-sm font-semibold text-white group-hover:text-brand-300 transition-colors">
              {prevGuide.title}
            </span>
          </Link>
        ) : (
          <div />
        )}

        {nextGuide && (
          <Link
            href={`/guides/${nextGuide.slug}/`}
            className="p-4 rounded-xl border border-slate-800 hover:border-slate-700 bg-slate-900/40 hover:bg-slate-900/80 transition-all flex flex-col group text-right sm:col-start-2"
          >
            <span className="text-[11px] font-medium text-slate-500 flex items-center justify-end gap-1 mb-1">
              Next Guide
              <ArrowRight className="w-3 h-3 group-hover:translate-x-0.5 transition-transform" />
            </span>
            <span className="text-sm font-semibold text-white group-hover:text-brand-300 transition-colors">
              {nextGuide.title}
            </span>
          </Link>
        )}
      </nav>
    </article>
  );
}
