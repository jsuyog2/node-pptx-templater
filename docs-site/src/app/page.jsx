'use client';

import React from 'react';
import Link from 'next/link';
import {
  ArrowRight,
  Zap,
  Cpu,
  Layers,
  FileSpreadsheet,
  CheckCircle2,
  XCircle,
  Sparkles,
  BookOpen,
  Terminal,
  Code2,
  ShieldCheck
} from 'lucide-react';
import HeroSection from '../components/HeroSection';
import StatusBar from '../components/PowerPointChrome/StatusBar';
import comparisonData from '../data/comparison.json';

const CORE_MODULES = [
  {
    title: 'Overview & Philosophy',
    href: '/docs',
    icon: BookOpen,
    desc: 'Understand the in-memory VFS, AST run healing, and template-driven mutation paradigm.',
  },
  {
    title: 'Installation & Setup',
    href: '/docs/installation',
    icon: Terminal,
    desc: 'Quickstart for npm, pnpm, yarn, dual ESM/CJS, and native TypeScript declarations.',
  },
  {
    title: 'Interactive API Reference',
    href: '/docs/api-reference',
    icon: Code2,
    desc: 'Complete catalog of 135 methods strictly organized across the 8 standard categories.',
  },
];

export default function HomePage() {
  return (
    <div className="pb-16">
      {/* Hero Section with ScrollTrigger & Recipe Playground */}
      <HeroSection />

      {/* Primary 4-Route Navigation Cards */}
      <section className="relative pt-6 pb-16 px-4 sm:px-6 lg:px-8 max-w-7xl mx-auto">
        <div className="hero-window-chrome ppt-card-surface rounded-xl overflow-hidden shadow-2xl border border-[var(--ppt-card-border)] mb-12">
          {/* Fake Desktop Title Bar */}
          <div className="h-8 bg-black/20 dark:bg-black/40 border-b border-[var(--ppt-card-border)] flex items-center justify-between px-3 text-xs">
            <div className="flex items-center gap-2">
              <div className="flex gap-1.5">
                <span className="w-2.5 h-2.5 rounded-full bg-[#D24726] inline-block opacity-80" />
                <span className="w-2.5 h-2.5 rounded-full bg-amber-500/80 inline-block" />
                <span className="w-2.5 h-2.5 rounded-full bg-emerald-500/80 inline-block" />
              </div>
              <span className="font-semibold text-[11px] text-[var(--ppt-text-secondary)]">
                Documentation • PowerPoint Template Engine
              </span>
            </div>
          </div>
          <div className="p-2 sm:p-4 lg:p-8 text-start mx-auto space-y-6">
            <div className="text-center max-w-2xl mx-auto mb-8">
              <h2 className="text-2xl sm:text-3xl font-extrabold text-[var(--ppt-text-primary)] tracking-tight">
                Comprehensive Documentation Suite
              </h2>
              <p className="text-sm text-[var(--ppt-text-secondary)] mt-2">
                Structured into 3 focused PowerPoint slide presentations for instant developer productivity.
              </p>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-3 lg:grid-cols-3 gap-4">
              {CORE_MODULES.map((mod) => {
                const Icon = mod.icon;
                return (
                  <Link
                    key={mod.href}
                    href={mod.href}
                    className="ppt-card-surface p-5 rounded-xl border border-[var(--ppt-card-border)] hover:border-[#D24726] transition-all group flex flex-col justify-between"
                  >
                    <div>
                      <div className="w-10 h-10 rounded-lg bg-[#D24726]/10 text-[#D24726] flex items-center justify-center mb-3 group-hover:scale-110 transition-transform">
                        <Icon className="w-5 h-5" />
                      </div>
                      <h3 className="font-bold text-base text-[var(--ppt-text-primary)] group-hover:text-[#D24726] transition-colors">
                        {mod.title}
                      </h3>
                      <p className="text-xs text-[var(--ppt-text-secondary)] mt-1.5 leading-relaxed">
                        {mod.desc}
                      </p>
                    </div>
                    <div className="flex items-center gap-1 text-xs font-semibold text-[#D24726] mt-4 pt-3 border-t border-[var(--ppt-card-border)]/50">
                      <span>Open Deck</span>
                      <ArrowRight className="w-3.5 h-3.5 group-hover:translate-x-1 transition-transform" />
                    </div>
                  </Link>
                );
              })}
            </div>
          </div>
        </div>
      </section>
      {/* Architecture & Performance Comparison Matrix */}
      <section className="relative pt-6 pb-16 px-4 sm:px-6 lg:px-8 max-w-7xl mx-auto">
        <div className="hero-window-chrome ppt-card-surface rounded-xl overflow-hidden shadow-2xl border border-[var(--ppt-card-border)] mb-12">
          {/* Fake Desktop Title Bar */}
          <div className="h-8 bg-black/20 dark:bg-black/40 border-b border-[var(--ppt-card-border)] flex items-center justify-between px-3 text-xs">
            <div className="flex items-center gap-2">
              <div className="flex gap-1.5">
                <span className="w-2.5 h-2.5 rounded-full bg-[#D24726] inline-block opacity-80" />
                <span className="w-2.5 h-2.5 rounded-full bg-amber-500/80 inline-block" />
                <span className="w-2.5 h-2.5 rounded-full bg-emerald-500/80 inline-block" />
              </div>
              <span className="font-semibold text-[11px] text-[var(--ppt-text-secondary)]">
                Comparison • PowerPoint Template Engine
              </span>
            </div>
          </div>
          <div className="p-6 sm:p-8 overflow-hidden">
            <div className="mb-6 flex flex-wrap items-center justify-between gap-4">
              <div>
                <div className="inline-flex items-center gap-1.5 text-xs font-semibold text-[#D24726] mb-1">
                  <ShieldCheck className="w-4 h-4" />
                  <span>Architectural Benchmarks</span>
                </div>
                <h2 className="text-xl sm:text-2xl font-bold text-[var(--ppt-text-primary)]">
                  Engine Comparison vs Legacy Alternatives
                </h2>
              </div>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs border-collapse">
                <thead>
                  <tr className="border-b border-[var(--ppt-card-border)] text-[var(--ppt-text-secondary)]">
                    <th className="py-3 px-4 font-semibold">Capability / Metric</th>
                    <th className="py-3 px-4 font-bold text-[#D24726] bg-[#D24726]/10 rounded-t">
                      node-pptx-templater
                    </th>
                    <th className="py-3 px-4 font-semibold">Headless LibreOffice</th>
                    <th className="py-3 px-4 font-semibold">Apache POI (Java)</th>
                    <th className="py-3 px-4 font-semibold">python-pptx</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-[var(--ppt-card-border)]/50">
                  {comparisonData.map((row, idx) => (
                    <tr key={idx} className="hover:bg-black/5 dark:hover:bg-white/5 transition-colors">
                      <td className="py-3 px-4 font-medium text-[var(--ppt-text-primary)]">
                        {row.feature}
                      </td>
                      <td className="py-3 px-4 font-bold text-[#D24726] bg-[#D24726]/5">
                        <div className="flex items-center gap-1.5">
                          <CheckCircle2 className="w-3.5 h-3.5 text-emerald-500 flex-shrink-0" />
                          <span>{row.pptxTemplater}</span>
                        </div>
                      </td>
                      <td className="py-3 px-4 text-[var(--ppt-text-secondary)]">
                        {row.libreOffice}
                      </td>
                      <td className="py-3 px-4 text-[var(--ppt-text-secondary)]">
                        {row.poi}
                      </td>
                      <td className="py-3 px-4 text-[var(--ppt-text-secondary)]">
                        {row.pythonPptx}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        </div>
      </section>

      {/* Global Status Bar */}
      <StatusBar
        currentSlide={1}
        totalSlides={4}
        viewMode="normal"
        zoomLevel={100}
      />
    </div>
  );
}
