'use client';

import React, { useRef, useEffect, useState } from 'react';
import Link from 'next/link';
import {
  ArrowRight,
  Sparkles,
  Zap,
  Cpu,
  Layers,
  FileSpreadsheet,
  Terminal,
  Copy,
  Check,
  Play,
  ShieldCheck,
  CheckCircle2
} from 'lucide-react';
import { animateHeroScrollTrigger } from '../animations/gsapSlideTransitions';
import CORE_MODULES from '../data/features.json';
import CodeBlock from './CodeBlock';
import Image from 'next/image';

const CODE_EXAMPLES = `import { PPTXTemplater } from 'node-pptx-templater';

// 1. Load PowerPoint template in memory (Zero /tmp disk I/O)
const ppt = await PPTXTemplater.load('./templates/executive-deck.pptx');

// 2. Perform safe, AST run-fragment healed text replacement
ppt.useSlide(1).replaceText({
  '{{COMPANY_NAME}}': 'Acme Global AI Corp',
  '{{QUARTER}}': 'Q4 2026',
  '{{REVENUE}}': '$52.4M ARR'
});

// 3. Export to file, buffer, or stream with zero repair warnings
await ppt.saveToFile('./output/Q4_Executive_Deck.pptx');`
export default function HeroSection() {
  const containerRef = useRef(null);
  const [activeTab, setActiveTab] = useState('text');
  const [copied, setCopied] = useState(false);

  useEffect(() => {
    const cleanup = animateHeroScrollTrigger(containerRef);
    return cleanup;
  }, []);

  const handleCopyInstall = async () => {
    try {
      await navigator.clipboard.writeText('npm install node-pptx-templater');
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      // ignore
    }
  };

  return (
    <section ref={containerRef} className="relative pt-6 pb-16 px-4 sm:px-6 lg:px-8 max-w-7xl mx-auto">
      {/* Decorative ambient background glows */}
      <div className="absolute top-1/4 left-1/2 -translate-x-1/2 -translate-y-1/2 w-[600px] h-[350px] bg-[#D24726]/10 blur-[120px] rounded-full pointer-events-none -z-10" />

      {/* Hero Window Chrome Wrapper */}
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
              PowerPoint Template Engine
            </span>
          </div>
        </div>

        {/* Hero Header Content */}
        <div className="p-4 sm:p-8 lg:p-12 text-center mx-auto space-y-6">
          <div className='grid grid-cols-1 lg:grid-cols-2 justify-center items-center gap-4'>
            <div className='text-center lg:text-start'>
              {/* Headline */}
              <div className="flex items-center justify-center lg:justify-start gap-2">
                <div className="relative w-16 h-16 sm:w-24 sm:h-24">
                  <Image
                    src="/logo.png"
                    fill
                    sizes="(max-width: 640px) 64px, 96px"
                    priority
                    alt="Logo"
                    className="object-contain"
                  />
                </div>
                <h1 className="hero-headline text-2xl sm:text-4xl lg:text-5xl font-black tracking-tight text-[var(--ppt-text-primary)] leading-[1.15]">
                  Automate Your
                  <span className="block mt-1 text-[#D24726]">
                    Slide Generation
                  </span>
                </h1>
              </div>

              {/* Subtitle */}
              <p className="hero-subheadline text-sm sm:text-md lg:text-lg text-base text-[var(--ppt-text-secondary)] max-w-2xl mx-auto leading-relaxed">
                Programmatically generate and edit <code className="font-mono text-[var(--ppt-text-primary)]">.pptx</code> decks. Dynamically inject text, media, tables, and charts into blank or pre-designed templates, then export to any format.
              </p>
              <div className="flex flex-wrap items-center justify-center lg:justify-start gap-3 pt-2">
                <button
                  type="button"
                  onClick={handleCopyInstall}
                  className="flex items-center gap-2 px-4 py-3 rounded-lg text-xs font-mono bg-black/10 dark:bg-black/30 border border-[var(--ppt-card-border)] hover:border-[#D24726] text-[var(--ppt-text-primary)] transition-all"
                  title="Click to copy npm install command"
                >
                  {copied ? (
                    <>
                      <Check className="w-4 h-4 text-emerald-500" />
                      <span className="text-emerald-500 font-sans font-medium">Copied to clipboard!</span>
                    </>
                  ) : (
                    <>
                      <span className="text-[#D24726] font-bold">$</span>
                      <span>npm i node-pptx-templater</span>
                      <Copy className="w-3.5 h-3.5 text-[var(--ppt-text-secondary)]" />
                    </>
                  )}
                </button>
              </div>
              {/* Action CTAs */}
              <div className="hero-action-buttons flex flex-wrap items-center justify-center lg:justify-start gap-3 pt-2">
                <Link
                  href="/docs"
                  className="flex items-center gap-2 px-6 py-3 rounded-lg text-sm font-semibold bg-[#D24726] text-white hover:bg-[#b83b1e] shadow-lg shadow-[#D24726]/20 transition-all hover:scale-[1.02]"
                >
                  <span>Explore Documentation</span>
                  <ArrowRight className="w-4 h-4" />
                </Link>

                <Link
                  href="/docs/installation"
                  className="flex items-center gap-2 px-6 py-3 rounded-lg text-sm font-semibold ppt-card-surface text-[var(--ppt-text-primary)] hover:border-[#D24726] transition-all"
                >
                  <Terminal className="w-4 h-4 text-[#D24726]" />
                  <span>Installation &amp; Setup</span>
                </Link>
              </div>
            </div>


            {/* Active Recipe Code Block */}
            <CodeBlock
              code={CODE_EXAMPLES}
              language="javascript"
              filename={`example.js`}
            />
          </div>
        </div>
      </div>
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
              Features • PowerPoint Template Engine
            </span>
          </div>
        </div>
        <div className="p-2 sm:p-4 lg:p-8 text-start mx-auto space-y-6">
          <div className="text-center max-w-2xl mx-auto mb-8">
            <h2 className="text-2xl sm:text-3xl font-extrabold text-[var(--ppt-text-primary)] tracking-tight">
              Built for Speed. Designed for Flexibility.
            </h2>
            <p className="text-sm text-[var(--ppt-text-secondary)] mt-2">
              Why Developers Choose <span className="text-[#D24726]">node-pptx-templater</span>
            </p>
          </div>
          {/* Bottom Feature Stats Ribbon */}
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
            {CORE_MODULES.map((mod) => {
              return (
                <div
                  key={mod.title}
                  className="ppt-card-surface p-5 rounded-xl border border-[var(--ppt-card-border)] hover:border-[#D24726] transition-all group flex flex-col justify-between"
                >
                  <div>
                    <h3 className="font-bold text-base text-[var(--ppt-text-primary)] group-hover:text-[#D24726] transition-colors">
                      {mod.title}
                    </h3>
                    <p className="text-xs text-[var(--ppt-text-secondary)] mt-1.5 leading-relaxed">
                      {mod.desc}
                    </p>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      </div>
    </section >
  );
}
