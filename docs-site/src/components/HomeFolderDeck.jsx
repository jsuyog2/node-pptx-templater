'use client';

import React, { useRef, useEffect, useState, useCallback } from 'react';
import Link from 'next/link';
import Image from 'next/image';
import {
  ArrowRight,
  Terminal,
  Copy,
  Check,
  ShieldCheck,
  CheckCircle2,
  BookOpen,
  Code2,
} from 'lucide-react';
import CodeBlock from './CodeBlock';
import comparisonData from '../data/comparison.json';
import featuresData from '../data/features.json';
import {
  slideToFolder,
  checkScrollBoundary,
  findScrollableContainer,
} from '../animations/homeFolderScroll';
import StatusBar from './PowerPointChrome/StatusBar';
import NavigationArrows from './NavigationArrows';

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
await ppt.saveToFile('./output/Q4_Executive_Deck.pptx');`;

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

const FOLDERS = [
  { id: 'hero', title: 'Folder 01 • Hero & Template Engine' },
  { id: 'features', title: 'Folder 02 • Features & Capabilities' },
  { id: 'architecture', title: 'Folder 03 • Architectural Benchmarks & Matrix' },
  { id: 'quickstart', title: 'Folder 04 • Documentation & Quickstart' },
];

export default function HomeFolderDeck() {
  const containerRef = useRef(null);
  const trackRef = useRef(null);
  const cardRefs = useRef([]);
  const [activeFolderIndex, setActiveFolderIndex] = useState(0);
  const [copied, setCopied] = useState(false);

  const [viewMode, setViewMode] = useState('normal');
  const [zoomLevel, setZoomLevel] = useState(100);

  const lastTransitionTimeRef = useRef(0);
  const touchStartYRef = useRef(0);
  const touchStartXRef = useRef(0);
  const wheelAccumulatorRef = useRef(0);
  const lastScrollDirectionRef = useRef(0);
  const boundaryReachedTimeRef = useRef(0);

  const COOLDOWN_MS = 380;
  const WHEEL_THRESHOLD = 50;

  const handleCopyInstall = async () => {
    try {
      await navigator.clipboard.writeText('npm install node-pptx-templater');
      setCopied(true);
      setTimeout(() => setCopied(false), 2000);
    } catch {
      // ignore
    }
  };

  const jumpToFolder = useCallback(
    (index) => {
      const now = Date.now();
      if (now - lastTransitionTimeRef.current < COOLDOWN_MS) return;
      if (index < 0 || index >= FOLDERS.length || index === activeFolderIndex) return;

      lastTransitionTimeRef.current = now;
      setActiveFolderIndex(index);
      slideToFolder(trackRef.current, index, FOLDERS.length, 0.45);

      // Reset scroll position of the destination card to top
      const nextCard = cardRefs.current[index];
      if (nextCard) {
        nextCard.scrollTop = 0;
      }
    },
    [activeFolderIndex]
  );

  const nextFolder = useCallback(() => {
    if (activeFolderIndex < FOLDERS.length - 1) {
      jumpToFolder(activeFolderIndex + 1);
    }
  }, [activeFolderIndex, jumpToFolder]);

  const prevFolder = useCallback(() => {
    if (activeFolderIndex > 0) {
      jumpToFolder(activeFolderIndex - 1);
    }
  }, [activeFolderIndex, jumpToFolder]);

  // Wheel Boundary Snapping with explicit damping
  const handleWheel = useCallback(
    (e) => {
      // On mobile and tablet (< 1024px), disable scroll-jacking to permit natural vertical scrolling
      if (typeof window !== 'undefined' && window.innerWidth < 1024) return;
      if (Math.abs(e.deltaY) < 10) return;

      const currentDirection = e.deltaY > 0 ? 1 : -1;
      if (lastScrollDirectionRef.current !== currentDirection) {
        wheelAccumulatorRef.current = 0;
        lastScrollDirectionRef.current = currentDirection;
        boundaryReachedTimeRef.current = 0;
      }

      const activeCard = cardRefs.current[activeFolderIndex];
      const scrollEl = findScrollableContainer(e.target, activeCard) || activeCard;
      const { isAtTop, isAtBottom, isScrollable } = checkScrollBoundary(scrollEl);
      const now = Date.now();

      if (isScrollable) {
        if (e.deltaY > 0) {
          // Scrolling down: if not at bottom, permit natural vertical scrolling
          if (!isAtBottom) {
            wheelAccumulatorRef.current = 0;
            boundaryReachedTimeRef.current = 0;
            return;
          }

          if (!boundaryReachedTimeRef.current) {
            boundaryReachedTimeRef.current = now;
            wheelAccumulatorRef.current = 0;
            return;
          }

          wheelAccumulatorRef.current += e.deltaY;
          if (
            wheelAccumulatorRef.current < WHEEL_THRESHOLD ||
            now - boundaryReachedTimeRef.current < 200
          ) {
            return;
          }

          wheelAccumulatorRef.current = 0;
          boundaryReachedTimeRef.current = 0;
          e.preventDefault();
          nextFolder();
        } else {
          // Scrolling up: if not at top, permit natural vertical scrolling
          if (!isAtTop) {
            wheelAccumulatorRef.current = 0;
            boundaryReachedTimeRef.current = 0;
            return;
          }

          if (!boundaryReachedTimeRef.current) {
            boundaryReachedTimeRef.current = now;
            wheelAccumulatorRef.current = 0;
            return;
          }

          wheelAccumulatorRef.current += Math.abs(e.deltaY);
          if (
            wheelAccumulatorRef.current < WHEEL_THRESHOLD ||
            now - boundaryReachedTimeRef.current < 200
          ) {
            return;
          }

          wheelAccumulatorRef.current = 0;
          boundaryReachedTimeRef.current = 0;
          e.preventDefault();
          prevFolder();
        }
      } else {
        // Not internally scrollable: transition after threshold
        wheelAccumulatorRef.current += Math.abs(e.deltaY);
        if (wheelAccumulatorRef.current < WHEEL_THRESHOLD) return;

        wheelAccumulatorRef.current = 0;
        boundaryReachedTimeRef.current = 0;

        if (e.deltaY > 0) {
          e.preventDefault();
          nextFolder();
        } else {
          e.preventDefault();
          prevFolder();
        }
      }
    },
    [activeFolderIndex, nextFolder, prevFolder]
  );

  // Touch Gesture Handling with Boundary Snapping
  const handleTouchStart = (e) => {
    if (e.touches && e.touches[0]) {
      touchStartYRef.current = e.touches[0].clientY;
      touchStartXRef.current = e.touches[0].clientX;
    }
  };

  const handleTouchMove = (e) => {
    if (!e.touches || !e.touches[0]) return;
    const currentY = e.touches[0].clientY;
    const currentX = e.touches[0].clientX;
    const deltaY = touchStartYRef.current - currentY; // positive = swipe up = scroll down
    const deltaX = touchStartXRef.current - currentX;

    // Horizontal swipe gesture prioritization
    if (Math.abs(deltaX) > Math.abs(deltaY) && Math.abs(deltaX) > 45) {
      if (deltaX > 0) {
        nextFolder();
      } else {
        prevFolder();
      }
      touchStartXRef.current = currentX;
      return;
    }

    // On mobile and tablet viewports (< 1024px), vertical touch scrolling remains completely
    // native and unhindered to prevent scroll-jacking or accidental folder jumps.
    if (typeof window !== 'undefined' && window.innerWidth < 1024) {
      return;
    }

    if (Math.abs(deltaY) < 25) return;

    const activeCard = cardRefs.current[activeFolderIndex];
    const scrollEl = findScrollableContainer(e.target, activeCard) || activeCard;
    const { isAtTop, isAtBottom, isScrollable } = checkScrollBoundary(scrollEl);

    if (isScrollable) {
      if (deltaY > 0) {
        if (!isAtBottom) return;
        nextFolder();
        touchStartYRef.current = currentY;
      } else {
        if (!isAtTop) return;
        prevFolder();
        touchStartYRef.current = currentY;
      }
    } else {
      if (deltaY > 0) {
        nextFolder();
        touchStartYRef.current = currentY;
      } else {
        prevFolder();
        touchStartYRef.current = currentY;
      }
    }
  };

  // Keyboard navigation
  useEffect(() => {
    const handleKeyDown = (e) => {
      if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') return;
      if (e.key === 'ArrowRight' || e.key === 'ArrowDown' || e.key === 'PageDown') {
        e.preventDefault();
        nextFolder();
      } else if (e.key === 'ArrowLeft' || e.key === 'ArrowUp' || e.key === 'PageUp') {
        e.preventDefault();
        prevFolder();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [nextFolder, prevFolder]);

  // Initial alignment
  useEffect(() => {
    if (trackRef.current) {
      slideToFolder(trackRef.current, 0, FOLDERS.length, 0);
    }
  }, []);

  return (
    <div
      ref={containerRef}
      onWheel={handleWheel}
      onTouchStart={handleTouchStart}
      onTouchMove={handleTouchMove}
      className="relative w-full overflow-hidden min-h-[calc(100vh-104px)] flex flex-col justify-between select-none"
    >
      {/* Horizontal Folders Presentation Track */}
      <div
        ref={trackRef}
        className="flex w-[400%] h-full flex-1 transition-transform"
      >
        {/* ==================== FOLDER 01: HERO & OVERVIEW ==================== */}
        <section className="w-screen h-full flex-shrink-0 px-4 sm:px-6 lg:px-8 py-3 sm:py-6 flex items-center justify-center overflow-hidden">
          <div
            ref={(el) => (cardRefs.current[0] = el)}
            style={{ transform: `scale(${zoomLevel / 100})`, transformOrigin: 'top center' }}
            className="w-full max-w-7xl ppt-card-surface rounded-xl overflow-hidden shadow-2xl border border-[var(--ppt-card-border)] max-h-[calc(100vh-170px)] overflow-y-auto overscroll-contain overscroll-y-contain isolate transition-transform duration-150"
          >
            {/* Folder Chrome Header */}
            <div className="h-8 bg-black/20 dark:bg-black/40 border-b border-[var(--ppt-card-border)] flex items-center justify-between px-3 text-xs flex-shrink-0 sticky top-0 z-10 backdrop-blur-md">
              <div className="flex items-center gap-2">
                <div className="flex gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-[#D24726] inline-block opacity-80" />
                  <span className="w-2.5 h-2.5 rounded-full bg-amber-500/80 inline-block" />
                  <span className="w-2.5 h-2.5 rounded-full bg-emerald-500/80 inline-block" />
                </div>
                <span className="font-semibold text-[11px] text-[var(--ppt-text-secondary)]">
                  {FOLDERS[0].title}
                </span>
              </div>
            </div>

            <div className="p-4 sm:p-8 lg:p-10 space-y-6">
              <div className="grid grid-cols-1 lg:grid-cols-2 justify-center items-center gap-6">
                <div className="text-center lg:text-start space-y-4">
                  {/* Headline */}
                  <div className="flex items-center justify-center lg:justify-start gap-3">
                    <div className="relative w-14 h-14 sm:w-20 sm:h-20 flex-shrink-0">
                      <Image
                        src="/logo.png"
                        fill
                        sizes="(max-width: 640px) 56px, 80px"
                        priority
                        alt="Logo"
                        className="object-contain"
                      />
                    </div>
                    <h1 className="text-2xl sm:text-4xl lg:text-5xl font-black tracking-tight text-[var(--ppt-text-primary)] leading-[1.15]">
                      Automate Your
                      <span className="block mt-1 text-[#D24726]">
                        Slide Generation
                      </span>
                    </h1>
                  </div>

                  {/* Subtitle */}
                  <p className="text-sm sm:text-base text-[var(--ppt-text-secondary)] max-w-2xl leading-relaxed">
                    Programmatically generate and edit <code className="font-mono text-[var(--ppt-text-primary)]">.pptx</code> decks. Dynamically inject text, media, tables, and charts into blank or pre-designed templates, then export to any format.
                  </p>

                  {/* Install Pill */}
                  <div className="flex flex-wrap items-center justify-center lg:justify-start gap-3 pt-1">
                    <button
                      type="button"
                      onClick={handleCopyInstall}
                      className="flex items-center gap-2 px-4 py-2.5 rounded-lg text-xs font-mono bg-black/10 dark:bg-black/30 border border-[var(--ppt-card-border)] hover:border-[#D24726] text-[var(--ppt-text-primary)] transition-all cursor-pointer"
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
                  <div className="flex flex-wrap items-center justify-center lg:justify-start gap-3 pt-1">
                    <Link
                      href="/docs"
                      className="flex items-center gap-2 px-5 py-2.5 rounded-lg text-xs font-semibold bg-[#D24726] text-white hover:bg-[#b83b1e] shadow-lg shadow-[#D24726]/20 transition-all hover:scale-[1.02]"
                    >
                      <span>Explore Documentation</span>
                      <ArrowRight className="w-3.5 h-3.5" />
                    </Link>

                    <Link
                      href="/docs/installation"
                      className="flex items-center gap-2 px-5 py-2.5 rounded-lg text-xs font-semibold ppt-card-surface text-[var(--ppt-text-primary)] hover:border-[#D24726] transition-all"
                    >
                      <Terminal className="w-3.5 h-3.5 text-[#D24726]" />
                      <span>Installation &amp; Setup</span>
                    </Link>
                  </div>
                </div>

                {/* Active Recipe Code Block */}
                <div className="w-full overflow-hidden">
                  <CodeBlock
                    code={CODE_EXAMPLES}
                    language="javascript"
                    filename="example.js"
                  />
                </div>
              </div>
            </div>
          </div>
        </section>

        {/* ==================== FOLDER 02: FEATURES ==================== */}
        <section className="w-screen h-full flex-shrink-0 px-4 sm:px-6 lg:px-8 py-3 sm:py-6 flex items-center justify-center overflow-hidden">
          <div
            ref={(el) => (cardRefs.current[1] = el)}
            style={{ transform: `scale(${zoomLevel / 100})`, transformOrigin: 'top center' }}
            className="w-full max-w-7xl ppt-card-surface rounded-xl overflow-hidden shadow-2xl border border-[var(--ppt-card-border)] max-h-[calc(100vh-170px)] overflow-y-auto overscroll-contain overscroll-y-contain isolate transition-transform duration-150"
          >
            {/* Folder Chrome Header */}
            <div className="h-8 bg-black/20 dark:bg-black/40 border-b border-[var(--ppt-card-border)] flex items-center justify-between px-3 text-xs flex-shrink-0 sticky top-0 z-10 backdrop-blur-md">
              <div className="flex items-center gap-2">
                <div className="flex gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-[#D24726] inline-block opacity-80" />
                  <span className="w-2.5 h-2.5 rounded-full bg-amber-500/80 inline-block" />
                  <span className="w-2.5 h-2.5 rounded-full bg-emerald-500/80 inline-block" />
                </div>
                <span className="font-semibold text-[11px] text-[var(--ppt-text-secondary)]">
                  {FOLDERS[1].title}
                </span>
              </div>
            </div>

            <div className="p-4 sm:p-8 space-y-6">
              <div className="text-center max-w-2xl mx-auto mb-4">
                <h2 className="text-2xl sm:text-3xl font-extrabold text-[var(--ppt-text-primary)] tracking-tight">
                  Built for Speed. Designed for Flexibility.
                </h2>
                <p className="text-sm text-[var(--ppt-text-secondary)] mt-1">
                  Why Developers Choose <span className="text-[#D24726] font-semibold">node-pptx-templater</span>
                </p>
              </div>

              {/* Feature Cards Grid */}
              <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                {featuresData.map((mod) => (
                  <div
                    key={mod.title}
                    className="ppt-card-surface p-4 sm:p-5 rounded-xl border border-[var(--ppt-card-border)] hover:border-[#D24726] transition-all flex flex-col justify-between"
                  >
                    <div>
                      <h3 className="font-bold text-sm sm:text-base text-[var(--ppt-text-primary)] hover:text-[#D24726] transition-colors">
                        {mod.title}
                      </h3>
                      <p className="text-xs text-[var(--ppt-text-secondary)] mt-1.5 leading-relaxed">
                        {mod.desc}
                      </p>
                    </div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </section>

        {/* ==================== FOLDER 03: ARCHITECTURE & BENCHMARKS ==================== */}
        <section className="w-screen h-full flex-shrink-0 px-4 sm:px-6 lg:px-8 py-3 sm:py-6 flex items-center justify-center overflow-hidden">
          <div
            ref={(el) => (cardRefs.current[2] = el)}
            style={{ transform: `scale(${zoomLevel / 100})`, transformOrigin: 'top center' }}
            className="w-full max-w-7xl ppt-card-surface rounded-xl overflow-hidden shadow-2xl border border-[var(--ppt-card-border)] max-h-[calc(100vh-170px)] overflow-y-auto overscroll-contain overscroll-y-contain isolate transition-transform duration-150"
          >
            {/* Folder Chrome Header */}
            <div className="h-8 bg-black/20 dark:bg-black/40 border-b border-[var(--ppt-card-border)] flex items-center justify-between px-3 text-xs flex-shrink-0 sticky top-0 z-10 backdrop-blur-md">
              <div className="flex items-center gap-2">
                <div className="flex gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-[#D24726] inline-block opacity-80" />
                  <span className="w-2.5 h-2.5 rounded-full bg-amber-500/80 inline-block" />
                  <span className="w-2.5 h-2.5 rounded-full bg-emerald-500/80 inline-block" />
                </div>
                <span className="font-semibold text-[11px] text-[var(--ppt-text-secondary)]">
                  {FOLDERS[2].title}
                </span>
              </div>
            </div>

            <div className="p-4 sm:p-8 space-y-6">
              <div className="flex flex-wrap items-center justify-between gap-3 mb-2">
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
                      <th className="py-2.5 px-3 font-semibold">Capability / Metric</th>
                      <th className="py-2.5 px-3 font-bold text-[#D24726] bg-[#D24726]/10 rounded-t">
                        node-pptx-templater
                      </th>
                      <th className="py-2.5 px-3 font-semibold">Headless LibreOffice</th>
                      <th className="py-2.5 px-3 font-semibold">Apache POI (Java)</th>
                      <th className="py-2.5 px-3 font-semibold">python-pptx</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-[var(--ppt-card-border)]/50">
                    {comparisonData.map((row, idx) => (
                      <tr key={idx} className="hover:bg-black/5 dark:hover:bg-white/5 transition-colors">
                        <td className="py-2 px-3 font-medium text-[var(--ppt-text-primary)]">
                          {row.feature}
                        </td>
                        <td className="py-2 px-3 font-bold text-[#D24726] bg-[#D24726]/5">
                          <div className="flex items-center gap-1.5">
                            <CheckCircle2 className="w-3.5 h-3.5 text-emerald-500 flex-shrink-0" />
                            <span>{row.pptxTemplater}</span>
                          </div>
                        </td>
                        <td className="py-2 px-3 text-[var(--ppt-text-secondary)]">
                          {row.libreOffice}
                        </td>
                        <td className="py-2 px-3 text-[var(--ppt-text-secondary)]">
                          {row.poi}
                        </td>
                        <td className="py-2 px-3 text-[var(--ppt-text-secondary)]">
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

        {/* ==================== FOLDER 04: DOCUMENTATION & QUICKSTART ==================== */}
        <section className="w-screen h-full flex-shrink-0 px-4 sm:px-6 lg:px-8 py-3 sm:py-6 flex items-center justify-center overflow-hidden">
          <div
            ref={(el) => (cardRefs.current[3] = el)}
            style={{ transform: `scale(${zoomLevel / 100})`, transformOrigin: 'top center' }}
            className="w-full max-w-7xl ppt-card-surface rounded-xl overflow-hidden shadow-2xl border border-[var(--ppt-card-border)] max-h-[calc(100vh-170px)] overflow-y-auto overscroll-contain overscroll-y-contain isolate transition-transform duration-150"
          >
            {/* Folder Chrome Header */}
            <div className="h-8 bg-black/20 dark:bg-black/40 border-b border-[var(--ppt-card-border)] flex items-center justify-between px-3 text-xs flex-shrink-0 sticky top-0 z-10 backdrop-blur-md">
              <div className="flex items-center gap-2">
                <div className="flex gap-1.5">
                  <span className="w-2.5 h-2.5 rounded-full bg-[#D24726] inline-block opacity-80" />
                  <span className="w-2.5 h-2.5 rounded-full bg-amber-500/80 inline-block" />
                  <span className="w-2.5 h-2.5 rounded-full bg-emerald-500/80 inline-block" />
                </div>
                <span className="font-semibold text-[11px] text-[var(--ppt-text-secondary)]">
                  {FOLDERS[3].title}
                </span>
              </div>
            </div>

            <div className="p-4 sm:p-8 space-y-6">
              <div className="text-center max-w-2xl mx-auto mb-4">
                <h2 className="text-2xl sm:text-3xl font-extrabold text-[var(--ppt-text-primary)] tracking-tight">
                  Comprehensive Documentation Suite
                </h2>
                <p className="text-sm text-[var(--ppt-text-secondary)] mt-1">
                  Structured into focused PowerPoint slide presentations for instant developer productivity.
                </p>
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
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
      </div>

      {/* Floating Directional Arrow Navigation Controls (Accessible on Mobile, Tablet & Desktop) */}
      <div className="fixed bottom-12 sm:bottom-14 right-4 sm:right-6 z-30">
        <NavigationArrows
          currentIndex={activeFolderIndex}
          totalItems={FOLDERS.length}
          onPrev={prevFolder}
          onNext={nextFolder}
          label="Folder"
        />
      </div>

      {/* Bottom Status Bar (Desktop) */}
      <StatusBar
        currentSlide={activeFolderIndex + 1}
        totalSlides={FOLDERS.length}
        viewMode={viewMode}
        onViewModeChange={setViewMode}
        zoomLevel={zoomLevel}
        onZoomChange={setZoomLevel}
        className="hidden lg:flex"
      />
    </div>
  );
}
