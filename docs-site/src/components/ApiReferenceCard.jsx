'use client';

import React, { useState, useMemo, useEffect, useRef } from 'react';
import { Search, ChevronRight, X } from 'lucide-react';
import CodeBlock from './CodeBlock';
import MobileFunctionReel from './MobileFunctionReel';

/**
 * ApiReferenceCard
 * Presentation card dedicated solely to API function details with strict containment
 *
 * @param {{
 *   categoryName: string,
 *   categorySlug: string,
 *   categoryDescription?: string,
 *   methods: Array<any>,
 *   initialMethodName?: string,
 *   slideNumber?: number,
 *   totalSlides?: number
 * }} props
 */
export default function ApiReferenceCard({
  categoryName,
  categorySlug,
  categoryDescription = '',
  methods = [],
  initialMethodName = '',
  slideNumber = 1,
  totalSlides = 8,
}) {
  const [search, setSearch] = useState('');
  const [selectedMethodName, setSelectedMethodName] = useState(
    initialMethodName || methods[0]?.name || ''
  );
  const detailContainerRef = useRef(null);

  // Sync with URL hash for deep-linking
  useEffect(() => {
    const handleHashSync = () => {
      if (typeof window === 'undefined') return;
      const hash = window.location.hash.replace(/^#/, '').trim();
      if (hash && methods.some((m) => m.name === hash)) {
        setSelectedMethodName(hash);
      }
    };

    handleHashSync();
    window.addEventListener('hashchange', handleHashSync);
    return () => window.removeEventListener('hashchange', handleHashSync);
  }, [methods]);

  // If initialMethodName changes from prop
  useEffect(() => {
    if (initialMethodName && methods.some((m) => m.name === initialMethodName)) {
      setSelectedMethodName(initialMethodName);
    }
  }, [initialMethodName, methods]);

  // Filter methods based on search input
  const filteredMethods = useMemo(() => {
    if (!search.trim()) return methods;
    const q = search.toLowerCase();
    return methods.filter(
      (m) =>
        m.name.toLowerCase().includes(q) ||
        (m.signature && m.signature.toLowerCase().includes(q)) ||
        (m.description && m.description.toLowerCase().includes(q))
    );
  }, [methods, search]);

  // Currently active method
  const activeMethod = useMemo(() => {
    return (
      filteredMethods.find((m) => m.name === selectedMethodName) ||
      filteredMethods[0] ||
      methods[0] ||
      {}
    );
  }, [filteredMethods, selectedMethodName, methods]);

  const handleSelectMethod = (name) => {
    setSelectedMethodName(name);
    if (typeof window !== 'undefined') {
      window.history.replaceState(null, '', `#${name}`);
    }
    // Scroll detail container back to top smoothly when method changes
    if (detailContainerRef.current) {
      detailContainerRef.current.scrollTop = 0;
    }
  };

  // Helper to stop touch propagation on horizontally scrollable elements
  const stopTouchPropagation = (e) => {
    e.stopPropagation();
  };

  return (
    <div className="flex flex-col flex-1 min-h-0 min-w-0 w-full max-w-full overflow-hidden overscroll-contain isolate">
      {/* =========================================================================
          Top Header Bar of the Slide Card:
          Unified header with Category Name + Search Input alongside each other,
          Slide indicator on the right, and Category Description beneath
         ========================================================================= */}
      <div className="pb-3 border-b border-[var(--ppt-card-border)]/70 flex-shrink-0 space-y-2">
        <div className="flex flex-wrap items-center justify-between gap-3">
          {/* Category Badge & Title */}
          <div className="flex items-center gap-2.5 flex-shrink-0">
            <span className="px-2.5 py-1 rounded text-xs font-bold bg-[#D24726] text-white tracking-wide shadow-sm">
              {categoryName}
            </span>
          </div>

          {/* Integrated Search / Filter Input alongside Category Name */}
          <div className="relative flex-1 sm:flex-initial sm:w-72 max-w-full">
            <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-[var(--ppt-text-secondary)] opacity-60 pointer-events-none" />
            <input
              type="text"
              placeholder={`Filter ${categoryName} methods...`}
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              className="w-full pl-8 pr-7 py-1.5 rounded text-xs ppt-card-surface border border-[var(--ppt-card-border)] text-[var(--ppt-text-primary)] placeholder-[var(--ppt-text-secondary)]/60 focus:border-[#D24726] focus:outline-none transition-colors"
            />
            {search && (
              <button
                type="button"
                onClick={() => setSearch('')}
                className="absolute right-2 top-1/2 -translate-y-1/2 text-[var(--ppt-text-secondary)] hover:text-[var(--ppt-text-primary)] cursor-pointer p-0.5"
                aria-label="Clear method filter"
              >
                <X className="w-3 h-3" />
              </button>
            )}
          </div>

          {/* Slide Indicator Badge */}
          <div className="hidden sm:flex items-center gap-2 text-xs text-[var(--ppt-text-secondary)] flex-shrink-0">
            <span className="font-mono text-[11px]">
              Slide #{slideNumber} of {totalSlides}
            </span>
            <div className="w-2 h-2 rounded-full bg-[#D24726]" />
          </div>
        </div>

        {/* Category Description cleanly beneath the category title */}
        {categoryName && (
          <h1 className="text-xl sm:text-2xl lg:text-3xl font-extrabold text-[var(--ppt-text-primary)] tracking-tight">
            {categoryName}
          </h1>
        )}
        {/* Category Description cleanly beneath the category title */}
        {categoryDescription && (
          <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)] leading-relaxed max-w-4xl">
            {categoryDescription}
          </p>
        )}
      </div>

      {/* =========================================================================
          Mobile Horizontal Function Reel (< 768px):
          Horizontal scrolling through method chips with touch isolation
         ========================================================================= */}
      <div
        className="flex-shrink-0"
        onTouchStart={stopTouchPropagation}
        onTouchMove={stopTouchPropagation}
        onTouchEnd={stopTouchPropagation}
      >
        <MobileFunctionReel
          methods={filteredMethods}
          activeMethodName={activeMethod.name}
          onSelectMethod={handleSelectMethod}
        />
      </div>

      {/* =========================================================================
          Main Content Layout:
          Two-column grid with strict containment (No overflow outside the slide card)
         ========================================================================= */}
      <div className="grid grid-cols-1 md:grid-cols-12 gap-3 sm:gap-4 pt-3 flex-1 min-h-0 min-w-0 overflow-hidden w-full max-w-full">
        {/* Desktop Method Selector Sidebar */}
        <div className="hidden md:flex md:col-span-4 lg:col-span-3 flex-col h-full min-h-0 overflow-y-auto space-y-1.5 pr-1 overscroll-contain isolate">
          {filteredMethods.map((m) => {
            const isSelected = m.name === activeMethod.name;
            return (
              <button
                key={m.name}
                type="button"
                onClick={() => handleSelectMethod(m.name)}
                className={`w-full text-left px-3 py-2 rounded text-xs transition-all flex items-center justify-between border cursor-pointer ${isSelected
                  ? 'border-[#D24726] bg-[#D24726]/10 text-[var(--ppt-text-primary)] font-semibold shadow-sm'
                  : 'border-transparent hover:border-[var(--ppt-card-border)] text-[var(--ppt-text-secondary)] hover:text-[var(--ppt-text-primary)] hover:bg-black/5 dark:hover:bg-white/5'
                  }`}
              >
                <div className="truncate mr-2">
                  <div className="font-mono text-[11px] truncate">{m.name}()</div>
                  <div className="text-[10px] opacity-70 truncate">{m.description}</div>
                </div>
                <ChevronRight
                  className={`w-3.5 h-3.5 flex-shrink-0 transition-transform ${isSelected ? 'text-[#D24726] translate-x-0.5' : 'opacity-40'
                    }`}
                />
              </button>
            );
          })}
          {filteredMethods.length === 0 && (
            <div className="text-center py-8 text-xs text-[var(--ppt-text-secondary)]">
              No matching methods found.
            </div>
          )}
        </div>

        {/* Function Details Card: Strictly Contained Viewport */}
        <div
          ref={detailContainerRef}
          className="col-span-1 md:col-span-8 lg:col-span-9 ppt-card-surface p-4 sm:p-5 rounded-lg border border-[var(--ppt-card-border)] space-y-4 h-full min-h-0 overflow-y-auto overscroll-contain isolate flex-1 min-w-0 max-w-full"
        >
          {activeMethod.name ? (
            <>
              {/* Method Title & Signature (Clean, with static/async pill badges removed) */}
              <div id={activeMethod.name} className="scroll-mt-4">
                <div className="flex items-center justify-between gap-2 flex-wrap">
                  <h3 className="font-mono text-base sm:text-lg font-bold text-[var(--ppt-text-primary)] flex items-center gap-1.5">
                    <span className="text-[#D24726] font-normal">
                      {activeMethod.isStatic ? 'PPTXTemplater.' : 'ppt.'}
                    </span>
                    <span>{activeMethod.name}()</span>
                  </h3>
                </div>

                {/* Full Signature Box with Horizontal Scrolling (touch-isolated) */}
                <div
                  onTouchStart={stopTouchPropagation}
                  onTouchMove={stopTouchPropagation}
                  onTouchEnd={stopTouchPropagation}
                  className="font-mono text-xs text-[#D24726] mt-2 bg-black/20 dark:bg-black/40 p-2.5 rounded border border-[var(--ppt-card-border)]/50 overflow-x-auto whitespace-nowrap scrollbar-none"
                >
                  {activeMethod.signature}
                </div>
              </div>

              {/* Description flows directly into parameter table */}
              <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)] leading-relaxed">
                {activeMethod.description}
              </p>

              {/* Parameter Table with Touch-Friendly Horizontal Scrolling (touch-isolated) */}
              {activeMethod.parameters && activeMethod.parameters.length > 0 && (
                <div className="space-y-1.5">
                  <div className="text-xs font-semibold text-[var(--ppt-text-primary)]">
                    Parameters:
                  </div>
                  <div
                    onTouchStart={stopTouchPropagation}
                    onTouchMove={stopTouchPropagation}
                    onTouchEnd={stopTouchPropagation}
                    className="overflow-x-auto rounded border border-[var(--ppt-card-border)]/60 scrollbar-none"
                  >
                    <table className="w-full text-left text-xs border-collapse min-w-[420px]">
                      <thead>
                        <tr className="border-b border-[var(--ppt-card-border)] bg-black/10 dark:bg-white/5 text-[var(--ppt-text-secondary)]">
                          <th className="py-2 px-3 font-mono font-semibold">Name</th>
                          <th className="py-2 px-3 font-mono font-semibold">Type</th>
                          <th className="py-2 px-3 font-semibold">Description</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[var(--ppt-card-border)]/40">
                        {activeMethod.parameters.map((param, pIdx) => (
                          <tr key={pIdx} className="hover:bg-black/5 dark:hover:bg-white/5">
                            <td className="py-2 px-3 font-mono font-medium text-[var(--ppt-text-primary)] whitespace-nowrap">
                              {param.name}
                              {param.optional && <span className="opacity-50">?</span>}
                            </td>
                            <td className="py-2 px-3 font-mono text-[#D24726] whitespace-nowrap">
                              {param.type}
                            </td>
                            <td className="py-2 px-3 text-[var(--ppt-text-secondary)] leading-relaxed">
                              {param.description}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* Return Value */}
              {activeMethod.returns && (
                <div className="text-xs text-[var(--ppt-text-secondary)] flex flex-wrap items-center gap-2 pt-2 border-t border-[var(--ppt-card-border)]/40">
                  <span className="font-semibold text-[var(--ppt-text-primary)]">Returns:</span>
                  <code className="font-mono text-[#D24726] bg-black/10 dark:bg-black/30 px-1.5 py-0.5 rounded">
                    {activeMethod.returns.type}
                  </code>
                  {activeMethod.returns.description && (
                    <span className="opacity-80">— {activeMethod.returns.description}</span>
                  )}
                </div>
              )}

              {/* Code Example with Horizontal Scrolling and Copy Button (touch-isolated) */}
              {activeMethod.example && (
                <div
                  onTouchStart={stopTouchPropagation}
                  onTouchMove={stopTouchPropagation}
                  onTouchEnd={stopTouchPropagation}
                  className="pt-2 space-y-1.5"
                >
                  <div className="text-xs font-semibold text-[var(--ppt-text-primary)]">
                    Usage Example:
                  </div>
                  <div className="overflow-x-auto scrollbar-none rounded-lg">
                    <CodeBlock
                      code={activeMethod.example}
                      language="javascript"
                      filename={`${activeMethod.name}-example.mjs`}
                    />
                  </div>
                </div>
              )}
            </>
          ) : (
            <div className="py-16 text-center text-xs text-[var(--ppt-text-secondary)]">
              Select a method from the list or reel above.
            </div>
          )}
        </div>
      </div>
    </div>
  );
}
