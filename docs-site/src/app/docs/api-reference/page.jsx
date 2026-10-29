'use client';

import React, { useState, useMemo, Suspense } from 'react';
import { useSearchParams } from 'next/navigation';
import { Code2, Search, Filter, Hash, Check, Copy, ChevronRight, Layers } from 'lucide-react';
import apiReferenceData from '../../../data/apiReference.json';
import PowerPointSlideDeck from '../../../components/PowerPointSlideDeck';
import CodeBlock from '../../../components/CodeBlock';

function CategorySlideContent({ categoryName, methods = [] }) {
  const [search, setSearch] = useState('');
  const [selectedMethodName, setSelectedMethodName] = useState(methods[0]?.name || '');

  const filteredMethods = useMemo(() => {
    if (!search.trim()) return methods;
    const q = search.toLowerCase();
    return methods.filter(
      (m) =>
        m.name.toLowerCase().includes(q) ||
        m.signature.toLowerCase().includes(q) ||
        m.description.toLowerCase().includes(q)
    );
  }, [methods, search]);

  const activeMethod = useMemo(() => {
    return (
      filteredMethods.find((m) => m.name === selectedMethodName) ||
      filteredMethods[0] ||
      methods[0] ||
      {}
    );
  }, [filteredMethods, selectedMethodName, methods]);

  return (
    <div className="space-y-4">
      {/* Category Header & Filter */}
      <div className="flex flex-wrap items-center justify-between gap-3 pb-3 border-b border-[var(--ppt-card-border)]/50">
        <div className="flex items-center gap-2">
          <span className="text-xs font-semibold text-[var(--ppt-text-secondary)]">
            {methods.length} Public API Methods
          </span>
          <span className="text-[10px] px-2 py-0.5 rounded-full bg-[#D24726]/10 text-[#D24726] font-mono">
            ISO/IEC 29500
          </span>
        </div>

        {/* Search Input */}
        <div className="relative">
          <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-[var(--ppt-text-secondary)] opacity-60" />
          <input
            type="text"
            placeholder={`Filter ${categoryName} methods...`}
            value={search}
            onChange={(e) => setSearch(e.target.value)}
            className="pl-8 pr-3 py-1 rounded text-xs ppt-card-surface border border-[var(--ppt-card-border)] text-[var(--ppt-text-primary)] placeholder-[var(--ppt-text-secondary)]/50 focus:border-[#D24726] focus:outline-none w-48 sm:w-60"
          />
        </div>
      </div>

      {/* Main Two-Column Layout: Method List & Method Detail Inspector */}
      <div className="grid grid-cols-1 lg:grid-cols-12 gap-4">
        {/* Left Column: Method Selector Chips */}
        <div className="lg:col-span-4 max-h-[360px] overflow-y-auto space-y-1.5 pr-1">
          {filteredMethods.map((m) => {
            const isSelected = m.name === activeMethod.name;
            return (
              <button
                key={m.name}
                type="button"
                onClick={() => setSelectedMethodName(m.name)}
                className={`w-full text-left px-3 py-2 rounded text-xs transition-all flex items-center justify-between border ${
                  isSelected
                    ? 'border-[#D24726] bg-[#D24726]/10 text-[var(--ppt-text-primary)] font-semibold'
                    : 'border-transparent hover:border-[var(--ppt-card-border)] text-[var(--ppt-text-secondary)] hover:text-[var(--ppt-text-primary)]'
                }`}
              >
                <div className="truncate mr-2">
                  <div className="font-mono text-[11px] truncate">{m.name}()</div>
                  <div className="text-[9px] opacity-70 truncate">{m.description}</div>
                </div>
                <ChevronRight
                  className={`w-3.5 h-3.5 flex-shrink-0 transition-transform ${
                    isSelected ? 'text-[#D24726] translate-x-0.5' : 'opacity-40'
                  }`}
                />
              </button>
            );
          })}
          {filteredMethods.length === 0 && (
            <div className="text-center py-6 text-xs text-[var(--ppt-text-secondary)]">
              No matching methods found.
            </div>
          )}
        </div>

        {/* Right Column: Active Method Detail Inspector */}
        <div className="lg:col-span-8 ppt-card-surface p-4 rounded-lg border border-[var(--ppt-card-border)] space-y-3">
          {activeMethod.name ? (
            <>
              {/* Method Title & Signature */}
              <div>
                <div className="flex items-center justify-between">
                  <h3 className="font-mono text-base font-bold text-[var(--ppt-text-primary)] flex items-center gap-1.5">
                    <span className="text-[#D24726] font-normal">ppt.</span>
                    <span>{activeMethod.name}()</span>
                  </h3>
                  {activeMethod.isAsync && (
                    <span className="text-[10px] font-mono px-2 py-0.5 rounded bg-amber-500/10 text-amber-500 border border-amber-500/20">
                      async
                    </span>
                  )}
                </div>
                <div className="font-mono text-xs text-[#D24726] mt-1 bg-black/20 dark:bg-black/40 p-2 rounded border border-[var(--ppt-card-border)]/50 overflow-x-auto">
                  {activeMethod.signature}
                </div>
              </div>

              {/* Description */}
              <p className="text-xs text-[var(--ppt-text-secondary)] leading-relaxed">
                {activeMethod.description}
              </p>

              {/* Parameters Table */}
              {activeMethod.parameters && activeMethod.parameters.length > 0 && (
                <div>
                  <div className="text-[11px] font-semibold text-[var(--ppt-text-primary)] mb-1">
                    Parameters:
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-[11px] border-collapse">
                      <thead>
                        <tr className="border-b border-[var(--ppt-card-border)] text-[var(--ppt-text-secondary)]">
                          <th className="py-1 px-2 font-mono">Name</th>
                          <th className="py-1 px-2 font-mono">Type</th>
                          <th className="py-1 px-2">Description</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[var(--ppt-card-border)]/40">
                        {activeMethod.parameters.map((param, pIdx) => (
                          <tr key={pIdx}>
                            <td className="py-1 px-2 font-mono font-medium text-[var(--ppt-text-primary)]">
                              {param.name}
                              {param.optional && <span className="opacity-50">?</span>}
                            </td>
                            <td className="py-1 px-2 font-mono text-[#D24726]">
                              {param.type}
                            </td>
                            <td className="py-1 px-2 text-[var(--ppt-text-secondary)]">
                              {param.description}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* Returns info */}
              {activeMethod.returns && (
                <div className="text-xs text-[var(--ppt-text-secondary)] flex items-center gap-2 pt-1 border-t border-[var(--ppt-card-border)]/40">
                  <span className="font-semibold text-[var(--ppt-text-primary)]">Returns:</span>
                  <code className="font-mono text-[#D24726]">{activeMethod.returns.type}</code>
                  {activeMethod.returns.description && (
                    <span className="opacity-80">— {activeMethod.returns.description}</span>
                  )}
                </div>
              )}

              {/* Code Example */}
              {activeMethod.example && (
                <div className="pt-1">
                  <div className="text-[11px] font-semibold text-[var(--ppt-text-primary)] mb-1">
                    Usage Example:
                  </div>
                  <CodeBlock
                    code={activeMethod.example}
                    language="javascript"
                    filename={`${activeMethod.name}-example.mjs`}
                  />
                </div>
              )}
            </>
          ) : (
            <div className="py-12 text-center text-xs text-[var(--ppt-text-secondary)]">
              Select a method from the left list.
            </div>
          )}
        </div>
      </div>
    </div>
  );
}

function ApiReferenceDeck() {
  const searchParams = useSearchParams();
  const categoryParam = searchParams.get('category');

  // Strict 8 Standardized Categories
  const categories = apiReferenceData.categories || [
    'Import/Export',
    'Slides',
    'Text & Placeholders',
    'Tables',
    'Charts',
    'Shapes',
    'Images & Media',
    'Elements',
  ];

  // Determine initial slide index from query param if provided
  const initialIndex = useMemo(() => {
    if (categoryParam) {
      const idx = categories.indexOf(categoryParam);
      if (idx !== -1) return idx;
    }
    return 0;
  }, [categoryParam, categories]);

  // Construct 8 slides, one for each standardized category
  const slides = useMemo(() => {
    return categories.map((cat, idx) => {
      const methods = apiReferenceData.methodsByCategory[cat] || [];
      return {
        id: `cat-${idx + 1}`,
        category: 'API Reference',
        title: `${cat}`,
        subtitle: `Category ${idx + 1} of 8 • ${methods.length} Standardized Methods`,
        signature: `${cat} Module`,
        content: <CategorySlideContent categoryName={cat} methods={methods} />,
      };
    });
  }, [categories]);

  return (
    <PowerPointSlideDeck
      slides={slides}
      initialIndex={initialIndex}
      categoryTitle="API Reference"
    />
  );
}

export default function ApiReferencePage() {
  return (
    <Suspense
      fallback={
        <div className="h-[calc(100vh-104px)] flex items-center justify-center text-xs text-[var(--ppt-text-secondary)]">
          Loading PowerPoint API Catalog...
        </div>
      }
    >
      <ApiReferenceDeck />
    </Suspense>
  );
}
