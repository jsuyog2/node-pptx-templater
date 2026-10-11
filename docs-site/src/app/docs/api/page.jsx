'use client';

import React, { useState, useMemo, useEffect, useRef, Suspense } from 'react';
import { useSearchParams } from 'next/navigation';
import { Search, Filter, Code2, Hash, Layers } from 'lucide-react';
import apiReferenceData from '../../../data/apiReference.json';
import ApiCard from '../../../components/ApiCard';
import { animateApiCards } from '../../../animations/gsapTransitions';

function ApiContent() {
  const searchParams = useSearchParams();
  const categoryParam = searchParams.get('category');

  const [selectedCategory, setSelectedCategory] = useState('All');
  const [searchQuery, setSearchQuery] = useState('');
  const cardsContainerRef = useRef(null);

  // Sync category param from URL
  useEffect(() => {
    if (categoryParam && apiReferenceData.categories.includes(categoryParam)) {
      setSelectedCategory(categoryParam);
    }
  }, [categoryParam]);

  // Flatten all methods
  const allMethods = useMemo(() => {
    const list = [];
    apiReferenceData.categories.forEach((cat) => {
      const methods = apiReferenceData.methodsByCategory[cat] || [];
      list.push(...methods);
    });
    return list;
  }, []);

  // Filter methods
  const filteredMethods = useMemo(() => {
    return allMethods.filter((method) => {
      const matchesCategory =
        selectedCategory === 'All' || method.category === selectedCategory;

      if (!matchesCategory) return false;

      if (!searchQuery.trim()) return true;

      const q = searchQuery.toLowerCase();
      return (
        method.name.toLowerCase().includes(q) ||
        method.signature.toLowerCase().includes(q) ||
        method.description.toLowerCase().includes(q) ||
        (method.parameters &&
          method.parameters.some((p) => p.name.toLowerCase().includes(q)))
      );
    });
  }, [allMethods, selectedCategory, searchQuery]);

  // Trigger GSAP animation when category or search changes
  useEffect(() => {
    if (cardsContainerRef.current) {
      animateApiCards(cardsContainerRef.current, '.api-method-card');
    }
  }, [selectedCategory, searchQuery]);

  return (
    <article className="space-y-8">
      {/* Title & Stats */}
      <div className="pb-6 border-b border-slate-200 dark:border-slate-800">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div>
            <div className="inline-flex items-center gap-2 px-3 py-1 rounded-full glass-pill text-xs font-semibold mb-3">
              <Code2 className="w-3.5 h-3.5" />
              <span>Full API Catalog</span>
            </div>
            <h1 className="text-3xl sm:text-4xl font-extrabold tracking-tight text-slate-900 dark:text-white">
              API Reference
            </h1>
            <p className="mt-2 text-sm sm:text-base text-slate-600 dark:text-slate-300">
              Complete catalog of 101 public methods across all 10 OpenXML engine domains.
            </p>
          </div>
          <div className="flex items-center gap-2">
            <span className="text-xs font-mono px-3 py-1.5 rounded-xl glass-card font-semibold text-brand-600 dark:text-brand-300">
              {filteredMethods.length} of {allMethods.length} Methods
            </span>
          </div>
        </div>
      </div>

      {/* Search & Category Filter Bar */}
      <div className="space-y-4">
        {/* Real-time search input */}
        <div className="relative">
          <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search by method name, signature, parameters, or description..."
            className="w-full pl-10 pr-10 py-3 text-xs sm:text-sm rounded-xl glass-card text-slate-900 dark:text-white placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-brand-500 transition-all shadow-sm"
          />
          {searchQuery && (
            <button
              onClick={() => setSearchQuery('')}
              className="absolute right-3.5 top-1/2 -translate-y-1/2 text-sm text-slate-400 hover:text-slate-600 dark:hover:text-white"
            >
              ×
            </button>
          )}
        </div>

        {/* Category Pills Filter */}
        <div className="flex items-center gap-1.5 overflow-x-auto pb-2 scrollbar-none">
          <button
            onClick={() => setSelectedCategory('All')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium whitespace-nowrap transition-all ${
              selectedCategory === 'All'
                ? 'bg-brand-600 text-white font-semibold shadow-sm'
                : 'glass-card hover:bg-slate-100 dark:hover:bg-slate-800'
            }`}
          >
            All Categories ({allMethods.length})
          </button>

          {apiReferenceData.categories.map((cat) => {
            const count = allMethods.filter((m) => m.category === cat).length;
            const isSelected = selectedCategory === cat;
            return (
              <button
                key={cat}
                onClick={() => setSelectedCategory(cat)}
                className={`px-3 py-1.5 rounded-lg text-xs whitespace-nowrap transition-all ${
                  isSelected
                    ? 'bg-brand-600 text-white font-semibold shadow-sm'
                    : 'glass-card hover:bg-slate-100 dark:hover:bg-slate-800'
                }`}
              >
                <span>{cat}</span>
                <span className="ml-1.5 opacity-70 font-mono text-[10px]">({count})</span>
              </button>
            );
          })}
        </div>
      </div>

      {/* Methods Cards Grid */}
      <div ref={cardsContainerRef} className="space-y-6">
        {filteredMethods.length === 0 ? (
          <div className="text-center py-16 px-4 rounded-2xl glass-card border-dashed">
            <Filter className="w-10 h-10 text-slate-400 mx-auto mb-3" />
            <h3 className="text-base font-semibold text-slate-900 dark:text-white">No APIs Found</h3>
            <p className="text-xs text-slate-500 dark:text-slate-400 mt-1 max-w-sm mx-auto">
              No methods matched &quot;{searchQuery}&quot; in category &quot;{selectedCategory}&quot;.
            </p>
            <button
              onClick={() => {
                setSelectedCategory('All');
                setSearchQuery('');
              }}
              className="mt-4 px-4 py-2 text-xs font-semibold rounded-xl bg-brand-600 text-white hover:bg-brand-500"
            >
              Reset Filters
            </button>
          </div>
        ) : (
          filteredMethods.map((method) => (
            <div key={method.name} className="api-method-card">
              <ApiCard method={method} />
            </div>
          ))
        )}
      </div>
    </article>
  );
}

export default function ApiDocsPage() {
  return (
    <Suspense fallback={<div className="py-12 text-center text-slate-400">Loading API Catalog...</div>}>
      <ApiContent />
    </Suspense>
  );
}
