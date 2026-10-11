'use client';

import React, { useState, useMemo, useEffect, Suspense } from 'react';
import { useSearchParams } from 'next/navigation';
import { Search, Filter, Hash, Layers, ChevronRight } from 'lucide-react';
import apiReferenceData from '../../data/apiReference.json';
import ApiCard from '../../components/ApiCard';
import Sidebar from '../../components/Sidebar';

function ApiReferenceContent() {
  const searchParams = useSearchParams();
  const categoryParam = searchParams.get('category');

  const [selectedCategory, setSelectedCategory] = useState('All');
  const [searchQuery, setSearchQuery] = useState('');

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

  // Filter methods by category and search
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

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8 sm:py-12">
      {/* Page Title & Subtitle */}
      <div className="mb-8 pb-6 border-b border-slate-800">
        <div className="flex flex-wrap items-center justify-between gap-4">
          <div>
            <h1 className="text-3xl sm:text-4xl font-extrabold text-white tracking-tight">
              API Reference
            </h1>
            <p className="mt-2 text-sm sm:text-base text-slate-400">
              Complete catalog of 101 public methods across all 10 OpenXML engine domains.
            </p>
          </div>
          <div className="flex items-center gap-2">
            <span className="text-xs font-mono px-3 py-1 rounded-full bg-slate-900 border border-slate-800 text-brand-300">
              Showing {filteredMethods.length} of {allMethods.length} APIs
            </span>
          </div>
        </div>
      </div>

      {/* Main 3-Column / Responsive Layout */}
      <div className="flex flex-col lg:flex-row gap-8 items-start">
        {/* Left Sidebar */}
        <Sidebar
          categories={apiReferenceData.categories}
          selectedCategory={selectedCategory}
          onSelectCategory={(cat) => setSelectedCategory(cat)}
          searchQuery={searchQuery}
          onSearchChange={(q) => setSearchQuery(q)}
          methods={allMethods}
        />

        {/* Center Main Methods List */}
        <div className="flex-1 min-w-0 space-y-8 w-full">
          {filteredMethods.length === 0 ? (
            <div className="text-center py-16 px-4 rounded-2xl border border-dashed border-slate-800 bg-slate-900/30">
              <Filter className="w-10 h-10 text-slate-500 mx-auto mb-3" />
              <h3 className="text-base font-semibold text-white">No methods found</h3>
              <p className="text-xs text-slate-400 mt-1 max-w-sm mx-auto">
                No APIs match your current search &quot;{searchQuery}&quot; in category &quot;{selectedCategory}&quot;.
              </p>
              <button
                onClick={() => {
                  setSelectedCategory('All');
                  setSearchQuery('');
                }}
                className="mt-4 px-4 py-2 text-xs font-semibold rounded-lg bg-slate-800 text-slate-200 hover:bg-slate-700"
              >
                Reset All Filters
              </button>
            </div>
          ) : (
            filteredMethods.map((method) => (
              <ApiCard key={method.name} method={method} />
            ))
          )}
        </div>

        {/* Right Quick On-Page Mini-TOC for Desktop (>1024px) */}
        <div className="hidden xl:block w-56 shrink-0">
          <div className="sticky top-20 max-h-[calc(100vh-6rem)] overflow-y-auto pr-1">
            <h4 className="text-xs font-semibold uppercase tracking-wider text-slate-400 mb-3">
              On This Page
            </h4>
            <div className="space-y-1">
              {filteredMethods.slice(0, 35).map((method) => (
                <a
                  key={method.name}
                  href={`#${method.name}`}
                  className="block text-[11px] font-mono text-slate-400 hover:text-brand-300 truncate py-0.5 hover:translate-x-0.5 transition-all"
                  title={method.name}
                >
                  {method.name}
                </a>
              ))}
              {filteredMethods.length > 35 && (
                <span className="block text-[10px] text-slate-500 italic pt-1">
                  + {filteredMethods.length - 35} more methods
                </span>
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default function ApiReferencePage() {
  return (
    <Suspense
      fallback={
        <div className="max-w-7xl mx-auto px-4 py-16 text-center text-slate-400">
          Loading API Reference...
        </div>
      }
    >
      <ApiReferenceContent />
    </Suspense>
  );
}
