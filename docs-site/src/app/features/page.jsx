'use client';

import React, { useState, useMemo, useEffect, useRef } from 'react';
import Link from 'next/link';
import {
  Sparkles,
  Search,
  CheckCircle2,
  Filter,
  Layers,
  ArrowRight,
  Zap,
  Star,
  Cpu
} from 'lucide-react';
import featuresData from '../../data/features.json';
import { animateFeaturesGrid } from '../../animations/gsapTransitions';

export default function FeaturesPage() {
  const [selectedCategory, setSelectedCategory] = useState('All');
  const [searchQuery, setSearchQuery] = useState('');
  const gridContainerRef = useRef(null);

  // Flatten features with category information
  const allFeatures = useMemo(() => {
    const list = [];
    featuresData.categories.forEach((cat) => {
      cat.items.forEach((item) => {
        list.push({
          ...item,
          categoryNumber: cat.number,
          categoryTitle: cat.title,
          isFlagship: cat.isFlagship
        });
      });
    });
    return list;
  }, []);

  // Filter features
  const filteredFeatures = useMemo(() => {
    return allFeatures.filter((item) => {
      const matchesCategory =
        selectedCategory === 'All' || item.categoryTitle === selectedCategory;

      if (!matchesCategory) return false;

      if (!searchQuery.trim()) return true;

      const q = searchQuery.toLowerCase();
      return (
        item.title.toLowerCase().includes(q) ||
        item.description.toLowerCase().includes(q) ||
        item.categoryTitle.toLowerCase().includes(q)
      );
    });
  }, [allFeatures, selectedCategory, searchQuery]);

  // Run GSAP entrance animation when category or search changes
  useEffect(() => {
    if (gridContainerRef.current) {
      animateFeaturesGrid(gridContainerRef.current, '.feature-card');
    }
  }, [selectedCategory, searchQuery]);

  return (
    <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-10 sm:py-14 space-y-12">
      {/* Header Banner */}
      <div className="text-center max-w-3xl mx-auto space-y-4">
        <div className="inline-flex items-center gap-2 px-3.5 py-1.5 rounded-full glass-pill text-xs font-semibold">
          <Sparkles className="w-4 h-4 text-brand-500" />
          <span>Complete Feature Catalog</span>
          <span className="w-1.5 h-1.5 rounded-full bg-brand-500"></span>
          <span className="font-mono">{featuresData.totalFeatures} Items</span>
        </div>

        <h1 className="text-4xl sm:text-5xl font-extrabold tracking-tight text-slate-900 dark:text-white leading-tight">
          Granular Capabilities &amp; Feature Inventory
        </h1>

        <p className="text-base sm:text-lg text-slate-600 dark:text-slate-300 leading-relaxed">
          An itemized, full-coverage index of all {featuresData.totalFeatures} architectural capabilities supported across {featuresData.totalCategories} specialized OpenXML manipulation domains.
        </p>

        {/* Stats Strip */}
        <div className="flex flex-wrap items-center justify-center gap-6 pt-2 text-xs font-medium text-slate-600 dark:text-slate-400">
          <div className="flex items-center gap-1.5">
            <CheckCircle2 className="w-4 h-4 text-emerald-500" />
            <span>86 Tested Capabilities</span>
          </div>
          <div className="flex items-center gap-1.5">
            <Cpu className="w-4 h-4 text-brand-500" />
            <span>12 Domain Categories</span>
          </div>
          <div className="flex items-center gap-1.5">
            <Zap className="w-4 h-4 text-amber-500" />
            <span>Zero Native Addons</span>
          </div>
        </div>
      </div>

      {/* Filter and Search Bar */}
      <div className="space-y-4 max-w-5xl mx-auto">
        <div className="relative">
          <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => setSearchQuery(e.target.value)}
            placeholder="Search all 86 capabilities (e.g. 'Excel', 'rowSpan', 'EMU', 'Z-Order', 'Streams')..."
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

        {/* Category Filter Buttons */}
        <div className="flex items-center gap-1.5 overflow-x-auto pb-2 scrollbar-none">
          <button
            onClick={() => setSelectedCategory('All')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium whitespace-nowrap transition-all ${selectedCategory === 'All'
              ? 'bg-brand-600 text-white font-semibold shadow-sm'
              : 'glass-card hover:bg-slate-100 dark:hover:bg-slate-800'
              }`}
          >
            All Categories ({allFeatures.length})
          </button>

          {featuresData.categories.map((cat) => {
            const isSelected = selectedCategory === cat.title;
            return (
              <button
                key={cat.title}
                onClick={() => setSelectedCategory(cat.title)}
                className={`px-3 py-1.5 rounded-lg text-xs whitespace-nowrap transition-all ${isSelected
                  ? 'bg-brand-600 text-white font-semibold shadow-sm'
                  : 'glass-card hover:bg-slate-100 dark:hover:bg-slate-800'
                  }`}
              >
                <span>{cat.title}</span>
                <span className="ml-1.5 opacity-70 font-mono text-[10px]">({cat.count})</span>
              </button>
            );
          })}
        </div>
      </div>

      {/* Grid of Feature Cards */}
      <div
        ref={gridContainerRef}
        className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6 max-w-7xl mx-auto"
      >
        {filteredFeatures.length === 0 ? (
          <div className="col-span-full text-center py-16 px-4 rounded-2xl glass-card border-dashed">
            <Filter className="w-10 h-10 text-slate-400 mx-auto mb-3" />
            <h3 className="text-base font-semibold text-slate-900 dark:text-white">No Features Found</h3>
            <p className="text-xs text-slate-500 dark:text-slate-400 mt-1 max-w-sm mx-auto">
              No capabilities matched your filter query &quot;{searchQuery}&quot;.
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
          filteredFeatures.map((feature) => (
            <div
              key={feature.id}
              className="feature-card glass-card p-6 rounded-2xl flex flex-col justify-between group hover:-translate-y-1 hover:border-brand-500/40 hover:shadow-xl transition-all duration-200"
            >
              <div className="space-y-3">
                <div className="flex items-center justify-between gap-2">
                  <span className="text-[11px] font-mono px-2 py-0.5 rounded-md bg-slate-200/80 dark:bg-slate-800 text-slate-700 dark:text-slate-300 font-semibold">
                    #{feature.id}
                  </span>
                  <span className="text-[10px] font-medium text-slate-500 dark:text-slate-400 truncate max-w-[170px]">
                    {feature.categoryTitle}
                  </span>
                </div>

                <h3 className="text-base font-bold text-slate-900 dark:text-white group-hover:text-brand-600 dark:group-hover:text-brand-300 transition-colors">
                  {feature.title}
                </h3>

                <p className="text-xs text-slate-600 dark:text-slate-300 leading-relaxed">
                  {feature.description}
                </p>
              </div>

              {feature.isFlagship && (
                <div className="mt-4 pt-3 border-t border-slate-200/60 dark:border-slate-800/60 flex items-center gap-1.5 text-[11px] font-semibold text-amber-600 dark:text-amber-400">
                  <Star className="w-3.5 h-3.5 fill-amber-400" />
                  <span>Flagship Innovation</span>
                </div>
              )}
            </div>
          ))
        )}
      </div>

      {/* Bottom CTA to API Reference */}
      <div className="text-center pt-10 border-t border-slate-200 dark:border-slate-800 max-w-xl mx-auto space-y-4">
        <h3 className="text-xl font-bold text-slate-900 dark:text-white">
          Ready to see the API signatures?
        </h3>
        <p className="text-xs text-slate-600 dark:text-slate-400">
          Check out the methods with parameter schemas, return types, and copyable code snippets.
        </p>
        <Link
          href="/docs/api"
          className="inline-flex items-center gap-2 px-6 py-3 rounded-xl font-semibold text-xs bg-brand-600 hover:bg-brand-500 text-white shadow-lg shadow-brand-500/25 transition-all"
        >
          <span>Explore Methods API Reference</span>
          <ArrowRight className="w-4 h-4" />
        </Link>
      </div>
    </div>
  );
}
