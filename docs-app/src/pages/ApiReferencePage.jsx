import React, { useState, useMemo, useEffect } from 'react';
import { Search, Filter, Layers, CheckCircle2, ChevronRight, Hash } from 'lucide-react';
import ApiCard from '../components/ApiCard';
import SEOHead from '../components/SEOHead';
import apiData from '../data/apiReference.json';

export default function ApiReferencePage({ selectedCategory = 'All', setSelectedCategory = () => { } }) {
  const [searchQuery, setSearchQuery] = useState('');
  const [filterType, setFilterType] = useState('all'); // 'all', 'static', 'async'

  const categories = apiData.categories;

  // Filter methods
  const filteredMethods = useMemo(() => {
    return apiData.methods.filter((method) => {
      // Category match
      if (selectedCategory !== 'All' && method.category !== selectedCategory) {
        return false;
      }

      // Filter modifier (static / async)
      if (filterType === 'static' && !method.isStatic) return false;
      if (filterType === 'async' && !method.isAsync) return false;

      // Search match
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase();
        const matchesName = method.name.toLowerCase().includes(q);
        const matchesDesc = (method.description || '').toLowerCase().includes(q);
        const matchesSig = (method.signature || '').toLowerCase().includes(q);
        const matchesParam = (method.parameters || []).some((p) =>
          p.name.toLowerCase().includes(q) || (p.description || '').toLowerCase().includes(q)
        );
        return matchesName || matchesDesc || matchesSig || matchesParam;
      }

      return true;
    });
  }, [selectedCategory, searchQuery, filterType]);

  return (
    <div>
      <SEOHead
        title="API Reference Manual (102 Methods) — node-pptx-templater"
        description="Comprehensive API reference manual for node-pptx-templater: Lifecycle, Slides, Text, Tables, Charts, Shapes, Cell Shapes, Images, Z-Order, Hyperlinks, and Validation."
        canonicalUrl="https://node-pptx-templater.pages.dev/api-reference"
      />

      {/* Header Banner */}
      <div className="mb-8 pb-6 border-b border-slate-200/80 dark:border-slate-800/80">
        <div className="flex flex-wrap items-center gap-2 mb-2">
          <span className="px-2.5 py-0.5 rounded-full text-xs font-semibold bg-blue-500/10 text-blue-600 dark:text-blue-400 border border-blue-500/20">
            API Reference Manual
          </span>
          <span className="text-xs text-slate-500">v3.0.0-alpha.0</span>
        </div>
        <h1 className="text-2xl sm:text-4xl font-extrabold tracking-tight text-slate-900 dark:text-slate-100 mb-3">
          Method & Function Catalog
        </h1>
        <p className="text-xs sm:text-sm text-slate-600 dark:text-slate-300 max-w-3xl leading-relaxed">
          Comprehensive, type-safe API documentation for all 102 public methods across 11 architectural subsystems.
          Every entry includes method signatures, parameter tables, return specifications, OpenXML behaviors, and copyable code examples.
        </p>
      </div>

      {/* Filter and Search Toolbar */}
      <div className="glass-card rounded-2xl p-4 sm:p-5 mb-8 border border-slate-200/80 dark:border-slate-800/80">
        <div className="flex flex-col md:flex-row gap-4 items-center justify-between">
          {/* Search Bar */}
          <div className="relative w-full md:w-96">
            <Search className="w-4 h-4 text-slate-400 absolute left-3.5 top-1/2 -translate-y-1/2 pointer-events-none" />
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="Search by method name, parameter, or signature..."
              className="w-full pl-10 pr-4 py-2 text-xs sm:text-sm rounded-xl bg-slate-100 dark:bg-slate-900 border border-slate-200/80 dark:border-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500/50 text-slate-900 dark:text-slate-100 placeholder-slate-400"
            />
          </div>

          {/* Modifier Toggle Buttons */}
          <div className="flex items-center gap-1.5 self-start md:self-auto">
            <span className="text-xs text-slate-500 mr-1 flex items-center gap-1">
              <Filter className="w-3.5 h-3.5" /> Filter:
            </span>
            <button
              onClick={() => setFilterType('all')}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${filterType === 'all'
                ? 'bg-blue-600 text-white'
                : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-700'
                }`}
            >
              All
            </button>
            <button
              onClick={() => setFilterType('static')}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${filterType === 'static'
                ? 'bg-amber-600 text-white'
                : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-700'
                }`}
            >
              Static
            </button>
            <button
              onClick={() => setFilterType('async')}
              className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${filterType === 'async'
                ? 'bg-purple-600 text-white'
                : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-300 hover:bg-slate-200 dark:hover:bg-slate-700'
                }`}
            >
              Async
            </button>
          </div>
        </div>

        {/* Category Pills */}
        <div className="flex flex-wrap gap-1.5 mt-4 pt-4 border-t border-slate-200/80 dark:border-slate-800/80">
          <button
            onClick={() => setSelectedCategory('All')}
            className={`px-2.5 py-1 rounded-lg text-xs font-medium transition-colors ${selectedCategory === 'All'
              ? 'bg-blue-600 text-white'
              : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-slate-700'
              }`}
          >
            All Categories ({apiData.methods.length})
          </button>
          {categories.map((cat) => {
            const count = apiData.methodsByCategory[cat]?.length || 0;
            return (
              <button
                key={cat}
                onClick={() => setSelectedCategory(cat)}
                className={`px-2.5 py-1 rounded-lg text-xs font-medium transition-colors ${selectedCategory === cat
                  ? 'bg-blue-600 text-white'
                  : 'bg-slate-100 dark:bg-slate-800 text-slate-600 dark:text-slate-400 hover:bg-slate-200 dark:hover:bg-slate-700'
                  }`}
              >
                {cat} ({count})
              </button>
            );
          })}
        </div>
      </div>

      {/* Results Header */}
      <div className="flex items-center justify-between mb-6 text-xs text-slate-500">
        <div>
          Showing <span className="font-semibold text-slate-900 dark:text-slate-100">{filteredMethods.length}</span> methods{' '}
          {selectedCategory !== 'All' && <span>in <strong className="text-blue-500">{selectedCategory}</strong></span>}
          {searchQuery && <span> matching "<strong>{searchQuery}</strong>"</span>}
        </div>
      </div>

      {/* Methods List */}
      {filteredMethods.length === 0 ? (
        <div className="glass-card rounded-2xl p-12 text-center text-slate-500 dark:text-slate-400">
          <p className="text-base font-medium mb-2">No methods match your search filter.</p>
          <button
            onClick={() => { setSelectedCategory('All'); setSearchQuery(''); setFilterType('all'); }}
            className="px-4 py-2 rounded-xl bg-blue-600 text-white text-xs font-semibold hover:bg-blue-500 transition-colors"
          >
            Reset Filters
          </button>
        </div>
      ) : (
        <div className="space-y-6">
          {filteredMethods.map((method) => (
            <ApiCard key={method.name} method={method} />
          ))}
        </div>
      )}
    </div>
  );
}
