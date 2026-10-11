'use client';

import React from 'react';
import Link from 'next/link';
import { Search, ChevronRight, Hash, Layers } from 'lucide-react';

export default function Sidebar({
  categories = [],
  selectedCategory = 'All',
  onSelectCategory = () => {},
  searchQuery = '',
  onSearchChange = () => {},
  methods = [],
}) {
  return (
    <aside className="w-full lg:w-72 shrink-0">
      <div className="sticky top-20 space-y-6 max-h-[calc(100vh-6rem)] overflow-y-auto pr-2">
        {/* Search Bar */}
        <div className="relative">
          <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
          <input
            type="text"
            value={searchQuery}
            onChange={(e) => onSearchChange(e.target.value)}
            placeholder="Search all 101 APIs..."
            className="w-full pl-9 pr-3 py-2 text-xs rounded-xl bg-slate-900 border border-slate-700/80 text-white placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-brand-500 focus:border-brand-500 transition-all"
          />
          {searchQuery && (
            <button
              onClick={() => onSearchChange('')}
              className="absolute right-2.5 top-1/2 -translate-y-1/2 text-xs text-slate-400 hover:text-white"
            >
              ×
            </button>
          )}
        </div>

        {/* Category List */}
        <div>
          <div className="flex items-center justify-between mb-2">
            <h3 className="text-xs font-semibold uppercase tracking-wider text-slate-400">
              API Categories
            </h3>
            <span className="text-[11px] text-slate-500 font-mono">
              {methods.length} methods
            </span>
          </div>

          <div className="space-y-1">
            <button
              onClick={() => onSelectCategory('All')}
              className={`w-full flex items-center justify-between px-3 py-2 rounded-lg text-xs font-medium transition-all ${
                selectedCategory === 'All'
                  ? 'bg-brand-500/20 text-brand-300 font-semibold border border-brand-500/30'
                  : 'text-slate-300 hover:bg-slate-800/60 hover:text-white'
              }`}
            >
              <span>All Categories</span>
              <span className="text-[10px] font-mono px-1.5 py-0.5 rounded bg-slate-800/80 text-slate-400">
                101
              </span>
            </button>

            {categories.map((cat) => {
              const count = methods.filter((m) => m.category === cat).length;
              const isSelected = selectedCategory === cat;
              return (
                <button
                  key={cat}
                  onClick={() => onSelectCategory(cat)}
                  className={`w-full flex items-center justify-between px-3 py-1.5 rounded-lg text-xs transition-all text-left ${
                    isSelected
                      ? 'bg-brand-500/20 text-brand-300 font-semibold border border-brand-500/30'
                      : 'text-slate-300 hover:bg-slate-800/60 hover:text-white'
                  }`}
                >
                  <span className="truncate pr-2">{cat}</span>
                  <span className="text-[10px] font-mono px-1.5 py-0.5 rounded bg-slate-800/80 text-slate-400 shrink-0">
                    {count}
                  </span>
                </button>
              );
            })}
          </div>
        </div>

        {/* Quick Guide Links */}
        <div className="pt-4 border-t border-slate-800">
          <h3 className="text-xs font-semibold uppercase tracking-wider text-slate-400 mb-2">
            Deep-Dive Guides
          </h3>
          <ul className="space-y-1.5 text-xs">
            <li>
              <Link
                href="/guides/quick-start/"
                className="text-slate-400 hover:text-white flex items-center gap-1.5 py-1"
              >
                <ChevronRight className="w-3.5 h-3.5 text-slate-500" />
                <span>5-Minute Quick Start</span>
              </Link>
            </li>
            <li>
              <Link
                href="/guides/serverless-deployment/"
                className="text-slate-400 hover:text-white flex items-center gap-1.5 py-1"
              >
                <ChevronRight className="w-3.5 h-3.5 text-slate-500" />
                <span>AWS Lambda & Cloudflare Workers</span>
              </Link>
            </li>
            <li>
              <Link
                href="/guides/run-fragment-healing/"
                className="text-slate-400 hover:text-white flex items-center gap-1.5 py-1"
              >
                <ChevronRight className="w-3.5 h-3.5 text-slate-500" />
                <span>Run Fragment Healing</span>
              </Link>
            </li>
            <li>
              <Link
                href="/guides/chart-dual-sync/"
                className="text-slate-400 hover:text-white flex items-center gap-1.5 py-1"
              >
                <ChevronRight className="w-3.5 h-3.5 text-slate-500" />
                <span>Chart Excel Dual-Syncing</span>
              </Link>
            </li>
            <li>
              <Link
                href="/guides/rectangular-table-spans/"
                className="text-slate-400 hover:text-white flex items-center gap-1.5 py-1"
              >
                <ChevronRight className="w-3.5 h-3.5 text-slate-500" />
                <span>Rectangular Table Spans</span>
              </Link>
            </li>
          </ul>
        </div>
      </div>
    </aside>
  );
}
