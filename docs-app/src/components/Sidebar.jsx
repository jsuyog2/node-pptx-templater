import React, { useState, useMemo } from 'react';
import { Search, ChevronRight, Book, Layers, Cpu, Sparkles, X, Compass, ExternalLink } from 'lucide-react';

export default function Sidebar({
  categories = [],
  methods = [],
  selectedCategory,
  onSelectCategory,
  activeRoute,
  onNavigate,
  isOpen,
  onClose
}) {
  const [searchTerm, setSearchTerm] = useState('');

  // Filter methods based on search term
  const filteredMethods = useMemo(() => {
    if (!searchTerm.trim()) return [];
    const term = searchTerm.toLowerCase();
    return methods.filter(
      m => m.name.toLowerCase().includes(term) ||
        (m.description && m.description.toLowerCase().includes(term)) ||
        m.category.toLowerCase().includes(term)
    ).slice(0, 15);
  }, [methods, searchTerm]);

  return (
    <>
      {/* Mobile Backdrop */}
      {isOpen && (
        <div
          onClick={onClose}
          className="fixed inset-0 bg-slate-950/60 backdrop-blur-sm z-40 md:hidden transition-opacity"
          aria-hidden="true"
        />
      )}

      {/* Sidebar Container */}
      <aside
        className={`fixed md:sticky top-0 md:top-16 left-0 h-full md:h-[calc(100vh-4rem)] w-72 sm:w-80 bg-white/95 dark:bg-[#0b0f19]/95 md:bg-transparent backdrop-blur-xl md:backdrop-blur-none border-r border-slate-200/80 dark:border-slate-800/80 z-50 md:z-10 p-4 sm:p-5 flex flex-col transform transition-transform duration-300 ease-in-out overflow-y-auto ${isOpen ? 'translate-x-0 shadow-2xl' : '-translate-x-full md:translate-x-0'
          }`}
      >
        {/* Mobile Header / Close Button */}
        <div className="flex items-center justify-between pb-3 mb-2 md:hidden border-b border-slate-200 dark:border-slate-800">
          <div className="flex items-center gap-2">
            <span className="font-bold text-sm tracking-tight text-slate-800 dark:text-slate-100">
              Documentation Menu
            </span>
          </div>
          <button
            onClick={onClose}
            className="p-1 rounded-lg text-slate-400 hover:text-slate-700 dark:hover:text-slate-200"
            aria-label="Close navigation sidebar"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Real-time Search Input */}
        <div className="relative mb-5">
          <div className="absolute inset-y-0 left-0 pl-3 flex items-center pointer-events-none text-slate-400">
            <Search className="w-4 h-4" />
          </div>
          <input
            type="text"
            value={searchTerm}
            onChange={(e) => setSearchTerm(e.target.value)}
            placeholder="Filter 102+ APIs & guides..."
            className="w-full pl-9 pr-8 py-2 text-xs sm:text-sm rounded-xl bg-slate-100 dark:bg-slate-900 border border-slate-200/80 dark:border-slate-800 focus:outline-none focus:ring-2 focus:ring-blue-500/50 text-slate-900 dark:text-slate-100 placeholder-slate-400"
          />
          {searchTerm && (
            <button
              onClick={() => setSearchTerm('')}
              className="absolute inset-y-0 right-0 pr-2.5 flex items-center text-slate-400 hover:text-slate-600 dark:hover:text-slate-200"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          )}
        </div>

        {/* Search Results Dropdown */}
        {searchTerm.trim() && (
          <div className="mb-6 p-2 rounded-xl bg-slate-100/90 dark:bg-slate-900/90 border border-slate-200 dark:border-slate-800">
            <div className="text-[10px] font-bold uppercase tracking-wider text-slate-400 px-2 py-1">
              Search Results ({filteredMethods.length})
            </div>
            {filteredMethods.length === 0 ? (
              <div className="px-2 py-3 text-xs text-slate-400">No matching APIs found.</div>
            ) : (
              <div className="space-y-1 mt-1 max-h-56 overflow-y-auto">
                {filteredMethods.map((m) => (
                  <button
                    key={m.name}
                    onClick={() => {
                      onNavigate('/api-reference');
                      onSelectCategory(m.category);
                      window.location.pathname = `api-${m.name}`;
                      setSearchTerm('');
                      if (onClose) onClose();
                    }}
                    className="w-full text-left px-2.5 py-1.5 rounded-lg text-xs hover:bg-blue-500/10 hover:text-blue-600 dark:hover:text-blue-400 flex items-center justify-between group transition-colors"
                  >
                    <span className="font-mono font-medium truncate">{m.name}</span>
                    <span className="text-[10px] text-slate-400 group-hover:text-blue-500">{m.category}</span>
                  </button>
                ))}
              </div>
            )}
          </div>
        )}

        {/* Section: Main Links */}
        <div className="mb-6">
          <div className="text-[11px] font-bold uppercase tracking-wider text-slate-400 dark:text-slate-500 px-2 mb-1.5 flex items-center gap-1.5">
            <Compass className="w-3.5 h-3.5" /> Navigation
          </div>
          <nav className="space-y-0.5">
            <button
              onClick={() => { onNavigate('/'); if (onClose) onClose(); }}
              className={`w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs sm:text-sm font-medium transition-colors ${activeRoute === '/'
                ? 'bg-blue-600 text-white font-semibold shadow-sm'
                : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800/60'
                }`}
            >
              <span>Overview & Quick Start</span>
            </button>
            <button
              onClick={() => { onNavigate('/guides'); if (onClose) onClose(); }}
              className={`w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs sm:text-sm font-medium transition-colors ${activeRoute === '/guides'
                ? 'bg-blue-600 text-white font-semibold shadow-sm'
                : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800/60'
                }`}
            >
              <span>Guides & Deep Dives</span>
            </button>
            <button
              onClick={() => { onNavigate('/api-reference'); if (onClose) onClose(); }}
              className={`w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs sm:text-sm font-medium transition-colors ${activeRoute === '/api-reference'
                ? 'bg-blue-600 text-white font-semibold shadow-sm'
                : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800/60'
                }`}
            >
              <span>API Reference</span>
            </button>
            <button
              onClick={() => { onNavigate('/showcase'); if (onClose) onClose(); }}
              className={`w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs sm:text-sm font-medium transition-colors ${activeRoute === '/showcase'
                ? 'bg-blue-600 text-white font-semibold shadow-sm'
                : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800/60'
                }`}
            >
              <span>Showcase</span>
            </button>
            <button
              onClick={() => { onNavigate('/features'); if (onClose) onClose(); }}
              className={`w-full flex items-center justify-between px-3 py-2 rounded-xl text-xs sm:text-sm font-medium transition-colors ${activeRoute === '/features'
                ? 'bg-blue-600 text-white font-semibold shadow-sm'
                : 'text-slate-700 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800/60'
                }`}
            >
              <span>Feature Catalog</span>
            </button>
          </nav>
        </div>

        {/* Section: API Categories */}
        <div className="flex-1">
          <div className="text-[11px] font-bold uppercase tracking-wider text-slate-400 dark:text-slate-500 px-2 mb-2 flex items-center justify-between">
            <span className="flex items-center gap-1.5">
              <Layers className="w-3.5 h-3.5" /> API Categories
            </span>
          </div>

          <div className="space-y-1">
            <button
              onClick={() => {
                onNavigate('/api-reference');
                onSelectCategory('All');
                if (onClose) onClose();
              }}
              className={`w-full flex items-center justify-between px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${selectedCategory === 'All' && activeRoute === '/api-reference'
                ? 'bg-blue-500/10 text-blue-600 dark:text-blue-400 font-semibold'
                : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800/40 hover:text-slate-900 dark:hover:text-slate-200'
                }`}
            >
              <span>All Categories</span>
              <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-slate-100 dark:bg-slate-800 text-slate-500">
                {methods.length}
              </span>
            </button>

            {categories.map((cat) => {
              const count = methods.filter((m) => m.category === cat).length;
              const isSelected = selectedCategory === cat && activeRoute === '/api-reference';
              return (
                <button
                  key={cat}
                  onClick={() => {
                    onNavigate(`/api-reference?category=${cat}`);
                    onSelectCategory(cat);
                    if (onClose) onClose();
                  }}
                  className={`w-full flex items-center justify-between px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${isSelected
                    ? 'bg-blue-500/10 text-blue-600 dark:text-blue-400 font-semibold'
                    : 'text-slate-600 dark:text-slate-400 hover:bg-slate-100 dark:hover:bg-slate-800/40 hover:text-slate-900 dark:hover:text-slate-200'
                    }`}
                >
                  <span className="truncate">{cat}</span>
                  <span className="text-[10px] px-1.5 py-0.5 rounded-full bg-slate-100 dark:bg-slate-800 text-slate-500">
                    {count}
                  </span>
                </button>
              );
            })}
          </div>
        </div>

        {/* Sidebar Footer */}
        <div className="mt-auto pt-4 border-t border-slate-200/80 dark:border-slate-800/80 flex items-center justify-between text-xs text-slate-400">
          <span>v3.0.0-alpha.0</span>
          <a
            href="https://github.com/jsuyog2/node-pptx-templater"
            target="_blank"
            rel="noopener noreferrer"
            className="flex items-center gap-1 hover:text-blue-500 transition-colors"
          >
            GitHub <ExternalLink className="w-3 h-3" />
          </a>
        </div>
      </aside>
    </>
  );
}
