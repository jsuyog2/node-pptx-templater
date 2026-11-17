'use client';

import React, { useState, useEffect, useRef } from 'react';
import { useRouter } from 'next/navigation';
import { Search, X, Code2, BookOpen, Layers, ArrowRight, CornerDownLeft } from 'lucide-react';
import { searchItems } from '../utils/searchIndex';

export default function GlobalSearchModal({ isOpen, onClose }) {
  const [query, setQuery] = useState('');
  const [results, setResults] = useState([]);
  const [selectedIndex, setSelectedIndex] = useState(0);
  const inputRef = useRef(null);
  const router = useRouter();

  // Handle Cmd+K / Ctrl+K and Escape
  useEffect(() => {
    const handleKeyDown = (e) => {
      if ((e.metaKey || e.ctrlKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault();
        if (isOpen) {
          onClose();
        }
      } else if (e.key === 'Escape' && isOpen) {
        e.preventDefault();
        onClose();
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [isOpen, onClose]);

  // Autofocus input when modal opens
  useEffect(() => {
    if (isOpen) {
      setQuery('');
      setResults(searchItems(''));
      setSelectedIndex(0);
      setTimeout(() => {
        inputRef.current?.focus();
      }, 50);
    }
  }, [isOpen]);

  // Update results on query change
  useEffect(() => {
    if (!isOpen) return;
    const res = searchItems(query, 12);
    setResults(res);
    setSelectedIndex(0);
  }, [query, isOpen]);

  const handleSelect = (item) => {
    if (!item) return;
    onClose();
    if (item.href) {
      router.push(item.href);
      if (typeof window !== 'undefined' && item.href.includes('#')) {
        const hash = item.href.slice(item.href.indexOf('#'));
        setTimeout(() => {
          window.location.hash = hash;
          window.dispatchEvent(new HashChangeEvent('hashchange'));
        }, 100);
      }
    }
  };

  // Keyboard navigation through search results
  const handleInputKeyDown = (e) => {
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setSelectedIndex((prev) => (results.length > 0 ? (prev + 1) % results.length : 0));
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setSelectedIndex((prev) => (results.length > 0 ? (prev - 1 + results.length) % results.length : 0));
    } else if (e.key === 'Enter') {
      e.preventDefault();
      const targetItem = results[selectedIndex] || results[0];
      if (targetItem) {
        handleSelect(targetItem);
      }
    }
  };

  if (!isOpen) return null;

  // Helper to highlight matching text in query
  const highlightMatch = (text, q) => {
    if (!q || !text) return text;
    const regex = new RegExp(`(${q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')})`, 'gi');
    const parts = text.split(regex);
    return parts.map((part, i) =>
      regex.test(part) ? (
        <span key={i} className="text-[#D24726] font-bold underline underline-offset-2">
          {part}
        </span>
      ) : (
        part
      )
    );
  };

  return (
    <div
      role="dialog"
      aria-modal="true"
      aria-label="Global Documentation Omnisearch"
      className="fixed inset-0 z-50 flex items-start justify-center pt-14 sm:pt-20 px-3 bg-black/60 backdrop-blur-sm animate-in fade-in duration-150"
      onClick={onClose}
    >
      <div
        className="w-full max-w-2xl ppt-card-surface rounded-xl border border-[var(--ppt-card-border)] shadow-2xl overflow-hidden flex flex-col bg-[#E6DEC8] dark:bg-[rgba(38,38,38,0.98)]"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Search Input Bar */}
        <div className="flex items-center px-4 py-3 border-b border-[var(--ppt-card-border)] gap-2.5">
          <Search className="w-4 h-4 text-[#D24726] flex-shrink-0" />
          <input
            ref={inputRef}
            type="text"
            value={query}
            onChange={(e) => setQuery(e.target.value)}
            onKeyDown={handleInputKeyDown}
            placeholder="Search 135+ methods, guides, categories... (e.g., useSlide, updateChart, table)"
            className="flex-1 bg-transparent text-sm text-[var(--ppt-text-primary)] placeholder-[var(--ppt-text-secondary)]/60 focus:outline-none"
          />
          {query && (
            <button
              type="button"
              onClick={() => setQuery('')}
              className="p-1 rounded text-[var(--ppt-text-secondary)] hover:text-[var(--ppt-text-primary)] cursor-pointer"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          )}
          <span className="hidden sm:inline text-[10px] px-1.5 py-0.5 rounded bg-black/10 dark:bg-white/10 font-mono text-[var(--ppt-text-secondary)]">
            ESC
          </span>
        </div>

        {/* Results List */}
        <div className="max-h-[380px] overflow-y-auto p-2 space-y-1 overscroll-contain">
          {results.length > 0 ? (
            results.map((item, index) => {
              const isSelected = index === selectedIndex;
              const isApi = item.type === 'api';
              const isCat = item.type === 'category';
              return (
                <div
                  key={item.id}
                  onClick={() => handleSelect(item)}
                  onMouseEnter={() => setSelectedIndex(index)}
                  className={`px-3 py-2.5 rounded-lg cursor-pointer transition-colors flex items-center justify-between border ${
                    isSelected
                      ? 'border-[#D24726] bg-[#D24726]/10 text-[var(--ppt-text-primary)]'
                      : 'border-transparent hover:bg-black/5 dark:hover:bg-white/5 text-[var(--ppt-text-secondary)]'
                  }`}
                >
                  <div className="flex items-center gap-2.5 overflow-hidden">
                    <div
                      className={`w-7 h-7 rounded flex items-center justify-center flex-shrink-0 ${
                        isApi
                          ? 'bg-[#D24726]/15 text-[#D24726]'
                          : isCat
                          ? 'bg-blue-500/15 text-blue-500'
                          : 'bg-emerald-500/15 text-emerald-500'
                      }`}
                    >
                      {isApi ? (
                        <Code2 className="w-3.5 h-3.5" />
                      ) : isCat ? (
                        <Layers className="w-3.5 h-3.5" />
                      ) : (
                        <BookOpen className="w-3.5 h-3.5" />
                      )}
                    </div>

                    <div className="truncate">
                      <div className="flex items-center gap-2">
                        <span className="text-xs font-bold text-[var(--ppt-text-primary)] font-mono">
                          {highlightMatch(item.title, query)}
                        </span>
                        <span className="text-[9px] px-1.5 py-0.2 rounded bg-black/10 dark:bg-white/10 font-semibold text-[#D24726] uppercase tracking-wider">
                          {item.category}
                        </span>
                      </div>
                      <p className="text-[11px] text-[var(--ppt-text-secondary)] truncate mt-0.5">
                        {item.subtitle}
                      </p>
                    </div>
                  </div>

                  <div className="flex items-center gap-1.5 text-xs text-[var(--ppt-text-secondary)] flex-shrink-0 ml-2">
                    {isSelected && (
                      <span className="hidden sm:flex items-center gap-0.5 text-[10px] text-[#D24726] font-mono">
                        <CornerDownLeft className="w-3 h-3" />
                      </span>
                    )}
                    <ArrowRight className="w-3.5 h-3.5 opacity-60" />
                  </div>
                </div>
              );
            })
          ) : (
            <div className="py-8 text-center text-xs text-[var(--ppt-text-secondary)]">
              No matching methods, topics, or guides found for &quot;{query}&quot;.
            </div>
          )}
        </div>

        {/* Footer shortcuts */}
        <div className="px-3 py-2 border-t border-[var(--ppt-card-border)] bg-black/5 dark:bg-black/20 flex items-center justify-between text-[10px] text-[var(--ppt-text-secondary)] select-none">
          <div className="flex items-center gap-3">
            <span>↑↓ Navigate</span>
            <span>↵ Select</span>
            <span>ESC Close</span>
          </div>
          <span className="font-mono text-[#D24726] font-semibold">Omnisearch v4.3</span>
        </div>
      </div>
    </div>
  );
}
