'use client';

import React, { useState } from 'react';
import Image from 'next/image';
import Link from 'next/link';
import { Save, Undo2, Redo2, PlaySquare, Check, Search, Minus, Square, X } from 'lucide-react';
import ThemeToggle from '../ThemeToggle';

export default function TitleBar({ title = 'node-pptx-templater.pptx — PowerPoint', onSlideShowClick }) {
  const [saveStatus, setSaveStatus] = useState('Save');
  const [saved, setSaved] = useState(false);

  const handleQuickSave = () => {
    setSaved(true);
    setSaveStatus('Saving...');
    setTimeout(() => {
      setSaveStatus('Saved');
      setSaved(false);
    }, 1200);
  };

  return (
    <header className="h-[36px] min-h-[36px] ppt-chrome-surface border-b border-[var(--ppt-chrome-border)] flex items-center justify-between px-2 text-xs select-none z-50 sticky top-0">
      {/* Left: PPT Badge, Quick Access Toolbar, AutoSave */}
      <div className="flex items-center gap-1.5 sm:gap-2 overflow-hidden">
        {/* PowerPoint App Badge */}
        <Link
          href="/"
          className="flex items-center gap-1.5 px-1.5 py-0.5 rounded font-bold text-white transition-colors"
          title="node-pptx-templater Home"
        >
          <Image src="/favicon.svg" width={16} height={16} alt="Logo" />
        </Link>

        {/* Quick Access Icons */}
        <div className="flex items-center gap-0.5 text-[var(--ppt-text-secondary)]">
          <button
            type="button"
            onClick={handleQuickSave}
            title={`Quick Save (${saveStatus})`}
            aria-label="Save presentation"
            className="p-1 rounded hover:bg-black/10 dark:hover:bg-white/10 hover:text-[var(--ppt-text-primary)] transition-colors"
          >
            {saved ? (
              <Check className="w-3.5 h-3.5 text-emerald-500 animate-pulse" />
            ) : (
              <Save className="w-3.5 h-3.5" />
            )}
          </button>
          <button
            type="button"
            title="Undo (Ctrl+Z)"
            aria-label="Undo"
            className="p-1 rounded hover:bg-black/10 dark:hover:bg-white/10 hover:text-[var(--ppt-text-primary)] transition-colors opacity-70 hover:opacity-100"
          >
            <Undo2 className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            title="Redo (Ctrl+Y)"
            aria-label="Redo"
            className="p-1 rounded hover:bg-black/10 dark:hover:bg-white/10 hover:text-[var(--ppt-text-primary)] transition-colors opacity-70 hover:opacity-100"
          >
            <Redo2 className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={onSlideShowClick}
            title="Start Slide Show from Beginning (F5)"
            aria-label="Start Slide Show"
            className="p-1 rounded hover:bg-black/10 dark:hover:bg-white/10 text-[#D24726] transition-colors"
          >
            <PlaySquare className="w-3.5 h-3.5" />
          </button>
        </div>

        <div className="hidden md:block h-3 w-px bg-[var(--ppt-chrome-border)] mx-1" />

        {/* AutoSave Tag */}
        <div className="hidden lg:flex items-center gap-1.5 text-[10px] text-[var(--ppt-text-secondary)]">
          <span className="w-1.5 h-1.5 rounded-full bg-emerald-500 animate-pulse" />
          <span className="font-medium">AutoSave On</span>
          <span className="opacity-60">• {saveStatus}</span>
        </div>
      </div>

      {/* Center: File Title */}
      <div className="flex-1 max-w-xl mx-2 text-center truncate">
        <span className="text-xs font-semibold text-[var(--ppt-text-primary)] tracking-tight">
          {title}
        </span>
      </div>

      {/* Right: Quick Search, ThemeToggle, Window Controls */}
      <div className="flex items-center gap-1">
        {/* Command Search hint */}
        <div className="hidden xl:flex items-center gap-1.5 px-2 py-0.5 rounded text-[11px] bg-black/5 dark:bg-white/5 text-[var(--ppt-text-secondary)] border border-[var(--ppt-chrome-border)]/50 mr-1">
          <Search className="w-3 h-3 opacity-60" />
          <span>Tell me what you want to do (Alt+Q)</span>
        </div>

        {/* Theme Toggle */}
        <div className="flex items-center">
          <ThemeToggle className="!p-1 !rounded !border-none !shadow-none !bg-transparent hover:!bg-black/10 dark:hover:!bg-white/10" />
        </div>

        {/* Window Controls */}
        <div className="flex items-center ml-1 -mr-2">
          <button
            type="button"
            aria-label="Minimize Window"
            title="Minimize"
            className="w-8 h-[36px] flex items-center justify-center text-[var(--ppt-text-secondary)] hover:bg-black/10 dark:hover:bg-white/10 hover:text-[var(--ppt-text-primary)] transition-colors"
          >
            <Minus className="w-3 h-3" />
          </button>
          <button
            type="button"
            aria-label="Maximize / Restore Window"
            title="Maximize"
            className="w-8 h-[36px] flex items-center justify-center text-[var(--ppt-text-secondary)] hover:bg-black/10 dark:hover:bg-white/10 hover:text-[var(--ppt-text-primary)] transition-colors"
          >
            <Square className="w-2.5 h-2.5" />
          </button>
          <button
            type="button"
            aria-label="Close Window"
            title="Close"
            className="w-9 h-[36px] flex items-center justify-center text-[var(--ppt-text-secondary)] hover:bg-rose-600 hover:text-white transition-colors"
          >
            <X className="w-3.5 h-3.5" />
          </button>
        </div>
      </div>
    </header>
  );
}
