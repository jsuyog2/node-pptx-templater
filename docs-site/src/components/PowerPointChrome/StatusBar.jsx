'use client';

import React from 'react';
import {
  LayoutTemplate,
  Grid,
  BookOpen,
  PlaySquare,
  Minus,
  Plus,
  Maximize2,
  CheckCircle2,
  Cpu
} from 'lucide-react';

export default function StatusBar({
  currentSlide = 1,
  totalSlides = 1,
  viewMode = 'normal',
  onViewModeChange = () => { },
  zoomLevel = 100,
  onZoomChange = () => { },
  className = '',
}) {
  return (
    <footer className={`h-[32px] min-h-[32px] ppt-chrome-surface border-t border-[var(--ppt-chrome-border)] items-center justify-between px-3 text-[11px] text-[var(--ppt-text-secondary)] select-none z-50 fixed bottom-0 left-0 right-0 w-full max-w-full overflow-hidden ${className || 'flex'}`}>
      {/* Left: Slide Counter & OpenXML Indicator */}
      <div className="flex items-center gap-3 sm:gap-5 overflow-hidden">
        {/* Slide Counter */}
        <div className="flex items-center gap-1 font-medium text-[var(--ppt-text-primary)]">
          <span>Slide</span>
          <span className="font-bold text-[#D24726]">{currentSlide}</span>
          <span>of</span>
          <span>{totalSlides}</span>
        </div>

        <div className="hidden sm:block h-3 w-px bg-[var(--ppt-chrome-border)]" />

        {/* Pure JavaScript OpenXML Indicator */}
        <div className="hidden sm:flex items-center gap-1.5 text-[10px]">
          <Cpu className="w-3 h-3 text-[#D24726]" />
          <span className="font-medium text-[var(--ppt-text-primary)]">Pure JavaScript Engine</span>
          <span className="hidden md:inline opacity-60">(Zero Office/Java Dependencies)</span>
        </div>
      </div>

      {/* Right: View Modes & Zoom Controls */}
      <div className="flex items-center gap-2 sm:gap-4">
        {/* View Mode Selectors */}
        <div className="flex items-center gap-0.5">
          <button
            type="button"
            onClick={() => onViewModeChange('normal')}
            title="Normal View"
            aria-label="Normal View"
            className={`p-1 rounded hover:bg-black/10 dark:hover:bg-white/10 transition-colors ${viewMode === 'normal'
              ? 'text-[#D24726] bg-black/5 dark:bg-white/5 font-semibold'
              : 'text-[var(--ppt-text-secondary)]'
              }`}
          >
            <LayoutTemplate className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => onViewModeChange('sorter')}
            title="Slide Sorter"
            aria-label="Slide Sorter"
            className={`p-1 rounded hover:bg-black/10 dark:hover:bg-white/10 transition-colors ${viewMode === 'sorter'
              ? 'text-[#D24726] bg-black/5 dark:bg-white/5 font-semibold'
              : 'text-[var(--ppt-text-secondary)]'
              }`}
          >
            <Grid className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => onViewModeChange('reading')}
            title="Reading View"
            aria-label="Reading View"
            className={`p-1 rounded hover:bg-black/10 dark:hover:bg-white/10 transition-colors ${viewMode === 'reading'
              ? 'text-[#D24726] bg-black/5 dark:bg-white/5 font-semibold'
              : 'text-[var(--ppt-text-secondary)]'
              }`}
          >
            <BookOpen className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => onViewModeChange('slideshow')}
            title="Slide Show (F5)"
            aria-label="Slide Show"
            className="p-1 rounded hover:bg-black/10 dark:hover:bg-white/10 text-[#D24726] transition-colors"
          >
            <PlaySquare className="w-3.5 h-3.5" />
          </button>
        </div>

        <div className="h-3 w-px bg-[var(--ppt-chrome-border)]" />

        {/* Zoom Controls */}
        <div className="flex items-center gap-1.5">
          <button
            type="button"
            onClick={() => onZoomChange(Math.max(50, zoomLevel - 10))}
            title="Zoom Out"
            aria-label="Zoom Out"
            className="p-0.5 rounded hover:bg-black/10 dark:hover:bg-white/10 transition-colors"
          >
            <Minus className="w-3 h-3" />
          </button>

          <input
            type="range"
            min="50"
            max="150"
            step="5"
            value={zoomLevel}
            onChange={(e) => onZoomChange(Number(e.target.value))}
            className="w-14 sm:w-20 h-1 accent-[#D24726] cursor-pointer"
            aria-label="Zoom Slider"
          />

          <button
            type="button"
            onClick={() => onZoomChange(Math.min(150, zoomLevel + 10))}
            title="Zoom In"
            aria-label="Zoom In"
            className="p-0.5 rounded hover:bg-black/10 dark:hover:bg-white/10 transition-colors"
          >
            <Plus className="w-3 h-3" />
          </button>

          <span className="w-8 text-right font-mono text-[10px] text-[var(--ppt-text-primary)]">
            {zoomLevel}%
          </span>

          <button
            type="button"
            onClick={() => onZoomChange(100)}
            title="Fit slide to current window"
            aria-label="Fit slide to window"
            className="p-0.5 rounded hover:bg-black/10 dark:hover:bg-white/10 transition-colors ml-0.5"
          >
            <Maximize2 className="w-3 h-3" />
          </button>
        </div>
      </div>
    </footer>
  );
}
