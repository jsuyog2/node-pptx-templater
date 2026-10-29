'use client';

import React from 'react';
import { ChevronLeft, ChevronRight } from 'lucide-react';

/**
 * Accessible Directional Navigation Arrow Controls
 * Provides Previous and Next navigation across PowerPoint slides and Home folder panels.
 * Fully accessible on mobile, tablet, and desktop viewports.
 *
 * @param {{
 *   currentIndex: number,
 *   totalItems: number,
 *   onPrev: () => void,
 *   onNext: () => void,
 *   label?: string,
 *   className?: string,
 *   showCounter?: boolean,
 * }} props
 */
export default function NavigationArrows({
  currentIndex = 0,
  totalItems = 1,
  onPrev,
  onNext,
  label = 'Slide',
  className = '',
  showCounter = true,
}) {
  const isFirst = currentIndex <= 0;
  const isLast = currentIndex >= totalItems - 1;

  return (
    <div
      role="group"
      aria-label={`${label} Navigation Controls`}
      className={`inline-flex items-center gap-1.5 sm:gap-2 select-none max-w-full lg:hidden ${className}`}
    >
      {/* Previous Button */}
      <button
        type="button"
        onClick={onPrev}
        disabled={isFirst}
        aria-label={`Previous ${label} (Slide #${currentIndex})`}
        title={`Previous ${label}`}
        className="p-1.5 sm:p-2.5 rounded-full ppt-card-surface border border-[var(--ppt-card-border)] text-[var(--ppt-text-primary)] hover:text-[#D24726] hover:border-[#D24726] disabled:opacity-30 disabled:cursor-not-allowed disabled:hover:text-[var(--ppt-text-primary)] disabled:hover:border-[var(--ppt-card-border)] transition-all shadow-md active:scale-95 cursor-pointer flex items-center justify-center focus:outline-none focus:ring-2 focus:ring-[#D24726]"
      >
        <ChevronLeft className="w-4 h-4 sm:w-4.5 sm:h-4.5" />
      </button>

      {/* Progress Counter */}
      {showCounter && (
        <span
          className="text-[11px] sm:text-xs font-mono font-medium px-2 py-1 rounded ppt-card-surface border border-[var(--ppt-card-border)] text-[var(--ppt-text-secondary)] shadow-sm"
          aria-live="polite"
          aria-atomic="true"
        >
          <span className="text-[var(--ppt-text-primary)] font-semibold">
            {currentIndex + 1}
          </span>
          <span className="mx-1 opacity-60">/</span>
          <span>{totalItems}</span>
        </span>
      )}

      {/* Next Button */}
      <button
        type="button"
        onClick={onNext}
        disabled={isLast}
        aria-label={`Next ${label} (Slide #${currentIndex + 2})`}
        title={`Next ${label}`}
        className="p-1.5 sm:p-2.5 rounded-full ppt-card-surface border border-[var(--ppt-card-border)] text-[var(--ppt-text-primary)] hover:text-[#D24726] hover:border-[#D24726] disabled:opacity-30 disabled:cursor-not-allowed disabled:hover:text-[var(--ppt-text-primary)] disabled:hover:border-[var(--ppt-card-border)] transition-all shadow-md active:scale-95 cursor-pointer flex items-center justify-center focus:outline-none focus:ring-2 focus:ring-[#D24726]"
      >
        <ChevronRight className="w-4 h-4 sm:w-4.5 sm:h-4.5" />
      </button>
    </div>
  );
}
