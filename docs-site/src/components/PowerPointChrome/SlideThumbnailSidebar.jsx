'use client';

import React, { useEffect, useRef } from 'react';
import { Layers, FileCode } from 'lucide-react';

export default function SlideThumbnailSidebar({
  slides = [],
  activeSlideIndex = 0,
  onSelectSlide = () => { },
  className = '',
}) {
  const activeItemRef = useRef(null);

  useEffect(() => {
    if (activeItemRef.current) {
      activeItemRef.current.scrollIntoView({
        behavior: 'smooth',
        block: 'nearest',
      });
    }
  }, [activeSlideIndex]);

  return (
    <aside
      className={`hidden lg:flex w-[280px] min-w-[280px] max-w-[280px] ppt-chrome-surface border-r border-[var(--ppt-chrome-border)] flex-col h-[calc(100vh-104px)] select-none relative ${className}`}
      aria-label="Presentation Slide Thumbnails"
    >
      {/* Sidebar Header with Opaque Theme Surface to prevent thumbnail collision */}
      <div className="sticky top-0 z-10 bg-[#E6DEC8] dark:bg-[rgba(38,38,38,0.98)] border-b border-[var(--ppt-chrome-border)] p-3 flex items-center justify-between shadow-sm flex-shrink-0">
        <div className="flex items-center gap-1.5 text-xs font-semibold text-[var(--ppt-text-primary)]">
          <Layers className="w-3.5 h-3.5 text-[#D24726]" />
          <span>Slides</span>
        </div>
        <span className="text-[10px] px-1.5 py-0.5 rounded bg-[#D24726]/10 text-[#D24726] font-medium font-mono">
          {slides.length} slides
        </span>
      </div>

      {/* Dedicated Scroll Container for Thumbnails with Explicit Padding */}
      <div className="flex-1 overflow-y-auto p-3 pt-3 space-y-3 overscroll-contain">
        {slides.map((slide, index) => {
          const isActive = index === activeSlideIndex;
          return (
            <div
              key={slide.id || index}
              ref={isActive ? activeItemRef : null}
              onClick={() => onSelectSlide(index)}
              className="flex items-start gap-2 cursor-pointer group"
              role="button"
              tabIndex={0}
              onKeyDown={(e) => {
                if (e.key === 'Enter' || e.key === ' ') {
                  e.preventDefault();
                  onSelectSlide(index);
                }
              }}
              aria-label={`Slide ${index + 1}: ${slide.title || 'Untitled Slide'}`}
            >
              {/* Slide Number */}
              <span
                className={`text-[11px] font-mono font-medium pt-1 w-4 text-right transition-colors ${isActive
                  ? 'text-[#D24726] font-bold'
                  : 'text-[var(--ppt-text-secondary)] group-hover:text-[var(--ppt-text-primary)]'
                  }`}
              >
                {index + 1}
              </span>

              {/* 16:9 Thumbnail Card */}
              <div
                className={`flex-1 aspect-[16/9] rounded-sm p-2 flex flex-col justify-between transition-all duration-200 relative overflow-hidden ppt-card-surface ${isActive
                  ? 'ppt-active-slide ring-1 ring-[#D24726]'
                  : 'hover:border-[#D24726]/60 opacity-80 hover:opacity-100'
                  }`}
              >
                {/* Mini Header / Category */}
                <div className="flex items-center justify-between gap-1">
                  <span className="text-[9px] font-semibold text-[#D24726] uppercase tracking-wider truncate">
                    {slide.category || 'Topic'}
                  </span>
                  {isActive && (
                    <span className="w-1.5 h-1.5 rounded-full bg-[#D24726] animate-pulse" />
                  )}
                </div>

                {/* Mini Slide Title */}
                <div className="my-auto">
                  <p className="text-[11px] font-bold text-[var(--ppt-text-primary)] leading-tight line-clamp-2">
                    {slide.title || `Slide ${index + 1}`}
                  </p>
                  {slide.subtitle && (
                    <p className="text-[9px] text-[var(--ppt-text-secondary)] truncate mt-0.5">
                      {slide.subtitle}
                    </p>
                  )}
                </div>

                {/* Mini Layout Mockup / Visual Indicator */}
                <div className="flex items-center justify-end pt-1 border-t border-[var(--ppt-card-border)]/50 text-[8px] text-[var(--ppt-text-secondary)]">

                  <FileCode className="w-2.5 h-2.5 opacity-60 flex-shrink-0" />
                </div>
              </div>
            </div>
          );
        })}
      </div>
    </aside>
  );
}
