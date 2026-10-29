'use client';

import React, { useState, useEffect, useRef } from 'react';
import { Layers, ChevronUp, ChevronDown } from 'lucide-react';

export default function MobileSlideFilmstrip({
  slides = [],
  activeSlideIndex = 0,
  activeOnfirst,
  onSelectSlide = () => { },
}) {
  const [isExpanded, setIsExpanded] = useState(activeOnfirst);
  const activeThumbnailRef = useRef(null);

  useEffect(() => {
    if (isExpanded && activeThumbnailRef.current) {
      activeThumbnailRef.current.scrollIntoView({
        behavior: 'smooth',
        inline: 'center',
        block: 'nearest',
      });
    }
  }, [activeSlideIndex, isExpanded]);

  const activeSlide = slides[activeSlideIndex] || slides[0] || {};

  return (
    <div
      className="lg:hidden fixed bottom-0 inset-x-0 w-full max-w-full overflow-hidden z-30 select-none transition-transform duration-300 ease-out"
      aria-label="Mobile Slide Navigation Filmstrip"
    >
      {/* Collapsible Drawer Container */}
      <div className="ppt-chrome-surface border-t border-[var(--ppt-card-border)] backdrop-blur-md shadow-2xl flex flex-col">
        {/* Compact Toggle Handle Bar */}
        <button
          type="button"
          onClick={() => setIsExpanded((prev) => !prev)}
          className="w-full flex items-center justify-between px-3.5 py-2 text-xs font-semibold text-[var(--ppt-text-primary)] hover:bg-black/5 dark:hover:bg-white/5 transition-colors cursor-pointer"
          aria-expanded={isExpanded}
          aria-label={isExpanded ? 'Collapse slide filmstrip' : 'Expand slide filmstrip'}
        >
          <div className="flex items-center gap-1.5 truncate mr-2">
            <Layers className="w-3.5 h-3.5 text-[#D24726] flex-shrink-0" />
            <span className="flex-shrink-0">
              Slide {activeSlideIndex + 1} of {slides.length}
            </span>
            <span className="text-[10px] text-[var(--ppt-text-secondary)] font-normal truncate">
              — {activeSlide.title || activeSlide.category || ''}
            </span>
          </div>

          <div className="flex items-center gap-1 text-[11px] text-[#D24726] font-medium flex-shrink-0">
            <span>{isExpanded ? 'Hide' : 'Slides'}</span>
            {isExpanded ? (
              <ChevronDown className="w-3.5 h-3.5 transition-transform duration-200" />
            ) : (
              <ChevronUp className="w-3.5 h-3.5 transition-transform duration-200" />
            )}
          </div>
        </button>

        {/* Expandable Horizontal Filmstrip Tray with Hardware-Accelerated Slide-Up / Slide-Down Animation */}
        <div
          className={`overflow-hidden transition-all duration-300 ease-[cubic-bezier(0.16,1,0.3,1)] transform ${isExpanded
            ? 'max-h-40 opacity-100 translate-y-0'
            : 'max-h-0 opacity-0 translate-y-4 pointer-events-none'
            }`}
          aria-hidden={!isExpanded}
        >
          <div className="flex items-center gap-2 overflow-x-auto max-w-full overscroll-x-contain px-3 py-2 no-scrollbar snap-x snap-mandatory border-t border-[var(--ppt-card-border)]/50 bg-black/10 dark:bg-black/20">
            {slides.map((slide, index) => {
              const isActive = index === activeSlideIndex;
              return (
                <button
                  key={slide.id || index}
                  ref={isActive ? activeThumbnailRef : null}
                  type="button"
                  onClick={() => {
                    setIsExpanded(() => {
                      setTimeout(() => {
                        onSelectSlide(index);
                      }, 300);
                      return false;
                    })
                  }}
                  className={`snap-center flex-shrink-0 w-28 sm:w-32 aspect-[16/9] rounded p-1.5 flex flex-col justify-between text-left transition-all duration-200 ppt-card-surface border cursor-pointer ${isActive
                    ? 'border-[#D24726] ring-2 ring-[#D24726] scale-105 shadow-md'
                    : 'border-[var(--ppt-card-border)] opacity-70 hover:opacity-100'
                    }`}
                  aria-label={`Go to slide ${index + 1}: ${slide.title || ''}`}
                >
                  <div className="flex items-center justify-between gap-1 w-full">
                    <span className="text-[8px] font-bold text-[#D24726] truncate uppercase">
                      #{index + 1}
                    </span>
                    <span className="text-[7px] text-[var(--ppt-text-secondary)] truncate">
                      {slide.category || 'Slide'}
                    </span>
                  </div>
                  <p className="text-[9px] font-semibold text-[var(--ppt-text-primary)] truncate leading-tight w-full">
                    {slide.title || `Slide ${index + 1}`}
                  </p>
                </button>
              );
            })}
          </div>
        </div>
      </div>
    </div>
  );
}
