'use client';

import React, { useEffect, useRef } from 'react';
import { Layers } from 'lucide-react';

export default function MobileSlideFilmstrip({
  slides = [],
  activeSlideIndex = 0,
  onSelectSlide = () => { },
}) {
  const activeThumbnailRef = useRef(null);

  useEffect(() => {
    if (activeThumbnailRef.current) {
      activeThumbnailRef.current.scrollIntoView({
        behavior: 'smooth',
        inline: 'center',
        block: 'nearest',
      });
    }
  }, [activeSlideIndex]);

  return (
    <div
      className="lg:hidden fixed bottom-0 inset-x-0 z-30 h-24 sm:h-24 ppt-chrome-surface border-t border-[var(--ppt-card-border)] backdrop-blur-md flex flex-col justify-center shadow-lg select-none"
      aria-label="Mobile Slide Navigation Filmstrip"
    >
      {/* Filmstrip Header Bar */}
      <div className="flex items-center justify-between px-3 pt-1 text-[10px] text-[var(--ppt-text-secondary)]">
        <div className="flex items-center gap-1 font-semibold text-[var(--ppt-text-primary)]">
          <Layers className="w-3 h-3 text-[#D24726]" />
          <span>Slide {activeSlideIndex + 1} of {slides.length}</span>
        </div>
        <span className="font-mono text-[#D24726] font-semibold">Tap thumbnail to switch</span>
      </div>

      {/* Horizontal Filmstrip Cards */}
      <div className="flex items-center gap-2 overflow-x-auto px-3 py-1.5 no-scrollbar snap-x snap-mandatory">
        {slides.map((slide, index) => {
          const isActive = index === activeSlideIndex;
          return (
            <button
              key={slide.id || index}
              ref={isActive ? activeThumbnailRef : null}
              type="button"
              onClick={() => onSelectSlide(index)}
              className={`snap-center flex-shrink-0 w-28 sm:w-32 aspect-[16/9] rounded p-1.5 flex flex-col justify-between text-left transition-all duration-200 ppt-card-surface border ${isActive
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
  );
}
