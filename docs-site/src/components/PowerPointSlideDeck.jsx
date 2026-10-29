'use client';

import React, { useState, useEffect, useRef, useCallback } from 'react';
import SlideThumbnailSidebar from './PowerPointChrome/SlideThumbnailSidebar';
import StatusBar from './PowerPointChrome/StatusBar';
import { transitionSlide } from '../animations/gsapSlideTransitions';
import { ChevronUp, ChevronDown, Maximize2, Sparkles, Layers } from 'lucide-react';

export default function PowerPointSlideDeck({
  slides = [],
  initialIndex = 0,
  categoryTitle = 'Documentation',
  onSlideChange = () => {},
}) {
  const [activeIndex, setActiveIndex] = useState(initialIndex);
  const [viewMode, setViewMode] = useState('normal');
  const [zoomLevel, setZoomLevel] = useState(100);
  const stageRef = useRef(null);
  const isTransitioningRef = useRef(false);

  const totalSlides = slides.length || 1;

  // Jump to specific slide with GSAP transition
  const goToSlide = useCallback(
    (newIndex, direction = 'next') => {
      if (newIndex < 0 || newIndex >= totalSlides || newIndex === activeIndex) return;
      if (isTransitioningRef.current) return;

      isTransitioningRef.current = true;
      setActiveIndex(newIndex);
      onSlideChange(newIndex);

      if (stageRef.current) {
        transitionSlide(stageRef.current, direction, () => {
          isTransitioningRef.current = false;
        });
      } else {
        isTransitioningRef.current = false;
      }
    },
    [activeIndex, totalSlides, onSlideChange]
  );

  const nextSlide = useCallback(() => {
    goToSlide(activeIndex + 1, 'next');
  }, [activeIndex, goToSlide]);

  const prevSlide = useCallback(() => {
    goToSlide(activeIndex - 1, 'prev');
  }, [activeIndex, goToSlide]);

  // Handle Mouse Wheel Scroll within Center Slide Stage
  const handleWheel = useCallback(
    (e) => {
      // If user scrolls inside a scrollable child (e.g. pre or code block), don't hijack unless boundary is reached
      const target = e.target;
      const isInsideScrollable = target.closest('pre, code, .overflow-y-auto');

      if (Math.abs(e.deltaY) > 35 && !isTransitioningRef.current) {
        if (e.deltaY > 0) {
          if (activeIndex < totalSlides - 1) {
            e.preventDefault();
            nextSlide();
          }
        } else {
          if (activeIndex > 0) {
            e.preventDefault();
            prevSlide();
          }
        }
      }
    },
    [activeIndex, totalSlides, nextSlide, prevSlide]
  );

  // Keyboard navigation
  useEffect(() => {
    const handleKeyDown = (e) => {
      if (e.key === 'ArrowDown' || e.key === 'PageDown' || e.key === ' ') {
        if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'TEXTAREA') {
          e.preventDefault();
          nextSlide();
        }
      } else if (e.key === 'ArrowUp' || e.key === 'PageUp') {
        if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'TEXTAREA') {
          e.preventDefault();
          prevSlide();
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [nextSlide, prevSlide]);

  const activeSlide = slides[activeIndex] || slides[0] || {};

  return (
    <div className="flex h-[calc(100vh-104px)] overflow-hidden relative">
      {/* Pinned Left Slide Thumbnail Sidebar */}
      <SlideThumbnailSidebar
        slides={slides}
        activeSlideIndex={activeIndex}
        onSelectSlide={(idx) => goToSlide(idx, idx > activeIndex ? 'next' : 'prev')}
      />

      {/* Center Slide Stage Canvas */}
      <main
        onWheel={handleWheel}
        className="flex-1 flex flex-col items-center justify-start p-4 sm:p-6 lg:p-8 overflow-y-auto pb-16 relative"
      >
        {/* Stage Floating Navigation Controls */}
        <div className="fixed right-6 top-[120px] hidden lg:flex flex-col gap-2 z-20">
          <button
            type="button"
            onClick={prevSlide}
            disabled={activeIndex === 0}
            aria-label="Previous Slide (Arrow Up)"
            title="Previous Slide (Arrow Up)"
            className="p-2 rounded-full ppt-card-surface hover:border-[#D24726] text-[var(--ppt-text-primary)] disabled:opacity-30 disabled:cursor-not-allowed transition-all shadow-md"
          >
            <ChevronUp className="w-4 h-4" />
          </button>
          <span className="text-[10px] font-mono text-center text-[var(--ppt-text-secondary)]">
            {activeIndex + 1}/{totalSlides}
          </span>
          <button
            type="button"
            onClick={nextSlide}
            disabled={activeIndex === totalSlides - 1}
            aria-label="Next Slide (Arrow Down)"
            title="Next Slide (Arrow Down)"
            className="p-2 rounded-full ppt-card-surface hover:border-[#D24726] text-[var(--ppt-text-primary)] disabled:opacity-30 disabled:cursor-not-allowed transition-all shadow-md"
          >
            <ChevronDown className="w-4 h-4" />
          </button>
        </div>

        {/* 16:9 Aspect Ratio Center Presentation Slide Card */}
        <div
          ref={stageRef}
          style={{ transform: `scale(${zoomLevel / 100})`, transformOrigin: 'top center' }}
          className="w-full max-w-5xl ppt-slide-card rounded-xl p-6 sm:p-10 border border-[var(--ppt-card-border)] relative my-auto min-h-[500px] flex flex-col justify-between transition-transform duration-150"
        >
          {/* Slide Top Accent Bar & Meta */}
          <div>
            <div className="flex items-center justify-between pb-4 mb-4 border-b border-[var(--ppt-card-border)]">
              <div className="flex items-center gap-2">
                <span className="px-2.5 py-0.5 rounded text-[11px] font-bold bg-[#D24726] text-white tracking-wide">
                  {activeSlide.category || categoryTitle}
                </span>
                <span className="text-xs text-[var(--ppt-text-secondary)] font-mono">
                  Slide #{activeIndex + 1}
                </span>
              </div>

              <div className="flex items-center gap-2 text-xs text-[var(--ppt-text-secondary)]">
                <span className="hidden sm:inline">16:9 Widescreen</span>
                <div className="w-2 h-2 rounded-full bg-[#D24726]" />
              </div>
            </div>

            {/* Slide Title & Subtitle */}
            <div className="space-y-1 mb-6">
              <h1 className="text-2xl sm:text-3xl font-extrabold text-[var(--ppt-text-primary)] tracking-tight">
                {activeSlide.title}
              </h1>
              {activeSlide.subtitle && (
                <p className="text-sm sm:text-base text-[var(--ppt-text-secondary)]">
                  {activeSlide.subtitle}
                </p>
              )}
            </div>

            {/* Slide Body / Main Content */}
            <div className="text-sm text-[var(--ppt-text-primary)] leading-relaxed space-y-4">
              {activeSlide.content}
            </div>
          </div>

          {/* Slide Footer Info Bar */}
          <div className="pt-6 mt-8 border-t border-[var(--ppt-card-border)]/60 flex items-center justify-between text-[11px] text-[var(--ppt-text-secondary)]">
            <div className="flex items-center gap-2">
              <span className="font-semibold text-[var(--ppt-text-primary)]">node-pptx-templater</span>
              <span>•</span>
              <span>ISO/IEC 29500 PresentationML</span>
            </div>

            <div className="font-mono text-xs text-[var(--ppt-text-primary)]">
              {activeIndex + 1}
            </div>
          </div>
        </div>
      </main>

      {/* Bottom Status Bar */}
      <StatusBar
        currentSlide={activeIndex + 1}
        totalSlides={totalSlides}
        viewMode={viewMode}
        onViewModeChange={setViewMode}
        zoomLevel={zoomLevel}
        onZoomChange={setZoomLevel}
      />
    </div>
  );
}
