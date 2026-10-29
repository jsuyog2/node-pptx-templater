'use client';

import React, { useState, useEffect, useRef, useCallback } from 'react';
import SlideThumbnailSidebar from './PowerPointChrome/SlideThumbnailSidebar';
import MobileSlideFilmstrip from './PowerPointChrome/MobileSlideFilmstrip';
import StatusBar from './PowerPointChrome/StatusBar';
import { transitionSlide } from '../animations/gsapSlideTransitions';
import { ChevronUp, ChevronDown } from 'lucide-react';

export default function PowerPointSlideDeck({
  slides = [],
  initialIndex = 0,
  categoryTitle = 'Documentation',
  onSlideChange = () => { },
}) {
  const [activeIndex, setActiveIndex] = useState(initialIndex);
  const [viewMode, setViewMode] = useState('normal');
  const [zoomLevel, setZoomLevel] = useState(100);
  const stageRef = useRef(null);
  const mainRef = useRef(null);
  const isTransitioningRef = useRef(false);
  const lastTransitionTimeRef = useRef(0);
  const touchStartYRef = useRef(0);
  const touchStartXRef = useRef(0);

  const COOLDOWN_MS = 380;
  const totalSlides = slides.length || 1;

  // Jump to specific slide with GSAP transition and cooldown debounce
  const goToSlide = useCallback(
    (newIndex, direction = 'next') => {
      const now = Date.now();
      if (now - lastTransitionTimeRef.current < COOLDOWN_MS) return;
      if (newIndex < 0 || newIndex >= totalSlides || newIndex === activeIndex) return;
      if (isTransitioningRef.current) return;

      lastTransitionTimeRef.current = now;
      isTransitioningRef.current = true;
      setActiveIndex(newIndex);
      onSlideChange(newIndex);

      // Reset scroll position to top for incoming slide
      if (stageRef.current) {
        stageRef.current.scrollTop = 0;
      }
      if (mainRef.current) {
        mainRef.current.scrollTop = 0;
      }

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

  // Helper to find scrollable container for event target
  const getScrollableElement = (target) => {
    let el = target;
    while (el && el !== document.body && el !== mainRef.current) {
      if (el.scrollHeight > el.clientHeight + 2) {
        const overflowY = window.getComputedStyle(el).overflowY;
        if (overflowY === 'auto' || overflowY === 'scroll') {
          return el;
        }
      }
      if (el === stageRef.current) break;
      el = el.parentElement;
    }
    return stageRef.current;
  };

  // Handle Mouse Wheel with Boundary Snapping
  const handleWheel = useCallback(
    (e) => {
      if (Math.abs(e.deltaY) < 15) return;

      const scrollEl = getScrollableElement(e.target);
      const isScrollable = scrollEl && scrollEl.scrollHeight > scrollEl.clientHeight + 2;
      const isAtTop = scrollEl ? scrollEl.scrollTop <= 0 : true;
      const isAtBottom = scrollEl
        ? scrollEl.scrollTop + scrollEl.clientHeight >= scrollEl.scrollHeight - 2
        : true;

      if (isScrollable) {
        if (e.deltaY > 0) {
          // Scrolling downward: if not yet at bottom of card, permit native card scroll
          if (!isAtBottom) return;
          // Reached bottom boundary: advance to next slide
          e.preventDefault();
          nextSlide();
        } else {
          // Scrolling upward: if not yet at top of card, permit native card scroll
          if (!isAtTop) return;
          // Reached top boundary: revert to previous slide
          e.preventDefault();
          prevSlide();
        }
      } else {
        // Not internally scrollable: transition directly
        if (e.deltaY > 0) {
          e.preventDefault();
          nextSlide();
        } else {
          e.preventDefault();
          prevSlide();
        }
      }
    },
    [nextSlide, prevSlide]
  );

  // Handle Touch Gestures with Boundary Snapping
  const handleTouchStart = (e) => {
    if (e.touches && e.touches[0]) {
      touchStartYRef.current = e.touches[0].clientY;
      touchStartXRef.current = e.touches[0].clientX;
    }
  };

  const handleTouchMove = (e) => {
    if (!e.touches || !e.touches[0]) return;
    const currentY = e.touches[0].clientY;
    const currentX = e.touches[0].clientX;
    const deltaY = touchStartYRef.current - currentY; // positive = swipe up = scroll down
    const deltaX = touchStartXRef.current - currentX;

    // Horizontal swipe gesture prioritization on mobile decks
    if (Math.abs(deltaX) > Math.abs(deltaY) && Math.abs(deltaX) > 45) {
      if (deltaX > 0) {
        nextSlide();
      } else {
        prevSlide();
      }
      touchStartXRef.current = currentX;
      return;
    }

    if (Math.abs(deltaY) < 30) return;

    const scrollEl = getScrollableElement(e.target);
    const isScrollable = scrollEl && scrollEl.scrollHeight > scrollEl.clientHeight + 2;
    const isAtTop = scrollEl ? scrollEl.scrollTop <= 0 : true;
    const isAtBottom = scrollEl
      ? scrollEl.scrollTop + scrollEl.clientHeight >= scrollEl.scrollHeight - 2
      : true;

    if (isScrollable) {
      if (deltaY > 0) {
        // Scrolling down: allow native card scroll until bottom boundary is hit
        if (!isAtBottom) return;
        nextSlide();
        touchStartYRef.current = currentY;
      } else {
        // Scrolling up: allow native card scroll until top boundary is hit
        if (!isAtTop) return;
        prevSlide();
        touchStartYRef.current = currentY;
      }
    } else {
      if (deltaY > 0) {
        nextSlide();
        touchStartYRef.current = currentY;
      } else {
        prevSlide();
        touchStartYRef.current = currentY;
      }
    }
  };

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
    <div className="flex h-[calc(100vh-104px)] overflow-hidden relative w-full">
      {/* Pinned Left Slide Thumbnail Sidebar (Desktop only) */}
      <SlideThumbnailSidebar
        slides={slides}
        activeSlideIndex={activeIndex}
        onSelectSlide={(idx) => goToSlide(idx, idx > activeIndex ? 'next' : 'prev')}
      />

      {/* Center Slide Stage Canvas */}
      <main
        ref={mainRef}
        onWheel={handleWheel}
        onTouchStart={handleTouchStart}
        onTouchMove={handleTouchMove}
        className="flex-1 flex flex-col items-center justify-start p-3 sm:p-6 lg:p-8 overflow-y-auto pb-28 lg:pb-16 relative w-full max-w-full"
      >
        {/* Stage Floating Navigation Controls (Desktop) */}
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
          className="w-full max-w-5xl ppt-slide-card rounded-xl p-4 sm:p-8 lg:p-10 border border-[var(--ppt-card-border)] overflow-y-auto relative my-auto min-h-[460px] max-h-[calc(100vh-140px)] flex flex-col justify-between transition-transform duration-150"
        >
          {/* Slide Top Accent Bar & Meta */}
          <div>
            <div className="flex items-center justify-between pb-3 sm:pb-4 mb-3 sm:mb-4 border-b border-[var(--ppt-card-border)]">
              <div className="flex items-center gap-2">
                <span className="px-2.5 py-0.5 rounded text-[10px] sm:text-[11px] font-bold bg-[#D24726] text-white tracking-wide">
                  {activeSlide.category || categoryTitle}
                </span>
              </div>

              <div className="flex items-center gap-2 text-xs text-[var(--ppt-text-secondary)]">
                <span className="font-mono text-[11px]">Slide #{activeIndex + 1} of {totalSlides}</span>
                <div className="w-2 h-2 rounded-full bg-[#D24726]" />
              </div>
            </div>

            {/* Slide Title & Subtitle */}
            <div className="space-y-1 mb-4">
              <h1 className="text-xl sm:text-2xl lg:text-3xl font-extrabold text-[var(--ppt-text-primary)] tracking-tight">
                {activeSlide.title}
              </h1>
              {activeSlide.subtitle && (
                <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)]">
                  {activeSlide.subtitle}
                </p>
              )}
            </div>

            {/* Slide Body / Main Content */}
            <div className="text-xs sm:text-sm text-[var(--ppt-text-primary)] leading-relaxed space-y-4">
              {activeSlide.content}
            </div>
          </div>

          {/* Slide Footer Info Bar */}
          <div className="pt-4 sm:pt-6 mt-6 border-t border-[var(--ppt-card-border)]/60 flex items-center justify-between text-[10px] sm:text-[11px] text-[var(--ppt-text-secondary)]">
            <div className="flex items-center gap-2">
              <span className="font-semibold text-[var(--ppt-text-primary)]">node-pptx-templater</span>
            </div>

            <div className="font-mono text-xs text-[var(--ppt-text-primary)]">
              {activeIndex + 1}
            </div>
          </div>
        </div>
      </main>

      {/* Bottom Horizontal Filmstrip for Mobile and Tablet (< 1024px) */}
      <MobileSlideFilmstrip
        slides={slides}
        activeSlideIndex={activeIndex}
        onSelectSlide={(idx) => goToSlide(idx, idx > activeIndex ? 'next' : 'prev')}
      />

      {/* Bottom Status Bar (Desktop) */}
      <StatusBar
        currentSlide={activeIndex + 1}
        totalSlides={totalSlides}
        viewMode={viewMode}
        onViewModeChange={setViewMode}
        zoomLevel={zoomLevel}
        onZoomChange={setZoomLevel}
        className="hidden lg:flex"
      />
    </div>
  );
}
