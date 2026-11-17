'use client';

import React, { useState, useEffect, useRef, useCallback } from 'react';
import { usePathname } from 'next/navigation';
import SlideThumbnailSidebar from './PowerPointChrome/SlideThumbnailSidebar';
import MobileSlideFilmstrip from './PowerPointChrome/MobileSlideFilmstrip';
import StatusBar from './PowerPointChrome/StatusBar';
import { transitionSlide } from '../animations/gsapSlideTransitions';
import NavigationArrows from './NavigationArrows';

export default function PowerPointSlideDeck({
  slides = [],
  initialIndex = 0,
  categoryTitle = 'Documentation',
  activeOnfirst,
  onSlideChange = () => { },
}) {
  const pathname = usePathname();
  const isApiReference = pathname?.includes('/docs/api-reference');

  const [activeIndex, setActiveIndex] = useState(initialIndex);
  const [activeFunctionIndex, setActiveFunctionIndex] = useState(0);
  const [viewMode, setViewMode] = useState('normal');
  const [zoomLevel, setZoomLevel] = useState(100);
  const stageRef = useRef(null);
  const mainRef = useRef(null);
  const isTransitioningRef = useRef(false);
  const lastTransitionTimeRef = useRef(0);

  const lastStepTimeRef = useRef(0);
  const STEP_COOLDOWN_MS = 280;

  const prevInitialIndexRef = useRef(initialIndex);

  // Synchronize activeIndex only when initialIndex prop actually changes (e.g. category route change in API Reference)
  useEffect(() => {
    if (prevInitialIndexRef.current !== initialIndex) {
      prevInitialIndexRef.current = initialIndex;
      setActiveIndex(initialIndex);
      if (isApiReference) {
        setActiveFunctionIndex(0);
      }
    }
  }, [initialIndex, isApiReference]);

  // Wheel damping and boundary threshold refs
  const wheelAccumulatorRef = useRef(0);
  const lastScrollDirectionRef = useRef(0);
  const boundaryReachedTimeRef = useRef(0);

  // Touch gesture tracking refs
  const touchStartXRef = useRef(0);
  const touchStartYRef = useRef(0);
  const touchStartTimeRef = useRef(0);
  const touchIgnoredRef = useRef(false);

  const COOLDOWN_MS = 380;
  const WHEEL_THRESHOLD = 50;
  const totalSlides = slides.length || 1;

  // Jump to specific slide with GSAP transition and cooldown debounce
  const goToSlide = useCallback(
    (newIndex, direction = 'next', targetFunctionIndex = 0) => {
      const now = Date.now();
      if (now - lastTransitionTimeRef.current < COOLDOWN_MS) return;
      if (newIndex < 0 || newIndex >= totalSlides || newIndex === activeIndex) return;
      if (isTransitioningRef.current) return;

      lastTransitionTimeRef.current = now;
      lastStepTimeRef.current = now;
      isTransitioningRef.current = true;
      wheelAccumulatorRef.current = 0;
      boundaryReachedTimeRef.current = 0;
      setActiveIndex(newIndex);
      setActiveFunctionIndex(targetFunctionIndex);
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
    goToSlide(activeIndex + 1, 'next', 0);
  }, [activeIndex, goToSlide]);

  const prevSlide = useCallback(() => {
    goToSlide(activeIndex - 1, 'prev', 0);
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

  // Helper to detect if touch originated inside horizontally scrollable element
  const isInsideHorizontalScrollable = (target) => {
    let el = target;
    while (el && el !== document.body && el !== stageRef.current) {
      if (el.tagName === 'PRE' || el.tagName === 'CODE' || el.tagName === 'TABLE') {
        return true;
      }
      const overflowX = window.getComputedStyle(el).overflowX;
      if (
        (overflowX === 'auto' || overflowX === 'scroll') &&
        el.scrollWidth > el.clientWidth + 2
      ) {
        return true;
      }
      el = el.parentElement;
    }
    return false;
  };

  // Handle Mouse Wheel with Explicit Threshold Damping and Direction Reset
  const handleWheel = useCallback(
    (e) => {
      if (Math.abs(e.deltaY) < 10) return;

      if (isTransitioningRef.current) {
        e.preventDefault();
        return;
      }

      const currentDirection = e.deltaY > 0 ? 1 : -1;
      // Reset accumulator if user reverses scroll direction
      if (lastScrollDirectionRef.current !== currentDirection) {
        wheelAccumulatorRef.current = 0;
        lastScrollDirectionRef.current = currentDirection;
        boundaryReachedTimeRef.current = 0;
      }

      const scrollEl = getScrollableElement(e.target);
      const isScrollable = scrollEl && scrollEl !== stageRef.current && scrollEl.scrollHeight > scrollEl.clientHeight + 2;
      const isAtTop = scrollEl ? scrollEl.scrollTop <= 1 : true;
      const isAtBottom = scrollEl
        ? scrollEl.scrollTop + scrollEl.clientHeight >= scrollEl.scrollHeight - 2
        : true;

      const now = Date.now();

      // =========================================================================
      // API REFERENCE SECTION: Function-First Step-Scrolling & Slide Boundary Handoff
      // =========================================================================
      if (isApiReference) {
        const currentSlide = slides[activeIndex] || {};
        const methods = currentSlide.methods || [];
        const totalFunctions = methods.length;

        // If active function card has internal overflow (long code blocks or parameter tables),
        // allow internal card scrolling first.
        if (isScrollable) {
          if (e.deltaY > 0) {
            // Scrolling downward: if not yet at bottom of card, permit free native scroll
            if (!isAtBottom) {
              wheelAccumulatorRef.current = 0;
              boundaryReachedTimeRef.current = 0;
              return;
            }
          } else {
            // Scrolling upward: if not yet at top of card, permit free native scroll
            if (!isAtTop) {
              wheelAccumulatorRef.current = 0;
              boundaryReachedTimeRef.current = 0;
              return;
            }
          }

          // Card boundary reached: require pause / debounce before stepping function
          if (!boundaryReachedTimeRef.current) {
            boundaryReachedTimeRef.current = now;
            wheelAccumulatorRef.current = 0;
            return;
          }
          if (now - boundaryReachedTimeRef.current < 200) {
            return;
          }
        }

        // Add a debounce threshold (~250ms–350ms cooldown) to prevent a single physical
        // wheel tick or trackpad flick from skipping multiple functions.
        if (now - lastStepTimeRef.current < STEP_COOLDOWN_MS) {
          return;
        }

        // Wheel threshold damping
        wheelAccumulatorRef.current += Math.abs(e.deltaY);
        if (wheelAccumulatorRef.current < WHEEL_THRESHOLD) {
          return;
        }

        wheelAccumulatorRef.current = 0;
        boundaryReachedTimeRef.current = 0;

        if (e.deltaY > 0) {
          // Downward Scroll (deltaY > 0)
          // Hard terminal lockout: Last Slide AND Last Function
          if (activeIndex === totalSlides - 1 && activeFunctionIndex >= totalFunctions - 1) {
            return;
          }

          e.preventDefault();

          if (activeFunctionIndex < totalFunctions - 1) {
            // Sub-Step A: Increment activeFunctionIndex by 1. Keep current slide unchanged.
            lastStepTimeRef.current = now;
            setActiveFunctionIndex((prev) => prev + 1);
          } else if (activeFunctionIndex === totalFunctions - 1 && activeIndex < totalSlides - 1) {
            // Sub-Step B (Handoff to Next Slide): Advance to next slide and set activeFunctionIndex = 0
            lastStepTimeRef.current = now;
            goToSlide(activeIndex + 1, 'next', 0);
          }
        } else {
          // Upward Scroll (deltaY < 0)
          // Hard terminal lockout: First Slide AND First Function
          if (activeIndex === 0 && activeFunctionIndex === 0) {
            return;
          }

          e.preventDefault();

          if (activeFunctionIndex > 0) {
            // Sub-Step A: Decrement activeFunctionIndex by 1. Keep current slide unchanged.
            lastStepTimeRef.current = now;
            setActiveFunctionIndex((prev) => prev - 1);
          } else if (activeFunctionIndex === 0 && activeIndex > 0) {
            // Sub-Step B (Handoff to Previous Slide): Revert to previous slide and select its last function
            lastStepTimeRef.current = now;
            const prevSlideMethods = slides[activeIndex - 1]?.methods || [];
            const targetFuncIdx = Math.max(0, prevSlideMethods.length - 1);
            goToSlide(activeIndex - 1, 'prev', targetFuncIdx);
          }
        }
        return;
      }

      // =========================================================================
      // NON-API REFERENCE SECTIONS (Overview, Installation, Home):
      // Direct slide-to-slide navigation with internal card scroll handling
      // =========================================================================
      // Maintain a 300ms–350ms cooldown debounce so a single scroll gesture advances exactly one slide at a time
      if (now - lastTransitionTimeRef.current < 320) {
        return;
      }

      if (isScrollable) {
        if (e.deltaY > 0) {
          // If the slide card content overflows vertically, permit native scrolling until reaching the boundary
          if (!isAtBottom) {
            wheelAccumulatorRef.current = 0;
            boundaryReachedTimeRef.current = 0;
            return;
          }

          // User reached the bottom boundary
          if (!boundaryReachedTimeRef.current) {
            boundaryReachedTimeRef.current = now;
            wheelAccumulatorRef.current = 0;
            return;
          }

          wheelAccumulatorRef.current += e.deltaY;

          // Require intentional extra delta and delay before advancing slide
          if (
            wheelAccumulatorRef.current < WHEEL_THRESHOLD ||
            now - boundaryReachedTimeRef.current < 160
          ) {
            return;
          }

          wheelAccumulatorRef.current = 0;
          boundaryReachedTimeRef.current = 0;

          // Downward Scroll (deltaY > 0): Advance to next slide or do nothing at last slide
          if (activeIndex < totalSlides - 1) {
            e.preventDefault();
            nextSlide();
          }
        } else {
          // If the slide card content overflows vertically, permit native scrolling until reaching the boundary
          if (!isAtTop) {
            wheelAccumulatorRef.current = 0;
            boundaryReachedTimeRef.current = 0;
            return;
          }

          // User reached the top boundary
          if (!boundaryReachedTimeRef.current) {
            boundaryReachedTimeRef.current = now;
            wheelAccumulatorRef.current = 0;
            return;
          }

          wheelAccumulatorRef.current += Math.abs(e.deltaY);

          // Require intentional extra delta and delay before reverting slide
          if (
            wheelAccumulatorRef.current < WHEEL_THRESHOLD ||
            now - boundaryReachedTimeRef.current < 160
          ) {
            return;
          }

          wheelAccumulatorRef.current = 0;
          boundaryReachedTimeRef.current = 0;

          // Upward Scroll (deltaY < 0): Revert to previous slide or do nothing at first slide
          if (activeIndex > 0) {
            e.preventDefault();
            prevSlide();
          }
        }
      } else {
        // Not internally scrollable: require deliberate wheel threshold
        wheelAccumulatorRef.current += Math.abs(e.deltaY);
        if (wheelAccumulatorRef.current < WHEEL_THRESHOLD) return;

        wheelAccumulatorRef.current = 0;
        boundaryReachedTimeRef.current = 0;

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
    [
      isApiReference,
      activeIndex,
      activeFunctionIndex,
      slides,
      totalSlides,
      goToSlide,
      nextSlide,
      prevSlide,
    ]
  );

  // Handle Touch Gestures with Horizontal Drag & Swipe Detection
  const handleTouchStart = (e) => {
    if (e.touches && e.touches[0]) {
      // If touch originates in horizontally scrollable table or code block, ignore slide swipe
      if (isInsideHorizontalScrollable(e.target)) {
        touchIgnoredRef.current = true;
        return;
      }
      touchIgnoredRef.current = false;
      touchStartXRef.current = e.touches[0].clientX;
      touchStartYRef.current = e.touches[0].clientY;
      touchStartTimeRef.current = Date.now();
    }
  };

  const handleTouchMove = (e) => {
    if (touchIgnoredRef.current || !e.touches || !e.touches[0]) return;
    // Native vertical scrolling occurs uninterrupted during move
  };

  const handleTouchEnd = (e) => {
    if (touchIgnoredRef.current) return;
    if (!e.changedTouches || !e.changedTouches[0]) return;

    const endX = e.changedTouches[0].clientX;
    const endY = e.changedTouches[0].clientY;
    const deltaX = touchStartXRef.current - endX; // positive = swipe right-to-left
    const deltaY = touchStartYRef.current - endY; // positive = swipe bottom-to-top
    const absX = Math.abs(deltaX);
    const absY = Math.abs(deltaY);

    // Mobile Horizontal Swipe: absX >= 60px and predominantly horizontal (absX > absY * 1.5)
    if (absX >= 60 && absX > absY * 1.5) {
      if (deltaX > 0) {
        // Swiping right-to-left -> Next Slide
        nextSlide();
      } else {
        // Swiping left-to-right -> Previous Slide
        prevSlide();
      }
      return;
    }

    // Mobile Vertical Boundary Swipe: absY >= 75px, predominantly vertical, and card at boundary
    if (absY >= 75 && absY > absX * 1.5) {
      const scrollEl = getScrollableElement(e.target);
      const isScrollable = scrollEl && scrollEl !== stageRef.current && scrollEl.scrollHeight > scrollEl.clientHeight + 2;
      const isAtTop = scrollEl ? scrollEl.scrollTop <= 1 : true;
      const isAtBottom = scrollEl
        ? scrollEl.scrollTop + scrollEl.clientHeight >= scrollEl.scrollHeight - 2
        : true;

      if (isApiReference) {
        const currentSlide = slides[activeIndex] || {};
        const methods = currentSlide.methods || [];
        const totalFunctions = methods.length;

        if (deltaY > 0 && (!isScrollable || isAtBottom)) {
          // Swipe up = scroll down
          if (activeIndex === totalSlides - 1 && activeFunctionIndex >= totalFunctions - 1) {
            return;
          }
          if (activeFunctionIndex < totalFunctions - 1) {
            setActiveFunctionIndex((prev) => prev + 1);
          } else if (activeFunctionIndex === totalFunctions - 1 && activeIndex < totalSlides - 1) {
            goToSlide(activeIndex + 1, 'next', 0);
          }
        } else if (deltaY < 0 && (!isScrollable || isAtTop)) {
          // Swipe down = scroll up
          if (activeIndex === 0 && activeFunctionIndex === 0) {
            return;
          }
          if (activeFunctionIndex > 0) {
            setActiveFunctionIndex((prev) => prev - 1);
          } else if (activeFunctionIndex === 0 && activeIndex > 0) {
            const prevSlideMethods = slides[activeIndex - 1]?.methods || [];
            const targetFuncIdx = Math.max(0, prevSlideMethods.length - 1);
            goToSlide(activeIndex - 1, 'prev', targetFuncIdx);
          }
        }
        return;
      }

      if (isScrollable) {
        if (deltaY > 0 && isAtBottom) {
          if (activeIndex < totalSlides - 1) {
            nextSlide();
          }
        } else if (deltaY < 0 && isAtTop) {
          if (activeIndex > 0) {
            prevSlide();
          }
        }
      } else {
        if (deltaY > 0) {
          if (activeIndex < totalSlides - 1) {
            nextSlide();
          }
        } else if (deltaY < 0) {
          if (activeIndex > 0) {
            prevSlide();
          }
        }
      }
    }
  };

  // Keyboard navigation
  useEffect(() => {
    const handleKeyDown = (e) => {
      if (e.target.tagName === 'INPUT' || e.target.tagName === 'TEXTAREA') {
        return;
      }

      if (e.key === 'ArrowDown' || e.key === 'PageDown' || e.key === ' ') {
        e.preventDefault();
        if (isApiReference) {
          const currentSlide = slides[activeIndex] || {};
          const methods = currentSlide.methods || [];
          const totalFunctions = methods.length;
          if (activeIndex === totalSlides - 1 && activeFunctionIndex >= totalFunctions - 1) {
            return;
          }
          if (activeFunctionIndex < totalFunctions - 1) {
            setActiveFunctionIndex((prev) => prev + 1);
          } else if (activeFunctionIndex === totalFunctions - 1 && activeIndex < totalSlides - 1) {
            goToSlide(activeIndex + 1, 'next', 0);
          }
        } else {
          nextSlide();
        }
      } else if (e.key === 'ArrowUp' || e.key === 'PageUp') {
        e.preventDefault();
        if (isApiReference) {
          if (activeIndex === 0 && activeFunctionIndex === 0) {
            return;
          }
          if (activeFunctionIndex > 0) {
            setActiveFunctionIndex((prev) => prev - 1);
          } else if (activeFunctionIndex === 0 && activeIndex > 0) {
            const prevSlideMethods = slides[activeIndex - 1]?.methods || [];
            const targetFuncIdx = Math.max(0, prevSlideMethods.length - 1);
            goToSlide(activeIndex - 1, 'prev', targetFuncIdx);
          }
        } else {
          prevSlide();
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [
    isApiReference,
    activeIndex,
    activeFunctionIndex,
    slides,
    totalSlides,
    goToSlide,
    nextSlide,
    prevSlide,
  ]);

  const activeSlide = slides[activeIndex] || slides[0] || {};

  return (
    <div className="flex h-[calc(100vh-104px)] overflow-hidden relative w-full max-w-full">
      {/* Pinned Left Slide Thumbnail Sidebar (Desktop only) */}
      <SlideThumbnailSidebar
        slides={slides}
        activeSlideIndex={activeIndex}
        onSelectSlide={(idx) => goToSlide(idx, idx > activeIndex ? 'next' : 'prev', 0)}
        sectionTitle={categoryTitle}
      />

      {/* Center Slide Stage Canvas */}
      <main
        ref={mainRef}
        onWheel={handleWheel}
        onTouchStart={handleTouchStart}
        onTouchMove={handleTouchMove}
        onTouchEnd={handleTouchEnd}
        className="flex-1 flex flex-col items-center justify-start p-2 sm:p-6 lg:p-8 overflow-hidden pb-16 relative w-full max-w-full overscroll-contain isolate box-border"
      >
        {/* Stage Directional Arrow Navigation Controls (Accessible on Mobile, Tablet & Desktop) */}
        <div className="fixed right-10 sm:right-20 bottom-11 sm:bottom-9 z-20">
          <NavigationArrows
            currentIndex={activeIndex}
            totalItems={totalSlides}
            onPrev={() => goToSlide(activeIndex - 1, 'prev', 0)}
            onNext={() => goToSlide(activeIndex + 1, 'next', 0)}
            label="Slide"
          />
        </div>

        {/* 16:9 Aspect Ratio Center Presentation Slide Card with Strict Boundary Containment */}
        <div
          ref={stageRef}
          style={{ transform: `scale(${zoomLevel / 100})`, transformOrigin: 'top center' }}
          className="w-full max-w-7xl ppt-slide-card rounded-xl p-2.5 sm:p-6 lg:p-8 border border-[var(--ppt-card-border)] relative my-auto min-h-[460px] max-h-[calc(100vh-140px)] flex flex-col justify-between overflow-hidden overscroll-contain isolate transition-transform duration-150 box-border"
        >
          {activeSlide.hideDefaultHeader ? (
            /* Custom Slide Card Header & Content (Used by ApiReferenceCard for Unified Header) */
            <div className="flex-1 min-h-0 min-w-0 flex flex-col overflow-hidden w-full max-w-full overscroll-contain isolate box-border">
              {React.isValidElement(activeSlide.content)
                ? React.cloneElement(activeSlide.content, {
                    activeFunctionIndex: isApiReference ? activeFunctionIndex : undefined,
                    onFunctionIndexChange: isApiReference ? setActiveFunctionIndex : undefined,
                  })
                : activeSlide.content}
            </div>
          ) : (
            /* Standard Presentation Slide Layout (Used by Guides & Installation Decks) */
            <div className="flex-1 min-h-0 overflow-y-auto overscroll-contain isolate w-full max-w-full box-border">
              {/* Slide Top Accent Bar & Meta */}
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

              {/* Slide Title & Subtitle / Category Description */}
              <div className="space-y-1 mb-4">
                <h1 className="text-xl sm:text-2xl lg:text-3xl font-extrabold text-[var(--ppt-text-primary)] tracking-tight">
                  {activeSlide.title}
                </h1>
                {(activeSlide.categoryDescription || activeSlide.subtitle) && (
                  <p className="text-xs sm:text-sm text-[var(--ppt-text-secondary)] leading-relaxed max-w-3xl">
                    {activeSlide.categoryDescription || activeSlide.subtitle}
                  </p>
                )}
              </div>

              {/* Slide Body / Main Content */}
              <div className="text-xs sm:text-sm text-[var(--ppt-text-primary)] leading-relaxed space-y-4">
                {activeSlide.content}
              </div>
            </div>
          )}

          {/* Slide Footer Info Bar */}
          <div className="pt-3 sm:pt-4 mt-3 sm:mt-4 border-t border-[var(--ppt-card-border)]/60 flex items-center justify-between text-[10px] sm:text-[11px] text-[var(--ppt-text-secondary)] flex-shrink-0">
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
        activeOnfirst={activeOnfirst}
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
