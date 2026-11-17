'use client';

import gsap from 'gsap';
import { ScrollTrigger } from 'gsap/ScrollTrigger';

if (typeof window !== 'undefined') {
  gsap.registerPlugin(ScrollTrigger);
}

/**
 * Checks vertical scroll boundary for a DOM container
 * @param {HTMLElement|null} element
 * @returns {{ isAtTop: boolean, isAtBottom: boolean, isScrollable: boolean }}
 */
export function checkScrollBoundary(element) {
  if (!element) {
    return { isAtTop: true, isAtBottom: true, isScrollable: false };
  }

  const isScrollable = element.scrollHeight > element.clientHeight + 2;
  const isAtTop = element.scrollTop <= 1;
  const isAtBottom = element.scrollTop + element.clientHeight >= element.scrollHeight - 2;

  return { isAtTop, isAtBottom, isScrollable };
}

/**
 * Finds the closest scrollable container for event target
 * @param {HTMLElement} target
 * @param {HTMLElement} boundary
 * @returns {HTMLElement|null}
 */
export function findScrollableContainer(target, boundary) {
  let current = target;
  while (current && current !== document.body) {
    if (current.scrollHeight > current.clientHeight + 2) {
      const overflowY = window.getComputedStyle(current).overflowY;
      if (overflowY === 'auto' || overflowY === 'scroll') {
        return current;
      }
    }
    if (boundary && current === boundary) break;
    current = current.parentElement;
  }
  return boundary;
}

/**
 * Programmatically animate horizontal folder track
 * @param {HTMLElement} trackEl
 * @param {number} targetIndex
 * @param {number} totalPanels
 * @param {number} duration
 * @returns {gsap.core.Tween|null}
 */
export function slideToFolder(trackEl, targetIndex = 0, totalPanels = 4, duration = 0.5) {
  if (!trackEl || typeof window === 'undefined') return null;

  const targetX = -(targetIndex * (100 / totalPanels));
  return gsap.to(trackEl, {
    xPercent: targetX,
    duration,
    ease: 'power2.out',
    overwrite: 'auto',
  });
}

/**
 * Responsive ScrollTrigger setup using gsap.matchMedia()
 * Desktop (>= 1024px): Enabled with pin/scrub
 * Mobile/Tablet (< 1024px): Completely killed and reverted for natural touch scrolling
 */
export function initResponsiveFolderScrollTrigger(
  containerEl,
  trackEl,
  panelCount = 4,
  onProgress = () => {}
) {
  if (!containerEl || !trackEl || typeof window === 'undefined') return () => {};

  const mm = gsap.matchMedia();

  // Desktop only (>= 1024px): pin & horizontal scrub
  mm.add('(min-width: 1024px)', () => {
    gsap.to(trackEl, {
      xPercent: -100 * (panelCount - 1) / panelCount,
      ease: 'none',
      scrollTrigger: {
        trigger: containerEl,
        pin: true,
        scrub: 0.8,
        start: 'top top',
        end: () => `+=${window.innerWidth * (panelCount - 1) * 0.85}`,
        onUpdate: (self) => {
          const activeIdx = Math.min(
            panelCount - 1,
            Math.floor(self.progress * panelCount)
          );
          onProgress(activeIdx, self.progress);
        },
      },
    });
  });

  // Mobile / Tablet (< 1024px): ScrollTrigger completely inactive
  mm.add('(max-width: 1023px)', () => {
    gsap.set(trackEl, { clearProps: 'transform' });
  });

  return () => mm.revert();
}
