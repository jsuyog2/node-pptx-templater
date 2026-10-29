'use client';

import gsap from 'gsap';
import { ScrollTrigger } from 'gsap/ScrollTrigger';

if (typeof window !== 'undefined') {
  gsap.registerPlugin(ScrollTrigger);
}

/**
 * Initializes GSAP ScrollTrigger horizontal translation with pin and scrub
 * @param {HTMLElement} containerEl
 * @param {HTMLElement} trackEl
 * @param {number} panelCount
 * @param {Function} onProgress
 */
export function initHorizontalFolderScrollTrigger(containerEl, trackEl, panelCount = 4, onProgress = () => { }) {
  if (!containerEl || !trackEl || typeof window === 'undefined') return () => { };

  // Only enable horizontal ScrollTrigger on desktop viewports
  if (window.innerWidth < 1024) return () => { };

  const ctx = gsap.context(() => {
    const totalPanels = panelCount;
    gsap.to(trackEl, {
      xPercent: -100.5 * (totalPanels - 1) / totalPanels,
      ease: 'none',
      scrollTrigger: {
        trigger: containerEl,
        pin: true,
        scrub: 0.8,
        start: 0,
        end: () => `+=${window.innerWidth * (totalPanels - 1) * 0.85}`,
        onUpdate: (self) => {
          const activeIdx = Math.min(
            totalPanels - 1,
            Math.floor(self.progress * totalPanels)
          );
          onProgress(activeIdx, self.progress);
        },
      },
    });
  }, containerEl);

  return () => ctx.revert();
}

/**
 * Programmatically animate horizontal track to a specific folder index
 * @param {HTMLElement} trackEl
 * @param {number} targetIndex
 * @param {number} totalPanels
 * @param {number} duration
 */
export function slideToFolder(trackEl, targetIndex = 0, totalPanels = 4, duration = 0.55) {
  if (!trackEl || typeof window === 'undefined') return;

  const targetX = -(targetIndex * (100 / totalPanels));
  gsap.to(trackEl, {
    xPercent: targetX,
    duration,
    ease: 'power2.out',
  });
}
