'use client';

import gsap from 'gsap';

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
