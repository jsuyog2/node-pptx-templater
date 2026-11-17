'use client';

import gsap from 'gsap';
import { ScrollTrigger } from 'gsap/ScrollTrigger';

if (typeof window !== 'undefined') {
  gsap.registerPlugin(ScrollTrigger);
}

/**
 * Orchestrates staggered reveals on Hero elements using ScrollTrigger
 * @param {React.RefObject} containerRef
 */
export function animateHeroScrollTrigger(containerRef) {
  if (!containerRef?.current || typeof window === 'undefined') return () => {};

  const mm = gsap.matchMedia();

  // Desktop only (>= 1024px): staggered reveal with ScrollTrigger
  mm.add('(min-width: 1024px)', () => {
    // 1. PowerPoint Window Chrome Reveal
    gsap.from('.hero-window-chrome', {
      opacity: 0,
      y: 35,
      scale: 0.98,
      duration: 0.8,
      ease: 'power3.out',
    });

    // 2. Headline & Subtitle Stagger
    gsap.from('.hero-headline', {
      opacity: 0,
      y: 25,
      duration: 0.7,
      delay: 0.2,
      ease: 'power3.out',
    });

    gsap.from('.hero-subheadline', {
      opacity: 0,
      y: 20,
      duration: 0.6,
      delay: 0.35,
      ease: 'power2.out',
    });

    // 3. CTA Action Buttons & Install Bar
    gsap.from('.hero-action-buttons', {
      opacity: 0,
      y: 15,
      duration: 0.5,
      delay: 0.5,
      ease: 'power2.out',
    });

    // 4. Live Code Playground Card
    gsap.from('.hero-playground-card', {
      opacity: 0,
      y: 30,
      duration: 0.8,
      delay: 0.6,
      ease: 'power2.out',
    });

    // 5. Stat Badges Stagger
    gsap.from('.hero-stat-badge', {
      opacity: 0,
      scale: 0.9,
      duration: 0.5,
      stagger: 0.1,
      delay: 0.75,
      ease: 'back.out(1.5)',
    });
  });

  // Mobile/Tablet (< 1024px): ScrollTrigger disabled, clear props for natural layout flow
  mm.add('(max-width: 1023px)', () => {
    gsap.set(
      [
        '.hero-window-chrome',
        '.hero-headline',
        '.hero-subheadline',
        '.hero-action-buttons',
        '.hero-playground-card',
        '.hero-stat-badge',
      ],
      { clearProps: 'all' }
    );
  });

  return () => mm.revert();
}

/**
 * Transitions presentation slide card using smooth Push / Slide-up effect
 * @param {HTMLElement} stageElement
 * @param {'next' | 'prev'} direction
 * @param {Function} onComplete
 */
export function transitionSlide(stageElement, direction = 'next', onComplete = () => {}) {
  if (!stageElement || typeof window === 'undefined') {
    onComplete();
    return;
  }

  gsap.killTweensOf(stageElement);

  const yOffset = direction === 'next' ? 40 : -40;

  gsap.fromTo(
    stageElement,
    {
      opacity: 0,
      y: yOffset,
      scale: 0.985,
    },
    {
      opacity: 1,
      y: 0,
      scale: 1,
      duration: 0.42,
      ease: 'power2.out',
      onComplete,
    }
  );
}

/**
 * Slide push transition for deck switching
 * @param {HTMLElement} currentElement
 * @param {HTMLElement} nextElement
 * @param {'next' | 'prev'} direction
 */
export function animateSlideDeckPush(currentElement, nextElement, direction = 'next') {
  if (typeof window === 'undefined') return;

  const yOut = direction === 'next' ? -50 : 50;
  const yIn = direction === 'next' ? 50 : -50;

  if (currentElement) {
    gsap.to(currentElement, {
      y: yOut,
      opacity: 0,
      duration: 0.35,
      ease: 'power2.in',
    });
  }

  if (nextElement) {
    gsap.fromTo(
      nextElement,
      { y: yIn, opacity: 0 },
      { y: 0, opacity: 1, duration: 0.42, ease: 'power2.out', delay: 0.1 }
    );
  }
}
