'use client';

import gsap from 'gsap';

/**
 * Animate hero elements on load: badge, title, subtitle, CTAs, install bar
 */
export function animateHero(containerRef) {
  if (!containerRef || typeof window === 'undefined') return;

  const ctx = gsap.context(() => {
    gsap.from('.hero-badge', {
      opacity: 0,
      y: -20,
      scale: 0.95,
      duration: 0.5,
      ease: 'back.out(1.7)'
    });

    gsap.from('.hero-title', {
      opacity: 0,
      y: 25,
      duration: 0.7,
      delay: 0.1,
      ease: 'power3.out'
    });

    gsap.from('.hero-description', {
      opacity: 0,
      y: 20,
      duration: 0.6,
      delay: 0.25,
      ease: 'power2.out'
    });

    gsap.fromTo('.hero-cta',
      {
        opacity: 0,
        y: 20
      },
      {
        opacity: 1,
        y: 0,
        stagger: 0.1,
        duration: 0.5,
        delay: 0.35,
        ease: 'power2.out'
      }
    );

    gsap.from('.hero-install', {
      opacity: 0,
      scale: 0.95,
      duration: 0.5,
      delay: 0.5,
      ease: 'power2.out'
    });
  }, containerRef);

  return () => ctx.revert();
}

/**
 * Animate feature grid cards when the features page mounts
 */
export function animateFeaturesGrid(containerRef, selector = '.feature-card') {
  if (!containerRef || typeof window === 'undefined') return;

  const ctx = gsap.context(() => {
    gsap.from(selector, {
      opacity: 0,
      y: 30,
      stagger: 0.06,
      duration: 0.55,
      ease: 'power2.out',
      clearProps: 'all'
    });
  }, containerRef);

  return () => ctx.revert();
}

/**
 * Animate API cards when category changes or filter updates
 */
export function animateApiCards(containerRef, selector = '.api-card') {
  if (!containerRef || typeof window === 'undefined') return;

  const ctx = gsap.context(() => {
    gsap.from(selector, {
      opacity: 0,
      y: 20,
      stagger: 0.04,
      duration: 0.45,
      ease: 'power2.out',
      clearProps: 'transform,opacity'
    });
  }, containerRef);

  return () => ctx.revert();
}

/**
 * Animate page container entrance
 */
export function animatePageEnter(containerRef) {
  if (!containerRef || typeof window === 'undefined') return;

  const ctx = gsap.context(() => {
    gsap.from(containerRef, {
      opacity: 0,
      y: 15,
      duration: 0.4,
      ease: 'power2.out'
    });
  }, containerRef);

  return () => ctx.revert();
}
