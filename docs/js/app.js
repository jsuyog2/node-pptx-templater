/**
 * @fileoverview Main Interactive Application Script for node-pptx-templater Docs
 *
 * Implements:
 *   - Dark/Light Theme Switcher with localStorage & system preference
 *   - Copy to Clipboard with accessible feedback and toasts
 *   - GSAP Subtle Animations with prefers-reduced-motion compliance
 *   - Active Navigation Scrollspy
 *   - Mobile Menu Toggle
 *   - Interactive Showcase & Tab Switching
 *   - Feature Catalog Filter & Search
 */

'use strict';

(function () {
  // ==========================================================================
  // 1. THEME MANAGEMENT
  // ==========================================================================

  function getPreferredTheme() {
    const saved = localStorage.getItem('theme');
    if (saved === 'dark' || saved === 'light') return saved;
    return window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  }

  function applyTheme(theme) {
    const root = document.documentElement;
    const themeIconLight = document.getElementById('theme-icon-light');
    const themeIconDark = document.getElementById('theme-icon-dark');

    if (theme === 'dark') {
      root.classList.add('dark');
      if (themeIconLight) themeIconLight.classList.remove('hidden');
      if (themeIconDark) themeIconDark.classList.add('hidden');
    } else {
      root.classList.remove('dark');
      if (themeIconLight) themeIconLight.classList.add('hidden');
      if (themeIconDark) themeIconDark.classList.remove('hidden');
    }
  }

  function toggleTheme() {
    const current = document.documentElement.classList.contains('dark') ? 'dark' : 'light';
    const next = current === 'dark' ? 'light' : 'dark';
    localStorage.setItem('theme', next);
    applyTheme(next);
    showToast(`Switched to ${next} mode`);
  }

  // ==========================================================================
  // 2. TOAST NOTIFICATIONS
  // ==========================================================================

  function showToast(message, duration = 2500) {
    const container = document.getElementById('toast-container');
    if (!container) return;

    const toast = document.createElement('div');
    toast.className = 'toast';
    toast.innerHTML = `
      <svg class="w-4 h-4 text-emerald-500" fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path>
      </svg>
      <span>${message}</span>
    `;

    container.appendChild(toast);

    setTimeout(() => {
      toast.style.opacity = '0';
      toast.style.transition = 'opacity 0.25s ease';
      setTimeout(() => toast.remove(), 250);
    }, duration);
  }

  // ==========================================================================
  // 3. COPY TO CLIPBOARD
  // ==========================================================================

  async function copyText(text, btnElement) {
    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(text);
      } else {
        const textarea = document.createElement('textarea');
        textarea.value = text;
        textarea.style.position = 'fixed';
        textarea.style.left = '-999999px';
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand('copy');
        textarea.remove();
      }

      if (btnElement) {
        btnElement.classList.add('copied');
        const originalContent = btnElement.innerHTML;
        btnElement.innerHTML = `
          <svg class="w-3.5 h-3.5 text-emerald-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M5 13l4 4L19 7"></path>
          </svg>
          <span>Copied!</span>
        `;
        setTimeout(() => {
          btnElement.classList.remove('copied');
          btnElement.innerHTML = originalContent;
        }, 2000);
      }

      showToast('Copied to clipboard!');
    } catch (err) {
      console.error('Failed to copy text: ', err);
      showToast('Copy failed — please copy manually', 3000);
    }
  }

  // ==========================================================================
  // 4. TAB CONTROLS
  // ==========================================================================

  function setupTabs() {
    const tabContainers = document.querySelectorAll('[data-tabs]');
    tabContainers.forEach((container) => {
      const buttons = container.querySelectorAll('[data-tab-target]');
      buttons.forEach((btn) => {
        btn.addEventListener('click', () => {
          const targetId = btn.getAttribute('data-tab-target');

          // Update active button
          buttons.forEach((b) => b.classList.remove('active'));
          btn.classList.add('active');

          // Update panels
          const group = btn.closest('[data-tabs-group]');
          if (group) {
            const panels = group.querySelectorAll('[data-tab-panel]');
            panels.forEach((p) => {
              if (p.id === targetId) {
                p.classList.remove('hidden');
              } else {
                p.classList.add('hidden');
              }
            });
          }
        });
      });
    });
  }

  // ==========================================================================
  // 5. SCROLLSPY & SIDEBAR NAVIGATION
  // ==========================================================================

  function setupScrollspy() {
    const navLinks = document.querySelectorAll('.nav-link[href^="#"]');
    if (navLinks.length === 0) return;

    const sections = Array.from(navLinks)
      .map((link) => {
        const id = link.getAttribute('href').slice(1);
        return document.getElementById(id);
      })
      .filter(Boolean);

    function onScroll() {
      const scrollPos = window.scrollY + 120;

      for (let i = sections.length - 1; i >= 0; i--) {
        const section = sections[i];
        if (section.offsetTop <= scrollPos) {
          const id = section.id;
          navLinks.forEach((link) => {
            if (link.getAttribute('href') === `#${id}`) {
              link.classList.add('active');
            } else {
              link.classList.remove('active');
            }
          });
          break;
        }
      }
    }

    window.addEventListener('scroll', onScroll, { passive: true });
    onScroll();
  }

  // ==========================================================================
  // 6. GSAP ANIMATIONS (RESTRAINED & ACCESSIBLE)
  // ==========================================================================

  function setupAnimations() {
    // Check if user requested reduced motion
    const prefersReducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
    if (prefersReducedMotion || typeof window.gsap === 'undefined') {
      // Show all elements immediately without transforms
      document.querySelectorAll('.animate-on-load, .reveal-card').forEach((el) => {
        el.style.opacity = '1';
        el.style.transform = 'none';
      });
      return;
    }

    const { gsap } = window;

    // Hero entrance timeline
    gsap.timeline({ defaults: { ease: 'power2.out', duration: 0.7 } })
      .from('.hero-badge', { opacity: 0, y: -15, delay: 0.1 })
      .from('.hero-title', { opacity: 0, y: 20 }, '-=0.5')
      .from('.hero-desc', { opacity: 0, y: 20 }, '-=0.5')
      .from('.hero-cta', { opacity: 0, y: 15 }, '-=0.4')
      .from('.hero-code', { opacity: 0, scale: 0.98, y: 25 }, '-=0.5');

    // Subtle scroll reveal for cards using IntersectionObserver
    const observer = new IntersectionObserver(
      (entries) => {
        entries.forEach((entry) => {
          if (entry.isIntersecting) {
            gsap.to(entry.target, {
              opacity: 1,
              y: 0,
              duration: 0.6,
              ease: 'power2.out',
              clearProps: 'transform'
            });
            observer.unobserve(entry.target);
          }
        });
      },
      { threshold: 0.1, rootMargin: '0px 0px -50px 0px' }
    );

    document.querySelectorAll('.reveal-card').forEach((card) => {
      card.style.opacity = '0';
      card.style.transform = 'translateY(20px)';
      observer.observe(card);
    });
  }

  // ==========================================================================
  // 7. FEATURE CATALOG FILTER & SEARCH
  // ==========================================================================

  function setupFeatureSearch() {
    const searchInput = document.getElementById('feature-search-input');
    const categorySelect = document.getElementById('feature-category-filter');
    const tableBody = document.getElementById('features-table-body');
    const countDisplay = document.getElementById('features-count-display');

    if (!searchInput || !tableBody) return;

    function filterFeatures() {
      const query = (searchInput.value || '').toLowerCase().trim();
      const selectedCategory = categorySelect ? categorySelect.value : 'all';

      const rows = tableBody.querySelectorAll('tr[data-feature]');
      let visibleCount = 0;

      rows.forEach((row) => {
        const featureName = (row.getAttribute('data-feature') || '').toLowerCase();
        const category = (row.getAttribute('data-category') || '').toLowerCase();
        const description = (row.getAttribute('data-desc') || '').toLowerCase();

        const matchesQuery = !query || featureName.includes(query) || description.includes(query);
        const matchesCategory = selectedCategory === 'all' || category === selectedCategory.toLowerCase();

        if (matchesQuery && matchesCategory) {
          row.classList.remove('hidden');
          visibleCount++;
        } else {
          row.classList.add('hidden');
        }
      });

      if (countDisplay) {
        countDisplay.textContent = `Showing ${visibleCount} of ${rows.length} features`;
      }
    }

    searchInput.addEventListener('input', filterFeatures);
    if (categorySelect) categorySelect.addEventListener('change', filterFeatures);
  }

  // ==========================================================================
  // 8. MOBILE NAVIGATION DRAWER
  // ==========================================================================

  function setupMobileMenu() {
    const toggleBtn = document.getElementById('mobile-menu-toggle');
    const closeBtn = document.getElementById('mobile-menu-close');
    const drawer = document.getElementById('mobile-drawer');
    const backdrop = document.getElementById('mobile-backdrop');

    if (!toggleBtn || !drawer || !backdrop) return;

    function openMenu() {
      drawer.classList.remove('-translate-x-full');
      backdrop.classList.remove('hidden');
      document.body.style.overflow = 'hidden';
    }

    function closeMenu() {
      drawer.classList.add('-translate-x-full');
      backdrop.classList.add('hidden');
      document.body.style.overflow = '';
    }

    toggleBtn.addEventListener('click', openMenu);
    if (closeBtn) closeBtn.addEventListener('click', closeMenu);
    backdrop.addEventListener('click', closeMenu);

    drawer.querySelectorAll('a').forEach((link) => {
      link.addEventListener('click', closeMenu);
    });
  }

  // ==========================================================================
  // INITIALIZATION ON DOM READY
  // ==========================================================================

  document.addEventListener('DOMContentLoaded', () => {
    // 1. Theme toggle hookup
    const themeToggleBtn = document.getElementById('theme-toggle-btn');
    if (themeToggleBtn) {
      themeToggleBtn.addEventListener('click', toggleTheme);
    }
    applyTheme(getPreferredTheme());

    // 2. Clipboard copy hookup
    document.querySelectorAll('[data-copy-target]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const targetId = btn.getAttribute('data-copy-target');
        const targetEl = document.getElementById(targetId);
        if (targetEl) {
          copyText(targetEl.textContent || targetEl.innerText, btn);
        }
      });
    });

    document.querySelectorAll('[data-copy-text]').forEach((btn) => {
      btn.addEventListener('click', () => {
        const text = btn.getAttribute('data-copy-text');
        copyText(text, btn);
      });
    });

    // 3. Initialize components
    setupTabs();
    setupScrollspy();
    setupAnimations();
    setupFeatureSearch();
    setupMobileMenu();

    // Listen for OS theme changes if user hasn't explicitly set one
    window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', (e) => {
      if (!localStorage.getItem('theme')) {
        applyTheme(e.matches ? 'dark' : 'light');
      }
    });
  });
})();
