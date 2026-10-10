/**
 * node-pptx-templater Documentation Web Application
 * Zero-dependency interactive controller
 */

document.addEventListener('DOMContentLoaded', () => {
  initSearch();
  initActiveNav();
});

// Installation tab switcher
function switchInstall(manager) {
  const codeEl = document.getElementById('install-cmd');
  const tabs = document.querySelectorAll('.install-tab');
  tabs.forEach(tab => tab.classList.remove('active'));

  if (manager === 'pnpm') {
    codeEl.textContent = 'pnpm add node-pptx-templater';
  } else if (manager === 'yarn') {
    codeEl.textContent = 'yarn add node-pptx-templater';
  } else {
    codeEl.textContent = 'npm install node-pptx-templater';
  }

  const activeTab = Array.from(tabs).find(t => t.textContent.toLowerCase() === manager);
  if (activeTab) activeTab.classList.add('active');
}

// Code snippet TS / JS tab switcher
function switchSnippet(group, lang) {
  const container = document.querySelector(`.code-tabs-wrapper`);
  if (!container) return;

  const buttons = container.querySelectorAll('.code-tab-btn');
  buttons.forEach(btn => btn.classList.remove('active'));

  const panes = container.querySelectorAll('.snippet-pane');
  panes.forEach(pane => pane.classList.remove('active'));

  const targetPane = document.getElementById(`${group}-${lang}`);
  if (targetPane) targetPane.classList.add('active');

  const activeBtn = Array.from(buttons).find(b =>
    (lang === 'ts' && b.textContent.includes('TypeScript')) ||
    (lang === 'js' && b.textContent.includes('JavaScript'))
  );
  if (activeBtn) activeBtn.classList.add('active');
}

// Real-time live search & filter
function initSearch() {
  const searchInput = document.getElementById('api-search');
  if (!searchInput) return;

  const apiCards = document.querySelectorAll('.api-card, .guide-item, .card');

  searchInput.addEventListener('input', (e) => {
    const query = e.target.value.trim().toLowerCase();

    if (!query) {
      apiCards.forEach(card => (card.style.display = ''));
      return;
    }

    apiCards.forEach(card => {
      const text = card.textContent.toLowerCase();
      if (text.includes(query)) {
        card.style.display = '';
        card.style.borderColor = '#38bdf8';
      } else {
        card.style.display = 'none';
        card.style.borderColor = '';
      }
    });
  });
}

// Active navigation highlight on scroll
function initActiveNav() {
  const navLinks = document.querySelectorAll('.sidebar-nav a');
  const sections = document.querySelectorAll('.doc-section');

  window.addEventListener('scroll', () => {
    let currentId = '';
    const scrollPos = window.scrollY + 100;

    sections.forEach(sec => {
      if (sec.offsetTop <= scrollPos) {
        currentId = sec.getAttribute('id');
      }
    });

    if (currentId) {
      navLinks.forEach(link => {
        link.classList.remove('active');
        if (link.getAttribute('href') === `#${currentId}`) {
          link.classList.add('active');
        }
      });
    }
  });
}
