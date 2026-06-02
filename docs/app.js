document.addEventListener('DOMContentLoaded', () => {
  // Navigation Router
  const navItems = document.querySelectorAll('.nav-item');
  const sections = document.querySelectorAll('.doc-section');

  function showSection(targetId) {
    let cleanId = targetId.replace('#', '');
    sections.forEach(s => s.classList.remove('active-section'));
    navItems.forEach(n => n.classList.remove('active'));

    const targetSection = document.getElementById(cleanId);
    if (targetSection) {
      targetSection.classList.add('active-section');
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }

    const activeLink = document.querySelector(`a[href="#${cleanId}"]`);
    if (activeLink) {
      activeLink.classList.add('active');
    }
  }

  navItems.forEach(item => {
    item.addEventListener('click', (e) => {
      e.preventDefault();
      const href = item.getAttribute('href');
      window.history.pushState(null, null, href);
      showSection(href);
    });
  });

  window.addEventListener('popstate', () => {
    const hash = window.location.hash || '#introduction';
    showSection(hash);
  });

  if (window.location.hash) {
    showSection(window.location.hash);
  }

  // Theme Toggler
  const themeToggle = document.getElementById('theme-toggle');
  const body = document.body;
  const savedTheme = localStorage.getItem('theme');

  if (savedTheme === 'light') {
    body.classList.remove('dark-theme');
    body.classList.add('light-theme', 'bg-white', 'text-slate-800');
    body.classList.remove('bg-[#0b0f19]', 'text-gray-200');
    themeToggle.querySelector('.toggle-icon').textContent = '☀️';
  }

  themeToggle.addEventListener('click', () => {
    if (body.classList.contains('light-theme')) {
      body.classList.remove('light-theme', 'bg-white', 'text-slate-800');
      body.classList.add('bg-[#0b0f19]', 'text-gray-200');
      themeToggle.querySelector('.toggle-icon').textContent = '🌙';
      localStorage.setItem('theme', 'dark');
    } else {
      body.classList.add('light-theme', 'bg-white', 'text-slate-800');
      body.classList.remove('bg-[#0b0f19]', 'text-gray-200');
      themeToggle.querySelector('.toggle-icon').textContent = '☀️';
      localStorage.setItem('theme', 'light');
    }
  });

  // Client Search Indexer
  const searchInput = document.getElementById('doc-search');
  searchInput.addEventListener('input', (e) => {
    const query = e.target.value.toLowerCase().trim();
    if (!query) {
      navItems.forEach(item => item.parentElement.style.display = 'block');
      return;
    }

    navItems.forEach(item => {
      const text = item.textContent.toLowerCase();
      const parent = item.parentElement;
      if (text.includes(query)) {
        parent.style.display = 'block';
      } else {
        parent.style.display = 'none';
      }
    });
  });

  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
      e.preventDefault();
      searchInput.focus();
    }
  });

  // Clipboard Copier
  const copyButtons = document.querySelectorAll('.copy-btn');
  copyButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const codeBlock = btn.previousElementSibling;
      if (codeBlock) {
        navigator.clipboard.writeText(codeBlock.textContent).then(() => {
          btn.textContent = 'Copied!';
          setTimeout(() => {
            btn.textContent = 'Copy';
          }, 2000);
        });
      }
    });
  });
});

// Learning Path selector
window.switchPath = function(pathName) {
  document.querySelectorAll('.path-content').forEach(el => el.classList.add('hidden'));
  document.querySelectorAll('.path-tab').forEach(el => {
    el.classList.remove('border-brand-500', 'text-white');
    el.classList.add('border-transparent', 'text-gray-400');
  });

  const content = document.getElementById('path-content-' + pathName);
  if (content) content.classList.remove('hidden');

  const tab = document.getElementById('tab-' + pathName);
  if (tab) {
    tab.classList.add('border-brand-500', 'text-white');
    tab.classList.remove('border-transparent', 'text-gray-400');
  }
}
