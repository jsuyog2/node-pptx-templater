document.addEventListener('DOMContentLoaded', () => {
  // Navigation Routing System
  const navItems = document.querySelectorAll('.nav-item');
  const sections = document.querySelectorAll('.doc-section');

  function showSection(targetId) {
    let cleanId = targetId.replace('#', '');
    
    // Deactivate all
    sections.forEach(s => s.classList.remove('active-section'));
    navItems.forEach(n => n.classList.remove('active'));

    // Activate selected
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

  // Bind sidebar nav links
  navItems.forEach(item => {
    item.addEventListener('click', (e) => {
      e.preventDefault();
      const href = item.getAttribute('href');
      window.history.pushState(null, null, href);
      showSection(href);
    });
  });

  // Handle hash changes on back/forward buttons
  window.addEventListener('popstate', () => {
    const hash = window.location.hash || '#introduction';
    showSection(hash);
  });

  // Load initial page hash if exists
  if (window.location.hash) {
    showSection(window.location.hash);
  }

  // Theme Toggle Mechanism
  const themeToggle = document.getElementById('theme-toggle');
  const body = document.body;
  
  // Set default theme from localStorage
  const savedTheme = localStorage.getItem('theme');
  if (savedTheme === 'light') {
    body.classList.remove('dark-theme');
    body.classList.add('light-theme');
    themeToggle.querySelector('.toggle-icon').textContent = '☀️';
  }

  themeToggle.addEventListener('click', () => {
    if (body.classList.contains('dark-theme')) {
      body.classList.remove('dark-theme');
      body.classList.add('light-theme');
      themeToggle.querySelector('.toggle-icon').textContent = '☀️';
      localStorage.setItem('theme', 'light');
    } else {
      body.classList.remove('light-theme');
      body.classList.add('dark-theme');
      themeToggle.querySelector('.toggle-icon').textContent = '🌙';
      localStorage.setItem('theme', 'dark');
    }
  });

  // Search Filter Mechanism
  const searchInput = document.getElementById('doc-search');
  searchInput.addEventListener('input', (e) => {
    const query = e.target.value.toLowerCase().trim();
    if (!query) {
      // Restore default sidebar list visibility
      document.querySelectorAll('.nav-section li').forEach(li => li.style.display = 'block');
      return;
    }

    // Filter items based on navigation tags
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

  // Global Ctrl + K search hotkey
  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key === 'k') {
      e.preventDefault();
      searchInput.focus();
    }
  });
});
