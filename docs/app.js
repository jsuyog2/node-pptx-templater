// Documentation Engine Client Logic
document.addEventListener('DOMContentLoaded', () => {
  const navItems = document.querySelectorAll('.nav-item');
  const sections = document.querySelectorAll('.doc-section, .class-docs-section');
  
  // Show/Hide page sections based on URL hashes
  function showSection(hash) {
    const cleanId = hash.replace('#', '');
    
    sections.forEach(sec => sec.classList.add('hidden'));
    navItems.forEach(item => item.classList.remove('active'));
    
    const targetSection = document.getElementById(cleanId);
    if (targetSection) {
      targetSection.classList.remove('hidden');
      window.scrollTo({ top: 0, behavior: 'smooth' });
      
      // GSAP Liquid Glass Reveal Animation
      gsap.fromTo(targetSection, 
        { opacity: 0, y: 15 },
        { opacity: 1, y: 0, duration: 0.4, ease: "power2.out" }
      );
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
  } else {
    showSection('#introduction');
  }

  // Liquid Glass Glow Cursor Tracker
  document.addEventListener('mousemove', e => {
    document.querySelectorAll('.glass-card, .method-card').forEach(card => {
      const rect = card.getBoundingClientRect();
      const x = e.clientX - rect.left;
      const y = e.clientY - rect.top;
      card.style.setProperty('--mouse-x', `${x}px`);
      card.style.setProperty('--mouse-y', `${y}px`);
    });
  });

  // GSAP Entrance Load Animations
  gsap.from("header", { opacity: 0, y: -20, duration: 0.6, ease: "power3.out" });
  gsap.from("aside", { opacity: 0, x: -30, duration: 0.6, ease: "power3.out", delay: 0.1 });
  gsap.from(".glow-orb", { opacity: 0, scale: 0.8, duration: 1.5, ease: "power2.out" });

  // Light/Dark Theme Switcher
  const themeToggle = document.getElementById('theme-toggle');
  const body = document.body;
  const savedTheme = localStorage.getItem('theme');

  if (savedTheme === 'light') {
    body.classList.remove('dark-theme');
    body.classList.add('light-theme', 'bg-white', 'text-slate-800');
    body.classList.remove('bg-[#080b11]', 'text-gray-200');
    themeToggle.querySelector('.toggle-icon').textContent = '☀️';
  }

  themeToggle.addEventListener('click', () => {
    if (body.classList.contains('light-theme')) {
      body.classList.remove('light-theme', 'bg-white', 'text-slate-800');
      body.classList.add('bg-[#080b11]', 'text-gray-200');
      themeToggle.querySelector('.toggle-icon').textContent = '🌙';
      localStorage.setItem('theme', 'dark');
    } else {
      body.classList.add('light-theme', 'bg-white', 'text-slate-800');
      body.classList.remove('bg-[#080b11]', 'text-gray-200');
      themeToggle.querySelector('.toggle-icon').textContent = '☀️';
      localStorage.setItem('theme', 'light');
    }
  });

  // Client Sidebar Navigation Filter Search
  const searchInput = document.getElementById('doc-search');
  searchInput.addEventListener('input', (e) => {
    const query = e.target.value.toLowerCase().trim();
    if (!query) {
      document.querySelectorAll('.nav-item').forEach(item => item.parentElement.style.display = 'block');
      document.querySelectorAll('.class-sidebar-group').forEach(group => group.style.display = 'block');
      return;
    }

    document.querySelectorAll('.nav-item').forEach(item => {
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

  // Clipboard Text Copier
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

// Interactive Sandbox tab switcher
window.switchSandbox = function(tabName) {
  const codeBlock = document.getElementById('sandbox-code-block');
  const previewBox = document.getElementById('sandbox-preview-box');
  
  document.querySelectorAll('.sb-tab').forEach(el => {
    el.classList.remove('bg-brand-500/10', 'text-brand-400', 'border', 'border-brand-500/20');
    el.classList.add('text-gray-400', 'hover:text-white');
  });
  
  const selectedTab = document.getElementById('sb-tab-' + tabName);
  selectedTab.classList.add('bg-brand-500/10', 'text-brand-400', 'border', 'border-brand-500/20');
  selectedTab.classList.remove('text-gray-400', 'hover:text-white');
  
  if (tabName === 'text') {
    codeBlock.textContent = `ppt.useSlide(1)\n   .replaceTextByTag('title', 'Q2 Report')\n   .replaceMultiple({\n     user: 'Acme Corp',\n     date: 'June 2026'\n   });`;
    previewBox.innerHTML = `<div class="border border-white/10 rounded bg-[#0e1526] p-3 w-full text-center space-y-2">\n      <div class="font-bold text-white text-[12px] border-b border-white/5 pb-1">Q2 Report</div>\n      <div class="text-[9px] text-gray-400">Owner: Acme Corp</div>\n      <div class="text-[9px] text-brand-400">Date: June 2026</div>\n    </div>`;
  } else if (tabName === 'charts') {
    codeBlock.textContent = `ppt.useSlide(2)\n   .updateChartData('sales-chart', {\n     categories: ['Q1', 'Q2', 'Q3'],\n     series: [\n       { name: 'Target', values: [80, 100, 120] },\n       { name: 'Actual', values: [95, 115, 130] }\n     ]\n   });`;
    previewBox.innerHTML = `<div class="w-full flex items-end justify-around h-32 px-4 border-b border-white/10">\n      <div class="flex flex-col items-center">\n        <div class="w-4 bg-brand-500/20 h-16 rounded-t"></div>\n        <div class="w-4 bg-brand-500 h-20 rounded-t -mt-2"></div>\n        <span class="text-[8px] text-gray-500 mt-1">Q1</span>\n      </div>\n      <div class="flex flex-col items-center">\n        <div class="w-4 bg-brand-500/20 h-20 rounded-t"></div>\n        <div class="w-4 bg-brand-500 h-24 rounded-t -mt-2"></div>\n        <span class="text-[8px] text-gray-500 mt-1">Q2</span>\n      </div>\n      <div class="flex flex-col items-center">\n        <div class="w-4 bg-brand-500/20 h-24 rounded-t"></div>\n        <div class="w-4 bg-brand-500 h-28 rounded-t -mt-2"></div>\n        <span class="text-[8px] text-gray-500 mt-1">Q3</span>\n      </div>\n    </div>\n    <div class="text-[8px] text-gray-400 mt-2 flex gap-3"><span class="flex items-center gap-1"><span class="w-2 h-2 bg-brand-500/20 rounded"></span>Target</span><span class="flex items-center gap-1"><span class="w-2 h-2 bg-brand-500 rounded"></span>Actual</span></div>`;
  } else if (tabName === 'tables') {
    codeBlock.textContent = `ppt.useSlide(3)\n   .updateTable('sales-table', [\n     ['Category', { value: 'Global Performance', colSpan: 2 }],\n     ['North Region', '120k', 'Growth: 8%'],\n     ['South Region', '150k', 'Growth: 12%']\n   ])\n   .mergeCells('sales-table', 1, 1, 2, 2);`;
    previewBox.innerHTML = `<div class="w-full border border-white/10 rounded-xl overflow-hidden text-[9px] bg-[#0e1526]">\n      <div class="bg-white/5 p-2 font-bold border-b border-white/10 text-center">Global Performance</div>\n      <div class="grid grid-cols-3 divide-x divide-white/10 border-b border-white/10 text-center">\n        <div class="p-2 text-gray-400">North Region</div>\n        <div class="p-2 col-span-2 text-brand-400 font-bold">120k / Growth: 8%</div>\n      </div>\n      <div class="grid grid-cols-3 divide-x divide-white/10 text-center">\n        <div class="p-2 text-gray-400">South Region</div>\n        <div class="p-2 col-span-2 text-brand-400 font-bold">150k / Growth: 12%</div>\n      </div>\n    </div>`;
  } else if (tabName === 'layers') {
    codeBlock.textContent = `ppt.useSlide(4)\n   .bringToFront('OverlayLogo')\n   .sendToBack('BackgroundShade');`;
    previewBox.innerHTML = `<div class="relative w-full h-32 border border-white/10 rounded-xl bg-[#0e1526] overflow-hidden">\n      <div class="absolute inset-0 bg-brand-500/5 flex items-center justify-center text-[10px] text-gray-500">BackgroundShade (zIndex: 1)</div>\n      <div class="absolute top-6 left-6 w-32 h-16 border border-white/10 bg-slate-900 flex items-center justify-center rounded shadow-lg text-[9px]">Text Container (zIndex: 2)</div>\n      <div class="absolute top-10 right-6 w-20 h-16 border border-brand-500/30 bg-brand-500/10 flex items-center justify-center rounded shadow-2xl text-[9px] text-brand-400 font-bold">OverlayLogo (zIndex: 3)</div>\n    </div>`;
  }

  // Sandbox reveal sweep GSAP
  gsap.fromTo([codeBlock, previewBox], 
    { opacity: 0.7, scale: 0.98 },
    { opacity: 1, scale: 1, duration: 0.3, ease: "power2.out" }
  );
}

// Sidebar sub-navigation toggle
window.toggleSidebarGroup = function(groupId) {
  const list = document.getElementById('sidebar-list-' + groupId);
  const arrow = document.getElementById('arrow-' + groupId);
  if (list && list.classList.contains('hidden')) {
    list.classList.remove('hidden');
    arrow.style.transform = 'rotate(0deg)';
  } else if (list) {
    list.classList.add('hidden');
    arrow.style.transform = 'rotate(-90deg)';
  }
}
