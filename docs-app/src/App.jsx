import React, { useState, useEffect } from 'react';
import Header from './components/Header';
import Sidebar from './components/Sidebar';
import Footer from './components/Footer';
import Home from './pages/Home';
import ApiReferencePage from './pages/ApiReferencePage';
import GuidesPage from './pages/GuidesPage';
import ShowcasePage from './pages/ShowcasePage';
import FeaturesPage from './pages/FeaturesPage';
import apiData from './data/apiReference.json';

export default function App() {
  const [route, setRoute] = useState(window.location.hash.replace('#', '') || '/');
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [selectedCategory, setSelectedCategory] = useState('All');
  const [guideSlug, setGuideSlug] = useState('getting-started');

  // Dark mode state
  const [darkMode, setDarkMode] = useState(() => {
    const saved = localStorage.getItem('theme');
    if (saved) return saved === 'dark';
    return window.matchMedia('(prefers-color-scheme: dark)').matches;
  });

  useEffect(() => {
    if (darkMode) {
      document.documentElement.classList.add('dark');
      localStorage.setItem('theme', 'dark');
    } else {
      document.documentElement.classList.remove('dark');
      localStorage.setItem('theme', 'light');
    }
  }, [darkMode]);

  // Handle hash changes for client routing
  useEffect(() => {
    const handleHashChange = () => {
      const hash = window.location.hash.replace('#', '') || '/';
      if (hash.startsWith('/guides/')) {
        const slug = hash.replace('/guides/', '');
        setRoute('/guides');
        setGuideSlug(slug);
      } else if (hash.startsWith('/api-reference')) {
        setRoute('/api-reference');
        const urlParams = new URLSearchParams(window.location.hash.split('?')[1] || '');
        const cat = urlParams.get('category');
        if (cat) setSelectedCategory(cat);
      } else if (hash.startsWith('/showcase')) {
        setRoute('/showcase');
      } else if (hash.startsWith('/features')) {
        setRoute('/features');
      } else if (hash.startsWith('/guides')) {
        setRoute('/guides');
      } else {
        setRoute('/');
      }
    };

    window.addEventListener('hashchange', handleHashChange);
    handleHashChange();

    return () => window.removeEventListener('hashchange', handleHashChange);
  }, []);

  const navigate = (path) => {
    window.location.hash = path;
    window.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const renderCurrentPage = () => {
    switch (route) {
      case '/api-reference':
        return <ApiReferencePage initialCategory={selectedCategory} />;
      case '/guides':
        return <GuidesPage initialGuideSlug={guideSlug} />;
      case '/showcase':
        return <ShowcasePage />;
      case '/features':
        return <FeaturesPage />;
      case '/':
      default:
        return <Home onNavigate={navigate} />;
    }
  };

  return (
    <div className="min-h-screen flex flex-col bg-slate-50 dark:bg-[#0b0f19] text-slate-900 dark:text-slate-100 transition-colors duration-200">
      {/* Ambient Glows */}
      <div className="hero-glow-1"></div>
      <div className="hero-glow-2"></div>

      {/* Header */}
      <Header
        activeRoute={route}
        onNavigate={navigate}
        onToggleSidebar={() => setSidebarOpen(!sidebarOpen)}
        darkMode={darkMode}
        onToggleDarkMode={() => setDarkMode(!darkMode)}
      />

      {/* Main Layout Container */}
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 w-full flex-1 flex flex-col md:flex-row relative z-10 pt-4">
        {/* Responsive Sidebar */}
        <Sidebar
          categories={apiData.categories}
          methods={apiData.methods}
          selectedCategory={selectedCategory}
          onSelectCategory={(cat) => {
            setSelectedCategory(cat);
            navigate('/api-reference');
          }}
          activeRoute={route}
          onNavigate={navigate}
          isOpen={sidebarOpen}
          onClose={() => setSidebarOpen(false)}
        />

        {/* Content Area */}
        <main className="flex-1 min-w-0 md:pl-8 lg:pl-10 pb-16">
          {renderCurrentPage()}
        </main>
      </div>

      {/* Footer */}
      <Footer />
    </div>
  );
}
