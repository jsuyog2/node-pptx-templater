import React, { useState, useEffect } from 'react';
import { Menu, Sun, Moon, Github, Box, ExternalLink } from 'lucide-react';

export default function Header({
  activeRoute,
  onNavigate,
  onToggleSidebar,
  darkMode,
  onToggleDarkMode
}) {
  return (
    <header className="sticky top-0 z-40 glass-nav h-16 transition-colors duration-200">
      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 h-full flex items-center justify-between">
        {/* Left: Hamburger & Brand */}
        <div className="flex items-center gap-3">
          <button
            onClick={onToggleSidebar}
            className="md:hidden p-2 rounded-xl text-slate-600 dark:text-slate-300 hover:bg-slate-100 dark:hover:bg-slate-800 transition-colors"
            aria-label="Toggle navigation drawer"
          >
            <Menu className="w-5 h-5" />
          </button>

          <button
            onClick={() => onNavigate('/')}
            className="flex items-center gap-2.5 group text-left"
          >
            <div className="w-8 h-8 rounded-xl bg-gradient-to-tr from-blue-600 to-indigo-600 flex items-center justify-center text-white font-black text-sm shadow-md group-hover:scale-105 transition-transform">
              P
            </div>
            <div>
              <span className="font-bold text-base sm:text-lg tracking-tight bg-gradient-to-r from-blue-600 to-indigo-600 dark:from-blue-400 dark:to-indigo-300 bg-clip-text text-transparent">
                node-pptx-templater
              </span>
            </div>
          </button>
          <span className="hidden sm:inline-flex px-2 py-0.5 rounded-full text-[11px] font-semibold bg-blue-500/10 text-blue-600 dark:text-blue-400 border border-blue-500/20">
            v3.0.0-alpha.0
          </span>
        </div>

        {/* Center: Desktop Navigation */}
        <nav className="hidden lg:flex items-center gap-1">
          <button
            onClick={() => onNavigate('/')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
              activeRoute === '/'
                ? 'bg-blue-500/10 text-blue-600 dark:text-blue-400 font-semibold'
                : 'text-slate-600 dark:text-slate-300 hover:text-slate-900 dark:hover:text-white hover:bg-slate-100 dark:hover:bg-slate-800/60'
            }`}
          >
            Overview
          </button>
          <button
            onClick={() => onNavigate('/guides')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
              activeRoute === '/guides'
                ? 'bg-blue-500/10 text-blue-600 dark:text-blue-400 font-semibold'
                : 'text-slate-600 dark:text-slate-300 hover:text-slate-900 dark:hover:text-white hover:bg-slate-100 dark:hover:bg-slate-800/60'
            }`}
          >
            Guides
          </button>
          <button
            onClick={() => onNavigate('/api-reference')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
              activeRoute === '/api-reference'
                ? 'bg-blue-500/10 text-blue-600 dark:text-blue-400 font-semibold'
                : 'text-slate-600 dark:text-slate-300 hover:text-slate-900 dark:hover:text-white hover:bg-slate-100 dark:hover:bg-slate-800/60'
            }`}
          >
            API Reference
          </button>
          <button
            onClick={() => onNavigate('/showcase')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
              activeRoute === '/showcase'
                ? 'bg-blue-500/10 text-blue-600 dark:text-blue-400 font-semibold'
                : 'text-slate-600 dark:text-slate-300 hover:text-slate-900 dark:hover:text-white hover:bg-slate-100 dark:hover:bg-slate-800/60'
            }`}
          >
            Showcase
          </button>
          <button
            onClick={() => onNavigate('/features')}
            className={`px-3 py-1.5 rounded-lg text-xs font-medium transition-colors ${
              activeRoute === '/features'
                ? 'bg-blue-500/10 text-blue-600 dark:text-blue-400 font-semibold'
                : 'text-slate-600 dark:text-slate-300 hover:text-slate-900 dark:hover:text-white hover:bg-slate-100 dark:hover:bg-slate-800/60'
            }`}
          >
            Features
          </button>
        </nav>

        {/* Right: Theme Toggle & Social Links */}
        <div className="flex items-center gap-1.5 sm:gap-2">
          {/* Dark / Light Toggle */}
          <button
            onClick={onToggleDarkMode}
            className="p-2 rounded-xl text-slate-500 hover:text-slate-900 dark:text-slate-400 dark:hover:text-slate-100 hover:bg-slate-100 dark:hover:bg-slate-800/70 transition-colors"
            aria-label="Toggle dark/light mode"
          >
            {darkMode ? <Sun className="w-4 h-4 text-amber-400" /> : <Moon className="w-4 h-4 text-slate-600" />}
          </button>

          {/* GitHub Link */}
          <a
            href="https://github.com/jsuyog2/node-pptx-templater"
            target="_blank"
            rel="noopener noreferrer"
            className="p-2 rounded-xl text-slate-500 hover:text-slate-900 dark:text-slate-400 dark:hover:text-slate-100 hover:bg-slate-100 dark:hover:bg-slate-800/70 transition-colors"
            title="GitHub Repository"
          >
            <Github className="w-4 h-4" />
          </a>

          {/* NPM Link */}
          <a
            href="https://www.npmjs.com/package/node-pptx-templater"
            target="_blank"
            rel="noopener noreferrer"
            className="p-2 rounded-xl text-slate-500 hover:text-slate-900 dark:text-slate-400 dark:hover:text-slate-100 hover:bg-slate-100 dark:hover:bg-slate-800/70 transition-colors"
            title="NPM Package"
          >
            <Box className="w-4 h-4 text-rose-500" />
          </a>
        </div>
      </div>
    </header>
  );
}
