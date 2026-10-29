'use client';

import React, { useRef, useEffect } from 'react';

/**
 * Mobile Horizontal Function Reel
 * Renders a touch-friendly horizontal ribbon of method chips on viewports < 768px
 *
 * @param {{
 *   methods: Array<{ name: string, isAsync?: boolean }>,
 *   activeMethodName: string,
 *   onSelectMethod: (name: string) => void
 * }} props
 */
export default function MobileFunctionReel({
  methods = [],
  activeMethodName = '',
  onSelectMethod = () => { },
}) {
  const containerRef = useRef(null);
  const activeBtnRef = useRef(null);

  // Auto-scroll active chip into view on selection change
  useEffect(() => {
    if (activeBtnRef.current) {
      activeBtnRef.current.scrollIntoView({
        behavior: 'smooth',
        block: 'nearest',
        inline: 'center',
      });
    }
  }, [activeMethodName]);

  if (!methods || methods.length === 0) return null;

  return (
    <div
      ref={containerRef}
      className="md:hidden w-full max-w-full overflow-x-auto overscroll-x-contain whitespace-nowrap scrollbar-none flex items-center gap-2 py-2 px-3 border-b border-[var(--ppt-card-border)]/50 bg-black/5 dark:bg-black/20"
      aria-label="API Method Reel"
    >
      {methods.map((method) => {
        const isActive = method.name === activeMethodName;
        return (
          <button
            key={method.name}
            ref={isActive ? activeBtnRef : null}
            type="button"
            onClick={() => onSelectMethod(method.name)}
            className={`flex-shrink-0 text-left px-3 py-1.5 rounded text-xs transition-all flex items-center justify-between border cursor-pointer ${isActive
              ? 'border-[#D24726] bg-[#D24726]/10 text-[var(--ppt-text-primary)] font-semibold shadow-sm'
              : 'border-transparent hover:border-[var(--ppt-card-border)] text-[var(--ppt-text-secondary)] hover:text-[var(--ppt-text-primary)] hover:bg-black/5 dark:hover:bg-white/5'
              }`}
          >
            <span>{method.name}()</span>
          </button>
        );
      })}
    </div>
  );
}
