import React from 'react';
import DocsSidebar from '../../components/DocsSidebar';

export default function DocsLayout({ children }) {
  return (
    <div className="min-h-[calc(100vh-4rem)]">
      <DocsSidebar />
      <div className="lg:pl-72 flex flex-col min-h-full">
        <div className="flex-1 max-w-5xl w-full mx-auto px-4 sm:px-6 lg:px-10 py-8 sm:py-12">
          {children}
        </div>
      </div>
    </div>
  );
}
