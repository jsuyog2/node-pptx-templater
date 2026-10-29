'use client';

import React, { Suspense } from 'react';
import { useSearchParams } from 'next/navigation';
import CategoryDeckClient from '../../../components/CategoryDeckClient';
import {
  CATEGORY_SLUG_MAP,
  SLUG_TO_CATEGORY_MAP,
} from '../../../utils/categoryMapping';

function ApiReferenceDefault() {
  const searchParams = useSearchParams();
  const categoryParam = searchParams.get('category');

  let activeSlug = 'import-export';
  if (categoryParam) {
    if (CATEGORY_SLUG_MAP[categoryParam]) {
      activeSlug = CATEGORY_SLUG_MAP[categoryParam];
    } else if (SLUG_TO_CATEGORY_MAP[categoryParam.toLowerCase()]) {
      activeSlug = categoryParam.toLowerCase();
    }
  }

  return <CategoryDeckClient activeCategorySlug={activeSlug} />;
}

export default function ApiReferencePage() {
  return (
    <Suspense
      fallback={
        <div className="h-[calc(100vh-104px)] flex items-center justify-center text-xs text-[var(--ppt-text-secondary)]">
          Loading PowerPoint API Catalog...
        </div>
      }
    >
      <main className="w-full h-full overflow-hidden">
        <ApiReferenceDefault />
      </main>
    </Suspense>
  );
}
