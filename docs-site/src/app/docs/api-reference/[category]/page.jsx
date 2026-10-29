import React from 'react';
import CategoryDeckClient from '../../../../components/CategoryDeckClient';
import { SLUG_TO_CATEGORY_MAP } from '../../../../utils/categoryMapping';

export function generateStaticParams() {
  return [
    { category: 'import-export' },
    { category: 'slides' },
    { category: 'text-placeholders' },
    { category: 'tables' },
    { category: 'charts' },
    { category: 'shapes' },
    { category: 'images-media' },
    { category: 'elements' },
  ];
}

export function generateMetadata({ params }) {
  const categorySlug = params?.category || 'import-export';
  const categoryName = SLUG_TO_CATEGORY_MAP[categorySlug] || 'Import/Export';

  return {
    title: `${categoryName} API Reference | node-pptx-templater`,
    description: `Complete API documentation, signatures, parameters, and examples for ${categoryName} in node-pptx-templater.`,
    alternates: {
      canonical: `https://node-pptx-templater.pages.dev/docs/api-reference/${categorySlug}`,
    },
    openGraph: {
      title: `${categoryName} API Reference — node-pptx-templater`,
      description: `Complete API documentation, signatures, parameters, and examples for ${categoryName}.`,
      url: `https://node-pptx-templater.pages.dev/docs/api-reference/${categorySlug}`,
    },
  };
}

export default function CategoryPage({ params }) {
  const categorySlug = params?.category || 'import-export';

  return (
    <main className="w-full h-full overflow-hidden">
      <CategoryDeckClient activeCategorySlug={categorySlug} />
    </main>
  );
}
