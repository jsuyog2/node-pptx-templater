'use client';

import React, { useMemo, useCallback } from 'react';
import { useRouter } from 'next/navigation';
import PowerPointSlideDeck from './PowerPointSlideDeck';
import ApiReferenceCard from './ApiReferenceCard';
import apiReferenceData from '../data/apiReference.json';
import {
  CATEGORIES,
  CATEGORY_DESCRIPTIONS,
  CATEGORY_SLUG_MAP,
  SLUG_TO_CATEGORY_MAP,
} from '../utils/categoryMapping';

export default function CategoryDeckClient({ activeCategorySlug = 'import-export' }) {
  const router = useRouter();

  // Normalize slug and find canonical category name
  const canonicalCategory = useMemo(() => {
    return SLUG_TO_CATEGORY_MAP[activeCategorySlug] || 'Import/Export';
  }, [activeCategorySlug]);

  const activeIndex = useMemo(() => {
    const idx = CATEGORIES.indexOf(canonicalCategory);
    return idx >= 0 ? idx : 0;
  }, [canonicalCategory]);

  const handleSlideChange = useCallback(
    (newIndex) => {
      const targetCategory = CATEGORIES[newIndex] || CATEGORIES[0];
      const targetSlug = CATEGORY_SLUG_MAP[targetCategory] || 'import-export';
      if (targetSlug !== activeCategorySlug) {
        router.push(`/docs/api-reference/${targetSlug}`);
      }
    },
    [activeCategorySlug, router]
  );

  // Construct 8 slides, one for each standardized category
  const slides = useMemo(() => {
    return CATEGORIES.map((cat, idx) => {
      const slug = CATEGORY_SLUG_MAP[cat];
      const methods = apiReferenceData.methodsByCategory?.[cat] || [];
      const description =
        CATEGORY_DESCRIPTIONS[cat] ||
        apiReferenceData.categoryDescriptions?.[cat] ||
        '';

      return {
        id: `slide-cat-${slug}`,
        category: 'API Reference',
        title: cat,
        subtitle: description,
        categoryDescription: description,
        signature: `${cat} Module`,
        hideDefaultHeader: true,
        content: (
          <ApiReferenceCard
            categoryName={cat}
            categorySlug={slug}
            categoryDescription={description}
            methods={methods}
            slideNumber={idx + 1}
            totalSlides={CATEGORIES.length}
          />
        ),
      };
    });
  }, []);

  return (
    <PowerPointSlideDeck
      slides={slides}
      initialIndex={activeIndex}
      categoryTitle="API Reference"
      onSlideChange={handleSlideChange}
    />
  );
}
