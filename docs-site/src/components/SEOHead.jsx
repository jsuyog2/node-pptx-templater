'use client';

import React from 'react';
import siteMeta from '../data/siteMeta.json';

export default function SEOHead({
  title = siteMeta.title,
  description = siteMeta.description,
  canonical = 'https://node-pptx-templater.pages.dev/',
  keywords = siteMeta.keywords,
  ogType = 'website',
  structuredData = null,
}) {
  const jsonLdData = structuredData || siteMeta.jsonLd;

  return (
    <>
      <link rel="canonical" href={canonical} />
      <meta name="description" content={description} />
      <meta name="keywords" content={Array.isArray(keywords) ? keywords.join(', ') : keywords} />

      {/* OpenGraph */}
      <meta property="og:title" content={title} />
      <meta property="og:description" content={description} />
      <meta property="og:type" content={ogType} />
      <meta property="og:url" content={canonical} />
      <meta property="og:image" content="https://node-pptx-templater.pages.dev/og-banner.png" />
      <meta property="og:image:width" content="1200" />
      <meta property="og:image:height" content="630" />
      <meta property="og:site_name" content="node-pptx-templater" />

      {/* Twitter */}
      <meta name="twitter:card" content="summary_large_image" />
      <meta name="twitter:title" content={title} />
      <meta name="twitter:description" content={description} />
      <meta name="twitter:image" content="https://node-pptx-templater.pages.dev/twitter-card.png" />
      <meta name="twitter:creator" content="@jsuyog2" />

      {/* Schema.org Structured Data */}
      {jsonLdData && (
        <script
          type="application/ld+json"
          dangerouslySetInnerHTML={{ __html: JSON.stringify(jsonLdData) }}
        />
      )}
    </>
  );
}
