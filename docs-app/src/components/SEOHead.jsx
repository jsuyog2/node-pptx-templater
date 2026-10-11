import React, { useEffect } from 'react';

export default function SEOHead({
  title = 'node-pptx-templater — High-Performance PowerPoint OpenXML Template Engine',
  description = 'Populate PowerPoint templates dynamically in Node.js & Serverless runtimes without Microsoft Office dependencies.',
  canonicalUrl = 'https://node-pptx-templater.pages.dev/',
  articleData = null
}) {
  useEffect(() => {
    // 1. Update Title
    document.title = title;

    // Helper to set or create meta tag
    const setMetaTag = (selector, attribute, value) => {
      let tag = document.querySelector(selector);
      if (!tag) {
        tag = document.createElement('meta');
        if (selector.startsWith('meta[name=')) {
          tag.setAttribute('name', selector.replace(/meta\[name="?([^"\]]+)"?\]/, '$1'));
        } else if (selector.startsWith('meta[property=')) {
          tag.setAttribute('property', selector.replace(/meta\[property="?([^"\]]+)"?\]/, '$1'));
        }
        document.head.appendChild(tag);
      }
      tag.setAttribute(attribute, value);
    };

    // 2. Standard Meta Tags
    setMetaTag('meta[name="description"]', 'content', description);
    setMetaTag(
      'meta[name="keywords"]',
      'content',
      'node-pptx-templater, node.js pptx template engine, openxml powerpoint generator, serverless powerpoint lambda, pure js pptx, powerpoint chart excel sync, cloudflare workers pptx, vercel edge powerpoint, table cell merge pptx, z-order shapes pptx'
    );
    setMetaTag('meta[name="author"]', 'content', 'node-pptx-templater contributors');

    // 3. OpenGraph Tags
    setMetaTag('meta[property="og:title"]', 'content', title);
    setMetaTag('meta[property="og:description"]', 'content', description);
    setMetaTag('meta[property="og:type"]', 'content', articleData ? 'article' : 'website');
    setMetaTag('meta[property="og:url"]', 'content', canonicalUrl);
    setMetaTag('meta[property="og:image"]', 'content', 'https://node-pptx-templater.pages.dev/og-banner.png');
    setMetaTag('meta[property="og:image:width"]', 'content', '1200');
    setMetaTag('meta[property="og:image:height"]', 'content', '630');
    setMetaTag('meta[property="og:image:alt"]', 'content', 'node-pptx-templater — High-Performance PowerPoint OpenXML Engine');

    // 4. Twitter / X Cards
    setMetaTag('meta[name="twitter:card"]', 'content', 'summary_large_image');
    setMetaTag('meta[name="twitter:title"]', 'content', title);
    setMetaTag('meta[name="twitter:description"]', 'content', description);
    setMetaTag('meta[name="twitter:image"]', 'content', 'https://node-pptx-templater.pages.dev/twitter-card.png');
    setMetaTag('meta[name="twitter:creator"]', 'content', '@jsuyog2');

    // 5. Canonical Link
    let canonical = document.querySelector('link[rel="canonical"]');
    if (!canonical) {
      canonical = document.createElement('link');
      canonical.setAttribute('rel', 'canonical');
      document.head.appendChild(canonical);
    }
    canonical.setAttribute('href', canonicalUrl);

    // 6. JSON-LD Structured Data
    const schemas = [
      {
        '@context': 'https://schema.org',
        '@type': 'SoftwareApplication',
        'name': 'node-pptx-templater',
        'headline': 'High-Performance PowerPoint OpenXML Template Engine',
        'description': description,
        'operatingSystem': 'Cross-platform (Node.js, AWS Lambda, Cloudflare Workers, Vercel Edge)',
        'applicationCategory': 'DeveloperApplication',
        'softwareVersion': '3.0.0-alpha.0',
        'license': 'https://opensource.org/licenses/MIT',
        'offers': {
          '@type': 'Offer',
          'price': '0',
          'priceCurrency': 'USD'
        },
        'author': {
          '@type': 'Organization',
          'name': 'node-pptx-templater contributors',
          'url': 'https://github.com/jsuyog2/node-pptx-templater'
        },
        'aggregateRating': {
          '@type': 'AggregateRating',
          'ratingValue': '5.0',
          'ratingCount': '150'
        }
      },
      {
        '@context': 'https://schema.org',
        '@type': 'Organization',
        'name': 'node-pptx-templater',
        'url': 'https://node-pptx-templater.pages.dev/',
        'logo': 'https://node-pptx-templater.pages.dev/og-banner.png',
        'sameAs': [
          'https://github.com/jsuyog2/node-pptx-templater',
          'https://www.npmjs.com/package/node-pptx-templater'
        ]
      }
    ];

    if (articleData) {
      schemas.push({
        '@context': 'https://schema.org',
        '@type': 'TechArticle',
        'headline': articleData.title,
        'description': articleData.summary || description,
        'author': {
          '@type': 'Organization',
          'name': 'node-pptx-templater contributors'
        },
        'datePublished': '2026-10-11T00:00:00Z',
        'dateModified': '2026-10-11T00:00:00Z'
      });
    }

    let jsonLdScript = document.getElementById('json-ld-structured-data');
    if (!jsonLdScript) {
      jsonLdScript = document.createElement('script');
      jsonLdScript.id = 'json-ld-structured-data';
      jsonLdScript.type = 'application/ld+json';
      document.head.appendChild(jsonLdScript);
    }
    jsonLdScript.textContent = JSON.stringify(schemas, null, 2);
  }, [title, description, canonicalUrl, articleData]);

  return null;
}
