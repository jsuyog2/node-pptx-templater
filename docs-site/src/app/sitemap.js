import siteMeta from '../data/siteMeta.json';

export default function sitemap() {
  const baseUrl = siteMeta.domainUrl.replace(/\/+$/, '');
  const now = new Date();

  // Strict 8 Standardized Categories
  const categorySlugs = [
    'import-export',
    'slides',
    'text-placeholders',
    'tables',
    'charts',
    'shapes',
    'images-media',
    'elements',
  ];

  const standardCategories = [
    'Import/Export',
    'Slides',
    'Text & Placeholders',
    'Tables',
    'Charts',
    'Shapes',
    'Images & Media',
    'Elements',
  ];

  const routes = [
    {
      url: `${baseUrl}/`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 1.0,
    },
    {
      url: `${baseUrl}/docs/`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.9,
    },
    {
      url: `${baseUrl}/docs/installation/`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.85,
    },
    {
      url: `${baseUrl}/docs/api-reference/`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.95,
    },
    {
      url: `${baseUrl}/api-reference/`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.9,
    },
    // Dedicated static category routes
    ...categorySlugs.map((slug) => ({
      url: `${baseUrl}/docs/api-reference/${slug}/`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.9,
    })),
    // Backward compatibility query entries
    ...standardCategories.map((category) => ({
      url: `${baseUrl}/docs/api-reference/?category=${encodeURIComponent(category)}`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.75,
    })),
  ];

  return routes;
}
