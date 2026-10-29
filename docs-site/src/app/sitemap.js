import siteMeta from '../data/siteMeta.json';

export default function sitemap() {
  const baseUrl = siteMeta.domainUrl.replace(/\/+$/, '');
  const now = new Date();

  // Strict 8 Standardized Categories
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
      url: `${baseUrl}/docs/capabilities/`,
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
    // Also include canonical /api-reference alias / query entries
    {
      url: `${baseUrl}/api-reference/`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.9,
    },
    ...standardCategories.map((category) => ({
      url: `${baseUrl}/docs/api-reference/?category=${encodeURIComponent(category)}`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.8,
    })),
  ];

  return routes;
}
