import siteMeta from '../data/siteMeta.json';
import guidesData from '../data/guides.json';

export default function sitemap() {
  const baseUrl = siteMeta.domainUrl.replace(/\/+$/, '');
  const now = new Date();

  // Exact 8 Standardized Categories
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
      priority: 0.8,
    },
    {
      url: `${baseUrl}/features/`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.8,
    },
    {
      url: `${baseUrl}/docs/api/`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.8,
    },
    ...standardCategories.map((category) => ({
      url: `${baseUrl}/docs/api/?category=${encodeURIComponent(category)}`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.8,
    })),
    {
      url: `${baseUrl}/docs/installation/`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.64,
    },
    {
      url: `${baseUrl}/docs/compatibility/`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.64,
    },
    ...guidesData.map((guide) => ({
      url: `${baseUrl}/guides/${guide.slug}/`,
      lastModified: now,
      changeFrequency: 'monthly',
      priority: 0.7,
    })),
  ];

  return routes;
}
