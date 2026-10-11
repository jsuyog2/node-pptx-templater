import siteMeta from '../data/siteMeta.json';
import guidesData from '../data/guides.json';

export default function sitemap() {
  const baseUrl = siteMeta.domainUrl.replace(/\/+$/, '');
  const now = new Date();

  // Categories with priority 0.8
  const primaryCategories = [
    'Lifecycle & I/O',
    'Slides Management',
    'Text & Placeholders',
    'Tables & Cell Shapes',
    'Charts & Series Sync',
    'Shapes & Geometry',
    'Images & Media',
    'Z-Order & Layering',
    'Diagnostics & Validation',
  ];

  // Categories with priority 0.64
  const secondaryCategories = [
    'Slides',
    'Hyperlinks',
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
    ...primaryCategories.map((category) => ({
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
    ...secondaryCategories.map((category) => ({
      url: `${baseUrl}/docs/api/?category=${encodeURIComponent(category)}`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.64,
    })),
  ];

  // Dynamic Guide Routes (if present in guidesData)
  if (Array.isArray(guidesData)) {
    guidesData.forEach((guide) => {
      routes.push({
        url: `${baseUrl}/guides/${guide.slug}/`,
        lastModified: now,
        changeFrequency: 'weekly',
        priority: 0.8,
      });
    });
  }

  return routes;
}