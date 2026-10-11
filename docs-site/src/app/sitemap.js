import siteMeta from '../data/siteMeta.json';
import guidesData from '../data/guides.json';

export default function sitemap() {
  const baseUrl = siteMeta.domainUrl;
  const now = new Date();

  const routes = [
    {
      url: `${baseUrl}/`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 1.0,
    },
    {
      url: `${baseUrl}/api-reference/`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.9,
    },
  ];

  // Dynamic Guide Routes
  guidesData.forEach((guide) => {
    routes.push({
      url: `${baseUrl}/guides/${guide.slug}/`,
      lastModified: now,
      changeFrequency: 'weekly',
      priority: 0.8,
    });
  });

  return routes;
}
