import siteMeta from '../data/siteMeta.json';

export default function robots() {
  return {
    rules: {
      userAgent: '*',
      allow: '/',
    },
    sitemap: `${siteMeta.domainUrl}/sitemap.xml`,
  };
}
