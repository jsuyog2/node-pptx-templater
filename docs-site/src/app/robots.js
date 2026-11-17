import siteMeta from '../data/siteMeta.json';

export default function robots() {
  return {
    rules: {
      userAgent: '*',
      allow: ['/', '/llms.txt'],
    },
    sitemap: `${siteMeta.domainUrl}/sitemap.xml`,
  };
}
