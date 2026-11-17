/** @type {import('next').NextConfig} */
const isGithubPages = process.env.GITHUB_PAGES === 'true' || process.env.DOCS_ENV === 'github-pages';
const basePath = process.env.DOCS_BASE_PATH ?? (isGithubPages ? '/node-pptx-templater' : '');

const nextConfig = {
  output: 'export',
  ...(basePath ? { basePath, assetPrefix: `${basePath}/` } : {}),
  images: {
    unoptimized: true,
  },
  trailingSlash: true,
};

export default nextConfig;
