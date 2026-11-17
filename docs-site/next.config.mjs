/** @type {import('next').NextConfig} */
const nextConfig = {
  output: 'export',
  basePath: '/node-pptx-templater',
  assetPrefix: '/node-pptx-templater/',
  images: {
    unoptimized: true,
  },
  trailingSlash: true,
};

export default nextConfig;

