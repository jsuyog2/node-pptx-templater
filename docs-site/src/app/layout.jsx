import './globals.css';
import Header from '../components/Header';
import Footer from '../components/Footer';
import siteMeta from '../data/siteMeta.json';

export const metadata = {
  metadataBase: new URL(siteMeta.domainUrl),
  title: {
    default: siteMeta.title,
    template: '%s | node-pptx-templater',
  },
  description: siteMeta.description,
  keywords: siteMeta.keywords,
  authors: [{ name: siteMeta.author, url: 'https://github.com/jsuyog2/node-pptx-templater' }],
  creator: siteMeta.author,
  publisher: siteMeta.author,
  alternates: {
    canonical: '/',
  },
  openGraph: {
    title: 'node-pptx-templater — High-Performance PowerPoint OpenXML Template Engine',
    description: 'Populate PowerPoint templates dynamically in Node.js & Serverless runtimes without Microsoft Office dependencies.',
    url: siteMeta.domainUrl,
    siteName: 'node-pptx-templater',
    images: [
      {
        url: '/og-banner.png',
        width: 1200,
        height: 630,
        alt: 'node-pptx-templater — High-Performance PowerPoint OpenXML Template Engine',
      },
    ],
    locale: 'en_US',
    type: 'website',
  },
  twitter: {
    card: 'summary_large_image',
    title: 'node-pptx-templater — High-Performance PowerPoint OpenXML Template Engine',
    description: 'Populate PowerPoint templates dynamically in Node.js & Serverless runtimes without Microsoft Office dependencies.',
    images: ['/twitter-card.png'],
    creator: '@jsuyog2',
  },
  robots: {
    index: true,
    follow: true,
    googleBot: {
      index: true,
      follow: true,
      'max-video-preview': -1,
      'max-image-preview': 'large',
      'max-snippet': -1,
    },
  },
  verification: {
    google: 'google1a674320340859aa',
  },
};

const jsonLd = {
  '@context': 'https://schema.org',
  '@type': 'SoftwareApplication',
  name: 'node-pptx-templater',
  applicationCategory: 'DeveloperApplication',
  operatingSystem: 'Node.js, AWS Lambda, Cloudflare Workers, Vercel Edge, Bun, Deno',
  softwareVersion: siteMeta.version,
  license: 'https://opensource.org/licenses/MIT',
  description: siteMeta.description,
  offers: {
    '@type': 'Offer',
    price: '0',
    priceCurrency: 'USD',
  },
  author: {
    '@type': 'Person',
    name: 'node-pptx-templater contributors',
    url: 'https://github.com/jsuyog2/node-pptx-templater',
  },
};

export default function RootLayout({ children }) {
  return (
    <html lang="en" className="dark">
      <head>
        <link rel="icon" href="/favicon.svg" type="image/svg+xml" />
        <script
          type="application/ld+json"
          dangerouslySetInnerHTML={{ __html: JSON.stringify(jsonLd) }}
        />
      </head>
      <body className="flex flex-col min-h-screen antialiased selection:bg-brand-500/30 selection:text-brand-200">
        <Header />
        <main className="flex-grow">{children}</main>
        <Footer />
      </body>
    </html>
  );
}
