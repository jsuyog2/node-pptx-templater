import './globals.css';
import { ThemeProvider } from '../context/ThemeContext';
import TitleBar from '../components/PowerPointChrome/TitleBar';
import RibbonHeader from '../components/PowerPointChrome/RibbonHeader';
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

export default function RootLayout({ children }) {
  return (
    <html lang="en" className="dark">
      <head>
        <link rel="icon" href="/favicon.svg" type="image/svg+xml" />
        <script
          type="application/ld+json"
          dangerouslySetInnerHTML={{ __html: JSON.stringify(siteMeta.jsonLd) }}
        />
      </head>
      <body className="flex flex-col min-h-screen antialiased selection:bg-[#D24726]/30 selection:text-white">
        <ThemeProvider>
          {/* PowerPoint Desktop Application Chrome */}
          <TitleBar />
          <RibbonHeader />

          {/* Main Application Canvas */}
          <main className="flex-1 w-full overflow-x-hidden">
            {children}
          </main>
        </ThemeProvider>
      </body>
    </html>
  );
}
