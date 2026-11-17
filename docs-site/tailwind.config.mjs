/** @type {import('tailwindcss').Config} */
export default {
  darkMode: 'class',
  content: [
    './src/pages/**/*.{js,ts,jsx,tsx,mdx}',
    './src/components/**/*.{js,ts,jsx,tsx,mdx}',
    './src/app/**/*.{js,ts,jsx,tsx,mdx}',
  ],
  theme: {
    extend: {
      colors: {
        ppt: {
          rust: '#D24726',
          vermillion: '#D24726',
          dark: {
            canvas: '#000000',
            chrome: 'rgba(38, 38, 38, 0.5)',
            card: 'rgba(54, 54, 54, 0.5)',
            border: 'rgba(210, 71, 38, 0.35)',
            text: '#E6E6E6',
            secondary: '#929299',
          },
          light: {
            canvas: '#FFFFFF',
            chrome: '#E6DEC8',
            card: 'rgba(245, 240, 230, 0.5)',
            border: 'rgba(121, 117, 104, 0.3)',
            text: '#302F28',
            secondary: '#797568',
          },
        },
        brand: {
          50: '#fff4f2',
          100: '#ffe6e1',
          200: '#ffd0c7',
          300: '#fcaaa0',
          400: '#f87b6c',
          500: '#D24726',
          600: '#b83b1e',
          700: '#993018',
          800: '#7e2a18',
          900: '#68271a',
          950: '#391109',
        },
      },
      fontFamily: {
        sans: ['Inter', '-apple-system', 'BlinkMacSystemFont', 'Segoe UI', 'Roboto', 'sans-serif'],
        mono: ['ui-monospace', 'SFMono-Regular', 'Menlo', 'Monaco', 'Consolas', 'monospace'],
      },
      boxShadow: {
        'rust-glow': '0 0 12px rgba(210, 71, 38, 0.4)',
        'rust-glow-lg': '0 0 24px rgba(210, 71, 38, 0.5)',
      },
    },
  },
  plugins: [],
};
