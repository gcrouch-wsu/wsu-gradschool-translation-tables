import type { Config } from 'tailwindcss'

const config: Config = {
  content: [
    './pages/**/*.{js,ts,jsx,tsx,mdx}',
    './components/**/*.{js,ts,jsx,tsx,mdx}',
    './app/**/*.{js,ts,jsx,tsx,mdx}',
  ],
  theme: {
    extend: {
      colors: {
        // WSU Brand Colors from brand.wsu.edu
        'wsu-crimson': '#A60F2D',
        'wsu-crimson-dark': '#8c0d25',
        'wsu-gray': '#4D4D4D',
        'wsu-gray-light': '#5E6A71',
        'wsu-text-dark': '#2A3033',
        'wsu-text-body': '#333333',
        'wsu-text-muted': '#5E6A71',
        'wsu-bg-light': '#f4f4f4',
        'wsu-bg-card': '#f9f9f9',
        'wsu-border-light': '#e0e0e0',
        'wsu-border-medium': '#d9d9d9',
      },
      fontFamily: {
        sans: ['Arial', 'Helvetica', 'sans-serif'],
      },
    },
  },
  plugins: [],
}
export default config


