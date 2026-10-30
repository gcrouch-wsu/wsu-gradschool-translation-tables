/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  outputFileTracingRoot: __dirname,
  // The translation table validator is a static page in public/validate; its files load by relative path.
  async redirects() {
    return [{ source: '/validate', destination: '/validate/index.html', permanent: false }]
  },
}

module.exports = nextConfig


