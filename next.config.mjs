/**
 * MyExcel is fully client-side: the site is a static export (out/) served by any static
 * server on the web and bundled into the Tauri desktop app. No API routes, SSR, server
 * actions or image optimization are used.
 * @type {import('next').NextConfig}
 */
const nextConfig = {
  output: 'export',
  reactStrictMode: false,
  images: { unoptimized: true },
  trailingSlash: false,
  poweredByHeader: false,
  webpack: (config) => {
    config.resolve.fallback = { ...config.resolve.fallback, fs: false, path: false, stream: false, crypto: false };
    return config;
  },
};
export default nextConfig;
