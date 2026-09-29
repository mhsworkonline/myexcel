/** @type {import('next').NextConfig} */
const isTauri = process.env.TAURI_BUILD === '1';
const nextConfig = {
  reactStrictMode: false,
  ...(isTauri ? { output: 'export', images: { unoptimized: true } } : {}),
  webpack: (config) => {
    config.resolve.fallback = { ...config.resolve.fallback, fs: false, path: false, stream: false, crypto: false };
    return config;
  },
};
export default nextConfig;
