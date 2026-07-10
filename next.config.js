/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  async rewrites() {
    return [
      { source: "/py", destination: "/catalogo-zenith.html" },
    ];
  },
};

module.exports = nextConfig;
