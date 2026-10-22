import type { NextConfig } from "next";

// Static export so the widget can be packed with ZET and hosted internally in Zoho CRM.
const nextConfig: NextConfig = {
  output: "export",
  trailingSlash: true,
  images: { unoptimized: true },
  // Relative asset paths are needed for Zoho hosting, but they stop `next dev` from hydrating.
  assetPrefix: process.env.NODE_ENV === "production" ? "./" : undefined,
  turbopack: {
    rules: {
      "*.css": {
        loaders: ["@tailwindcss/turbopack"],
        as: "*.css",
      },
    },
  },
};

export default nextConfig;
