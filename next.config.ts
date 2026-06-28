import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // pdf-parse uses pdfjs-dist which references pdf.worker.mjs as a separate
  // file. When Turbopack bundles it the worker path breaks. Opting out of
  // bundling lets Node.js resolve both files from node_modules directly.
  serverExternalPackages: ["pdf-parse"],
};

export default nextConfig;
