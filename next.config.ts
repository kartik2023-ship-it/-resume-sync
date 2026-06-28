import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // pdf-parse uses pdfjs-dist which references pdf.worker.mjs as a separate
  // file. When Turbopack bundles it the worker path breaks. Opting out of
  // bundling lets Node.js resolve both files from node_modules directly.
  serverExternalPackages: ["pdf-parse"],
  // Ensure DOCX template files are included in Vercel output file tracing
  outputFileTracingIncludes: {
    "/api/download/*": ["./templates/**/*"],
  },
};

export default nextConfig;
