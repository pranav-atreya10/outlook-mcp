import type { NextConfig } from "next";
import path from "node:path";

const stub = path.resolve(process.cwd(), "lib/cloudflare-workers-stub.ts");

const nextConfig: NextConfig = {
  // The sign-in library imports "cloudflare:workers"; swap in a tiny stand-in.
  webpack(config, { webpack }) {
    config.plugins.push(
      new webpack.NormalModuleReplacementPlugin(/^cloudflare:workers$/, (resource: { request: string }) => {
        resource.request = stub;
      }),
    );
    return config;
  },
};

export default nextConfig;
