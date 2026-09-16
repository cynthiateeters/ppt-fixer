import { defineConfig } from "vite";

// Same policy as netlify.toml, so `pnpm preview` catches anything the deployed site would block.
const CSP =
  "default-src 'self'; img-src 'self' blob: data:; worker-src 'self' blob:; style-src 'self'; script-src 'self'; connect-src 'self'";

export default defineConfig({
  // Relative asset paths, so the built page works from any folder or host.
  base: "./",
  worker: { format: "es" },
  preview: {
    port: 4174,
    strictPort: true,
    headers: { "Content-Security-Policy": CSP },
  },
});
