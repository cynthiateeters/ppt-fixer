import { defineConfig } from "vite";

// The page's security policy. netlify.toml sends the same policy as a header (keep them in sync),
// `pnpm preview` sends it too, and the build also writes it into index.html as a <meta> tag,
// so the page stays locked down on a host that doesn't send headers.
//
// - blob: in img-src is for picture previews, which are created in the page.
// - frame-ancestors only works as a header; browsers ignore it in a <meta> tag.
const POLICY = [
  "default-src 'self'",
  "img-src 'self' blob:",
  "worker-src 'self'",
  "style-src 'self'",
  "script-src 'self'",
  "connect-src 'self'",
  "object-src 'none'",
  "base-uri 'none'",
  "form-action 'none'",
];
const HEADER_CSP = [...POLICY, "frame-ancestors 'none'"].join("; ");
const META_CSP = POLICY.join("; ");

const cspMetaTag = {
  name: "csp-meta-tag",
  apply: "build",
  transformIndexHtml: () => [
    {
      tag: "meta",
      attrs: { "http-equiv": "Content-Security-Policy", content: META_CSP },
      injectTo: "head-prepend",
    },
    { tag: "meta", attrs: { name: "referrer", content: "no-referrer" }, injectTo: "head-prepend" },
  ],
};

export default defineConfig({
  // Relative asset paths, so the built page works from any folder or host.
  base: "./",
  worker: { format: "es" },
  plugins: [cspMetaTag],
  preview: {
    port: 4174,
    strictPort: true,
    headers: {
      "Content-Security-Policy": HEADER_CSP,
      "X-Frame-Options": "DENY",
    },
  },
});
