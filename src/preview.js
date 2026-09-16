// Turns picture bytes from a deck into something an <img> can show.

import { sniffImage, tiffDimensions, MAX_PREVIEW_PIXELS } from "./pptx.js";

const NATIVE = new Set(["image/png", "image/jpeg", "image/gif", "image/bmp", "image/webp"]);
const cache = new Map();
let worker = null;
let nextId = 0;
const pending = new Map();

const DECODE_TIMEOUT_MS = 15000;

// A worker that crashed, failed to load or hung can't be trusted with more jobs:
// fail everything waiting on it and start a fresh worker next time.
function resetWorker(reason) {
  worker?.terminate();
  worker = null;
  for (const job of pending.values()) {
    clearTimeout(job.timer);
    job.reject(new Error(reason));
  }
  pending.clear();
}

function tiffWorker() {
  if (!worker) {
    worker = new Worker(new URL("./tiff-worker.js", import.meta.url), { type: "module" });
    worker.onmessage = ({ data }) => {
      const job = pending.get(data.id);
      if (!job) return;
      clearTimeout(job.timer);
      pending.delete(data.id);
      if (data.error) job.reject(new Error(data.error));
      else job.resolve(data);
    };
    worker.onerror = () => resetWorker("The image decoder stopped working.");
    worker.onmessageerror = () => resetWorker("The image decoder sent something unreadable.");
  }
  return worker;
}

async function tiffToUrl(bytes) {
  // Check the size in the header first: a crafted TIFF can claim dimensions that would exhaust memory.
  const dims = tiffDimensions(bytes);
  if (!dims || dims.width * dims.height > MAX_PREVIEW_PIXELS) return null;
  const id = nextId++;
  const copy = bytes.slice();
  const { width, height, rgba } = await new Promise((resolve, reject) => {
    const timer = setTimeout(
      () => resetWorker("The image took too long to decode."),
      DECODE_TIMEOUT_MS,
    );
    pending.set(id, { resolve, reject, timer });
    tiffWorker().postMessage({ id, bytes: copy }, [copy.buffer]);
  });
  const canvas = document.createElement("canvas");
  canvas.width = width;
  canvas.height = height;
  canvas
    .getContext("2d")
    .putImageData(new ImageData(new Uint8ClampedArray(rgba.buffer), width, height), 0, 0);
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, "image/png"));
  return URL.createObjectURL(blob);
}

// Resolves to an object URL, or null when the format can't be previewed.
export function previewUrl(mediaPath, bytes) {
  if (!bytes) return Promise.resolve(null);
  if (!cache.has(mediaPath)) {
    const type = sniffImage(bytes);
    const job = NATIVE.has(type)
      ? Promise.resolve(URL.createObjectURL(new Blob([bytes], { type })))
      : type === "image/tiff"
        ? tiffToUrl(bytes).catch(() => null)
        : Promise.resolve(null);
    cache.set(mediaPath, job);
  }
  return cache.get(mediaPath);
}

export function clearPreviews() {
  for (const job of cache.values()) job.then((url) => url && URL.revokeObjectURL(url));
  cache.clear();
}
