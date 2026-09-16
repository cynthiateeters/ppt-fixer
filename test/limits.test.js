// Invalid XML characters and size limits. Most tests build their own inputs and always run;
// the ones that need the reference deck skip without it (see test/sample.js).

import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { zipSync, unzipSync, strFromU8, strToU8 } from "fflate";
import { sample, skip } from "./sample.js";
import {
  loadDeck,
  applyEdits,
  cleanXmlText,
  tiffDimensions,
  DeckTooLargeError,
  LIMITS,
  MAX_PREVIEW_PIXELS,
} from "../src/pptx.js";

// Built from char codes so this file contains no control characters itself.
const ch = (...codes) => String.fromCharCode(...codes);

test("cleanXmlText removes characters XML doesn't allow and keeps the rest", () => {
  const input = `a${ch(0x0)}b${ch(0x1)}c${ch(0x8)}d${ch(0xb)}e${ch(0xc)}f${ch(0x1f)}g${ch(0xfffe)}h${ch(0xffff)}i`;
  assert.equal(cleanXmlText(input), "abcdefghi");
  assert.equal(
    cleanXmlText(`tab${ch(0x9)}nl${ch(0xa)}cr${ch(0xd)}`),
    `tab${ch(0x9)}nl${ch(0xa)}cr${ch(0xd)}`,
  );
  assert.equal(
    cleanXmlText("café, naïve, 日本, 🎨"),
    "café, naïve, 日本, 🎨",
    "accents, CJK and emoji pairs survive",
  );
  assert.equal(cleanXmlText(`x${ch(0xd83c)}y`), "xy", "lone high surrogate removed");
  assert.equal(cleanXmlText(`x${ch(0xdfa8)}y`), "xy", "lone low surrogate removed");
});

test("invalid characters never reach the saved deck", { skip }, () => {
  const deck = loadDeck(new Uint8Array(readFileSync(sample)));
  const dirty = `Pasted${ch(0x1)} title${ch(0xb)}`;
  const bytes = applyEdits(deck, {
    slides: {
      8: { title: dirty, pictures: { 0: { alt: `Alt${ch(0x0)} text`, decorative: false } } },
      9: { title: `${ch(0x1)}${ch(0x2)} ` }, // only invalid characters: counts as empty
    },
    docTitle: `Doc${ch(0x1f)} title`,
  });
  const files = unzipSync(bytes);
  for (const name of ["ppt/slides/slide8.xml", "ppt/slides/slide9.xml", "docProps/core.xml"]) {
    const xml = strFromU8(files[name]);
    const bad = [...xml].find((c) => {
      const code = c.charCodeAt(0);
      return (
        (code < 0x20 && code !== 0x9 && code !== 0xa && code !== 0xd) ||
        code === 0xfffe ||
        code === 0xffff
      );
    });
    assert.equal(bad, undefined, `${name} has no control characters`);
  }
  const fixed = loadDeck(bytes);
  assert.equal(fixed.slides[7].title, "Pasted title");
  assert.equal(fixed.slides[7].pictures[0].alt, "Alt text");
  assert.equal(fixed.slides[8].title, "", "a title of only invalid characters isn't written");
  assert.equal(fixed.docTitle, "Doc title");
});

test("rejects a file over the input size limit", () => {
  const bytes = zipSync({ "ppt/presentation.xml": strToU8("<p:presentation/>") });
  assert.throws(() => loadDeck(bytes, { ...LIMITS, inputBytes: 10 }), DeckTooLargeError);
});

test("rejects a small zip that expands past the limit, before expanding it", () => {
  const big = new Uint8Array(20 * 1024 * 1024).fill(0x61); // 20 MB of "a"
  const bomb = zipSync({ "ppt/presentation.xml": big }, { level: 9 });
  assert.ok(bomb.length < 100_000, `bomb is small on disk (${bomb.length} bytes)`);
  assert.throws(
    () => loadDeck(bomb, { ...LIMITS, expandedBytes: 5 * 1024 * 1024 }),
    (e) => e instanceof DeckTooLargeError && /unpacks to more than 5 MB/.test(e.message),
  );
});

test("rejects a zip with too many parts", () => {
  const parts = Object.fromEntries(
    Array.from({ length: 50 }, (_, i) => [`x/${i}.xml`, strToU8("<a/>")]),
  );
  assert.throws(() => loadDeck(zipSync(parts), { ...LIMITS, entries: 20 }), DeckTooLargeError);
});

test("reads TIFF dimensions from the header without decoding", { skip }, () => {
  const deck = loadDeck(new Uint8Array(readFileSync(sample)));
  assert.deepEqual(tiffDimensions(deck.files["ppt/media/image2.tiff"]), {
    width: 485,
    height: 438,
  });
  assert.equal(tiffDimensions(deck.files["ppt/media/image1.tiff"]), null, "PNG data named .tiff");
});

test("a TIFF header claiming huge dimensions is over the preview cap", () => {
  // Little-endian TIFF, one directory with ImageWidth and ImageLength as LONG values.
  const buf = new Uint8Array(8 + 2 + 2 * 12 + 4);
  const v = new DataView(buf.buffer);
  buf.set([0x49, 0x49]);
  v.setUint16(2, 42, true);
  v.setUint32(4, 8, true);
  v.setUint16(8, 2, true);
  const entry = (at, tag, value) => {
    v.setUint16(at, tag, true);
    v.setUint16(at + 2, 4, true);
    v.setUint32(at + 4, 1, true);
    v.setUint32(at + 8, value, true);
  };
  entry(10, 256, 100000);
  entry(22, 257, 100000);
  const dims = tiffDimensions(buf);
  assert.deepEqual(dims, { width: 100000, height: 100000 });
  assert.ok(dims.width * dims.height > MAX_PREVIEW_PIXELS);
  assert.equal(tiffDimensions(buf.slice(0, 12)), null, "truncated header");
});

test("real decks load within the default limits", { skip }, () => {
  assert.doesNotThrow(() => loadDeck(new Uint8Array(readFileSync(sample))));
});
