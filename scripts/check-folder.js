// Robustness check: load every .pptx in a folder, fill in every missing title and alt text,
// write the fixed deck to memory, reload it, and confirm nothing is left missing.
// Usage: node scripts/check-folder.js <folder> [out-folder]

import { readdirSync, readFileSync, writeFileSync, mkdirSync } from "node:fs";
import { join } from "node:path";
import { loadDeck, applyEdits, summarize, needsTitle, needsAlt } from "../src/pptx.js";

const [dir, outDir] = process.argv.slice(2);
if (!dir) {
  console.error("Usage: node scripts/check-folder.js <folder> [out-folder]");
  process.exit(2);
}
if (outDir) mkdirSync(outDir, { recursive: true });

let failures = 0;
for (const name of readdirSync(dir)
  .filter((n) => n.endsWith(".pptx"))
  .sort()) {
  try {
    const deck = loadDeck(new Uint8Array(readFileSync(join(dir, name))));
    const before = summarize(deck);
    const edits = { slides: {} };
    for (const s of deck.slides) {
      const e = {};
      if (needsTitle(s)) e.title = `Test title ${s.number}`;
      const pics = s.pictures.filter(needsAlt);
      if (pics.length)
        e.pictures = Object.fromEntries(
          pics.map((p) => [p.index, { alt: `Test alt ${s.number}.${p.index}`, decorative: false }]),
        );
      edits.slides[s.number] = e;
    }
    const bytes = applyEdits(deck, edits);
    const after = summarize(loadDeck(bytes));
    const ok =
      after.untitled === 0 &&
      after.missingAlt === 0 &&
      after.slides === before.slides &&
      after.pictures === before.pictures;
    if (!ok) failures++;
    const imageSlides = deck.slides.filter((s) => s.imageSlide && needsTitle(s)).length;
    console.log(
      `${ok ? "OK  " : "FAIL"} ${name}: ${before.slides} slides, untitled ${before.untitled}->${after.untitled} (${imageSlides} off-slide), alt ${before.missingAlt}->${after.missingAlt}, other visuals ${deck.slides.reduce((n, s) => n + s.otherVisuals, 0)}`,
    );
    if (outDir) writeFileSync(join(outDir, name), bytes);
  } catch (e) {
    failures++;
    console.log(`ERR  ${name}: ${e.message}`);
  }
}
process.exit(failures ? 1 : 0);
