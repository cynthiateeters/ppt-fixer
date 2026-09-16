// Tests that read the reference deck. See test/sample.js.

import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { unzipSync, strFromU8 } from "fflate";
import { loadDeck, applyEdits, summarize, sniffImage, needsTitle, needsAlt } from "../src/pptx.js";
import { sample, skip } from "./sample.js";

test("reads the reference deck the way the audit measured it", { skip }, () => {
  const deck = loadDeck(new Uint8Array(readFileSync(sample)));
  assert.deepEqual(summarize(deck), {
    slides: 25,
    untitled: 11,
    pictures: 16,
    missingAlt: 16,
    needWork: 13,
  });
  assert.equal(deck.docTitle, "Chapter 8");
  assert.deepEqual(
    deck.slides.filter(needsTitle).map((s) => s.number),
    [8, 9, 11, 12, 14, 15, 16, 22, 23, 24, 25],
  );
  assert.equal(deck.slides[13].hasTitleBox, false, "slide 14 has no title box");
});

test("sniffs images by content, not extension", { skip }, () => {
  const deck = loadDeck(new Uint8Array(readFileSync(sample)));
  assert.equal(sniffImage(deck.files["ppt/media/image1.tiff"]), "image/png");
  assert.equal(sniffImage(deck.files["ppt/media/image2.tiff"]), "image/tiff");
  assert.equal(sniffImage(deck.files["ppt/media/image8.JPG"]), "image/jpeg");
});

test("writes titles and alt text, moving titles off image slides", { skip }, () => {
  const deck = loadDeck(new Uint8Array(readFileSync(sample)));
  const edits = { slides: {}, docTitle: "Chapter 9: Platform and Posture" };
  for (const s of deck.slides) {
    const e = {};
    if (needsTitle(s)) e.title = `Slide ${s.number} <test> & "quotes"`;
    if (s.pictures.length) {
      e.pictures = Object.fromEntries(
        s.pictures.map((p) => [
          p.index,
          p.index === 1 && s.number === 23
            ? { alt: "", decorative: true }
            : { alt: `Picture ${p.index} on slide ${s.number} & more`, decorative: false },
        ]),
      );
    }
    edits.slides[s.number] = e;
  }
  const fixedBytes = applyEdits(deck, edits);
  const fixed = loadDeck(fixedBytes);
  const sum = summarize(fixed);
  assert.equal(sum.untitled, 0);
  assert.equal(sum.missingAlt, 0);
  assert.equal(fixed.docTitle, "Chapter 9: Platform and Posture");
  assert.equal(fixed.slides[7].title, 'Slide 8 <test> & "quotes"');
  assert.equal(fixed.slides[22].pictures[1].decorative, true);
  assert.equal(fixed.slides[7].pictures[0].alt, "Picture 0 on slide 8 & more");

  const files = unzipSync(fixedBytes);
  assert.equal(Object.keys(files)[0], "[Content_Types].xml");
  for (const n of [8, 14, 24]) {
    const xml = strFromU8(files[`ppt/slides/slide${n}.xml`]);
    const y = Number(
      xml.match(/type="title"[\s\S]*?<a:off x="-?\d+" y="(-?\d+)"\/><a:ext cx="\d+" cy="(\d+)"/)[1],
    );
    assert.ok(y < 0, `slide ${n} title sits above the slide`);
  }
  // Text slides that already had titles are untouched.
  assert.equal(
    strFromU8(files["ppt/slides/slide7.xml"]),
    strFromU8(deck.files["ppt/slides/slide7.xml"]),
  );
  assert.equal(needsAlt(fixed.slides[7].pictures[0]), false);
});

test("keeps $ patterns in typed text literally", { skip }, () => {
  const deck = loadDeck(new Uint8Array(readFileSync(sample)));
  const tricky = "Cost $& vs $` and $' or $1 $$";
  const fixed = loadDeck(
    applyEdits(deck, {
      slides: {
        8: { title: `${tricky} a`, pictures: { 0: { alt: `${tricky} b`, decorative: false } } },
        14: { title: `${tricky} c` },
        7: { title: `${tricky} d` },
      },
      docTitle: `${tricky} e`,
    }),
  );
  assert.equal(fixed.slides[7].title, `${tricky} a`, "existing empty title box, image slide");
  assert.equal(fixed.slides[7].pictures[0].alt, `${tricky} b`);
  assert.equal(fixed.slides[13].title, `${tricky} c`, "slide with no title box");
  assert.equal(fixed.slides[6].title, `${tricky} d`, "retitled text slide");
  assert.equal(fixed.docTitle, `${tricky} e`);
});
