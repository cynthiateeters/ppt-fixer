// Valid XML that PowerPoint doesn't write itself but other tools might, plus crafted part names.
// Each test edits a copy of the reference deck (see test/sample.js).

import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { zipSync, unzipSync, strFromU8, strToU8 } from "fflate";
import { loadDeck, applyEdits, mediaBytes } from "../src/pptx.js";
import { sample, skip } from "./sample.js";

// Returns a deck built from the reference deck with one part's text changed.
function variant(path, change) {
  const files = unzipSync(new Uint8Array(readFileSync(sample)));
  files[path] = strToU8(change(strFromU8(files[path])));
  return zipSync(files);
}

test("reads and replaces single-quoted alt text without duplicating it", { skip }, () => {
  const bytes = variant("ppt/slides/slide8.xml", (x) =>
    x.replace('name="Picture 3"', "name='Picture 3' descr='old alt'"),
  );
  const deck = loadDeck(bytes);
  assert.equal(deck.slides[7].pictures[0].alt, "old alt");
  const out = unzipSync(
    applyEdits(deck, { slides: { 8: { pictures: { 0: { alt: "new alt", decorative: false } } } } }),
  );
  const tag = strFromU8(out["ppt/slides/slide8.xml"]).match(/<p:cNvPr id="4"[^>]*>/)[0];
  assert.equal((tag.match(/\sdescr=/g) ?? []).length, 1, tag);
  assert.equal(loadDeck(zipSync(out)).slides[7].pictures[0].alt, "new alt");
});

test("finds a title placeholder written with single quotes", { skip }, () => {
  const bytes = variant("ppt/slides/slide7.xml", (x) =>
    x.replace('<p:ph type="title"/>', "<p:ph type='title'/>"),
  );
  const deck = loadDeck(bytes);
  assert.equal(deck.slides[6].title, "Sovereign Posture (1/2)");
  assert.equal(deck.slides[6].hasTitleBox, true);
});

test("reads and replaces a document title that has attributes", { skip }, () => {
  const bytes = variant("docProps/core.xml", (x) =>
    x.replace("<dc:title>", '<dc:title xml:lang="en-US">'),
  );
  const deck = loadDeck(bytes);
  assert.equal(deck.docTitle, "Chapter 8");
  const core = strFromU8(
    unzipSync(applyEdits(deck, { slides: {}, docTitle: "Chapter 9" }))["docProps/core.xml"],
  );
  assert.equal((core.match(/<dc:title\b/g) ?? []).length, 1, "one title element");
  assert.match(core, /<dc:title xml:lang="en-US">Chapter 9<\/dc:title>/);
});

test("parts named like JavaScript built-ins are never mistaken for files", { skip }, () => {
  const files = unzipSync(new Uint8Array(readFileSync(sample)));
  files["toString"] = strToU8("<x/>");
  const rels = "ppt/slides/_rels/slide8.xml.rels";
  files[rels] = strToU8(
    strFromU8(files[rels]).replace("../media/image1.tiff", "../../constructor"),
  );
  const deck = loadDeck(zipSync(files));
  const pic = deck.slides[7].pictures[0];
  assert.equal(pic.mediaPath, "constructor");
  assert.equal(mediaBytes(deck, pic.mediaPath), null, "no built-in function returned");
  assert.ok(
    mediaBytes(deck, "toString") instanceof Uint8Array,
    "a real part named toString is still readable",
  );
  assert.doesNotThrow(() => applyEdits(deck, { slides: { 8: { title: "Still works" } } }));
});
