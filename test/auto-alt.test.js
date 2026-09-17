// PowerPoint's automatic alt text counts as missing until someone checks it.

import { test } from "node:test";
import assert from "node:assert/strict";
import { isAutoAlt, altMissing, needsAlt, AUTO_ALT_MARKERS } from "../src/pptx.js";

test("recognizes PowerPoint's automatic descriptions, with and without confidence", () => {
  for (const alt of [
    "A picture containing text\n\nDescription automatically generated",
    "A plant in a room\n\nDescription automatically generated with low confidence",
    "A person looking at art\n\nDescription automatically generated with medium confidence",
    "A diagram\n\nDescription automatically generated with high confidence",
    "A chart. Description automatically generated.  ",
  ])
    assert.equal(isAutoAlt(alt), true, alt);
});

test("leaves written descriptions and mentions of the phrase alone", () => {
  for (const alt of [
    "A surreal painting of two goldfish near an ice block",
    'The caption "Description automatically generated" is PowerPoint\'s stock text, shown here',
    "",
  ])
    assert.equal(isAutoAlt(alt), false, alt);
});

test("a guess needs work until the marker line is gone or the picture is decorative", () => {
  const guess = "A painting of a city on fire\n\nDescription automatically generated";
  assert.equal(needsAlt({ alt: guess, decorative: false }), true);
  assert.equal(altMissing("A painting of a city on fire", false), false);
  assert.equal(altMissing(guess, true), false);
  assert.equal(altMissing("   ", false), true);
});

test("only the languages asked for are matched", () => {
  const guess = "A chart\n\nDescription automatically generated";
  assert.equal(isAutoAlt(guess, []), false);
  assert.equal(isAutoAlt(guess, ["xx"]), false);
  assert.ok(AUTO_ALT_MARKERS.en instanceof RegExp);
});
