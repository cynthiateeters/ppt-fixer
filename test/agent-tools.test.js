// The WebMCP tools an AI agent calls. The deck is built in memory, so no course file is needed.

import { test } from "node:test";
import assert from "node:assert/strict";
import {
  listWork,
  setTitle,
  setDescription,
  registerAgentTools,
  MAX_TITLE,
} from "../src/agent-tools.js";

const guess = "A painting\n\nDescription automatically generated";

// Slide 1 is done, slide 2 needs a title and has a PowerPoint guess,
// slide 3 has a decorative picture and the same picture as slide 2.
function makeDeck() {
  const deck = {
    slides: [
      {
        number: 1,
        title: "Welcome",
        titleHidden: false,
        bodyText: ["Week 1"],
        notes: [],
        pictures: [{ alt: "A red chair", decorative: false, mediaPath: "ppt/media/a.png" }],
      },
      {
        number: 2,
        title: "",
        titleHidden: false,
        bodyText: ["Oil on canvas"],
        notes: ["Talk about color"],
        pictures: [
          { alt: guess, decorative: false, mediaPath: "ppt/media/b.png" },
          { alt: "", decorative: false, mediaPath: "ppt/media/c.png" },
        ],
      },
      {
        number: 3,
        title: "Detail",
        titleHidden: false,
        bodyText: [],
        notes: [],
        pictures: [
          { alt: "", decorative: true, mediaPath: "ppt/media/d.png" },
          { alt: "", decorative: false, mediaPath: "ppt/media/b.png" },
        ],
      },
    ],
  };
  const edits = {
    slides: deck.slides.map((s) => ({
      title: s.title,
      pictures: s.pictures.map((p) => ({ alt: p.alt, decorative: p.decorative, touched: false })),
    })),
  };
  return { deck, edits };
}

test("lists only the slides that still need work, with their context", () => {
  const { deck, edits } = makeDeck();
  const result = listWork(deck, edits);
  assert.equal(result.ok, true);
  assert.deepEqual(
    result.slides.map((s) => s.slide),
    [2, 3],
  );
  const two = result.slides[0];
  assert.equal(two.titleNeeded, true);
  assert.deepEqual(two.slideText, ["Oil on canvas"]);
  assert.equal(two.speakerNotes, "Talk about color");
  assert.equal(two.pictures[0].powerpointGuess, true);
  assert.equal(two.pictures[0].needsDescription, true);
  assert.equal(result.slides[1].pictures[0].needsDescription, false, "decorative needs nothing");
  assert.equal(listWork(deck, edits, { includeDone: true }).slides.length, 3);
});

test("a slide drops off the list once its title and descriptions are filled", () => {
  const { deck, edits } = makeDeck();
  assert.equal(setTitle(deck, edits, 2, "  Two paintings  ").ok, true);
  assert.equal(edits.slides[1].title, "Two paintings");
  setDescription(deck, edits, 2, 1, "Oil painting of a harbor at dusk");
  setDescription(deck, edits, 2, 2, "Charcoal drawing of a hand");
  assert.deepEqual(
    listWork(deck, edits).slides.map((s) => s.slide),
    [],
    "slide 3's copy of picture b was filled too",
  );
});

test("a description reaches copies of the same picture, but never overwrites typed ones", () => {
  const { deck, edits } = makeDeck();
  const result = setDescription(deck, edits, 2, 1, "Oil painting of a harbor");
  assert.deepEqual(result.alsoUsedOn, [{ slide: 3, picture: 2 }]);
  assert.equal(edits.slides[2].pictures[1].alt, "Oil painting of a harbor");

  const typed = makeDeck();
  typed.edits.slides[2].pictures[1] = {
    alt: "Typed by a person",
    decorative: false,
    touched: true,
  };
  setDescription(typed.deck, typed.edits, 2, 1, "Oil painting of a harbor");
  assert.equal(typed.edits.slides[2].pictures[1].alt, "Typed by a person");

  const unshared = makeDeck();
  const r = setDescription(unshared.deck, unshared.edits, 2, 1, "Oil painting", false);
  assert.deepEqual(r.alsoUsedOn, []);
  assert.equal(unshared.edits.slides[2].pictures[1].alt, "");
});

test("refuses bad input with a reason the agent can act on, and changes nothing", () => {
  const { deck, edits } = makeDeck();
  const before = structuredClone(edits);
  for (const result of [
    setTitle(deck, edits, 9, "Nope"),
    setTitle(deck, edits, 2, "   "),
    setTitle(deck, edits, 2, "x".repeat(MAX_TITLE + 1)),
    setTitle(deck, edits, "2", "Not a number"),
    setDescription(deck, edits, 2, 5, "No such picture"),
    setDescription(deck, edits, 3, 1, "Decorative, so refused"),
    setDescription(deck, edits, 2, 1, ""),
    listWork(null, null),
  ]) {
    assert.equal(result.ok, false);
    assert.equal(typeof result.error, "string");
  }
  assert.deepEqual(edits, before);
});

test("registers four tools, none of which download or mark pictures decorative", async () => {
  const registered = [];
  globalThis.document = { modelContext: { registerTool: (t) => registered.push(t) } };
  try {
    const { deck, edits } = makeDeck();
    const calls = [];
    const ok = registerAgentTools({
      getDeck: () => deck,
      getEdits: () => edits,
      showSlide: (n) => calls.push(["show", n]),
      shareChecked: () => true,
      refreshTitle: (n) => calls.push(["title", n]),
      refreshPicture: (n, i) => calls.push(["picture", n, i]),
      changed: () => calls.push(["changed"]),
    });
    assert.equal(ok, true);
    assert.deepEqual(
      registered.map((t) => t.name),
      ["list_slides_needing_work", "show_slide", "set_slide_title", "set_picture_description"],
    );
    const tool = (name) => registered.find((t) => t.name === name);

    const reply = await tool("set_picture_description").execute({
      slide: 2,
      picture: 1,
      description: "Oil painting of a harbor",
    });
    assert.equal(reply.ok, true);
    assert.equal("changed" in reply, false, "internal bookkeeping stays out of the reply");
    assert.deepEqual(calls, [["picture", 2, 0], ["picture", 3, 1], ["changed"]]);

    calls.length = 0;
    await tool("set_slide_title").execute({ slide: 9, title: "Nope" });
    assert.deepEqual(calls, [], "a refused call touches nothing on the page");
    await tool("show_slide").execute({ slide: 3 });
    assert.deepEqual(calls, [["show", 3]]);
  } finally {
    delete globalThis.document;
  }
});

test("does nothing in a browser without WebMCP", () => {
  assert.equal(registerAgentTools({}), false);
});
