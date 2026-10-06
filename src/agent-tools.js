// WebMCP tools: named actions a browser AI agent can call instead of clicking around the page.
// Agents that support WebMCP (ChatGPT's desktop browser today) find these when the page loads.
// In browsers without WebMCP, registerAgentTools does nothing.
//
// The agent can read what needs work, scroll a slide into view, and fill title and
// description boxes. It can't download, open a file or mark a picture decorative:
// a person checks the work and does those.

import { needsTitle, needsAlt, altMissing, isAutoAlt } from "./pptx.js";

export const MAX_TITLE = 250;
export const MAX_DESCRIPTION = 1000;

const fail = (error) => ({ ok: false, error });

function findSlide(deck, slideNumber) {
  if (!deck) return { error: "No deck is open. Ask the person to open a .pptx file first." };
  const slide = Number.isInteger(slideNumber) ? deck.slides[slideNumber - 1] : undefined;
  if (!slide)
    return { error: `There's no slide ${slideNumber}. The deck has ${deck.slides.length} slides.` };
  return { slide };
}

// ---------- logic with no DOM, so tests can run it ----------

export function listWork(deck, edits, { includeDone = false } = {}) {
  if (!deck) return fail("No deck is open. Ask the person to open a .pptx file first.");
  const slides = [];
  for (const slide of deck.slides) {
    const e = edits.slides[slide.number - 1];
    const titleNeeded = needsTitle(slide) && !e.title.trim();
    const pictures = slide.pictures.map((p, i) => {
      const pe = e.pictures[i];
      return {
        picture: i + 1,
        description: pe.alt,
        needsDescription: needsAlt(p) && altMissing(pe.alt, pe.decorative),
        powerpointGuess: isAutoAlt(pe.alt),
        decorative: pe.decorative,
      };
    });
    const open = titleNeeded || pictures.some((p) => p.needsDescription);
    if (!open && !includeDone) continue;
    slides.push({
      slide: slide.number,
      title: e.title,
      titleNeeded,
      slideText: slide.bodyText.slice(0, 12),
      speakerNotes: slide.notes.join(" "),
      pictures,
    });
  }
  return {
    ok: true,
    totalSlides: deck.slides.length,
    slidesListed: slides.length,
    slides,
  };
}

export function setTitle(deck, edits, slideNumber, title) {
  const { slide, error } = findSlide(deck, slideNumber);
  if (error) return fail(error);
  const text = typeof title === "string" ? title.trim() : "";
  if (!text) return fail("The title is empty.");
  if (text.length > MAX_TITLE) return fail(`Keep the title under ${MAX_TITLE} characters.`);
  edits.slides[slide.number - 1].title = text;
  return { ok: true, slide: slide.number, title: text };
}

// shareWithCopies mirrors the page's "Use this description for the same picture" checkbox:
// copies of the same picture on other slides get the text too, unless someone already typed there.
export function setDescription(
  deck,
  edits,
  slideNumber,
  pictureNumber,
  description,
  shareWithCopies = true,
) {
  const { slide, error } = findSlide(deck, slideNumber);
  if (error) return fail(error);
  const index = Number.isInteger(pictureNumber) ? pictureNumber - 1 : -1;
  const pic = slide.pictures[index];
  if (!pic)
    return fail(
      `Slide ${slide.number} has no picture ${pictureNumber}. It has ${slide.pictures.length}.`,
    );
  const pe = edits.slides[slide.number - 1].pictures[index];
  if (pe.decorative)
    return fail(
      `A person marked slide ${slide.number}, picture ${pictureNumber} as decorative, so it needs no description. Leave it.`,
    );
  const text = typeof description === "string" ? description.trim() : "";
  if (!text) return fail("The description is empty.");
  if (text.length > MAX_DESCRIPTION)
    return fail(`Keep the description under ${MAX_DESCRIPTION} characters.`);

  pe.alt = text;
  pe.touched = true;
  const changed = [{ slide: slide.number, index }];
  if (shareWithCopies && pic.mediaPath)
    for (const s of deck.slides)
      s.pictures.forEach((p, i) => {
        if (p.mediaPath !== pic.mediaPath || (s.number === slide.number && i === index)) return;
        const other = edits.slides[s.number - 1].pictures[i];
        if (other.touched || other.decorative) return;
        other.alt = text;
        changed.push({ slide: s.number, index: i });
      });
  return {
    ok: true,
    slide: slide.number,
    picture: pictureNumber,
    alsoUsedOn: changed.slice(1).map((c) => ({ slide: c.slide, picture: c.index + 1 })),
    changed,
  };
}

// ---------- registration ----------

const SLIDE = { type: "integer", minimum: 1, description: "Slide number, starting at 1." };
const PICTURE = {
  type: "integer",
  minimum: 1,
  description: "Picture number on that slide, starting at 1, as the page labels it.",
};

// page supplies: getDeck, getEdits, showSlide(n), shareChecked(n, index),
// refreshTitle(n), refreshPicture(n, index), changed().
export function registerAgentTools(page) {
  const mc = globalThis.document?.modelContext ?? globalThis.navigator?.modelContext;
  if (typeof mc?.registerTool !== "function") return false;

  const tools = [
    {
      name: "list_slides_needing_work",
      description:
        "List the slides that still need a slide title or a picture description, with each slide's text and speaker notes for context. Call this first, and again to check nothing is left.",
      inputSchema: {
        type: "object",
        properties: {
          includeDone: { type: "boolean", description: "Also list slides that are already done." },
        },
        additionalProperties: false,
      },
      annotations: { readOnlyHint: true },
      execute: async ({ includeDone = false } = {}) =>
        listWork(page.getDeck(), page.getEdits(), { includeDone }),
    },
    {
      name: "show_slide",
      description:
        "Scroll a slide onto the screen so its pictures load and can be seen. Look at each picture before describing it.",
      inputSchema: {
        type: "object",
        properties: { slide: SLIDE },
        required: ["slide"],
        additionalProperties: false,
      },
      annotations: { readOnlyHint: true },
      execute: async ({ slide }) => {
        const found = findSlide(page.getDeck(), slide);
        if (found.error) return fail(found.error);
        page.showSlide(slide);
        return { ok: true, slide, note: "Picture previews can take a moment to appear." };
      },
    },
    {
      name: "set_slide_title",
      description:
        "Fill in one slide's title box. A person reviews every title before downloading the fixed file.",
      inputSchema: {
        type: "object",
        properties: {
          slide: SLIDE,
          title: {
            type: "string",
            description: "A few plain words saying what the slide is about.",
          },
        },
        required: ["slide", "title"],
        additionalProperties: false,
      },
      execute: async ({ slide, title }) => {
        const result = setTitle(page.getDeck(), page.getEdits(), slide, title);
        if (result.ok) {
          page.refreshTitle(slide);
          page.changed();
        }
        return result;
      },
    },
    {
      name: "set_picture_description",
      description:
        "Fill in the description (alt text) box for one picture on a slide. If the same picture appears on other slides, they get it too unless someone already wrote one there. A person reviews every description before downloading.",
      inputSchema: {
        type: "object",
        properties: {
          slide: SLIDE,
          picture: PICTURE,
          description: {
            type: "string",
            description: "What the picture shows, for someone who can't see it.",
          },
        },
        required: ["slide", "picture", "description"],
        additionalProperties: false,
      },
      execute: async ({ slide, picture, description }) => {
        const share = page.shareChecked(slide, picture - 1);
        const result = setDescription(
          page.getDeck(),
          page.getEdits(),
          slide,
          picture,
          description,
          share,
        );
        if (!result.ok) return result;
        for (const c of result.changed) page.refreshPicture(c.slide, c.index);
        page.changed();
        const { changed, ...reply } = result;
        return reply;
      },
    },
  ];

  // The API is a draft and may return a promise; a refusal shouldn't break the page.
  for (const tool of tools)
    try {
      Promise.resolve(mc.registerTool(tool)).catch((e) =>
        console.warn(`WebMCP didn't register ${tool.name}:`, e),
      );
    } catch (e) {
      console.warn(`WebMCP didn't register ${tool.name}:`, e);
    }
  return true;
}
