import { loadDeck, applyEdits, needsTitle, needsAlt, mediaBytes, LIMITS } from "./pptx.js";
import { previewUrl, clearPreviews } from "./preview.js";
import {
  fingerprint,
  loadSaved,
  scheduleSave,
  saveNow,
  clearSaved,
  cancelPendingSave,
  pruneOld,
  storageAvailable,
} from "./saved-work.js";

const $ = (id) => document.getElementById(id);

// Build an element with text only; never innerHTML.
function h(tag, props = {}, ...children) {
  const el = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (v === undefined || v === null || v === false) continue;
    if (k === "class") el.className = v;
    else if (k.startsWith("on")) el.addEventListener(k.slice(2), v);
    else if (k in el && typeof v !== "string") el[k] = v;
    else el.setAttribute(k, v === true ? "" : v);
  }
  for (const c of children.flat()) if (c !== null && c !== undefined && c !== false) el.append(c);
  return el;
}

const state = {
  deck: null,
  fileName: "",
  edits: null,
  current: 0,
  onlyNeedsWork: true,
  id: null, // fingerprint of the open file
  dirty: false, // typed something since opening or restoring
  saving: true, // false after "Delete saved work" or when storage is blocked
};

// ---------- edit state ----------

function freshEdits(deck) {
  return {
    docTitle: deck.docTitle,
    slides: deck.slides.map((s) => ({
      title: s.titleHidden ? s.title : s.title,
      pictures: s.pictures.map((p) => ({ alt: p.alt, decorative: p.decorative, touched: false })),
    })),
  };
}

const editFor = (slide) => state.edits.slides[slide.number - 1];

function slideOpen(slide) {
  const e = editFor(slide);
  const titleOpen = needsTitle(slide) && !e.title.trim();
  const altOpen = slide.pictures.some((p, i) => {
    const pe = e.pictures[i];
    return needsAlt(p) && !pe.decorative && !pe.alt.trim();
  });
  return { titleOpen, altOpen, open: titleOpen || altOpen };
}

const slideNeededWork = (slide) => needsTitle(slide) || slide.pictures.some(needsAlt);

function otherUses(mediaPath, except) {
  const uses = [];
  for (const s of state.deck.slides)
    s.pictures.forEach((p, i) => {
      if (p.mediaPath === mediaPath && !(s.number === except.slide && i === except.index))
        uses.push({ slide: s, index: i });
    });
  return uses;
}

function visibleSlides() {
  const all = state.deck.slides;
  return state.onlyNeedsWork ? all.filter(slideNeededWork) : all;
}

// ---------- rendering ----------

function renderStatus() {
  const slides = state.deck.slides;
  const open = slides.map(slideOpen);
  const titles = open.filter((o) => o.titleOpen).length;
  const pics = slides.reduce((n, s, i) => {
    const e = state.edits.slides[i];
    return (
      n +
      s.pictures.filter(
        (p, j) => needsAlt(p) && !e.pictures[j].decorative && !e.pictures[j].alt.trim(),
      ).length
    );
  }, 0);
  const parts = [];
  parts.push(
    titles
      ? `${titles} slide${titles === 1 ? "" : "s"} still need a title`
      : "Every slide has a title",
  );
  parts.push(
    pics
      ? `${pics} picture${pics === 1 ? "" : "s"} still need alt text`
      : "every picture has alt text or is marked decorative",
  );
  $("deck-status").textContent = `${slides.length} slides. ${parts.join(", ")}.`;
}

function renderList() {
  const list = $("slide-list");
  list.replaceChildren();
  const slides = visibleSlides();
  if (!slides.length) {
    list.append(h("li", { class: "empty" }, "No slides need work. You can still download a copy."));
    return;
  }
  for (const s of slides) {
    const { open } = slideOpen(s);
    const label = editFor(s).title.trim() || "Untitled";
    list.append(
      h(
        "li",
        {},
        h(
          "button",
          {
            type: "button",
            class: `slide-link${s.number - 1 === state.current ? " current" : ""}${open ? " open" : " done"}`,
            "aria-current": s.number - 1 === state.current ? "true" : null,
            onclick: () => goTo(s.number - 1),
          },
          h("span", { class: "num" }, String(s.number)),
          h("span", { class: "label" }, label),
          h("span", { class: "state" }, open ? "Needs work" : slideNeededWork(s) ? "Done" : "OK"),
        ),
      ),
    );
  }
}

const formatTime = (ms) =>
  new Date(ms).toLocaleString(undefined, {
    month: "short",
    day: "numeric",
    hour: "numeric",
    minute: "2-digit",
  });

function setSaveStatus(text) {
  $("save-status").textContent = text;
}

// Called after every edit the person makes.
function changed() {
  state.dirty = true;
  refresh();
  if (!state.saving) return;
  const queued = scheduleSave(state.id, state.fileName, state.edits, (ok) =>
    setSaveStatus(
      ok
        ? `Saved in this browser at ${formatTime(Date.now())}`
        : "Couldn't save in this browser. Download before closing this tab.",
    ),
  );
  if (!queued) {
    state.saving = false;
    setSaveStatus("This browser isn't saving your work. Download before closing this tab.");
  }
}

function refresh() {
  renderStatus();
  renderList();
}

function titleSection(slide) {
  const e = editFor(slide);
  const id = `title-${slide.number}`;
  const hintId = `${id}-hint`;
  let hint;
  if (!needsTitle(slide))
    hint = "This slide already has a title. Edit it only if it needs better wording.";
  else if (slide.imageSlide)
    hint =
      "This slide is mostly picture, so the title goes just above the slide where it won't show when you present. Screen readers and Ally still find it, so give it real words.";
  else hint = "The title will appear on the slide in the layout's usual title spot.";

  const input = h("input", {
    type: "text",
    id,
    value: e.title,
    "aria-describedby": hintId,
    oninput: (ev) => {
      e.title = ev.target.value;
      changed();
    },
  });
  return h(
    "section",
    { class: "field-group" },
    h("h3", {}, "Slide title", needsTitle(slide) ? h("span", { class: "flag" }, " missing") : null),
    h("label", { for: id, class: "visually-hidden" }, `Title for slide ${slide.number}`),
    input,
    h("p", { class: "hint", id: hintId }, hint),
  );
}

function pictureCard(slide, pic, index) {
  const e = editFor(slide).pictures[index];
  const id = `alt-${slide.number}-${index}`;
  const img = h("img", { alt: "", class: "thumb" });
  const frame = h(
    "div",
    { class: "thumb-frame" },
    h("span", { class: "loading" }, "Loading preview..."),
  );
  previewUrl(pic.mediaPath, mediaBytes(state.deck, pic.mediaPath)).then((url) => {
    if (url) {
      img.src = url;
      frame.replaceChildren(img);
    } else
      frame.replaceChildren(
        h(
          "span",
          { class: "loading" },
          "No preview for this picture. Check it in PowerPoint.",
        ),
      );
  });

  const uses = pic.mediaPath ? otherUses(pic.mediaPath, { slide: slide.number, index }) : [];
  const shareId = `${id}-share`;
  const share = uses.length
    ? h(
        "label",
        { class: "check" },
        h("input", { type: "checkbox", id: shareId, checked: true }),
        ` Use this description for the same picture on slide${uses.length === 1 ? "" : "s"} ${[...new Set(uses.map((u) => u.slide.number))].join(", ")}`,
      )
    : null;

  const textarea = h("textarea", {
    id,
    rows: 3,
    disabled: e.decorative,
    oninput: (ev) => {
      e.alt = ev.target.value;
      e.touched = true;
      if (share && share.querySelector("input").checked) {
        for (const u of uses) {
          const other = state.edits.slides[u.slide.number - 1].pictures[u.index];
          if (!other.touched) other.alt = e.alt;
        }
      }
      changed();
    },
  });
  textarea.value = e.alt;

  const decorative = h("input", {
    type: "checkbox",
    checked: e.decorative,
    onchange: (ev) => {
      e.decorative = ev.target.checked;
      textarea.disabled = e.decorative;
      changed();
    },
  });

  return h(
    "section",
    { class: "picture" },
    frame,
    h(
      "div",
      { class: "picture-fields" },
      h(
        "h3",
        {},
        `Picture ${index + 1}`,
        needsAlt(pic) ? h("span", { class: "flag" }, " needs alt text") : null,
      ),
      h("label", { for: id }, "Describe what this picture shows students"),
      textarea,
      h("label", { class: "check" }, decorative, " Decorative, no description needed"),
      share,
    ),
  );
}

function renderSlide() {
  const slides = state.deck.slides;
  const slide = slides[state.current];
  const article = $("slide");
  const pos = visibleSlides().indexOf(slide);
  const seq = visibleSlides();

  const context = h(
    "section",
    { class: "context" },
    h("h3", {}, "Text on this slide"),
    slide.bodyText.length
      ? h(
          "ul",
          {},
          slide.bodyText.slice(0, 12).map((t) => h("li", {}, t)),
        )
      : h("p", { class: "hint" }, "No other text on this slide."),
    slide.notes.length
      ? h("details", {}, h("summary", {}, "Speaker notes"), h("p", {}, slide.notes.join(" ")))
      : null,
  );

  const extras = [];
  if (slide.background)
    extras.push(
      h(
        "p",
        { class: "note" },
        "This slide uses a picture as its background. Backgrounds can't hold alt text, so if the picture matters, describe it in the slide title or the speaker notes.",
      ),
    );
  if (slide.otherVisuals)
    extras.push(
      h(
        "p",
        { class: "note" },
        "This slide also has charts, tables or picture-filled shapes this page can't edit. Check them in PowerPoint's Accessibility Assistant.",
      ),
    );

  const prev = h(
    "button",
    { type: "button", disabled: pos <= 0, onclick: () => goTo(seq[pos - 1].number - 1) },
    "Previous",
  );
  const next = h(
    "button",
    {
      type: "button",
      class: "primary",
      disabled: pos < 0 || pos >= seq.length - 1,
      onclick: () => goTo(seq[pos + 1].number - 1),
    },
    "Next slide",
  );

  const parts = [
    h("h2", {}, `Slide ${slide.number} of ${slides.length}`),
    context,
    titleSection(slide),
    slide.pictures.length
      ? h("h3", { class: "section-head" }, slide.pictures.length === 1 ? "Picture" : "Pictures")
      : null,
    slide.pictures.map((p, i) => pictureCard(slide, p, i)),
    extras,
    h("div", { class: "pager" }, prev, next),
  ];
  article.replaceChildren(...parts.flat().filter(Boolean));
}

function goTo(index) {
  state.current = index;
  renderList();
  renderSlide();
  $("slide").focus();
}

// ---------- file handling ----------

async function openFile(file) {
  const error = $("start-error");
  error.hidden = true;
  if (!file) return;
  try {
    // Check the size before reading, so a huge file is never loaded into memory.
    if (file.size > LIMITS.inputBytes)
      throw new Error(
        `It's ${Math.round(file.size / 1048576)} MB, and this page handles files up to ${Math.round(LIMITS.inputBytes / 1048576)} MB.`,
      );
    const bytes = new Uint8Array(await file.arrayBuffer());
    const deck = loadDeck(bytes);
    cancelPendingSave();
    clearPreviews();
    state.deck = deck;
    state.fileName = file.name;
    state.id = await fingerprint(bytes);
    state.dirty = false;
    state.saving = storageAvailable();
    const saved = state.saving ? loadSaved(state.id, deck) : null;
    state.edits = saved ? saved.edits : freshEdits(deck);
    if (typeof state.edits.docTitle !== "string") state.edits.docTitle = deck.docTitle;
    $("restored").hidden = !saved;
    if (saved)
      $("restored-text").textContent =
        `Picked up where you left off. Your work on this file was saved in this browser on ${formatTime(saved.savedAt)}.`;
    setSaveStatus(
      state.saving
        ? saved
          ? `Saved in this browser at ${formatTime(saved.savedAt)}`
          : "Your typing is saved in this browser as you go."
        : "This browser isn't saving your work. Download before closing this tab.",
    );
    const first = deck.slides.find(slideNeededWork) ?? deck.slides[0];
    state.current = first ? first.number - 1 : 0;
    state.onlyNeedsWork = deck.slides.some(slideNeededWork);
    $("only-needs-work").checked = state.onlyNeedsWork;
    $("deck-name").textContent = file.name;
    $("doc-title").value = state.edits.docTitle;
    $("start").hidden = true;
    $("editor").hidden = false;
    refresh();
    renderSlide();
    $("deck-name").focus();
  } catch (e) {
    error.textContent = `Couldn't open that file. ${e.message}`;
    error.hidden = false;
  }
}

function buildEdits() {
  const out = {
    slides: {},
    docTitle: state.edits.docTitle !== state.deck.docTitle ? state.edits.docTitle : undefined,
  };
  state.deck.slides.forEach((s, i) => {
    const e = state.edits.slides[i];
    const slideEdit = {};
    if (e.title.trim() && (e.title.trim() !== s.title || s.titleHidden)) slideEdit.title = e.title;
    const pics = {};
    s.pictures.forEach((p, j) => {
      const pe = e.pictures[j];
      if (pe.alt.trim() !== p.alt.trim() || pe.decorative !== p.decorative)
        pics[j] = { alt: pe.alt, decorative: pe.decorative };
    });
    if (Object.keys(pics).length) slideEdit.pictures = pics;
    if (Object.keys(slideEdit).length) out.slides[s.number] = slideEdit;
  });
  return out;
}

function download() {
  const bytes = applyEdits(state.deck, buildEdits());
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  });
  const url = URL.createObjectURL(blob);
  const a = h("a", {
    href: url,
    download: state.fileName.replace(/\.pptx$/i, "") + " (fixed).pptx",
  });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 10000);
}

$("file").addEventListener("change", (e) => openFile(e.target.files[0]));
const drop = $("drop");
drop.addEventListener("dragover", (e) => {
  e.preventDefault();
  drop.classList.add("over");
});
drop.addEventListener("dragleave", () => drop.classList.remove("over"));
drop.addEventListener("drop", (e) => {
  e.preventDefault();
  drop.classList.remove("over");
  openFile(e.dataTransfer.files[0]);
});
$("only-needs-work").addEventListener("change", (e) => {
  state.onlyNeedsWork = e.target.checked;
  renderList();
  renderSlide();
});
$("doc-title").addEventListener("input", (e) => {
  state.edits.docTitle = e.target.value;
  changed();
});
$("download").addEventListener("click", download);
$("close-deck").addEventListener("click", () => {
  if (state.saving && state.dirty) saveNow(state.id, state.fileName, state.edits);
  cancelPendingSave();
  state.deck = null;
  clearPreviews();
  $("file").value = "";
  $("editor").hidden = true;
  $("start").hidden = false;
});

// Throw away restored work and start from the file as it is.
$("start-over").addEventListener("click", () => {
  clearSaved(state.id);
  state.edits = freshEdits(state.deck);
  state.dirty = false;
  $("doc-title").value = state.edits.docTitle;
  $("restored").hidden = true;
  setSaveStatus(state.saving ? "Your typing is saved in this browser as you go." : "");
  refresh();
  renderSlide();
});

// For a computer login other people also use: remove this file's saved work and stop saving for now.
$("forget").addEventListener("click", () => {
  clearSaved(state.id);
  state.saving = false;
  $("restored").hidden = true;
  setSaveStatus("Saved work deleted. This browser won't save more until you open the file again.");
});

window.addEventListener("beforeunload", (e) => {
  if (!state.deck || !state.dirty) return;
  if (state.saving && saveNow(state.id, state.fileName, state.edits)) return;
  e.preventDefault();
  e.returnValue = "";
});

pruneOld();
