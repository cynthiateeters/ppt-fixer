// Read a .pptx, report slide titles and picture alt text, and write fixes back.
// Works in the browser and in Node: no DOM APIs, only string edits on the slide XML.

import { unzipSync, zipSync, strFromU8, strToU8 } from "fflate";

const TITLE_SP =
  /<p:sp>(?:(?!<\/p:sp>)[\s\S])*?<p:ph\b[^>]*type=["'](?:title|ctrTitle)["'][\s\S]*?<\/p:sp>/;
// XML allows either quote around attribute values. PowerPoint writes double quotes,
// but other tools may not, so every attribute match below accepts both.
const PIC = /<p:pic>[\s\S]*?<\/p:pic>/g;
const DECORATIVE_URI = "{C183D7F6-B498-43B3-948B-1728B52AA6E4}";

// A picture covering this share of the slide makes it an "image slide":
// a title there gets flagged for contrast by Ally, even behind the picture.
export const IMAGE_SLIDE_COVERAGE = 0.5;

const NAMED_ENTITIES = { lt: "<", gt: ">", quot: '"', apos: "'", amp: "&" };
// One pass, so "&amp;#xA;" decodes to the text "&#xA;" rather than a line break.
// PowerPoint writes line breaks in alt text as numeric entities like "&#xA;".
const decodeXml = (s) =>
  s.replace(/&(?:#x([0-9a-f]+)|#([0-9]+)|(lt|gt|quot|apos|amp));/gi, (m, hex, dec, name) => {
    if (name) return NAMED_ENTITIES[name] ?? m;
    const code = hex ? parseInt(hex, 16) : Number(dec);
    return code <= 0x10ffff ? String.fromCodePoint(code) : m;
  });
// Characters XML 1.0 doesn't allow: most control characters, U+FFFE, U+FFFF and unpaired surrogates.
// They mostly arrive by pasting from PDFs or web pages, and PowerPoint treats them as corruption.
const esc = (code) => `\\u${code.toString(16).padStart(4, "0")}`;
const INVALID_XML_CHARS = new RegExp(
  `[${esc(0x0)}-${esc(0x8)}${esc(0xb)}${esc(0xc)}${esc(0xe)}-${esc(0x1f)}${esc(0xfffe)}${esc(0xffff)}]` +
    `|[${esc(0xd800)}-${esc(0xdbff)}](?![${esc(0xdc00)}-${esc(0xdfff)}])` +
    `|(?<![${esc(0xd800)}-${esc(0xdbff)}])[${esc(0xdc00)}-${esc(0xdfff)}]`,
  "g",
);
export const cleanXmlText = (s) => s.replace(INVALID_XML_CHARS, "");

const encodeText = (s) =>
  cleanXmlText(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
// Line breaks and tabs stay entities in attributes, or XML parsers turn them into spaces.
const encodeAttr = (s) =>
  encodeText(s)
    .replace(/"/g, "&quot;")
    .replace(/\n/g, "&#xA;")
    .replace(/\r/g, "&#xD;")
    .replace(/\t/g, "&#x9;");

const attr = (tag, name) => {
  const m = tag.match(new RegExp(`\\s${name}=(?:"([^"]*)"|'([^']*)')`));
  return m ? decodeXml(m[1] ?? m[2]) : null;
};

const textOf = (xml) =>
  [...xml.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)]
    .map((p) => [...p[1].matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((t) => decodeXml(t[1])).join(""))
    .filter((line) => line.trim());

const resolvePath = (base, target) => {
  const parts = base.split("/").slice(0, -1);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else if (seg !== ".") parts.push(seg);
  }
  return parts.join("/");
};

const parseRels = (xml) =>
  Object.fromEntries(
    [...(xml ?? "").matchAll(/<Relationship\b[^>]*>/g)].map((m) => [
      attr(m[0], "Id"),
      attr(m[0], "Target"),
    ]),
  );

// Largest TIFF the preview will decode: a decoder allocates width x height x 4 bytes.
export const MAX_PREVIEW_PIXELS = 50_000_000;

// Width and height from a classic TIFF's first image directory, read without decoding.
// Returns null when the header can't be read, so the caller can skip the preview.
export function tiffDimensions(bytes) {
  if (bytes.length < 8) return null;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const little = bytes[0] === 0x49 && bytes[1] === 0x49;
  if (!little && !(bytes[0] === 0x4d && bytes[1] === 0x4d)) return null;
  if (view.getUint16(2, little) !== 42) return null;
  const ifd = view.getUint32(4, little);
  if (ifd + 2 > bytes.length) return null;
  const entries = view.getUint16(ifd, little);
  let width = null;
  let height = null;
  for (let i = 0; i < entries; i++) {
    const at = ifd + 2 + i * 12;
    if (at + 12 > bytes.length) return null;
    const tag = view.getUint16(at, little);
    const type = view.getUint16(at + 2, little);
    const value = type === 3 ? view.getUint16(at + 8, little) : view.getUint32(at + 8, little);
    if (tag === 256) width = value;
    if (tag === 257) height = value;
  }
  return width && height ? { width, height } : null;
}

// Identify an image by its first bytes; PowerPoint file extensions can't be trusted.
export function sniffImage(bytes) {
  const b = bytes;
  if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4e && b[3] === 0x47) return "image/png";
  if (b[0] === 0xff && b[1] === 0xd8 && b[2] === 0xff) return "image/jpeg";
  if (b[0] === 0x47 && b[1] === 0x49 && b[2] === 0x46) return "image/gif";
  if (
    (b[0] === 0x49 && b[1] === 0x49 && b[2] === 0x2a) ||
    (b[0] === 0x4d && b[1] === 0x4d && b[3] === 0x2a)
  )
    return "image/tiff";
  if (b[0] === 0x42 && b[1] === 0x4d) return "image/bmp";
  if (b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) return "image/webp";
  if (b[0] === 0x3c) return "image/svg+xml";
  return "application/octet-stream";
}

// Position and size of a shape: { x, y, w, h } in EMUs, or null.
const ownRect = (xml) => {
  const m = xml.match(
    /<p:spPr>\s*<a:xfrm\b[^>]*>\s*<a:off x=["'](-?\d+)["'] y=["'](-?\d+)["']\s*\/>\s*<a:ext cx=["'](\d+)["'] cy=["'](\d+)["']/,
  );
  return m ? { x: +m[1], y: +m[2], w: +m[3], h: +m[4] } : null;
};

// Rect of a placeholder in a layout or master, matched by idx, else by type.
function placeholderRect(xml, ph) {
  if (!xml || !ph) return null;
  const idx = attr(ph, "idx");
  const type = attr(ph, "type");
  for (const sp of xml.match(/<p:sp>[\s\S]*?<\/p:sp>/g) ?? []) {
    const other = sp.match(/<p:ph\b[^>]*>/)?.[0];
    if (!other) continue;
    const match = idx ? attr(other, "idx") === idx : type && attr(other, "type") === type;
    const rect = match && ownRect(sp);
    if (rect) return rect;
  }
  return null;
}

// A shape without its own position inherits it: slide, then layout, then master.
const inheritedRect = (xml, layoutXml, masterXml) => {
  const ph = xml.match(/<p:ph\b[^>]*>/)?.[0];
  return ownRect(xml) ?? placeholderRect(layoutXml, ph) ?? placeholderRect(masterXml, ph);
};

const overlaps = (a, b) =>
  a && b && a.x < b.x + b.w && b.x < a.x + a.w && a.y < b.y + b.h && b.y < a.y + a.h;

function readPicture(xml, index, rels, slidePath, slideArea, layoutXml, masterXml) {
  const cNvPr = xml.match(/<p:cNvPr\b[^>]*\/?>/)?.[0] ?? "";
  const rId = attr(xml.match(/<a:blip\b[^>]*>/)?.[0] ?? "", "r:embed");
  const target = rId ? rels[rId] : null;
  const rect = inheritedRect(xml, layoutXml, masterXml);
  const area = rect ? rect.w * rect.h : 0;
  return {
    index,
    name: attr(cNvPr, "name") ?? "",
    alt: attr(cNvPr, "descr") ?? "",
    decorative: xml.includes(DECORATIVE_URI),
    mediaPath: target && !/^https?:/.test(target) ? resolvePath(slidePath, target) : null,
    coverage: slideArea ? area / slideArea : 0,
    rect,
  };
}

// Limits that keep a malformed or crafted file from exhausting the tab's memory.
export const LIMITS = {
  inputBytes: 250 * 1024 * 1024, // the .pptx file itself
  // All parts once unzipped. Pictures barely compress, so a real deck unpacks to little more than
  // its file size. Saving holds a second copy, so this keeps peak memory well under 1 GB.
  expandedBytes: 500 * 1024 * 1024,
  entries: 20000, // parts inside the zip
};

export class DeckTooLargeError extends Error {}

const mb = (n) =>
  n >= 1024 ** 3 ? `${+(n / 1024 ** 3).toFixed(1)} GB` : `${Math.round(n / (1024 * 1024))} MB`;

// fflate reads each part's declared expanded size before inflating it and allocates
// exactly that much, so capping the declared sizes caps memory.
function unzipWithinLimits(bytes, limits) {
  if (bytes.length > limits.inputBytes)
    throw new DeckTooLargeError(
      `This file is ${mb(bytes.length)}. The limit is ${mb(limits.inputBytes)}.`,
    );
  let total = 0;
  let count = 0;
  // Copy into an object with no prototype, so a part named "constructor" or "toString"
  // can't be confused with a built-in JavaScript property.
  const parts = unzipSync(bytes, {
    filter: ({ originalSize }) => {
      count += 1;
      total += originalSize;
      if (count > limits.entries)
        throw new DeckTooLargeError(`This file has more than ${limits.entries} parts inside it.`);
      if (total > limits.expandedBytes)
        throw new DeckTooLargeError(
          `This file unpacks to more than ${mb(limits.expandedBytes)}, which is more than this page can handle.`,
        );
      return true;
    },
  });
  return Object.assign(Object.create(null), parts);
}

const hasPart = (files, path) => typeof path === "string" && Object.hasOwn(files, path);

export function loadDeck(bytes, limits = LIMITS) {
  const files = unzipWithinLimits(bytes, limits);
  const read = (p) => (hasPart(files, p) ? strFromU8(files[p]) : null);

  const presentation = read("ppt/presentation.xml");
  if (!presentation)
    throw new Error("This file doesn't look like a PowerPoint deck (no ppt/presentation.xml).");
  const size = presentation.match(/<p:sldSz cx=["'](\d+)["'] cy=["'](\d+)["']/);
  const slideWidth = size ? Number(size[1]) : 12192000;
  const slideHeight = size ? Number(size[2]) : 6858000;
  const presRels = parseRels(read("ppt/_rels/presentation.xml.rels"));

  const slides = [...presentation.matchAll(/<p:sldId\b[^>]*>/g)].map((m, i) => {
    const path = resolvePath("ppt/presentation.xml", presRels[attr(m[0], "r:id")]);
    const xml = read(path);
    const rels = parseRels(read(path.replace("slides/", "slides/_rels/") + ".rels"));
    const titleSp = xml.match(TITLE_SP)?.[0] ?? null;
    const title = titleSp ? textOf(titleSp).join(" ").trim() : "";
    const titleHidden = !!titleSp && /<p:cNvPr\b[^>]*\shidden=["'](?:1|true)["']/.test(titleSp);
    const layoutPath = Object.values(rels).find((t) => /slideLayout/.test(t));
    const layoutFull = layoutPath ? resolvePath(path, layoutPath) : null;
    const layoutXml = layoutFull ? read(layoutFull) : null;
    const layoutRels = layoutFull
      ? parseRels(read(layoutFull.replace("slideLayouts/", "slideLayouts/_rels/") + ".rels"))
      : {};
    const masterPath = Object.values(layoutRels).find((t) => /slideMaster/.test(t));
    const masterXml = masterPath ? read(resolvePath(layoutFull, masterPath)) : null;
    const pictures = [...xml.matchAll(PIC)].map((p, j) =>
      readPicture(p[0], j, rels, path, slideWidth * slideHeight, layoutXml, masterXml),
    );
    const titleRect = inheritedRect(
      titleSp ?? '<p:sp><p:ph type="title"/></p:sp>',
      layoutXml,
      masterXml,
    );
    // A picture used as the slide's own background. It can't carry alt text.
    const bgBlip = xml.match(/<p:bg>[\s\S]*?(<a:blip\b[^>]*>)[\s\S]*?<\/p:bg>/)?.[1];
    const bgRId = bgBlip ? attr(bgBlip, "r:embed") : null;
    const background = bgRId && rels[bgRId] ? { mediaPath: resolvePath(path, rels[bgRId]) } : null;
    const bodyText = textOf(xml.replace(TITLE_SP, "").replace(PIC, ""));
    const hasNotes = Object.values(rels).some((t) => /notesSlide/.test(t));
    const notesPath = hasNotes
      ? resolvePath(
          path,
          Object.values(rels).find((t) => /notesSlide/.test(t)),
        )
      : null;
    const notes =
      notesPath && read(notesPath)
        ? textOf(read(notesPath)).filter((l) => !/^\d+$/.test(l.trim()))
        : [];
    return {
      number: i + 1,
      path,
      title,
      titleHidden,
      hasTitleBox: !!titleSp,
      pictures,
      bodyText,
      notes,
      background,
      // Titles on these slides go off the slide: Ally flags title text that shares space with a picture.
      imageSlide:
        !!background ||
        pictures.some((p) => p.coverage >= IMAGE_SLIDE_COVERAGE || overlaps(p.rect, titleRect)),
      otherVisuals:
        (xml.match(/<p:graphicFrame>/g) ?? []).length +
        (xml.match(/<p:sp>(?:(?!<\/p:sp>)[\s\S])*?<a:blip\b/g) ?? []).length,
    };
  });

  const core = read("docProps/core.xml") ?? "";
  // The title element may carry attributes, such as xml:lang.
  const docTitle = decodeXml(core.match(/<dc:title\b[^>]*>([^<]*)<\/dc:title>/)?.[1] ?? "");

  return { files, slides, slideWidth, slideHeight, docTitle };
}

export const needsTitle = (slide) => !slide.title || slide.titleHidden;
// PowerPoint's automatic alt text ends with a marker line in the language Office was set to.
// To support another language, add its marker here, copied exactly from a real deck,
// and add its code to AUTO_ALT_LANGUAGES.
export const AUTO_ALT_MARKERS = {
  en: /\bDescription automatically generated(?: with (?:very high|high|medium|low) confidence)?\.?\s*$/i,
};
export const AUTO_ALT_LANGUAGES = ["en"];

export const isAutoAlt = (text, languages = AUTO_ALT_LANGUAGES) =>
  languages.some((lang) => AUTO_ALT_MARKERS[lang]?.test(text));

// Alt text is missing when it's empty or is still PowerPoint's untouched guess.
export const altMissing = (alt, decorative) => !decorative && (!alt.trim() || isAutoAlt(alt));
export const needsAlt = (pic) => altMissing(pic.alt, pic.decorative);
export const slideNeedsWork = (slide) => needsTitle(slide) || slide.pictures.some(needsAlt);

export function summarize(deck) {
  const pics = deck.slides.flatMap((s) => s.pictures);
  return {
    slides: deck.slides.length,
    untitled: deck.slides.filter(needsTitle).length,
    pictures: pics.length,
    missingAlt: pics.filter(needsAlt).length,
    needWork: deck.slides.filter(slideNeedsWork).length,
  };
}

export function mediaBytes(deck, mediaPath) {
  return hasPart(deck.files, mediaPath) ? deck.files[mediaPath] : null;
}

// Position for a title that must not show: just above the top edge of the slide.
function offSlideXfrm(deck) {
  const height = Math.round(deck.slideHeight * 0.13);
  const width = Math.round(deck.slideWidth * 0.86);
  const x = Math.round((deck.slideWidth - width) / 2);
  return `<a:xfrm><a:off x="${x}" y="${-height - Math.round(deck.slideHeight * 0.02)}"/><a:ext cx="${width}" cy="${height}"/></a:xfrm>`;
}

function setTitle(xml, text, offSlide, deck) {
  const tx = `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${encodeText(text)}</a:t></a:r></a:p></p:txBody>`;
  const spPr = offSlide ? `<p:spPr>${offSlideXfrm(deck)}</p:spPr>` : null;
  const m = xml.match(TITLE_SP);
  if (m) {
    // Replacer functions throughout: a replacement string would treat $& or $' in typed text as patterns.
    let sp = m[0].replace(/(<p:cNvPr\b[^>]*?)\shidden=["'](?:1|true)["']/, "$1");
    sp = sp.replace(/<p:txBody>[\s\S]*<\/p:txBody>/, () => tx);
    if (spPr) sp = sp.replace(/<p:spPr\/>|<p:spPr>[\s\S]*?<\/p:spPr>/, () => spPr);
    return xml.replace(m[0], () => sp);
  }
  const ids = [...xml.matchAll(/<p:cNvPr\b[^>]*>/g)].map((x) => Number(attr(x[0], "id")) || 0);
  const sp = `<p:sp><p:nvSpPr><p:cNvPr id="${Math.max(0, ...ids) + 1}" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>${spPr ?? "<p:spPr/>"}${tx}</p:sp>`;
  // Insert first so the title is also first in reading order.
  return xml.replace("</p:grpSpPr>", () => "</p:grpSpPr>" + sp);
}

function setPicture(picXml, { alt, decorative }) {
  return picXml.replace(
    /<p:cNvPr\b([^>]*?)(\/>|>([\s\S]*?)<\/p:cNvPr>)/,
    (whole, attrs, _end, inner = "") => {
      let a = attrs.replace(/\sdescr=(?:"[^"]*"|'[^']*')/g, "");
      const text = cleanXmlText(alt).trim();
      if (!decorative && text) a += ` descr="${encodeAttr(text)}"`;
      let body = inner.replace(
        new RegExp(`<a:ext uri="${DECORATIVE_URI.replace(/[{}]/g, "\\$&")}">[\\s\\S]*?<\\/a:ext>`),
        "",
      );
      if (decorative) {
        const ext = `<a:ext uri="${DECORATIVE_URI}"><adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="1"/></a:ext>`;
        body = body.includes("<a:extLst>")
          ? body.replace("<a:extLst>", () => `<a:extLst>${ext}`)
          : `<a:extLst>${ext}</a:extLst>${body}`;
      }
      body = body.replace("<a:extLst></a:extLst>", "");
      return body ? `<p:cNvPr${a}>${body}</p:cNvPr>` : `<p:cNvPr${a}/>`;
    },
  );
}

// edits: { slides: { [number]: { title?, pictures?: { [index]: { alt, decorative } } } }, docTitle? }
export function applyEdits(deck, edits) {
  const out = { ...deck.files };
  for (const slide of deck.slides) {
    const e = edits.slides?.[slide.number];
    if (!e) continue;
    let xml = strFromU8(deck.files[slide.path]);
    // Clean before trimming, so text made only of invalid characters counts as empty.
    const title = typeof e.title === "string" ? cleanXmlText(e.title).trim() : "";
    const titleChanged = title && (title !== slide.title || slide.titleHidden);
    if (titleChanged) xml = setTitle(xml, title, slide.imageSlide, deck);
    if (e.pictures) {
      let i = 0;
      xml = xml.replace(PIC, (pic) => {
        const p = e.pictures[i++];
        return p ? setPicture(pic, p) : pic;
      });
    }
    out[slide.path] = strToU8(xml);
  }
  const docTitle = typeof edits.docTitle === "string" ? cleanXmlText(edits.docTitle).trim() : "";
  if (docTitle && out["docProps/core.xml"]) {
    const core = strFromU8(out["docProps/core.xml"]);
    const text = encodeText(docTitle);
    // Keep any attributes on an existing title element, and never add a second one.
    const existing = /<dc:title\b([^>]*?)(?:\/>|>[^<]*<\/dc:title>)/;
    out["docProps/core.xml"] = strToU8(
      existing.test(core)
        ? core.replace(existing, (_, attrs) => `<dc:title${attrs}>${text}</dc:title>`)
        : core.replace(
            /(<cp:coreProperties\b[^>]*>)/,
            (open) => `${open}<dc:title>${text}</dc:title>`,
          ),
    );
  }
  // [Content_Types].xml goes first, matching how PowerPoint writes the zip.
  const ordered = { "[Content_Types].xml": out["[Content_Types].xml"], ...out };
  return zipSync(ordered, { level: 6 });
}
