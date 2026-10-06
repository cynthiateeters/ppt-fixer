# ppt-fixer

A proof of concept: a web page that fixes two problems Ally flags in image-heavy PowerPoint decks, missing slide titles and missing alt text.

You open a .pptx, the page shows every slide on one scrolling page, you type titles and picture descriptions where they're missing, and it gives you a fixed copy to upload to Canvas. The file never leaves your computer.

## Beginner's guide

### What it does

- **Slide titles.** Every untitled slide gets a box to type a title. On slides where a picture covers the slide or overlaps the title area, the title is placed just above the slide. It doesn't show when presenting, but screen readers and Ally still find it.
- **Alt text.** Every picture gets a preview and a text box. You can mark a picture as decorative instead. When the same picture appears on several slides, one description can fill them all.
- **Every slide on one page.** Each slide has its own section with its text, title box and picture boxes, one after another. There's no Next button. The list beside it jumps to a slide and highlights the one at the top of the screen. Typing never rebuilds the page, so the boxes stay where they are. That matters for a browser assistant like Gemini in Chrome filling in the whole deck, which got stuck on the Next button when the page showed one slide at a time. Each click replaced the whole slide and moved the button, sometimes below the bottom of the screen.
- **Slide status.** Each slide in the list says what it's missing ("Needs title", "Needs alt text"), "Filled in" once you've added it, or "Nothing missing". It never says a slide is fine, because the page checks titles and alt text only, and it says so above the list. "Only slides missing a title or alt text" hides the rest. It's off by default, because a slide that needs a description often depends on the ones around it, such as a section slide or a caption.
- **PowerPoint's guesses.** Older PowerPoint saved automatic alt text ending in "Description automatically generated". The page treats that as missing and flags it "PowerPoint guessed this" until you rewrite it or remove that line. Text you don't touch is saved unchanged. Only English is recognized for now. Other languages can be added in `AUTO_ALT_MARKERS` in `src/pptx.js`, using the exact line from a real deck.
- **File properties.** The document title can be corrected. Ally copies it into the PDF and HTML versions students download.
- **Download.** The fixed copy is saved as `<name> (fixed).pptx`. The original isn't changed.
- **Saved work.** What you type is saved in the browser's localStorage as you go, keyed by a SHA-256 fingerprint of the file.
  - Opening the exact same file again picks up where you left off, and "Start over from the file" discards that.
  - A different version of the deck won't match, so it starts fresh.
  - Saved work older than 30 days is removed when the page loads.
  - If other people use the same computer login, "Delete saved work for this file" removes it and stops saving until the file is reopened.
  - If the browser blocks storage, the page says so and warns before the tab closes.
- **Safe text.** Characters XML doesn't allow, which usually arrive by pasting from a PDF, are removed before they're written, so they can't corrupt the file.
- **Help.** The header's Help link walks through fixing a file, writing titles and alt text, drafting them with a browser AI assistant (with a copyable example prompt for art slides), what the slide list labels mean, and common problems. Help is a section of `index.html`, shown in place of the fixer, so an open deck stays open while it's read. The browser's Back button and "Back to the fixer" both return to the same spot. When a label, limit or message in `src/main.js` changes, update the help text too.
- **Size limits.** Files over 250 MB, files that unpack to more than 500 MB, and files with more than 20,000 parts are refused with a plain message, so a broken or crafted file can't freeze the tab. TIFF previews over 50 million pixels are skipped.

### What it doesn't do

- Charts, tables, SmartArt and picture-filled shapes. The page points them out, and you fix them in PowerPoint.
- Pictures used as slide backgrounds can't hold alt text. The page says so on those slides.
- Text contrast, reading order and anything else Ally checks.

### Run it on your computer

You need [Node.js](https://nodejs.org/) and [pnpm](https://pnpm.io/).

```bash
pnpm install
pnpm dev
```

Then open the address it prints, usually `http://localhost:5173`.

To build the static site into `dist/`:

```bash
pnpm build
pnpm preview
```

### Tests

```bash
pnpm test
```

Tests that build their own inputs always run. Tests that need a real deck read it from `PPTX_SAMPLE`, because course decks aren't committed to this repo. Put the path in a `.env.test` file, which is gitignored:

```bash
PPTX_SAMPLE=/path/to/reference-deck.pptx
```

Without it, those tests skip. Their assertions expect the reference deck used during development, described in `test/sample.js`.

To check a whole folder of decks, filling every gap with test text and confirming the result reloads cleanly:

```bash
node scripts/check-folder.js <folder-of-pptx-files>
```

## Why titles go above the slide

This rule comes from sandbox tests in Ally on 2026-09-16, starting with one image-heavy 25-slide deck:

- A title typed into the slide's own title box on a full-picture slide was flagged for insufficient contrast. That happened even when the picture covered the title.
- The same titles moved just above the slide cleared the contrast flag.
- With titles above the slide and alt text on every picture, the deck scored 100%.
- A second deck from a different course, fixed by hand in the live page, also scored 100%. It had 18 slides and 31 pictures, 2 of them marked decorative, with 7 titles placed above the slide.
- In Ally's HTML and tagged PDF versions, those titles came through as headings and the alt text came through word for word.
- A fixed deck opened in PowerPoint for Mac with no repair prompt (2026-09-17).

## Accessibility of the page itself

Checked 2026-09-22 in Chrome with a practice deck:

- Lighthouse accessibility scores 100 on the start page and the editor.
- At 320 pixels wide, the width of a 1280-pixel window at 400% zoom, the page reflows with no sideways scrolling.
- All text is sized in rem and nothing is smaller than the browser's base size, so browser zoom and the browser's font size setting enlarge everything.
- Every picture's fields sit in a group named for that picture, and the labels include its number, so a screen reader can tell pictures apart.
- The "missing" and "needs alt text" flags disappear once fixed, and the status line only updates when its counts change.
- Every visible control shows a focus outline.
- A description box that's switched off, because its picture is marked decorative, looks switched off and says why right beneath it.

Rechecked 2026-09-29 after the move to one page, with the same practice deck:

- Lighthouse accessibility still scores 100 on the editor.
- At 320 pixels wide there's still no sideways scrolling. This needed a fix: a URL in one slide's text has no spaces to wrap at, and it widened every slide. Slide text now wraps anywhere.

Not yet tried with a real screen reader.

## Not tested yet

- **Decks from other departments.** It's only been run against decks from Arts and Design courses.

## How it works

- `src/pptx.js` reads the deck with [fflate](https://github.com/101arrowz/fflate), finds titles and pictures in each slide's XML, and writes the edits back. It has no DOM code, so the same module runs in the tests.
- `src/preview.js` shows pictures. It checks each file's first bytes, because PowerPoint's file extensions can't be trusted. PNG, JPEG and GIF display natively. Previews are decoded only as they scroll near the screen, so a long deck doesn't decode every picture when it opens.
- `src/tiff-worker.js` decodes TIFF previews with [image-in-browser](https://github.com/yegor-pelykh/image-in-browser) in a worker, so the page doesn't freeze. It only loads when a deck has TIFFs.
- `src/saved-work.js` saves and restores typed edits in localStorage. Every storage call is guarded, because storage can be blocked or full.
- `src/main.js` is the interface, built with plain DOM calls.
Dependencies are pinned to exact versions.

## Contributing

This is a proof of concept for the Arts and Design digital accessibility work. Issues and suggestions are welcome. Before opening a pull request:

1. Run `pnpm test` with a sample deck.
2. Run `node scripts/check-folder.js` on a folder of real decks.
3. Upload one fixed deck to a sandbox course and check its Ally score before claiming a fix works.

## License

MIT License. Copyright (c) 2026 Cynthia Teeters. See [LICENSE](LICENSE).
