# ppt-fixer

A proof of concept: a web page that fixes two problems Ally flags in image-heavy PowerPoint decks, missing slide titles and missing alt text.

You open a .pptx, the page walks you through each slide that needs a title or a picture description, you type them in, and it gives you a fixed copy to upload to Canvas. The file never leaves your computer.

## Beginner's guide

### What it does

- **Slide titles.** Every untitled slide gets a box to type a title. On slides where a picture covers the slide or overlaps the title area, the title is placed just above the slide. It doesn't show when presenting, but screen readers and Ally still find it.
- **Alt text.** Every picture gets a preview and a text box. You can mark a picture as decorative instead. When the same picture appears on several slides, one description can fill them all.
- **File properties.** The document title can be corrected. Ally copies it into the PDF and HTML versions students download.
- **Download.** The fixed copy is saved as `<name> (fixed).pptx`. The original isn't changed.
- **Saved work.** What you type is saved in the browser's localStorage as you go, keyed by a SHA-256 fingerprint of the file.
  - Opening the exact same file again picks up where you left off, and "Start over from the file" discards that.
  - A different version of the deck won't match, so it starts fresh.
  - Saved work older than 30 days is removed when the page loads.
  - If other people use the same computer login, "Delete saved work for this file" removes it and stops saving until the file is reopened.
  - If the browser blocks storage, the page says so and warns before the tab closes.
- **Safe text.** Characters XML doesn't allow, which usually arrive by pasting from a PDF, are removed before they're written, so they can't corrupt the file.
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

## Not tested yet

- **Decks from other departments.** It's only been run against decks from Arts and Design courses.

## How it works

- `src/pptx.js` reads the deck with [fflate](https://github.com/101arrowz/fflate), finds titles and pictures in each slide's XML, and writes the edits back. It has no DOM code, so the same module runs in the tests.
- `src/preview.js` shows pictures. It checks each file's first bytes, because PowerPoint's file extensions can't be trusted. PNG, JPEG and GIF display natively.
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
