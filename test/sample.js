// The reference deck some tests read. Course decks aren't committed, so its path comes from
// PPTX_SAMPLE, usually set in a local .env.test file. Without it, those tests skip.
//
// The assertions expect the reference deck used during development: 25 slides, 11 without
// titles, 16 pictures without alt text, a slide with no title box (14), and a mix of real TIFF
// and PNG data saved with .tiff names.

import { existsSync } from "node:fs";

export const sample = process.env.PPTX_SAMPLE ?? "";

export const skip =
  sample && existsSync(sample)
    ? false
    : "set PPTX_SAMPLE to the reference deck (see README, Tests)";
