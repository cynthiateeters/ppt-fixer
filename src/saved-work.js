// Keeps typed titles and alt text in this browser's localStorage, keyed by a fingerprint
// of the exact file, so work survives a closed tab. Every storage call can throw
// (private windows, blocked site data, full quota), so all of them are guarded.

const PREFIX = "ppt-fixer:v1:";
const MAX_AGE_MS = 30 * 24 * 60 * 60 * 1000;
const SAVE_DELAY_MS = 500;

export async function fingerprint(bytes) {
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  return [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, "0")).join("");
}

function storage() {
  try {
    const s = window.localStorage;
    const probe = `${PREFIX}probe`;
    s.setItem(probe, "1");
    s.removeItem(probe);
    return s;
  } catch {
    return null;
  }
}

export const storageAvailable = () => storage() !== null;

// Saved edits only fit a deck with the same slides and picture counts.
function fits(edits, deck) {
  return (
    edits &&
    Array.isArray(edits.slides) &&
    edits.slides.length === deck.slides.length &&
    edits.slides.every(
      (s, i) =>
        typeof s?.title === "string" &&
        Array.isArray(s.pictures) &&
        s.pictures.length === deck.slides[i].pictures.length &&
        s.pictures.every((p) => typeof p?.alt === "string" && typeof p?.decorative === "boolean"),
    )
  );
}

export function loadSaved(id, deck) {
  const s = storage();
  if (!s) return null;
  try {
    const record = JSON.parse(s.getItem(PREFIX + id) ?? "null");
    if (!record || !fits(record.edits, deck)) return null;
    return record;
  } catch {
    return null;
  }
}

let timer = null;

// Returns false when the browser won't store anything.
export function scheduleSave(id, fileName, edits, onSaved) {
  const s = storage();
  if (!s) return false;
  clearTimeout(timer);
  timer = setTimeout(() => {
    try {
      s.setItem(PREFIX + id, JSON.stringify({ fileName, savedAt: Date.now(), edits }));
      onSaved?.(true);
    } catch {
      onSaved?.(false);
    }
  }, SAVE_DELAY_MS);
  return true;
}

// Immediate save, for closing the tab or the deck. Returns true on success.
export function saveNow(id, fileName, edits) {
  cancelPendingSave();
  try {
    const s = storage();
    if (!s) return false;
    s.setItem(PREFIX + id, JSON.stringify({ fileName, savedAt: Date.now(), edits }));
    return true;
  } catch {
    return false;
  }
}

export function cancelPendingSave() {
  clearTimeout(timer);
  timer = null;
}

export function clearSaved(id) {
  cancelPendingSave();
  try {
    storage()?.removeItem(PREFIX + id);
  } catch {
    // Nothing to clear.
  }
}

export function pruneOld(now = Date.now()) {
  const s = storage();
  if (!s) return;
  try {
    for (let i = s.length - 1; i >= 0; i--) {
      const key = s.key(i);
      if (!key?.startsWith(PREFIX)) continue;
      try {
        const record = JSON.parse(s.getItem(key));
        if (!record?.savedAt || now - record.savedAt > MAX_AGE_MS) s.removeItem(key);
      } catch {
        s.removeItem(key);
      }
    }
  } catch {
    // Storage went away mid-loop; try again next time.
  }
}
