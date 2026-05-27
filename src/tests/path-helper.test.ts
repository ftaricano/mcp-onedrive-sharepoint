import test from "node:test";
import assert from "node:assert/strict";

import { sanitizeFileName, analyzePath } from "../tools/utils/path-helper.js";

test("sanitizeFileName preserves spaces in folder names (regression #12)", () => {
  assert.equal(sanitizeFileName("Marco Costa"), "Marco Costa");
});

test("sanitizeFileName preserves spaces in file names", () => {
  assert.equal(
    sanitizeFileName("Recibo de Pagamento 042026.pdf"),
    "Recibo de Pagamento 042026.pdf",
  );
});

test("sanitizeFileName still replaces invalid characters with underscores", () => {
  assert.equal(sanitizeFileName("invalid<chars>here.pdf"), "invalid_chars_here.pdf");
});

test("sanitizeFileName still prepends file_ to reserved Windows names", () => {
  assert.equal(sanitizeFileName("CON"), "file_CON");
});

test("sanitizeFileName collapses internal whitespace and trims edges", () => {
  // Current pipeline: invalid-char pass → /\s+/g → " " (collapse) → trim → strip leading/trailing dots.
  // No further space-to-underscore step after the fix, so the single space before ".pdf" survives.
  assert.equal(
    sanitizeFileName("  multiple   spaces  .pdf"),
    "multiple spaces .pdf",
  );
});

test("analyzePath preserves spaces across folder + file segments", () => {
  const result = analyzePath("/folder with spaces/file with spaces.pdf");
  assert.equal(result.folderPath, "folder with spaces");
  assert.equal(result.fileName, "file with spaces.pdf");
  assert.equal(
    result.sanitizedPath,
    "folder with spaces/file with spaces.pdf",
  );
});
