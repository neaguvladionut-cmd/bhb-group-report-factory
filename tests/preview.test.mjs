import assert from "node:assert/strict";
import { access, readFile } from "node:fs/promises";
import { dirname, resolve } from "node:path";
import { fileURLToPath } from "node:url";
import test from "node:test";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
test("PDF-first contract has one full-width review surface and no HTML slide preview", async () => {
  const index = await readFile(resolve(root, "src/index.html"), "utf8"), css = await readFile(resolve(root, "src/styles.css"), "utf8"), app = await readFile(resolve(root, "src/app.js"), "utf8");
  assert.doesNotMatch(index, /id="preview"|report-preview|stage-summary/u);
  assert.match(index, /class="workspace pdf-first-workspace"/u);
  assert.match(css, /\.pdf-preview-panel\{/u);
  assert.doesNotMatch(app, /mountPreview|renderSlidePreview|renderResultsPreview|renderReportOutline/u);
});
test("PDF preview controls and vendored offline viewer assets are present", async () => {
  const index = await readFile(resolve(root, "src/index.html"), "utf8"), build = await readFile(resolve(root, "tools/build.mjs"), "utf8");
  assert.match(index, /id="pdf-preview-panel" class="pdf-preview-panel"/u); for (const id of ["pdf-canvas", "refresh-pdf-preview", "pdf-prev", "pdf-next", "pdf-open-fallback", "pdf"]) assert.match(index, new RegExp(`id="${id}"`, "u"));
  assert.match(index, /pdf-lib\.min\.js/u); assert.match(build, /pdf\.js/u); assert.doesNotMatch(build, /wrap\(await code\("preview\.js"/u);
  await access(resolve(root, "src/assets/vendor/pdf.min.mjs")); await access(resolve(root, "src/assets/vendor/pdf.worker.min.mjs"));
});
test("all section toggles preserve reportPlan as the single PDF inventory", async () => {
  const app = await readFile(resolve(root, "src/app.js"), "utf8"), pdf = await readFile(resolve(root, "src/pdf.js"), "utf8");
  assert.match(app, /reportPlan\(payload, \{ scope: "whole", ignoreSections: true \}\)/u);
  assert.match(pdf, /reportPlan\(payload, \{ scope \}\)/u);
  assert.match(app, /ui\.sections\[section\.key\] = !isOn/u);
});
