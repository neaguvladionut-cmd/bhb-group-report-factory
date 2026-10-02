import assert from "node:assert/strict";
import { access, readFile } from "node:fs/promises";
import { resolve } from "node:path";
import vm from "node:vm";
import test from "node:test";

const root = resolve(new URL("..", import.meta.url).pathname);

test("deploy is the source-owned one-page Romanian builder with the approved four-step spine", async () => {
  const source = await readFile(resolve(root, "src/rebuild-index.html"), "utf8");
  const deploy = await readFile(resolve(root, "deploy/index.html"), "utf8");
  const app = await readFile(resolve(root, "deploy/app.js"), "utf8");
  const css = await readFile(resolve(root, "deploy/styles.css"), "utf8");
  assert.equal(deploy, await readFile(resolve(root, "src/rebuild-index.html"), "utf8"));
  assert.match(deploy, /id="split-groups"/u);
  assert.match(deploy, /name="annex"/u);
  assert.match(deploy, /evaluation-sheet-template\.csv/u);
  assert.match(deploy, /id="bundle"/u);
  assert.match(deploy, /id="download-fallback"/u);
  assert.equal((deploy.match(/class="workflow-section/g) || []).length, 4);
  assert.ok((deploy.match(/class="field-context"/gu) || []).length >= 6);
  assert.match(deploy, /Ce faci aici/u);
  assert.match(deploy, /Ce urmează/u);
  assert.match(deploy, /Apare în/u);
  assert.match(deploy, /Pachet BHB \(ZIP\)/u);
  const visibleCopy = deploy.replace(/<[^>]+>/gu, " ");
  assert.doesNotMatch(visibleCopy, /\b(?:schema|upload|API)\b/iu);
  assert.match(app, /downloadBundle/u);
  assert.match(app, /downloadPptx/u);
  assert.match(app, /groupedIssues\(blockers, "blocker"\)/u);
  assert.match(app, /data-issue-search/u);
  assert.match(app, /Confirmă toate avertismentele/u);
  assert.doesNotMatch(app, /PptxGenJS/u);
  assert.match(css, /\.issue-panel\{/u);
  assert.match(css, /\.issue-toggle/u);
  assert.doesNotMatch(css, /\.issue-items\{[^}]*overflow/u);
  assert.match(css, /\.workflow-rail\{/u);
  assert.match(source, /id="step-upload"/u);
  assert.match(source, /id="step-download"/u);
  const scripts = [...deploy.matchAll(/<script(?:(?:\s+src="([^"]+)")?)>([\s\S]*?)<\/script>/gu)];
  assert.equal(scripts.length, 5);
  for (const [, src] of scripts) if (src) await access(resolve(root, "deploy", src.split("?")[0]));
});

test("source-owned build emits the new normalized payload and package-preserving route", async () => {
  const core = await readFile(resolve(root, "src/rebuild-core.js"), "utf8");
  const plan = await readFile(resolve(root, "src/rebuild-report-plan.js"), "utf8");
  const template = await readFile(resolve(root, "src/template-pptx.js"), "utf8");
  assert.match(core, /CODE/u);
  assert.match(core, /region-missing/u);
  assert.match(core, /objective_text_score_0/u);
  assert.match(plan, /rankBehaviors/u);
  assert.match(plan, /splitGroups/u);
  assert.match(template, /ppt\/embeddings/u);
  assert.match(template, /generateBundle/u);
  assert.doesNotMatch(template, /PptxGenJS/u);
});

test("built deploy bundle initializes without throwing in a DOM-like vm", async () => {
  const bundle = await readFile(resolve(root, "deploy/app.js"), "utf8");
  const document = { body: { dataset: {} }, querySelector: () => null, querySelectorAll: () => [] };
  const window = { XLSX: {}, JSZip: {}, __grfBooted() {} };
  assert.doesNotThrow(() => vm.runInNewContext(bundle, { Blob, Intl, Map, Set, URL, document, window }));
  assert.equal(typeof window.__grfDownload, "function");
});
