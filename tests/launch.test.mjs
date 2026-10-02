import assert from "node:assert/strict";
import { access, readFile } from "node:fs/promises";
import { resolve } from "node:path";
import test from "node:test";

const root = resolve(new URL("..", import.meta.url).pathname);

test("deploy is a classic offline entry with the Trend rebuild controls", async () => {
  const source = await readFile(resolve(root, "src/rebuild-index.html"), "utf8");
  const deploy = await readFile(resolve(root, "deploy/index.html"), "utf8");
  const app = await readFile(resolve(root, "deploy/app.js"), "utf8");
  const css = await readFile(resolve(root, "deploy/styles.css"), "utf8");
  assert.equal(deploy, await readFile(resolve(root, "src/rebuild-index.html"), "utf8"));
  assert.match(deploy, /id="split-groups"/u);
  assert.match(deploy, /id="annex-setting"/u);
  assert.match(deploy, /evaluation-sheet-template\.csv/u);
  assert.match(deploy, /id="bundle"/u);
  assert.match(deploy, /id="download-fallback"/u);
  assert.match(app, /downloadBundle/u);
  assert.match(app, /downloadPptx/u);
  assert.doesNotMatch(app, /PptxGenJS/u);
  assert.match(css, /\.workspace\{/u);
  assert.match(source, /Trend este singurul PPTX/u);
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
