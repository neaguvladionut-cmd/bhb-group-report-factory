import assert from "node:assert/strict";
import { access, readFile } from "node:fs/promises";
import { execFile } from "node:child_process";
import { promisify } from "node:util";
import { resolve } from "node:path";
import test from "node:test";
import { FIXED_TEMPLATE_LABELS, fillFixedTemplateLabels } from "../src/template-pptx.js";

const run = promisify(execFile);
const root = resolve(new URL("..", import.meta.url).pathname);
const asset = resolve(root, "src/assets/trend/template-raport-de-grup-RO.pptx");

test("Trend route uses the cleaned native template, not a generated picture deck", async () => {
  await access(asset);
  const source = await readFile(resolve(root, "src/template-pptx.js"), "utf8");
  const wrapper = await readFile(resolve(root, "src/pptx.js"), "utf8");
  assert.doesNotMatch(source, /PptxGenJS/u);
  assert.doesNotMatch(wrapper, /PptxGenJS/u);
  assert.match(source, /ppt\/charts/u);
  assert.match(source, /ppt\/embeddings/u);
  const { stdout } = await run("unzip", ["-l", asset]);
  assert.match(stdout, /ppt\/slideMasters\/slideMaster1\.xml/u);
  assert.match(stdout, /ppt\/slideLayouts\/slideLayout1\.xml/u);
  assert.match(stdout, /ppt\/charts\/chart1\.xml/u);
  assert.match(stdout, /ppt\/embeddings\/Microsoft_Excel_Worksheet\.xlsx/u);
  assert.doesNotMatch(stdout, /ppt\/notesSlides\//u);
  assert.doesNotMatch(stdout, /customXml\//u);
  assert.doesNotMatch(stdout, /ppt\/changesInfos\//u);
  assert.doesNotMatch(stdout, /ppt\/revisionInfo\.xml/u);
  assert.doesNotMatch(stdout, /ppt\/fonts\//u);
  const presentation = (await run("unzip", ["-p", asset, "ppt/presentation.xml"])).stdout;
  assert.doesNotMatch(presentation, /notesMasterIdLst/u);
});

test("clean-template declares the public-package scrub and local-only source boundary", async () => {
  const script = await readFile(resolve(root, "tools/clean-template.mjs"), "utf8");
  assert.match(script, /notesSlides/u);
  assert.match(script, /customXml/u);
  assert.match(script, /changesInfos/u);
  assert.match(script, /dc:title/u);
  assert.match(script, /selfCheck/u);
  assert.match(script, /embeddedFontLst/u);
  assert.match(script, /ppt\/fonts/u);
  assert.match(script, /git-ignored/iu);
});

test("Trend generator re-inserts every fixed template label into generated slide XML", () => {
  const generated = fillFixedTemplateLabels("<p:sld><p:spTree></p:spTree></p:sld>", FIXED_TEMPLATE_LABELS);
  for (const label of FIXED_TEMPLATE_LABELS) assert.match(generated, new RegExp(label.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&"), "u"));
});
